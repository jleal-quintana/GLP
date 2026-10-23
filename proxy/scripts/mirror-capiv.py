#!/usr/bin/env python3
"""Mirror the official Capítulo IV sources into the proxy's blob container.

datos.energia.gob.ar resets connections coming from Azure (and other cloud)
IP ranges, so the Function App can no longer download the CSVs itself. This
script runs on an on-prem server with an Argentine IP and publishes what the
proxy needs:

  catalog.json                     area catalog (same shape as ?catalog=1)
  <year>.csv                       raw annual production CSV
  filtered/<year>/<area>.ndjson    production rows per area
  filtered/<year>/_complete.json   {"scannedRows": n, "generatedAt": iso}

Usage:
  mirror-capiv.py --sas-file ~/.capiv-mirror-sas            # catalog + current and previous year
  mirror-capiv.py --sas-file ... --years 2006-2026           # backfill

The SAS file holds a container SAS URL with create/write permissions, e.g.
https://<account>.blob.core.windows.net/production-cache?sv=...&sig=...
"""

import argparse
import csv
import datetime as dt
import io
import json
import os
import re
import sys
import tempfile
import time
import unicodedata
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

DATASET_URL = 'https://datos.gob.ar/api/3/action/package_show?id=produccion-de-petroleo-y-gas-por-pozo'
USER_AGENT = 'Quintana-CapIV-Mirror/1.0'
UPLOAD_WORKERS = 8

csv.field_size_limit(sys.maxsize)


def log(message):
    print(f'{dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")} {message}', flush=True)


def normalize(value):
    text = unicodedata.normalize('NFD', str(value or ''))
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn').lower()
    return re.sub(r'[^a-z0-9]+', ' ', text).strip()


def text_value(row, *keys):
    for key in keys:
        value = row.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ''


def http_get(url, timeout=120):
    request = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
    return urllib.request.urlopen(request, timeout=timeout)


def newest(resources):
    return max(resources, key=lambda r: r.get('last_modified') or '', default=None)


def discover_resources():
    """Same selection rules as discoverCapivResources() in src/services/capiv.ts."""
    with http_get(DATASET_URL) as response:
        dataset = json.load(response)
    resources = [r for r in dataset['result']['resources'] if str(r.get('format', '')).upper() == 'CSV']
    wells = newest([r for r in resources if normalize(r.get('name')) == 'capitulo iv pozos'])
    if not wells:
        raise RuntimeError('no se encontró el recurso Capítulo IV - Pozos')
    by_year = {}
    for resource in resources:
        name = normalize(resource.get('name'))
        if 'produccion de pozos de gas y petroleo' not in name:
            continue
        if 'ddjj abiertas y cerradas' in name or 'no convencional' in name:
            continue
        match = re.search(r'\b(20\d{2}|19\d{2})\b', name)
        if match:
            by_year.setdefault(int(match.group(1)), []).append(resource)
    return wells, {year: newest(candidates) for year, candidates in by_year.items()}


def download(url, attempts=3):
    for attempt in range(1, attempts + 1):
        handle = tempfile.NamedTemporaryFile(prefix='capiv-', suffix='.csv', delete=False)
        try:
            with http_get(url, timeout=300) as response, handle:
                while chunk := response.read(1 << 20):
                    handle.write(chunk)
            if os.path.getsize(handle.name) == 0:
                raise RuntimeError('descarga vacía')
            return handle.name
        except Exception as error:
            os.unlink(handle.name)
            if attempt == attempts:
                raise
            log(f'  reintento {attempt} de {url}: {error}')
            time.sleep(10 * attempt)


def read_rows(path):
    with open(path, newline='', encoding='utf-8-sig', errors='replace') as handle:
        reader = csv.reader(handle)
        headers = [h.strip().lower() for h in next(reader)]
        for values in reader:
            if not any(values):
                continue
            yield dict(zip(headers, values))


class Container:
    def __init__(self, sas_url):
        parsed = urllib.parse.urlsplit(sas_url.strip())
        self.base = f'{parsed.scheme}://{parsed.netloc}{parsed.path.rstrip("/")}'
        self.query = parsed.query

    def put(self, name, data, content_type):
        url = f'{self.base}/{urllib.parse.quote(name, safe="/")}?{self.query}'
        is_file = isinstance(data, str)
        length = os.path.getsize(data) if is_file else len(data)
        headers = {
            'x-ms-blob-type': 'BlockBlob',
            'x-ms-version': '2021-08-06',
            'Content-Type': content_type,
            'Content-Length': str(length),
        }
        for attempt in range(1, 4):
            body = open(data, 'rb') if is_file else data
            try:
                request = urllib.request.Request(url, data=body, method='PUT', headers=headers)
                with urllib.request.urlopen(request, timeout=600) as response:
                    response.read()
                return
            except Exception as error:
                if attempt == 3:
                    raise RuntimeError(f'no se pudo subir {name}: {error}') from error
                time.sleep(5 * attempt)
            finally:
                if is_file:
                    body.close()


def mirror_catalog(container, wells):
    log(f'Catálogo: {wells["url"]}')
    path = download(wells['url'])
    try:
        by_area = {}
        for row in read_rows(path):
            area_id = text_value(row, 'cod_area', 'idareapermisoconcesion')
            area_name = text_value(row, 'area', 'areapermisoconcesion')
            if not area_id or not area_name:
                continue
            company = text_value(row, 'empresa')
            existing = by_area.get(area_id)
            if existing:
                if company and company not in existing['companies']:
                    existing['companies'].append(company)
            else:
                by_area[area_id] = {
                    'province': text_value(row, 'provincia') or 'Sin provincia',
                    'areaId': area_id,
                    'areaName': area_name,
                    'basin': text_value(row, 'cuenca'),
                    'companies': [company] if company else [],
                }
    finally:
        os.unlink(path)
    if len(by_area) < 50:
        raise RuntimeError(f'el catálogo tiene sólo {len(by_area)} áreas; no se publica')
    catalog = sorted(by_area.values(), key=lambda a: (normalize(a['province']), normalize(a['areaName'])))
    container.put('catalog.json', json.dumps(catalog, ensure_ascii=False).encode('utf-8'),
                  'application/json; charset=utf-8')
    log(f'Catálogo publicado: {len(catalog)} áreas')


def mirror_year(container, year, resource):
    log(f'{year}: {resource["url"]}')
    path = download(resource['url'])
    try:
        scanned = 0
        by_area = {}
        for row in read_rows(path):
            scanned += 1
            area_id = text_value(row, 'idareapermisoconcesion', 'cod_area')
            if not area_id:
                continue
            by_area.setdefault(area_id, io.StringIO()).write(json.dumps({
                'idareapermisoconcesion': area_id,
                'idpozo': text_value(row, 'idpozo'),
                'sigla': text_value(row, 'sigla'),
                'anio': text_value(row, 'anio'),
                'mes': text_value(row, 'mes'),
                'prod_pet': text_value(row, 'prod_pet', 'prod_petroleo', 'petroleo'),
                'prod_gas': text_value(row, 'prod_gas', 'gas'),
                'prod_agua': text_value(row, 'prod_agua', 'agua'),
                'iny_agua': text_value(row, 'iny_agua', 'agua_iny', 'inyeccion_agua'),
                'iny_gas': text_value(row, 'iny_gas', 'gas_iny', 'inyeccion_gas'),
            }, ensure_ascii=False) + '\n')
        if scanned == 0:
            raise RuntimeError(f'{year}: el CSV no tiene filas; no se publica')

        container.put(f'{year}.csv', path, 'text/csv; charset=utf-8')

        def upload(item):
            area_id, buffer = item
            # Match encodeURIComponent() in proxy.js so both write the same blob names.
            blob_area = urllib.parse.quote(area_id, safe="!'()*")
            container.put(f'filtered/{year}/{blob_area}.ndjson',
                          buffer.getvalue().encode('utf-8'), 'application/x-ndjson; charset=utf-8')

        with ThreadPoolExecutor(UPLOAD_WORKERS) as pool:
            list(pool.map(upload, by_area.items()))
        marker = {'scannedRows': scanned, 'generatedAt': dt.datetime.now(dt.timezone.utc).isoformat()}
        container.put(f'filtered/{year}/_complete.json', json.dumps(marker).encode('utf-8'), 'application/json')
        log(f'{year} publicado: {scanned} filas, {len(by_area)} áreas')
    finally:
        os.unlink(path)


def parse_years(spec):
    years = set()
    for part in spec.split(','):
        if '-' in part:
            start, end = part.split('-')
            years.update(range(int(start), int(end) + 1))
        elif part:
            years.add(int(part))
    return sorted(years)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--sas-file', required=True, help='archivo con la URL SAS del contenedor')
    parser.add_argument('--years', help='años a publicar, ej. 2025,2026 o 2006-2026 (por defecto: actual y anterior)')
    parser.add_argument('--skip-catalog', action='store_true')
    args = parser.parse_args()

    with open(os.path.expanduser(args.sas_file), encoding='utf-8') as handle:
        container = Container(handle.read())
    current = dt.datetime.now(dt.timezone.utc).year
    years = parse_years(args.years) if args.years else [current - 1, current]

    wells, by_year = discover_resources()
    failures = 0
    if not args.skip_catalog:
        try:
            mirror_catalog(container, wells)
        except Exception as error:
            failures += 1
            log(f'ERROR catálogo: {error}')
    for year in years:
        resource = by_year.get(year)
        if not resource:
            log(f'{year}: sin recurso publicado; se omite')
            continue
        try:
            mirror_year(container, year, resource)
        except Exception as error:
            failures += 1
            log(f'ERROR {year}: {error}')
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
