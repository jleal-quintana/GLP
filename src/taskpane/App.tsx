import React, { useMemo, useState } from 'react';
import { buildWorkbook } from '../excel/workbookBuilder';
import type { AreaCatalogItem, AreaSelection, AreaWorkbookPlan, ForecastDefaults } from '../models/types';
import { fetchAreaCatalog } from '../services/capiv';
import { appendDebug, createDebugEntry } from '../excel/debugSheet';

const DEFAULTS: ForecastDefaults = {
  startYear: 2015,
  horizonYears: 10,
  grossMethod: 'Constante',
  oilMethod: 'Declinación Exp.',
  gasMethod: 'RGP',
  takeInitialFromHistory: true,
};

export function App() {
  const [catalog, setCatalog] = useState<AreaCatalogItem[]>([]);
  const [province, setProvince] = useState('Todas');
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState<AreaSelection[]>([]);
  const [defaults, setDefaults] = useState<ForecastDefaults>(DEFAULTS);
  const [mode, setMode] = useState<'update' | 'regenerate'>('update');
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('Catalogo pendiente de actualizar');

  const provinces = useMemo(() => {
    const values = [...new Set(catalog.map((item) => item.province).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    return ['Todas', ...values];
  }, [catalog]);

  const matchingAreas = useMemo(() => {
    const needle = query.trim().toUpperCase();
    return catalog
      .filter((item) => province === 'Todas' || item.province === province)
      .filter((item) => !needle || `${item.areaId} ${item.areaName} ${item.companies.join(' ')}`.toUpperCase().includes(needle));
  }, [catalog, province, query]);
  const visibleAreas = useMemo(() => matchingAreas.slice(0, 80), [matchingAreas]);

  async function refreshCatalog() {
    setBusy(true);
    setMessage('Actualizando catalogo desde Capitulo IV - Pozos');
    try {
      await appendDebug(createDebugEntry('Catalogo', 'info', 'Inicio descarga de catalogo'));
      const items = await fetchAreaCatalog();
      setCatalog(items);
      setMessage(`${items.length} areas disponibles`);
      await appendDebug(createDebugEntry('Catalogo', 'ok', `${items.length} areas disponibles`));
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      setMessage(detail);
      await appendDebug(createDebugEntry('Catalogo', 'error', detail));
    } finally {
      setBusy(false);
    }
  }

  function toggleArea(area: AreaCatalogItem) {
    setSelected((current) => {
      if (current.some((item) => item.areaId === area.areaId)) {
        return current.filter((item) => item.areaId !== area.areaId);
      }
      return [...current, { ...area, startYearOverride: defaults.startYear }];
    });
  }

  function selectMatchingAreas() {
    setSelected((current) => {
      const byId = new Map(current.map((item) => [item.areaId, item]));
      for (const area of matchingAreas) {
        if (!byId.has(area.areaId)) byId.set(area.areaId, { ...area, startYearOverride: defaults.startYear });
      }
      return [...byId.values()];
    });
  }

  function updateAreaStartYear(areaId: string, startYearOverride: number) {
    setSelected((current) => current.map((item) => (item.areaId === areaId ? { ...item, startYearOverride } : item)));
  }

  async function runBuild() {
    if (selected.length === 0) {
      setMessage('Seleccionar al menos un area');
      return;
    }
    setBusy(true);
    setMessage('Generando workbook');
    try {
      const plans: AreaWorkbookPlan[] = selected.map((selection) => ({
        selection,
        defaults,
        mode,
      }));
      await buildWorkbook(plans);
      setMessage('Workbook actualizado');
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      setMessage(detail);
      await appendDebug(createDebugEntry('Workbook', 'error', detail));
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <img src="/assets/branding/logo_isotipo.png" alt="" />
        <div>
          <h1>GLP</h1>
          <p>Historico, proyeccion y graficos por area</p>
        </div>
      </header>

      <section className="panel">
        <div className="section-header">
          <h2>Areas</h2>
          <button type="button" onClick={refreshCatalog} disabled={busy}>Actualizar catalogo</button>
        </div>
        <label>
          Provincia
          <select value={province} onChange={(event) => setProvince(event.target.value)}>
            {provinces.map((item) => <option key={item}>{item}</option>)}
          </select>
        </label>
        <label>
          Buscar area
          <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Codigo, nombre o empresa" />
        </label>
        <div className="area-list">
          {visibleAreas.map((area) => {
            const checked = selected.some((item) => item.areaId === area.areaId);
            return (
              <button
                type="button"
                key={area.areaId}
                className={checked ? 'area-item selected' : 'area-item'}
                onClick={() => toggleArea(area)}
              >
                <strong>{area.areaId}</strong>
                <span>{area.areaName}</span>
                <small>{area.province}</small>
              </button>
            );
          })}
        </div>
        <div className="list-actions">
          <button type="button" onClick={selectMatchingAreas} disabled={busy || matchingAreas.length === 0}>
            Seleccionar filtradas ({matchingAreas.length})
          </button>
          <button type="button" onClick={() => setSelected([])} disabled={busy || selected.length === 0}>Limpiar seleccion</button>
        </div>
      </section>

      <section className="panel">
        <h2>Configuracion inicial</h2>
        <label>
          Ano de inicio
          <input
            type="number"
            min="2006"
            max="2030"
            value={defaults.startYear}
            onChange={(event) => setDefaults({ ...defaults, startYear: Number(event.target.value) })}
          />
        </label>
        <label>
          Horizonte (anos)
          <input
            type="number"
            min="1"
            max="40"
            value={defaults.horizonYears}
            onChange={(event) => setDefaults({ ...defaults, horizonYears: Number(event.target.value) })}
          />
        </label>
        <label>
          Bruta
          <select value={defaults.grossMethod} onChange={(event) => setDefaults({ ...defaults, grossMethod: event.target.value as ForecastDefaults['grossMethod'] })}>
            <option>Constante</option>
            <option>HypMod</option>
            <option>Declinación Hip.</option>
            <option>Declinación Exp.</option>
          </select>
        </label>
        <label>
          Petroleo
          <select value={defaults.oilMethod} onChange={(event) => setDefaults({ ...defaults, oilMethod: event.target.value as ForecastDefaults['oilMethod'] })}>
            <option>Constante</option>
            <option>HypMod</option>
            <option>Declinación Hip.</option>
            <option>Declinación Exp.</option>
            <option>Rap Np</option>
          </select>
        </label>
        <label>
          Gas
          <select value={defaults.gasMethod} onChange={(event) => setDefaults({ ...defaults, gasMethod: event.target.value as ForecastDefaults['gasMethod'] })}>
            <option>Constante</option>
            <option>HypMod</option>
            <option>Declinación Hip.</option>
            <option>Declinación Exp.</option>
            <option>RGP</option>
          </select>
        </label>
        <label className="check-row">
          <input
            type="checkbox"
            checked={defaults.takeInitialFromHistory}
            onChange={(event) => setDefaults({ ...defaults, takeInitialFromHistory: event.target.checked })}
          />
          Tomar inicial desde historia
        </label>
      </section>

      <section className="panel">
        <h2>Salida</h2>
        <div className="segmented">
          <button type="button" className={mode === 'update' ? 'active' : ''} onClick={() => setMode('update')}>Actualizar datos</button>
          <button type="button" className={mode === 'regenerate' ? 'active' : ''} onClick={() => setMode('regenerate')}>Regenerar area</button>
        </div>
        <p className="selected-count">{selected.length} areas seleccionadas</p>
        {selected.length > 0 && (
          <div className="selected-areas">
            {selected.map((area) => (
              <div className="selected-row" key={area.areaId}>
                <div>
                  <strong>{area.areaId}</strong>
                  <span>{area.areaName}</span>
                </div>
                <label>
                  Inicio
                  <input
                    type="number"
                    min="2006"
                    max="2030"
                    value={area.startYearOverride ?? defaults.startYear}
                    onChange={(event) => updateAreaStartYear(area.areaId, Number(event.target.value))}
                  />
                </label>
                <button type="button" onClick={() => toggleArea(area)} disabled={busy}>Quitar</button>
              </div>
            ))}
          </div>
        )}
      </section>

      <footer className="actionbar">
        <p>{message}</p>
        <button type="button" className="primary" disabled={busy || selected.length === 0} onClick={runBuild}>
          {busy ? 'Procesando' : 'Generar hojas'}
        </button>
      </footer>
    </main>
  );
}
