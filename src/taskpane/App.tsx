import React, { useEffect, useMemo, useState } from 'react';
import { buildWorkbook } from '../excel/workbookBuilder';
import { appendDebug, createDebugEntry } from '../excel/debugSheet';
import { readSavedWorkbookPlans } from '../excel/workbookState';
import type {
  AreaCatalogItem,
  AreaForecastOverride,
  AreaSelection,
  AreaWorkbookPlan,
  BuildProgress,
  ForecastDefaults,
  MissingMonthsDecision,
} from '../models/types';
import { fetchAreaCatalog } from '../services/capiv';

const DEFAULTS: ForecastDefaults = {
  startYear: 2015,
  horizonYears: 10,
  grossMethod: 'Constante',
  oilMethod: 'Declinación Exp.',
  gasMethod: 'RGP',
  takeInitialFromHistory: true,
};

const GROSS_METHODS: ForecastDefaults['grossMethod'][] = ['Constante', 'HypMod', 'Declinación Hip.', 'Declinación Exp.'];
const OIL_METHODS: ForecastDefaults['oilMethod'][] = [...GROSS_METHODS, 'Rap Np'];
const GAS_METHODS: ForecastDefaults['gasMethod'][] = [...GROSS_METHODS, 'RGP'];

type StatusTone = 'neutral' | 'success' | 'warning' | 'error';

interface StatusMessage {
  tone: StatusTone;
  text: string;
}

export function App() {
  const [catalog, setCatalog] = useState<AreaCatalogItem[]>([]);
  const [province, setProvince] = useState('Todas');
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState<AreaSelection[]>([]);
  const [overrides, setOverrides] = useState<Record<string, AreaForecastOverride>>({});
  const [defaults, setDefaults] = useState<ForecastDefaults>(DEFAULTS);
  const [mode, setMode] = useState<'update' | 'regenerate'>('update');
  const [catalogBusy, setCatalogBusy] = useState(false);
  const [buildBusy, setBuildBusy] = useState(false);
  const [status, setStatus] = useState<StatusMessage>({ tone: 'neutral', text: 'Conectando con Capítulo IV…' });
  const [progress, setProgress] = useState<BuildProgress | null>(null);
  const [missingDecision, setMissingDecision] = useState<{
    request: MissingMonthsDecision;
    resolve: (policy: 'blank' | 'zero') => void;
  } | null>(null);
  const busy = catalogBusy || buildBusy;

  const provinces = useMemo(() => {
    const values = [...new Set(catalog.map((item) => item.province).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    return ['Todas', ...values];
  }, [catalog]);

  const matchingAreas = useMemo(() => {
    const needle = query.trim().toLocaleUpperCase('es-AR');
    return catalog
      .filter((item) => province === 'Todas' || item.province === province)
      .filter((item) => !needle || `${item.areaId} ${item.areaName} ${item.companies.join(' ')}`.toLocaleUpperCase('es-AR').includes(needle));
  }, [catalog, province, query]);
  const visibleAreas = useMemo(() => matchingAreas.slice(0, 100), [matchingAreas]);

  useEffect(() => {
    void refreshCatalog();
  }, []);

  async function refreshCatalog() {
    setCatalogBusy(true);
    setStatus({ tone: 'neutral', text: 'Actualizando catálogo oficial…' });
    try {
      await logDebugSafely('Catálogo', 'info', 'Inicio de descarga del catálogo');
      const items = await fetchAreaCatalog();
      setCatalog(items);
      setStatus({ tone: 'success', text: `${items.length} áreas disponibles` });
      await logDebugSafely('Catálogo', 'ok', `${items.length} áreas disponibles`);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      setStatus({ tone: 'error', text: detail });
      await logDebugSafely('Catálogo', 'error', detail);
    } finally {
      setCatalogBusy(false);
    }
  }

  function toggleArea(area: AreaCatalogItem) {
    setSelected((current) => {
      if (current.some((item) => item.areaId === area.areaId)) {
        setOverrides((currentOverrides) => {
          const next = { ...currentOverrides };
          delete next[area.areaId];
          return next;
        });
        return current.filter((item) => item.areaId !== area.areaId);
      }
      return [...current, { ...area }];
    });
  }

  function selectMatchingAreas() {
    setSelected((current) => {
      const byId = new Map(current.map((item) => [item.areaId, item]));
      for (const area of matchingAreas) {
        if (!byId.has(area.areaId)) byId.set(area.areaId, { ...area });
      }
      return [...byId.values()];
    });
  }

  function clearSelection() {
    setSelected([]);
    setOverrides({});
  }

  function updateOverride(areaId: string, patch: Partial<AreaForecastOverride>) {
    setOverrides((current) => ({ ...current, [areaId]: { ...current[areaId], ...patch, areaId } }));
  }

  function clearOverride(areaId: string, key?: keyof Omit<AreaForecastOverride, 'areaId'>) {
    setOverrides((current) => {
      const next = { ...current };
      if (!key) {
        delete next[areaId];
        return next;
      }
      const areaOverride = { ...next[areaId] };
      delete areaOverride[key];
      if (Object.keys(areaOverride).every((item) => item === 'areaId')) delete next[areaId];
      else next[areaId] = areaOverride as AreaForecastOverride;
      return next;
    });
  }

  async function runBuild() {
    if (selected.length === 0) {
      setStatus({ tone: 'warning', text: 'Seleccioná al menos un área.' });
      return;
    }
    if (typeof Excel === 'undefined') {
      setStatus({ tone: 'error', text: 'Para generar hojas, abrí GLP desde Microsoft Excel.' });
      return;
    }
    setBuildBusy(true);
    setStatus({ tone: 'neutral', text: 'Preparando el workbook…' });
    setProgress({ completed: 0, total: 1, percent: 0, message: 'Preparando generación' });
    try {
      const plans: AreaWorkbookPlan[] = selected.map((selection) => ({
        selection,
        defaults,
        override: overrides[selection.areaId],
        mode,
      }));
      await buildWorkbook(
        plans,
        (nextProgress) => {
          setProgress(nextProgress);
          setStatus({ tone: 'neutral', text: nextProgress.message });
        },
        (request) => new Promise((resolve) => setMissingDecision({ request, resolve })),
      );
      setStatus({ tone: 'success', text: 'Workbook actualizado correctamente.' });
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      setStatus({ tone: 'error', text: detail });
      await logDebugSafely('Workbook', 'error', detail);
    } finally {
      setBuildBusy(false);
    }
  }

  async function runSavedWorkbookUpdate() {
    if (typeof Excel === 'undefined') {
      setStatus({ tone: 'error', text: 'Abrí GLP desde Microsoft Excel para actualizar el libro.' });
      return;
    }
    setBuildBusy(true);
    setStatus({ tone: 'neutral', text: 'Buscando áreas generadas anteriormente…' });
    setProgress({ completed: 0, total: 1, percent: 0, message: 'Leyendo estado del libro' });
    try {
      const saved = await readSavedWorkbookPlans();
      if (!saved) {
        setStatus({ tone: 'warning', text: 'Este libro todavía no tiene áreas generadas por GLP.' });
        setProgress(null);
        return;
      }
      setSelected(saved.plans.map((plan) => plan.selection));
      await appendDebug(createDebugEntry('Actualización automática', 'info', `Áreas detectadas: ${saved.plans.map((plan) => plan.selection.areaId).join(', ')}`));
      await buildWorkbook(
        saved.plans,
        (nextProgress) => {
          setProgress(nextProgress);
          setStatus({ tone: 'neutral', text: nextProgress.message });
        },
        (request) => new Promise((resolve) => setMissingDecision({ request, resolve })),
      );
      setStatus({
        tone: 'success',
        text: `${saved.plans.length} ${saved.plans.length === 1 ? 'área actualizada' : 'áreas actualizadas'} con la información oficial más reciente.`,
      });
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      setStatus({ tone: 'error', text: detail });
      await logDebugSafely('Actualización automática', 'error', detail);
    } finally {
      setBuildBusy(false);
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <img src="/assets/branding/logo_isotipo.png" alt="Quintana Energy" />
        <div className="brand-copy">
          <div className="product-line"><h1>GLP</h1><span>v0.2</span></div>
          <p>Capítulo IV · Histórico y pronóstico</p>
        </div>
        <span className={catalog.length ? 'connection-dot online' : 'connection-dot'} title={catalog.length ? 'Catálogo conectado' : 'Conectando'} />
      </header>

      <div className="content">
        <section className="quick-update-card">
          <div>
            <span className="quick-update-kicker">¿Ya usaste GLP en este libro?</span>
            <h2>Traé los meses nuevos</h2>
            <p>Detecta las áreas existentes, refresca la serie oficial y conserva tus supuestos.</p>
          </div>
          <button type="button" onClick={runSavedWorkbookUpdate} disabled={busy}>Actualizar libro</button>
        </section>
        <section className="panel">
          <SectionHeading step="1" title="Elegí las áreas" description="Fuente oficial de Capítulo IV" />
          <button className="catalog-button" type="button" onClick={refreshCatalog} disabled={busy}>
            <span>{catalogBusy ? 'Actualizando…' : 'Actualizar catálogo'}</span>
            <small>{catalog.length ? `${catalog.length} áreas` : 'Sin datos locales'}</small>
          </button>
          <div className="field-grid">
            <label>
              Provincia
              <select value={province} onChange={(event) => setProvince(event.target.value)} disabled={!catalog.length || busy}>
                {provinces.map((item) => <option key={item}>{item}</option>)}
              </select>
            </label>
            <label>
              Buscar
              <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Área, código o empresa" disabled={!catalog.length || busy} />
            </label>
          </div>
          <div className="area-list" aria-label="Áreas disponibles">
            {!catalogBusy && visibleAreas.length === 0 && (
              <div className="empty-state">{catalog.length ? 'No hay áreas que coincidan con el filtro.' : 'El catálogo aparecerá acá cuando termine la conexión.'}</div>
            )}
            {visibleAreas.map((area) => {
              const checked = selected.some((item) => item.areaId === area.areaId);
              return (
                <button type="button" key={area.areaId} className={checked ? 'area-item selected' : 'area-item'} onClick={() => toggleArea(area)} disabled={buildBusy} aria-pressed={checked}>
                  <span className="selection-mark">{checked ? '✓' : ''}</span>
                  <span className="area-copy"><strong>{area.areaName}</strong><small>{area.areaId} · {area.province}</small></span>
                </button>
              );
            })}
          </div>
          {matchingAreas.length > visibleAreas.length && <p className="helper-text">Se muestran las primeras {visibleAreas.length} de {matchingAreas.length}. Usá los filtros para acotar.</p>}
          <div className="list-actions">
            <button type="button" onClick={selectMatchingAreas} disabled={busy || matchingAreas.length === 0}>Seleccionar filtradas ({matchingAreas.length})</button>
            <button type="button" className="ghost" onClick={clearSelection} disabled={busy || selected.length === 0}>Limpiar</button>
          </div>
        </section>

        <section className="panel">
          <SectionHeading step="2" title="Definí el pronóstico" description="Valores globales; se pueden ajustar por área" />
          <div className="field-grid two-columns">
            <label>
              Año de inicio
              <input type="number" min="2006" max={new Date().getFullYear()} value={defaults.startYear} onChange={(event) => setDefaults({ ...defaults, startYear: boundedNumber(event.target.value, 2006, new Date().getFullYear(), defaults.startYear) })} />
            </label>
            <label>
              Horizonte (años)
              <input type="number" min="1" max="40" value={defaults.horizonYears} onChange={(event) => setDefaults({ ...defaults, horizonYears: boundedNumber(event.target.value, 1, 40, defaults.horizonYears) })} />
            </label>
          </div>
          <div className="method-grid">
            <MethodSelect label="Bruta" value={defaults.grossMethod} options={GROSS_METHODS} onChange={(value) => setDefaults({ ...defaults, grossMethod: value as ForecastDefaults['grossMethod'] })} />
            <MethodSelect label="Petróleo" value={defaults.oilMethod} options={OIL_METHODS} onChange={(value) => setDefaults({ ...defaults, oilMethod: value as ForecastDefaults['oilMethod'] })} />
            <MethodSelect label="Gas" value={defaults.gasMethod} options={GAS_METHODS} onChange={(value) => setDefaults({ ...defaults, gasMethod: value as ForecastDefaults['gasMethod'] })} />
          </div>
          <label className="check-row">
            <input type="checkbox" checked={defaults.takeInitialFromHistory} onChange={(event) => setDefaults({ ...defaults, takeInitialFromHistory: event.target.checked })} />
            <span><strong>Tomar valores iniciales del histórico</strong><small>Los supuestos quedan editables en Excel.</small></span>
          </label>
        </section>

        <section className="panel">
          <SectionHeading step="3" title="Revisá la salida" description={`${selected.length} ${selected.length === 1 ? 'área seleccionada' : 'áreas seleccionadas'}`} />
          <div className="segmented" role="group" aria-label="Modo de generación">
            <button type="button" className={mode === 'update' ? 'active' : ''} onClick={() => setMode('update')} aria-pressed={mode === 'update'}>Actualizar</button>
            <button type="button" className={mode === 'regenerate' ? 'active' : ''} onClick={() => setMode('regenerate')} aria-pressed={mode === 'regenerate'}>Regenerar</button>
          </div>
          <p className="mode-description">{mode === 'update' ? 'Actualiza datos y recalcula, conservando los supuestos editables existentes.' : 'Elimina y crea nuevamente todas las hojas de cada área.'}</p>
          {selected.length === 0 ? (
            <div className="empty-state compact">Seleccioná al menos un área para habilitar la generación.</div>
          ) : (
            <div className="selected-areas">
              {selected.map((area) => {
                const areaOverride = overrides[area.areaId];
                return (
                  <article className="selected-card" key={area.areaId}>
                    <div className="selected-card-header">
                      <div><strong>{area.areaName}</strong><span>{area.areaId} · {area.province}</span></div>
                      <button type="button" className="icon-button" onClick={() => toggleArea(area)} disabled={busy} aria-label={`Quitar ${area.areaName}`}>×</button>
                    </div>
                    <div className="override-row">
                      <label>
                        Inicio
                        <input type="number" min="2006" max={new Date().getFullYear()} value={areaOverride?.startYear ?? defaults.startYear} onChange={(event) => updateOverride(area.areaId, { startYear: boundedNumber(event.target.value, 2006, new Date().getFullYear(), defaults.startYear) })} />
                      </label>
                      {areaOverride?.startYear !== undefined && <button type="button" className="text-button" onClick={() => clearOverride(area.areaId, 'startYear')}>Usar global</button>}
                    </div>
                    <details>
                      <summary>Ajustes avanzados del área</summary>
                      <div className="override-methods">
                        <OverrideSelect label="Bruta" value={areaOverride?.grossMethod} globalValue={defaults.grossMethod} options={GROSS_METHODS} onChange={(value) => value ? updateOverride(area.areaId, { grossMethod: value as ForecastDefaults['grossMethod'] }) : clearOverride(area.areaId, 'grossMethod')} />
                        <OverrideSelect label="Petróleo" value={areaOverride?.oilMethod} globalValue={defaults.oilMethod} options={OIL_METHODS} onChange={(value) => value ? updateOverride(area.areaId, { oilMethod: value as ForecastDefaults['oilMethod'] }) : clearOverride(area.areaId, 'oilMethod')} />
                        <OverrideSelect label="Gas" value={areaOverride?.gasMethod} globalValue={defaults.gasMethod} options={GAS_METHODS} onChange={(value) => value ? updateOverride(area.areaId, { gasMethod: value as ForecastDefaults['gasMethod'] }) : clearOverride(area.areaId, 'gasMethod')} />
                        <label>
                          Valor inicial
                          <select value={areaOverride?.takeInitialFromHistory === undefined ? '' : areaOverride.takeInitialFromHistory ? 'history' : 'manual'} onChange={(event) => event.target.value ? updateOverride(area.areaId, { takeInitialFromHistory: event.target.value === 'history' }) : clearOverride(area.areaId, 'takeInitialFromHistory')}>
                            <option value="">Global ({defaults.takeInitialFromHistory ? 'histórico' : 'manual'})</option>
                            <option value="history">Desde histórico</option>
                            <option value="manual">Manual en Excel</option>
                          </select>
                        </label>
                      </div>
                      {areaOverride && <button type="button" className="text-button reset" onClick={() => clearOverride(area.areaId)}>Restablecer ajustes del área</button>}
                    </details>
                  </article>
                );
              })}
            </div>
          )}
        </section>
      </div>

      <footer className="actionbar">
        {progress && buildBusy && (
          <div className="progress-box" aria-live="polite">
            <div className="progress-meta"><strong>{progress.percent}%</strong><span>{progress.completed}/{progress.total}</span></div>
            <div className="progress-track" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={progress.percent}><div className="progress-fill" style={{ width: `${progress.percent}%` }} /></div>
          </div>
        )}
        <p className={`status-message ${status.tone}`} role={status.tone === 'error' ? 'alert' : 'status'}><span />{status.text}</p>
        <button type="button" className="primary" disabled={busy || selected.length === 0} onClick={runBuild}>{buildBusy ? 'Procesando…' : `Generar ${selected.length || ''} ${selected.length === 1 ? 'área' : 'áreas'}`}</button>
      </footer>

      {missingDecision && (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="missing-title">
          <div className="decision-modal">
            <span className="modal-kicker">Revisión de datos</span>
            <h2 id="missing-title">Hay meses intermedios sin información</h2>
            <p><strong>{missingDecision.request.areaId}</strong> · {missingDecision.request.areaName}</p>
            <p>Elegí cómo representar estos meses en el histórico:</p>
            <p className="missing-list">{missingDecision.request.months.join(', ')}</p>
            <div className="modal-actions">
              <button type="button" onClick={() => { missingDecision.resolve('blank'); setMissingDecision(null); }}>Dejar vacíos</button>
              <button type="button" className="primary" onClick={() => { missingDecision.resolve('zero'); setMissingDecision(null); }}>Completar con 0</button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

function SectionHeading({ step, title, description }: { step: string; title: string; description: string }) {
  return <div className="section-heading"><span>{step}</span><div><h2>{title}</h2><p>{description}</p></div></div>;
}

function MethodSelect({ label, value, options, onChange }: { label: string; value: string; options: readonly string[]; onChange: (value: string) => void }) {
  return <label>{label}<select value={value} onChange={(event) => onChange(event.target.value)}>{options.map((option) => <option key={option}>{option}</option>)}</select></label>;
}

function OverrideSelect({ label, value, globalValue, options, onChange }: { label: string; value?: string; globalValue: string; options: readonly string[]; onChange: (value: string) => void }) {
  return <label>{label}<select value={value ?? ''} onChange={(event) => onChange(event.target.value)}><option value="">Global ({globalValue})</option>{options.map((option) => <option key={option}>{option}</option>)}</select></label>;
}

function boundedNumber(raw: string, min: number, max: number, fallback: number): number {
  const value = Number(raw);
  if (!Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, Math.round(value)));
}

async function logDebugSafely(step: string, status: 'info' | 'ok' | 'warning' | 'error', detail: string): Promise<void> {
  if (typeof Excel === 'undefined') return;
  try {
    await appendDebug(createDebugEntry(step, status, detail));
  } catch {
    // El log nunca debe bloquear la operación principal.
  }
}
