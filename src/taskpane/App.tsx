import React, { useEffect, useMemo, useState } from 'react';
import { appendDebug, createDebugEntry } from '../excel/debugSheet';
import { captureSelectedCell, createNewDataSheetTarget } from '../excel/databaseSheet';
import { buildForecastWorkbook, downloadWorkbookData } from '../excel/workbookBuilder';
import { readSavedWorkbookPlans } from '../excel/workbookState';
import type {
  AreaCatalogItem,
  AreaForecastOverride,
  AreaSelection,
  AreaWorkbookPlan,
  BuildProgress,
  DataGranularity,
  DataOutputTarget,
  ForecastDefaults,
  MissingMonthsDecision,
  OverwriteWarning,
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

type Workflow = 'data' | 'forecast';
type DestinationMode = 'new-sheet' | 'selected-cell';
type StatusTone = 'neutral' | 'success' | 'warning' | 'error';

interface StatusMessage {
  tone: StatusTone;
  text: string;
}

interface SavedInfo {
  areaCount: number;
  dataSavedAt?: string;
  forecastSavedAt?: string;
}

export function App() {
  const [workflow, setWorkflow] = useState<Workflow>('data');
  const [catalog, setCatalog] = useState<AreaCatalogItem[]>([]);
  const [province, setProvince] = useState('Todas');
  const [query, setQuery] = useState('');
  const [dataSelected, setDataSelected] = useState<AreaSelection[]>([]);
  const [forecastSelected, setForecastSelected] = useState<AreaSelection[]>([]);
  const [overrides, setOverrides] = useState<Record<string, AreaForecastOverride>>({});
  const [defaults, setDefaults] = useState<ForecastDefaults>(DEFAULTS);
  const [startYearDraft, setStartYearDraft] = useState(String(DEFAULTS.startYear));
  const [granularity, setGranularity] = useState<DataGranularity>('area');
  const [destinationMode, setDestinationMode] = useState<DestinationMode>('new-sheet');
  const [dataOutput, setDataOutput] = useState<DataOutputTarget | null>(null);
  const [forecastMode, setForecastMode] = useState<'update' | 'regenerate'>('update');
  const [savedInfo, setSavedInfo] = useState<SavedInfo | null>(null);
  const [catalogBusy, setCatalogBusy] = useState(false);
  const [buildBusy, setBuildBusy] = useState(false);
  const [status, setStatus] = useState<StatusMessage>({ tone: 'neutral', text: 'Conectando con Capítulo IV…' });
  const [progress, setProgress] = useState<BuildProgress | null>(null);
  const [missingDecision, setMissingDecision] = useState<{
    request: MissingMonthsDecision;
    resolve: (policy: 'blank' | 'zero') => void;
  } | null>(null);
  const [overwriteDecision, setOverwriteDecision] = useState<{
    warning: OverwriteWarning;
    resolve: (accepted: boolean) => void;
  } | null>(null);
  const busy = catalogBusy || buildBusy;
  const startYearsValid = isValidStartYear(startYearDraft);

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

  useEffect(() => {
    setStartYearDraft(String(defaults.startYear));
  }, [defaults.startYear]);

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
      await showError('Catálogo', error);
    } finally {
      setCatalogBusy(false);
    }
  }

  function toggleDataArea(area: AreaCatalogItem | AreaSelection) {
    setDataSelected((current) => {
      if (current.some((item) => item.areaId === area.areaId)) return current.filter((item) => item.areaId !== area.areaId);
      return [...current, { ...area }];
    });
  }

  function selectMatchingAreas() {
    setDataSelected((current) => {
      const byId = new Map(current.map((item) => [item.areaId, item]));
      for (const area of matchingAreas) if (!byId.has(area.areaId)) byId.set(area.areaId, { ...area });
      return [...byId.values()];
    });
  }

  function clearDataSelection() {
    setDataSelected([]);
  }

  function changeGlobalStartYear(raw: string) {
    if (!isYearDraft(raw)) return;
    setStartYearDraft(raw);
    if (isValidStartYear(raw)) setDefaults((current) => ({ ...current, startYear: Number(raw) }));
  }

  function toggleForecastArea(area: AreaSelection) {
    setForecastSelected((current) => current.some((item) => item.areaId === area.areaId)
      ? current.filter((item) => item.areaId !== area.areaId)
      : [...current, area]);
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

  function dataPlans(selections: AreaSelection[], mode: 'update' | 'regenerate'): AreaWorkbookPlan[] {
    return selections.map((selection) => ({ selection, defaults, mode }));
  }

  function forecastPlans(): AreaWorkbookPlan[] {
    return forecastSelected.map((selection) => ({
      selection,
      defaults,
      override: overrides[selection.areaId],
      mode: forecastMode,
    }));
  }

  function reportProgress(nextProgress: BuildProgress) {
    setProgress(nextProgress);
    setStatus({ tone: 'neutral', text: nextProgress.message });
  }

  function requestMissingMonths(request: MissingMonthsDecision): Promise<'blank' | 'zero'> {
    return new Promise((resolve) => setMissingDecision({ request, resolve }));
  }

  function requestOverwrite(warning: OverwriteWarning): Promise<boolean> {
    return new Promise((resolve) => setOverwriteDecision({ warning, resolve }));
  }

  async function chooseDestination() {
    if (!ensureExcel('Seleccioná una celda en Excel y volvé a intentarlo.')) return;
    try {
      const target = await captureSelectedCell(granularity);
      setDataOutput(target);
      setDestinationMode('selected-cell');
      setStatus({ tone: 'success', text: `Destino elegido: ${target.sheetName}!${target.startAddress}` });
    } catch (error) {
      await showError('Celda de destino', error);
    }
  }

  function changeGranularity(value: DataGranularity) {
    setGranularity(value);
    setDataOutput((current) => current ? { ...current, granularity: value } : current);
  }

  async function runDataDownload() {
    if (!startYearsValid) {
      setStatus({ tone: 'warning', text: `Ingresá un año de 4 dígitos entre ${START_YEAR_MIN} y ${START_YEAR_MAX}.` });
      return;
    }
    if (dataSelected.length === 0) {
      setStatus({ tone: 'warning', text: 'Seleccioná al menos un área para descargar.' });
      return;
    }
    if (destinationMode === 'selected-cell' && !dataOutput) {
      setStatus({ tone: 'warning', text: 'Seleccioná en Excel la celda donde querés crear la tabla.' });
      return;
    }
    if (!ensureExcel()) return;
    setBuildBusy(true);
    setProgress({ completed: 0, total: 1, percent: 0, message: 'Preparando descarga' });
    setStatus({ tone: 'neutral', text: 'Preparando descarga…' });
    try {
      const output = destinationMode === 'new-sheet'
        ? await createNewDataSheetTarget(granularity)
        : dataOutput!;
      const plans = dataPlans(dataSelected, 'update');
      await downloadWorkbookData(plans, output, reportProgress, requestMissingMonths, requestOverwrite);
      setDataOutput(output);
      setForecastSelected(dataSelected);
      setSavedInfo({ areaCount: plans.length, dataSavedAt: new Date().toISOString() });
      setStatus({ tone: 'success', text: `Tabla ${granularity === 'area' ? 'mensual por área' : 'pozo-mes'} generada en ${output.sheetName}!${output.startAddress}.` });
    } catch (error) {
      await showError('Descarga', error);
    } finally {
      setBuildBusy(false);
    }
  }

  async function runSavedWorkbookUpdate() {
    if (!ensureExcel()) return;
    setBuildBusy(true);
    setStatus({ tone: 'neutral', text: 'Buscando áreas descargadas anteriormente…' });
    setProgress({ completed: 0, total: 1, percent: 0, message: 'Leyendo estado del libro' });
    try {
      const saved = await readSavedWorkbookPlans();
      if (!saved || saved.plans.length === 0 || !saved.dataOutput) {
        setStatus({ tone: 'warning', text: 'Este libro todavía no tiene una tabla descargada por CapIV.' });
        setProgress(null);
        return;
      }
      const plans = saved.plans.map((plan) => ({ ...plan, mode: 'update' as const }));
      setDataOutput(saved.dataOutput);
      setGranularity(saved.dataOutput.granularity);
      setDataSelected(plans.map((plan) => plan.selection));
      setForecastSelected(plans.map((plan) => plan.selection));
      await appendDebug(createDebugEntry('Actualización de datos', 'info', `Áreas detectadas: ${plans.map((plan) => plan.selection.areaId).join(', ')}`));
      await downloadWorkbookData(plans, saved.dataOutput, reportProgress, requestMissingMonths, requestOverwrite);
      setSavedInfo({ areaCount: plans.length, dataSavedAt: new Date().toISOString() });
      setStatus({ tone: 'success', text: `Tabla actualizada en ${saved.dataOutput.sheetName}!${saved.dataOutput.startAddress}. Los pronósticos no se modificaron.` });
    } catch (error) {
      await showError('Actualización de datos', error);
    } finally {
      setBuildBusy(false);
    }
  }

  async function loadWorkbookData(switchFlow = false) {
    if (switchFlow) setWorkflow('forecast');
    if (!ensureExcel('Abrí CapIV desde Microsoft Excel para leer los datos guardados.')) return;
    setBuildBusy(true);
    setStatus({ tone: 'neutral', text: 'Leyendo datos guardados en el libro…' });
    try {
      const saved = await readSavedWorkbookPlans();
      if (!saved || saved.data.length === 0) {
        setForecastSelected([]);
        setSavedInfo(null);
        setStatus({ tone: 'warning', text: 'No hay datos descargados. Completá primero el flujo Datos.' });
        return;
      }
      const availableIds = new Set(saved.data.map((item) => item.areaId));
      const availablePlans = saved.plans.filter((plan) => availableIds.has(plan.selection.areaId));
      setForecastSelected(availablePlans.map((plan) => plan.selection));
      if (availablePlans[0]) setDefaults(availablePlans[0].defaults);
      setOverrides(Object.fromEntries(availablePlans.filter((plan) => plan.override).map((plan) => [plan.selection.areaId, plan.override!])));
      setSavedInfo({
        areaCount: availablePlans.length,
        dataSavedAt: saved.dataSavedAt,
        forecastSavedAt: saved.forecastSavedAt,
      });
      setStatus({ tone: 'success', text: `${availablePlans.length} ${availablePlans.length === 1 ? 'área disponible' : 'áreas disponibles'} para pronosticar.` });
    } catch (error) {
      await showError('Datos del libro', error);
    } finally {
      setBuildBusy(false);
    }
  }

  async function runForecast() {
    if (forecastSelected.length === 0) {
      setStatus({ tone: 'warning', text: 'Cargá y seleccioná datos del libro antes de pronosticar.' });
      return;
    }
    if (!ensureExcel()) return;
    setBuildBusy(true);
    setProgress({ completed: 0, total: 1, percent: 0, message: 'Preparando pronósticos' });
    setStatus({ tone: 'neutral', text: 'Preparando pronósticos…' });
    try {
      const plans = forecastPlans();
      await buildForecastWorkbook(plans, reportProgress);
      setSavedInfo((current) => ({ areaCount: plans.length, dataSavedAt: current?.dataSavedAt, forecastSavedAt: new Date().toISOString() }));
      setStatus({ tone: 'success', text: `${plans.length} ${plans.length === 1 ? 'pronóstico generado' : 'pronósticos generados'} sin volver a descargar datos.` });
    } catch (error) {
      await showError('Pronósticos', error);
    } finally {
      setBuildBusy(false);
    }
  }

  function ensureExcel(message = 'Abrí CapIV desde Microsoft Excel para trabajar con el libro.'): boolean {
    if (typeof Excel !== 'undefined') return true;
    setStatus({ tone: 'error', text: message });
    return false;
  }

  async function showError(step: string, error: unknown) {
    const detail = error instanceof Error ? error.message : String(error);
    setStatus({ tone: 'error', text: detail });
    await logDebugSafely(step, 'error', detail);
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <img src="assets/branding/logo_isotipo.png" alt="Quintana Energy" />
        <div className="brand-copy">
          <div className="product-line"><h1>CapIV</h1><span>v0.4</span></div>
          <p>Capítulo IV · Datos y pronósticos</p>
        </div>
        <span className={catalog.length ? 'connection-dot online' : 'connection-dot'} title={catalog.length ? 'Catálogo conectado' : 'Conectando'} />
      </header>

      <nav className="workflow-switch" aria-label="Flujo de trabajo">
        <button type="button" className={workflow === 'data' ? 'active' : ''} onClick={() => setWorkflow('data')} aria-pressed={workflow === 'data'}>
          <span>1</span><strong>Datos</strong><small>Descargar y actualizar</small>
        </button>
        <button type="button" className={workflow === 'forecast' ? 'active' : ''} onClick={() => void loadWorkbookData(true)} aria-pressed={workflow === 'forecast'}>
          <span>2</span><strong>Pronósticos</strong><small>Opcional · Modelar</small>
        </button>
      </nav>

      <div className="content">
        {workflow === 'data' ? (
          <>
            <section className="quick-update-card">
              <div>
                <span className="quick-update-kicker">Libro existente</span>
                <h2>Traé los meses nuevos</h2>
                <p>Detecta la tabla anterior y trae el último mes publicado, sin duplicados.</p>
              </div>
              <button type="button" onClick={runSavedWorkbookUpdate} disabled={busy}>Actualizar datos</button>
            </section>

            <section className="panel">
              <SectionHeading step="1" title="Elegí las áreas" description="Fuente oficial de Capítulo IV" />
              <button className="catalog-button" type="button" onClick={refreshCatalog} disabled={busy}>
                <span>{catalogBusy ? 'Actualizando…' : 'Actualizar catálogo'}</span>
                <small>{catalog.length ? `${catalog.length} áreas` : 'Sin datos locales'}</small>
              </button>
              <div className="field-grid">
                <label>Provincia<select value={province} onChange={(event) => setProvince(event.target.value)} disabled={!catalog.length || busy}>{provinces.map((item) => <option key={item}>{item}</option>)}</select></label>
                <label>Buscar<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Área, código o empresa" disabled={!catalog.length || busy} /></label>
              </div>
              <div className="area-list" aria-label="Áreas disponibles">
                {!catalogBusy && visibleAreas.length === 0 && <div className="empty-state">{catalog.length ? 'No hay áreas que coincidan con el filtro.' : 'El catálogo aparecerá acá cuando termine la conexión.'}</div>}
                {visibleAreas.map((area) => {
                  const checked = dataSelected.some((item) => item.areaId === area.areaId);
                  return (
                    <button type="button" key={area.areaId} className={checked ? 'area-item selected' : 'area-item'} onClick={() => toggleDataArea(area)} disabled={buildBusy} aria-pressed={checked}>
                      <span className="selection-mark">{checked ? '✓' : ''}</span>
                      <span className="area-copy"><strong>{area.areaName}</strong><small title={`${area.areaId} · ${area.province} · ${companyNames(area)}`}>{area.areaId} · {area.province} · {companyNames(area)}</small></span>
                    </button>
                  );
                })}
              </div>
              {matchingAreas.length > visibleAreas.length && <p className="helper-text">Se muestran las primeras {visibleAreas.length} de {matchingAreas.length}. Usá los filtros para acotar.</p>}
              <div className="list-actions">
                <button type="button" onClick={selectMatchingAreas} disabled={busy || matchingAreas.length === 0}>Seleccionar filtradas ({matchingAreas.length})</button>
                <button type="button" className="ghost" onClick={clearDataSelection} disabled={busy || dataSelected.length === 0}>Limpiar</button>
              </div>
            </section>

            <section className="panel">
              <SectionHeading step="2" title="Elegí el nivel de la base" description="Una sola tabla, lista para filtrar o analizar" />
              <StartYearField id="start-year-global" label="Año de inicio" value={startYearDraft} onChange={changeGlobalStartYear} />
              <div className="segmented level-selector" role="group" aria-label="Nivel de detalle">
                <button type="button" className={granularity === 'area' ? 'active' : ''} onClick={() => changeGranularity('area')} aria-pressed={granularity === 'area'}><strong>Por área</strong><small>Un registro por mes; el add-in calcula los totales.</small></button>
                <button type="button" className={granularity === 'well' ? 'active' : ''} onClick={() => changeGranularity('well')} aria-pressed={granularity === 'well'}><strong>Por pozo</strong><small>Detalle completo de cada pozo y mes.</small></button>
              </div>
              {dataSelected.length === 0 ? <div className="empty-state compact">Seleccioná áreas para habilitar la descarga.</div> : (
                <div className="selected-areas data-selection">
                  {dataSelected.map((area) => {
                    return (
                      <article className="selected-card" key={area.areaId}>
                        <div className="selected-card-header">
                          <div><strong>{area.areaName}</strong><span title={`${area.areaId} · ${area.province} · ${companyNames(area)}`}>{area.areaId} · {area.province} · {companyNames(area)}</span></div>
                          <button type="button" className="icon-button" onClick={() => toggleDataArea(area)} disabled={busy} aria-label={`Quitar ${area.areaName}`}>×</button>
                        </div>
                      </article>
                    );
                  })}
                </div>
              )}
              <div className="segmented destination-selector" role="group" aria-label="Destino de la tabla">
                <button type="button" className={destinationMode === 'new-sheet' ? 'active' : ''} onClick={() => setDestinationMode('new-sheet')} aria-pressed={destinationMode === 'new-sheet'}><strong>Nueva hoja</strong><small>Automático · recomendado</small></button>
                <button type="button" className={destinationMode === 'selected-cell' ? 'active' : ''} onClick={() => setDestinationMode('selected-cell')} aria-pressed={destinationMode === 'selected-cell'}><strong>Celda actual</strong><small>Ubicación personalizada</small></button>
              </div>
              {destinationMode === 'new-sheet' ? (
                <div className="automatic-destination"><strong>CapIV crea la hoja por vos</strong><span>La base comenzará en A1, en CapIV_Datos o el siguiente nombre libre.</span></div>
              ) : (
                <>
                  <div className="destination-card">
                    <div><strong>Celda de destino</strong><span>{dataOutput ? `${dataOutput.sheetName}!${dataOutput.startAddress}` : 'Seleccioná una celda vacía en Excel'}</span></div>
                    <button type="button" onClick={() => void chooseDestination()} disabled={busy}>{dataOutput ? 'Cambiar celda' : 'Usar celda seleccionada'}</button>
                  </div>
                  <p className="helper-text">Si el rango contiene datos, CapIV pedirá confirmación antes de sobrescribir.</p>
                </>
              )}
            </section>
          </>
        ) : (
          <>
            <section className="flow-source-card">
              <div>
                <span className="quick-update-kicker">Fuente: este Excel</span>
                <h2>Usá los datos ya descargados</h2>
                <p>{savedInfo ? `${savedInfo.areaCount} áreas disponibles · Datos ${formatSavedDate(savedInfo.dataSavedAt)}` : 'Primero leé el estado guardado dentro del libro.'}</p>
              </div>
              <button type="button" onClick={() => void loadWorkbookData()} disabled={busy}>Leer datos del libro</button>
            </section>

            <section className="panel">
              <SectionHeading step="1" title="Definí el pronóstico" description="No se realiza ninguna descarga" />
              <div className="forecast-scope-note">
                <strong>1 pronóstico por cada área / concesión</strong>
                <span>CapIV calcula cada selección por separado; sólo las reúne en Resumen_Areas.</span>
              </div>
              <div className="field-grid two-columns">
                <label>Horizonte (años)<input type="number" min="1" max="40" value={defaults.horizonYears} onChange={(event) => setDefaults({ ...defaults, horizonYears: boundedNumber(event.target.value, 1, 40, defaults.horizonYears) })} /></label>
                <label>Datos<input value={savedInfo ? `${savedInfo.areaCount} áreas` : 'Sin cargar'} disabled /></label>
              </div>
              <div className="method-grid">
                <MethodSelect label="Bruta" value={defaults.grossMethod} options={GROSS_METHODS} onChange={(value) => setDefaults({ ...defaults, grossMethod: value as ForecastDefaults['grossMethod'] })} />
                <MethodSelect label="Petróleo" value={defaults.oilMethod} options={OIL_METHODS} onChange={(value) => setDefaults({ ...defaults, oilMethod: value as ForecastDefaults['oilMethod'] })} />
                <MethodSelect label="Gas" value={defaults.gasMethod} options={GAS_METHODS} onChange={(value) => setDefaults({ ...defaults, gasMethod: value as ForecastDefaults['gasMethod'] })} />
              </div>
              <div className="forecast-output-note">
                <strong>Gráficos técnicos separados</strong>
                <span>Producción, relaciones, inyección, acumuladas, pozos y RAP vs. Np; cada visual conserva unidades compatibles.</span>
              </div>
              <label className="check-row">
                <input type="checkbox" checked={defaults.takeInitialFromHistory} onChange={(event) => setDefaults({ ...defaults, takeInitialFromHistory: event.target.checked })} />
                <span><strong>Tomar valores iniciales del histórico</strong><small>Los supuestos quedan editables en Excel.</small></span>
              </label>
            </section>

            <section className="panel">
              <SectionHeading step="2" title="Elegí qué pronosticar" description={`${forecastSelected.length} ${forecastSelected.length === 1 ? 'área activa' : 'áreas activas'}`} />
              <ModeSelector mode={forecastMode} onChange={setForecastMode} updateText="Conserva los supuestos editados en Prono y Pozos." regenerateText="Reconstruye pronósticos, gráficos y resumen desde cero." />
              {forecastSelected.length === 0 ? <div className="empty-state compact">No hay datos cargados. Volvé al flujo Datos o leé el estado del libro.</div> : (
                <div className="selected-areas">
                  {forecastSelected.map((area) => {
                    const areaOverride = overrides[area.areaId];
                    return (
                      <article className="selected-card" key={area.areaId}>
                        <div className="selected-card-header">
                          <div><strong>{area.areaName}</strong><span>{area.areaId} · {area.province}</span></div>
                          <button type="button" className="icon-button" onClick={() => toggleForecastArea(area)} disabled={busy} aria-label={`Excluir ${area.areaName}`}>×</button>
                        </div>
                        <details>
                          <summary>Ajustes específicos del área</summary>
                          <div className="override-methods">
                            <OverrideSelect label="Bruta" value={areaOverride?.grossMethod} globalValue={defaults.grossMethod} options={GROSS_METHODS} onChange={(value) => value ? updateOverride(area.areaId, { grossMethod: value as ForecastDefaults['grossMethod'] }) : clearOverride(area.areaId, 'grossMethod')} />
                            <OverrideSelect label="Petróleo" value={areaOverride?.oilMethod} globalValue={defaults.oilMethod} options={OIL_METHODS} onChange={(value) => value ? updateOverride(area.areaId, { oilMethod: value as ForecastDefaults['oilMethod'] }) : clearOverride(area.areaId, 'oilMethod')} />
                            <OverrideSelect label="Gas" value={areaOverride?.gasMethod} globalValue={defaults.gasMethod} options={GAS_METHODS} onChange={(value) => value ? updateOverride(area.areaId, { gasMethod: value as ForecastDefaults['gasMethod'] }) : clearOverride(area.areaId, 'gasMethod')} />
                            <label>Valor inicial<select value={areaOverride?.takeInitialFromHistory === undefined ? '' : areaOverride.takeInitialFromHistory ? 'history' : 'manual'} onChange={(event) => event.target.value ? updateOverride(area.areaId, { takeInitialFromHistory: event.target.value === 'history' }) : clearOverride(area.areaId, 'takeInitialFromHistory')}><option value="">Global ({defaults.takeInitialFromHistory ? 'histórico' : 'manual'})</option><option value="history">Desde histórico</option><option value="manual">Manual en Excel</option></select></label>
                          </div>
                          {areaOverride && <button type="button" className="text-button reset" onClick={() => clearOverride(area.areaId)}>Restablecer ajustes del área</button>}
                        </details>
                      </article>
                    );
                  })}
                </div>
              )}
            </section>
          </>
        )}
      </div>

      <footer className="actionbar">
        {progress && buildBusy && (
          <div className="progress-box" aria-live="polite">
            <div className="progress-meta"><strong>{progress.percent}%</strong><span>{progress.completed}/{progress.total}</span></div>
            <div className="progress-track" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={progress.percent}><div className="progress-fill" style={{ width: `${progress.percent}%` }} /></div>
          </div>
        )}
        <p className={`status-message ${status.tone}`} role={status.tone === 'error' ? 'alert' : 'status'}><span />{status.text}</p>
        {workflow === 'data' ? (
          <button type="button" className="primary" disabled={busy || dataSelected.length === 0 || (destinationMode === 'selected-cell' && !dataOutput) || !startYearsValid} onClick={runDataDownload}>{buildBusy ? 'Descargando…' : `Crear tabla ${dataSelected.length ? `(${dataSelected.length} áreas)` : ''}`}</button>
        ) : (
          <button type="button" className="primary forecast-action" disabled={busy || forecastSelected.length === 0} onClick={runForecast}>{buildBusy ? 'Generando…' : `Generar pronósticos ${forecastSelected.length ? `(${forecastSelected.length})` : ''}`}</button>
        )}
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

      {overwriteDecision && (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="overwrite-title">
          <div className="decision-modal">
            <span className="modal-kicker">Confirmación necesaria</span>
            <h2 id="overwrite-title">La tabla va a sobrescribir datos</h2>
            <p>Destino: <strong>{overwriteDecision.warning.rangeAddress}</strong></p>
            <p>Hay {overwriteDecision.warning.occupiedCells} celdas con contenido{overwriteDecision.warning.overlappingTables.length ? ` y ${overwriteDecision.warning.overlappingTables.length} tabla(s) existente(s)` : ''}.</p>
            <p>Continuar borrará ese contenido para crear la nueva base.</p>
            <div className="modal-actions">
              <button type="button" onClick={() => { overwriteDecision.resolve(false); setOverwriteDecision(null); }}>Cancelar</button>
              <button type="button" className="primary danger-action" onClick={() => { overwriteDecision.resolve(true); setOverwriteDecision(null); }}>Sobrescribir</button>
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

function ModeSelector({ mode, onChange, updateText, regenerateText }: { mode: 'update' | 'regenerate'; onChange: (mode: 'update' | 'regenerate') => void; updateText: string; regenerateText: string }) {
  return <><div className="segmented" role="group" aria-label="Modo de escritura"><button type="button" className={mode === 'update' ? 'active' : ''} onClick={() => onChange('update')} aria-pressed={mode === 'update'}>Actualizar</button><button type="button" className={mode === 'regenerate' ? 'active' : ''} onClick={() => onChange('regenerate')} aria-pressed={mode === 'regenerate'}>Regenerar</button></div><p className="mode-description">{mode === 'update' ? updateText : regenerateText}</p></>;
}

function MethodSelect({ label, value, options, onChange }: { label: string; value: string; options: readonly string[]; onChange: (value: string) => void }) {
  return <label>{label}<select value={value} onChange={(event) => onChange(event.target.value)}>{options.map((option) => <option key={option}>{option}</option>)}</select></label>;
}

function OverrideSelect({ label, value, globalValue, options, onChange }: { label: string; value?: string; globalValue: string; options: readonly string[]; onChange: (value: string) => void }) {
  return <label>{label}<select value={value ?? ''} onChange={(event) => onChange(event.target.value)}><option value="">Global ({globalValue})</option>{options.map((option) => <option key={option}>{option}</option>)}</select></label>;
}

function StartYearField({ id, label, value, onChange }: { id: string; label: string; value: string; onChange: (value: string) => void }) {
  const valid = isValidStartYear(value);
  const errorId = `${id}-error`;
  return (
    <label>{label}
      <input
        id={id}
        type="text"
        inputMode="numeric"
        autoComplete="off"
        maxLength={4}
        pattern="[0-9]{4}"
        value={value}
        aria-invalid={!valid}
        aria-describedby={!valid ? errorId : undefined}
        onChange={(event) => onChange(event.target.value)}
      />
      {!valid && <small id={errorId} className="field-error">Usá 4 dígitos ({START_YEAR_MIN}–{START_YEAR_MAX}).</small>}
    </label>
  );
}

function boundedNumber(raw: string, min: number, max: number, fallback: number): number {
  const value = Number(raw);
  if (!Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, Math.round(value)));
}

const START_YEAR_MIN = 2006;
const START_YEAR_MAX = new Date().getFullYear();

export function isYearDraft(value: string): boolean {
  return /^\d{0,4}$/.test(value);
}

export function isValidStartYear(value: string): boolean {
  if (!/^\d{4}$/.test(value)) return false;
  const year = Number(value);
  return year >= START_YEAR_MIN && year <= START_YEAR_MAX;
}

function formatSavedDate(value?: string): string {
  if (!value) return 'sin fecha';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 'sin fecha' : date.toLocaleString('es-AR', { dateStyle: 'short', timeStyle: 'short' });
}

function companyNames(area: Pick<AreaCatalogItem, 'companies'>): string {
  return area.companies.filter(Boolean).join(', ') || 'Empresa no informada';
}

async function logDebugSafely(step: string, status: 'info' | 'ok' | 'warning' | 'error', detail: string): Promise<void> {
  if (typeof Excel === 'undefined') return;
  try {
    await appendDebug(createDebugEntry(step, status, detail));
  } catch {
    // El log nunca debe bloquear la operación principal.
  }
}
