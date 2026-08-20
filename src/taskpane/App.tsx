import React, { useEffect, useMemo, useState } from 'react';
import { createAssetId, normalizeAssetGroups } from '../domain/assets';
import { resolveAreaParams } from '../domain/forecast';
import { appendDebug, createDebugEntry } from '../excel/debugSheet';
import { captureSelectedCell, createNewDataSheetTarget } from '../excel/databaseSheet';
import { buildForecastWorkbook, downloadWorkbookData } from '../excel/workbookBuilder';
import { assetSheetName } from '../excel/names';
import { readSavedWorkbookPlans } from '../excel/workbookState';
import type {
  AreaCatalogItem,
  AreaForecastOverrideField,
  AreaForecastParams,
  AreaSelection,
  AreaWorkbookPlan,
  AssetGroup,
  BuildProgress,
  DataGranularity,
  DataOutputTarget,
  ForecastDefaults,
  MissingMonthsDecision,
  OverwriteWarning,
} from '../models/types';
import { fetchAreaCatalog, type CatalogLoadProgress } from '../services/capiv';

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

const MIXED = '__mixed__';
const METHOD_SHORT: Record<string, string> = {
  'Constante': 'Const',
  'HypMod': 'HypMod',
  'Declinación Hip.': 'Hip',
  'Declinación Exp.': 'Exp',
  'Rap Np': 'Rap Np',
  'RGP': 'RGP',
};
const USES_DI = new Set(['HypMod', 'Declinación Hip.', 'Declinación Exp.', 'Rap Np']);
const USES_B = new Set(['HypMod', 'Declinación Hip.', 'Rap Np']);

type Workflow = 'data' | 'forecast';
type DestinationMode = 'new-sheet' | 'selected-cell';
type StatusTone = 'neutral' | 'success' | 'warning' | 'error';
type MethodParamField = 'grossMethod' | 'oilMethod' | 'gasMethod';
type NumericParamField = 'grossDi' | 'grossB' | 'oilDi' | 'oilB' | 'gasDi' | 'gasB';

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
  const [provinceFilters, setProvinceFilters] = useState<string[]>([]);
  const [companyFilters, setCompanyFilters] = useState<string[]>([]);
  const [openFilter, setOpenFilter] = useState<'province' | 'company' | null>(null);
  const [query, setQuery] = useState('');
  const [dataSelected, setDataSelected] = useState<AreaSelection[]>([]);
  const [forecastSelected, setForecastSelected] = useState<AreaSelection[]>([]);
  const [params, setParams] = useState<Record<string, AreaForecastParams>>({});
  const [dirtyOverrideFields, setDirtyOverrideFields] = useState<Record<string, AreaForecastOverrideField[]>>({});
  const [excludedAreaIds, setExcludedAreaIds] = useState<string[]>([]);
  const [selectedParamAreaIds, setSelectedParamAreaIds] = useState<string[]>([]);
  const [forecastQuery, setForecastQuery] = useState('');
  const [numericDrafts, setNumericDrafts] = useState<Partial<Record<NumericParamField, string>>>({});
  const [assetGroups, setAssetGroups] = useState<AssetGroup[]>([]);
  const [assetFormOpen, setAssetFormOpen] = useState(false);
  const [assetNameDraft, setAssetNameDraft] = useState('');
  const [defaults, setDefaults] = useState<ForecastDefaults>(DEFAULTS);
  const [startYearDraft, setStartYearDraft] = useState(String(DEFAULTS.startYear));
  const [granularity, setGranularity] = useState<DataGranularity>('area');
  const [destinationMode, setDestinationMode] = useState<DestinationMode>('new-sheet');
  const [dataOutput, setDataOutput] = useState<DataOutputTarget | null>(null);
  const [forecastMode, setForecastMode] = useState<'update' | 'regenerate'>('update');
  const [savedInfo, setSavedInfo] = useState<SavedInfo | null>(null);
  const [catalogBusy, setCatalogBusy] = useState(false);
  const [catalogProgress, setCatalogProgress] = useState<CatalogLoadProgress | null>(null);
  const [catalogElapsed, setCatalogElapsed] = useState(0);
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

  const provinces = useMemo(
    () => [...new Set(catalog.map((item) => item.province).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es')),
    [catalog],
  );
  const companies = useMemo(
    () => [...new Set(catalog.flatMap((item) => item.companies).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es')),
    [catalog],
  );

  useEffect(() => {
    if (!catalog.length) return;
    setProvinceFilters((current) => current.filter((value) => provinces.includes(value)));
    setCompanyFilters((current) => current.filter((value) => companies.includes(value)));
  }, [catalog.length, provinces, companies]);

  const matchingAreas = useMemo(() => {
    const needle = query.trim().toLocaleUpperCase('es-AR');
    return catalog
      .filter((item) => provinceFilters.length === 0 || provinceFilters.includes(item.province))
      .filter((item) => companyFilters.length === 0 || item.companies.some((company) => companyFilters.includes(company)))
      .filter((item) => !needle || `${item.areaId} ${item.areaName} ${item.companies.join(' ')}`.toLocaleUpperCase('es-AR').includes(needle));
  }, [catalog, provinceFilters, companyFilters, query]);
  const visibleAreas = useMemo(() => matchingAreas.slice(0, 100), [matchingAreas]);
  const assignedAssetByArea = useMemo(() => new Map(assetGroups.flatMap((group) => group.areaIds.map((areaId) => [areaId, group] as const))), [assetGroups]);
  const areaGroups = useMemo(() => {
    const needle = forecastQuery.trim().toLocaleUpperCase('es-AR');
    const matches = (area: AreaSelection, groupName: string) => !needle
      || `${area.areaId} ${area.areaName} ${area.province} ${area.companies.join(' ')} ${groupName}`.toLocaleUpperCase('es-AR').includes(needle);
    const grouped = assetGroups.map((group) => ({
      id: group.id,
      name: group.name,
      areas: forecastSelected.filter((area) => group.areaIds.includes(area.areaId) && matches(area, group.name)),
    })).filter((group) => group.areas.length > 0);
    const unassigned = forecastSelected.filter((area) => !assignedAssetByArea.has(area.areaId) && matches(area, 'Sin activo'));
    if (unassigned.length > 0) grouped.push({ id: 'unassigned', name: 'Sin activo', areas: unassigned });
    return grouped;
  }, [assetGroups, assignedAssetByArea, forecastQuery, forecastSelected]);
  const includedAreas = useMemo(() => forecastSelected.filter((area) => !excludedAreaIds.includes(area.areaId)), [forecastSelected, excludedAreaIds]);
  const visibleIncludedAreaIds = useMemo(
    () => areaGroups.flatMap((group) => group.areas.map((area) => area.areaId)).filter((areaId) => !excludedAreaIds.includes(areaId)),
    [areaGroups, excludedAreaIds],
  );
  const selectedAreas = useMemo(() => includedAreas.filter((area) => selectedParamAreaIds.includes(area.areaId)), [includedAreas, selectedParamAreaIds]);
  const selectionHasAssigned = useMemo(() => selectedAreas.some((area) => assignedAssetByArea.has(area.areaId)), [selectedAreas, assignedAssetByArea]);
  const dirtyAreaCount = useMemo(
    () => forecastSelected.filter((area) => (dirtyOverrideFields[area.areaId]?.length ?? 0) > 0 && !excludedAreaIds.includes(area.areaId)).length,
    [forecastSelected, dirtyOverrideFields, excludedAreaIds],
  );

  useEffect(() => {
    void refreshCatalog();
  }, []);

  useEffect(() => {
    setStartYearDraft(String(defaults.startYear));
  }, [defaults.startYear]);

  useEffect(() => {
    if (!catalogBusy) return undefined;
    const startedAt = Date.now();
    setCatalogElapsed(0);
    const timer = window.setInterval(() => setCatalogElapsed(Math.floor((Date.now() - startedAt) / 1000)), 1000);
    return () => window.clearInterval(timer);
  }, [catalogBusy]);

  useEffect(() => {
    const availableIds = new Set(forecastSelected.map((area) => area.areaId));
    setSelectedParamAreaIds((current) => current.filter((areaId) => availableIds.has(areaId)));
    setExcludedAreaIds((current) => current.filter((areaId) => availableIds.has(areaId)));
  }, [forecastSelected]);

  useEffect(() => {
    setNumericDrafts({});
  }, [selectedParamAreaIds]);

  async function refreshCatalog() {
    setCatalogBusy(true);
    setCatalogProgress({ step: 1, message: 'Iniciando conexión con CapIV' });
    setStatus({ tone: 'neutral', text: 'Actualizando catálogo oficial…' });
    try {
      await logDebugSafely('Catálogo', 'info', 'Inicio de descarga del catálogo');
      const items = await fetchAreaCatalog((progress) => {
        setCatalogProgress(progress);
        setStatus({ tone: 'neutral', text: progress.message });
      });
      setCatalog(items);
      setStatus({ tone: 'success', text: `${items.length} áreas disponibles` });
      await logDebugSafely('Catálogo', 'ok', `${items.length} áreas disponibles`);
    } catch (error) {
      await showError('Catálogo', error);
    } finally {
      setCatalogBusy(false);
      setCatalogProgress(null);
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

  function markOverrideFields(areaIds: string[], fields: AreaForecastOverrideField[]) {
    if (fields.length === 0) return;
    setDirtyOverrideFields((current) => {
      const next = { ...current };
      for (const areaId of areaIds) next[areaId] = [...new Set([...(next[areaId] ?? []), ...fields])];
      return next;
    });
  }

  function selectionValue<K extends keyof AreaForecastParams>(key: K): AreaForecastParams[K] | undefined {
    let value: AreaForecastParams[K] | undefined;
    for (const area of selectedAreas) {
      const areaParams = params[area.areaId] ?? resolveAreaParams(defaults);
      if (value === undefined) value = areaParams[key];
      else if (value !== areaParams[key]) return undefined;
    }
    return value;
  }

  function applyParamsPatch(patch: Partial<AreaForecastParams>) {
    const targets = selectedAreas.map((area) => area.areaId);
    if (targets.length === 0) return;
    setParams((current) => {
      const next = { ...current };
      for (const areaId of targets) next[areaId] = { ...(next[areaId] ?? resolveAreaParams(defaults)), ...patch };
      return next;
    });
    markOverrideFields(targets, Object.keys(patch) as AreaForecastOverrideField[]);
    setStatus({ tone: 'success', text: `Cambio aplicado a ${targets.length} ${targets.length === 1 ? 'concesión' : 'concesiones'}. Se escribe al generar.` });
  }

  function changeSelectionMethod(field: MethodParamField, value: string) {
    if (!value || value === MIXED) return;
    applyParamsPatch({ [field]: value } as Partial<AreaForecastParams>);
  }

  function changeSelectionInitial(value: string) {
    if (!value || value === MIXED) return;
    applyParamsPatch({ takeInitialFromHistory: value === 'history' });
  }

  function commitNumericDraft(field: NumericParamField) {
    const raw = (numericDrafts[field] ?? '').trim();
    if (!raw) return;
    const value = Number(raw.replace(',', '.'));
    const isB = field.endsWith('B');
    if (!Number.isFinite(value) || value < 0 || (isB ? value > 2 || value === 0 : value > 5)) {
      setStatus({ tone: 'warning', text: isB ? 'Los valores b deben ser mayores que 0 y no superar 2.' : 'Las declinaciones Di deben estar entre 0 y 5.' });
      return;
    }
    setNumericDrafts((current) => {
      const next = { ...current };
      delete next[field];
      return next;
    });
    applyParamsPatch({ [field]: value } as Partial<AreaForecastParams>);
  }

  function toggleParamArea(areaId: string) {
    setSelectedParamAreaIds((current) => current.includes(areaId) ? current.filter((item) => item !== areaId) : [...current, areaId]);
  }

  function toggleParamGroup(areaIds: string[]) {
    setSelectedParamAreaIds((current) => areaIds.every((areaId) => current.includes(areaId))
      ? current.filter((areaId) => !areaIds.includes(areaId))
      : [...new Set([...current, ...areaIds])]);
  }

  function toggleExcludedArea(areaId: string) {
    setExcludedAreaIds((current) => current.includes(areaId) ? current.filter((item) => item !== areaId) : [...current, areaId]);
    setSelectedParamAreaIds((current) => current.filter((item) => item !== areaId));
  }

  function createAssetFromSelection() {
    const name = assetNameDraft.trim();
    if (!name) {
      setStatus({ tone: 'warning', text: 'Escribí un nombre para el activo.' });
      return;
    }
    if (assetGroups.some((group) => group.name.localeCompare(name, 'es', { sensitivity: 'base' }) === 0)) {
      setStatus({ tone: 'warning', text: `Ya existe un activo llamado ${name}.` });
      return;
    }
    if (assetGroups.some((group) => assetSheetName(group.name) === assetSheetName(name))) {
      setStatus({ tone: 'warning', text: 'Ese nombre produciría la misma hoja que otro activo. Elegí un nombre más distinto.' });
      return;
    }
    const areaIds = selectedAreas.map((area) => area.areaId);
    if (areaIds.length === 0) {
      setStatus({ tone: 'warning', text: 'Marcá al menos una concesión para agrupar.' });
      return;
    }
    setAssetGroups((current) => normalizeAssetGroups([
      ...current.map((group) => ({ ...group, areaIds: group.areaIds.filter((areaId) => !areaIds.includes(areaId)) })),
      { id: createAssetId(name, current), name, areaIds },
    ]));
    setAssetNameDraft('');
    setAssetFormOpen(false);
    setStatus({ tone: 'success', text: `Activo ${name} creado con ${areaIds.length} ${areaIds.length === 1 ? 'concesión' : 'concesiones'}.` });
  }

  function removeSelectionFromAssets() {
    const ids = new Set(selectedAreas.map((area) => area.areaId));
    setAssetGroups((current) => normalizeAssetGroups(current.map((group) => ({
      ...group,
      areaIds: group.areaIds.filter((areaId) => !ids.has(areaId)),
    }))));
    setStatus({ tone: 'success', text: `${ids.size} ${ids.size === 1 ? 'concesión movida' : 'concesiones movidas'} a Sin activo.` });
  }

  function removeAssetGroup(groupId: string) {
    setAssetGroups((current) => current.filter((group) => group.id !== groupId));
  }

  function dataPlans(selections: AreaSelection[], mode: 'update' | 'regenerate'): AreaWorkbookPlan[] {
    return selections.map((selection) => ({ selection, defaults, mode }));
  }

  function forecastPlans(): AreaWorkbookPlan[] {
    return includedAreas.map((selection) => ({
      selection,
      defaults,
      override: { areaId: selection.areaId, ...(params[selection.areaId] ?? resolveAreaParams(defaults)) },
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
      setParams((current) => {
        const next = { ...current };
        for (const area of dataSelected) if (!next[area.areaId]) next[area.areaId] = resolveAreaParams(defaults);
        return next;
      });
      setAssetGroups((current) => normalizeAssetGroups(current, dataSelected.map((area) => area.areaId)));
      setSavedInfo({ areaCount: plans.length, dataSavedAt: new Date().toISOString() });
      setStatus({ tone: 'success', text: `Tabla ${granularity === 'area' ? 'mensual por área' : 'pozo-mes'} generada en ${output.sheetName}!${output.startAddress}. Las áreas quedaron listas en la pestaña Pronósticos.` });
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
      setParams(Object.fromEntries(plans.map((plan) => [plan.selection.areaId, resolveAreaParams(plan.defaults, plan.override)])));
      setAssetGroups(normalizeAssetGroups(saved.assetGroups, plans.map((plan) => plan.selection.areaId)));
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
        setAssetGroups([]);
        setSavedInfo(null);
        setStatus({ tone: 'warning', text: 'No hay datos descargados. Completá primero el flujo Datos.' });
        return;
      }
      const availableIds = new Set(saved.data.map((item) => item.areaId));
      const availablePlans = saved.plans.filter((plan) => availableIds.has(plan.selection.areaId));
      setForecastSelected(availablePlans.map((plan) => plan.selection));
      if (availablePlans[0]) setDefaults(availablePlans[0].defaults);
      setParams(Object.fromEntries(availablePlans.map((plan) => [plan.selection.areaId, resolveAreaParams(plan.defaults, plan.override)])));
      setDirtyOverrideFields({});
      setExcludedAreaIds([]);
      setSelectedParamAreaIds([]);
      setAssetGroups(normalizeAssetGroups(saved.assetGroups, availablePlans.map((plan) => plan.selection.areaId)));
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
    if (includedAreas.length === 0) {
      setStatus({ tone: 'warning', text: 'Cargá datos del libro e incluí al menos una concesión antes de pronosticar.' });
      return;
    }
    if (!ensureExcel()) return;
    setBuildBusy(true);
    setProgress({ completed: 0, total: 1, percent: 0, message: 'Preparando pronósticos' });
    setStatus({ tone: 'neutral', text: 'Preparando pronósticos…' });
    try {
      const plans = forecastPlans();
      await buildForecastWorkbook(plans, reportProgress, { assetGroups, applyOverrideFields: dirtyOverrideFields });
      setDirtyOverrideFields({});
      setSavedInfo((current) => ({ areaCount: plans.length, dataSavedAt: current?.dataSavedAt, forecastSavedAt: new Date().toISOString() }));
      setStatus({ tone: 'success', text: `${plans.length} ${plans.length === 1 ? 'pronóstico generado' : 'pronósticos generados'}${assetGroups.length ? ` y ${assetGroups.length} ${assetGroups.length === 1 ? 'activo resumido' : 'activos resumidos'}` : ''}.` });
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

  function renderMethodRow(
    label: string,
    methodField: MethodParamField,
    options: readonly string[],
    diField: NumericParamField,
    bField: NumericParamField,
  ) {
    const method = selectionValue(methodField);
    const diValue = selectionValue(diField);
    const bValue = selectionValue(bField);
    const diDisabled = buildBusy || (method !== undefined && !USES_DI.has(method));
    const bDisabled = buildBusy || (method !== undefined && !USES_B.has(method));
    return (
      <React.Fragment key={methodField}>
        <span className="editor-row-label">{label}</span>
        <select value={method ?? MIXED} disabled={buildBusy} onChange={(event) => changeSelectionMethod(methodField, event.target.value)} aria-label={`Método ${label}`}>
          {method === undefined && <option value={MIXED}>— Varios —</option>}
          {options.map((option) => <option key={option}>{option}</option>)}
        </select>
        <input
          inputMode="decimal"
          value={numericDrafts[diField] ?? (diValue !== undefined ? formatDecimal(diValue) : '')}
          placeholder={diValue === undefined ? 'varios' : '0,12'}
          disabled={diDisabled}
          onChange={(event) => setNumericDrafts((current) => ({ ...current, [diField]: event.target.value }))}
          onBlur={() => commitNumericDraft(diField)}
          onKeyDown={(event) => { if (event.key === 'Enter') commitNumericDraft(diField); }}
          aria-label={`Di ${label}`}
        />
        <input
          inputMode="decimal"
          value={numericDrafts[bField] ?? (bValue !== undefined ? formatDecimal(bValue) : '')}
          placeholder={bValue === undefined ? 'varios' : '0,70'}
          disabled={bDisabled}
          onChange={(event) => setNumericDrafts((current) => ({ ...current, [bField]: event.target.value }))}
          onBlur={() => commitNumericDraft(bField)}
          onKeyDown={(event) => { if (event.key === 'Enter') commitNumericDraft(bField); }}
          aria-label={`b ${label}`}
        />
      </React.Fragment>
    );
  }

  function renderFilterGroup(
    key: 'province' | 'company',
    label: string,
    options: string[],
    selected: string[],
    setSelected: React.Dispatch<React.SetStateAction<string[]>>,
  ) {
    const open = openFilter === key;
    return (
      <div className="filter-group">
        <button
          type="button"
          className={open ? 'filter-toggle open' : 'filter-toggle'}
          onClick={() => setOpenFilter(open ? null : key)}
          disabled={!catalog.length || busy}
          aria-expanded={open}
        >
          <span>{label}</span>
          <small>{selected.length === 0 ? 'Todas' : selected.length === 1 ? '1 elegida' : `${selected.length} elegidas`}</small>
        </button>
        {open && (
          <div className="filter-list" role="group" aria-label={`Filtrar por ${label.toLocaleLowerCase('es-AR')}`}>
            {selected.length > 0 && (
              <button type="button" className="filter-clear" onClick={() => setSelected([])}>Limpiar filtro ({selected.length})</button>
            )}
            {options.map((option) => {
              const checked = selected.includes(option);
              return (
                <label key={option} className="filter-option" title={option}>
                  <input
                    type="checkbox"
                    checked={checked}
                    disabled={busy}
                    onChange={() => setSelected((current) => (checked ? current.filter((value) => value !== option) : [...current, option]))}
                  />
                  <span>{option}</span>
                </label>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  const initialUniform = selectionValue('takeInitialFromHistory');

  return (
    <main className="app-shell">
      <header className="topbar">
        <img src="assets/branding/logo_isotipo.png" alt="Quintana Energy" />
        <div className="brand-copy">
          <div className="product-line"><h1>CapIV</h1><span>v0.6.0</span></div>
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
                <small>{catalogBusy ? `${catalogElapsed} s` : catalog.length ? `${catalog.length} áreas` : 'Sin datos locales'}</small>
              </button>
              {catalogBusy && (
                <div className="catalog-progress" role="status" aria-live="polite">
                  <div className="catalog-progress-copy">
                    <span className="catalog-pulse" aria-hidden="true" />
                    <div><strong>{catalogProgress?.message ?? 'Actualizando catálogo oficial'}</strong><small>Seguimos trabajando · {catalogElapsed} s transcurridos</small></div>
                  </div>
                  <div className="catalog-progress-track" aria-hidden="true"><span /></div>
                  <div className="catalog-progress-steps" aria-hidden="true">
                    {['Conectar', 'Recibir', 'Preparar'].map((label, index) => {
                      const step = index + 1;
                      return <span key={label} className={step < (catalogProgress?.step ?? 1) ? 'done' : step === (catalogProgress?.step ?? 1) ? 'active' : ''}>{step < (catalogProgress?.step ?? 1) ? '✓' : step} {label}</span>;
                    })}
                  </div>
                </div>
              )}
              <div className="field-grid data-filters">
                <label>Buscar<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Área, código o empresa" disabled={!catalog.length || busy} /></label>
                <div className="filter-groups">
                  {renderFilterGroup('province', 'Provincias', provinces, provinceFilters, setProvinceFilters)}
                  {renderFilterGroup('company', 'Empresas', companies, companyFilters, setCompanyFilters)}
                </div>
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
              {destinationMode === 'selected-cell' && (
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
            <section className="panel">
              <SectionHeading
                step="1"
                title="Configurá las concesiones"
                description={forecastSelected.length ? `${includedAreas.length} de ${forecastSelected.length} se van a pronosticar` : 'Cada concesión tiene su propio pronóstico'}
              />
              {forecastSelected.length === 0 ? (
                <div className="empty-state compact">No hay datos cargados. Volvé al flujo Datos para crear la base en este libro.</div>
              ) : (
                <>
                  <div className="param-toolbar">
                    <label>Buscar concesión<input value={forecastQuery} onChange={(event) => setForecastQuery(event.target.value)} placeholder="Código, nombre, provincia o activo" /></label>
                    <div className="param-shortcuts">
                      <button type="button" onClick={() => setSelectedParamAreaIds(includedAreas.map((area) => area.areaId))} disabled={buildBusy}>Todas</button>
                      <button type="button" onClick={() => setSelectedParamAreaIds([])} disabled={buildBusy}>Ninguna</button>
                      <button type="button" onClick={() => setSelectedParamAreaIds(visibleIncludedAreaIds)} disabled={buildBusy || visibleIncludedAreaIds.length === 0}>Solo visibles</button>
                    </div>
                  </div>
                  <div className="param-groups">
                    {areaGroups.map((group) => {
                      const selectableIds = group.areas.filter((area) => !excludedAreaIds.includes(area.areaId)).map((area) => area.areaId);
                      const selectedCount = selectableIds.filter((areaId) => selectedParamAreaIds.includes(areaId)).length;
                      return (
                        <section className="param-group" key={group.id}>
                          <div className="param-group-header">
                            <label className="param-group-toggle">
                              <input
                                type="checkbox"
                                checked={selectableIds.length > 0 && selectedCount === selectableIds.length}
                                disabled={buildBusy || selectableIds.length === 0}
                                ref={(element) => { if (element) element.indeterminate = selectedCount > 0 && selectedCount < selectableIds.length; }}
                                onChange={() => toggleParamGroup(selectableIds)}
                              />
                              <strong>{group.name}</strong>
                              <small>{selectedCount}/{selectableIds.length}</small>
                            </label>
                            {group.id !== 'unassigned' && <button type="button" className="text-button" onClick={() => removeAssetGroup(group.id)} disabled={buildBusy}>Desagrupar</button>}
                          </div>
                          {group.areas.map((area) => {
                            const excluded = excludedAreaIds.includes(area.areaId);
                            const dirty = !excluded && (dirtyOverrideFields[area.areaId]?.length ?? 0) > 0;
                            return (
                              <div className={excluded ? 'param-row excluded' : 'param-row'} key={area.areaId}>
                                <label className="param-row-main" title={`${area.areaId} · ${area.areaName} · ${area.province}`}>
                                  <input
                                    type="checkbox"
                                    checked={selectedParamAreaIds.includes(area.areaId)}
                                    disabled={buildBusy || excluded}
                                    onChange={() => toggleParamArea(area.areaId)}
                                  />
                                  <span className="param-copy">
                                    <span className="param-name"><strong>{area.areaId}</strong><small>{area.areaName}</small></span>
                                    <span className="param-chips">
                                      {excluded ? 'Excluida: no se pronostica' : paramSummary(params[area.areaId])}
                                      {dirty && <i className="dirty-dot" title="Cambios pendientes: se escriben en Excel al generar" />}
                                    </span>
                                  </span>
                                </label>
                                <button
                                  type="button"
                                  className="icon-button"
                                  onClick={() => toggleExcludedArea(area.areaId)}
                                  disabled={buildBusy}
                                  title={excluded ? 'Volver a incluir en el pronóstico' : 'Excluir del pronóstico'}
                                  aria-label={excluded ? `Volver a incluir ${area.areaName}` : `Excluir ${area.areaName}`}
                                >{excluded ? '+' : '×'}</button>
                              </div>
                            );
                          })}
                        </section>
                      );
                    })}
                    {areaGroups.length === 0 && <div className="empty-state compact">No hay concesiones que coincidan con la búsqueda.</div>}
                  </div>
                  {dirtyAreaCount > 0 && (
                    <p className="helper-text dirty-note"><i className="dirty-dot" /> {dirtyAreaCount} {dirtyAreaCount === 1 ? 'concesión tiene' : 'concesiones tienen'} cambios que se escriben en Excel al generar.</p>
                  )}

                  <div className="selection-editor">
                    <div className="selection-editor-heading">
                      <strong>{selectedAreas.length ? `Aplicar a ${selectedAreas.length} ${selectedAreas.length === 1 ? 'seleccionada' : 'seleccionadas'}` : 'Editor de parámetros'}</strong>
                      <span>{selectedAreas.length ? 'Cada cambio queda aplicado al instante; en Excel se escribe al generar.' : 'Marcá una o más concesiones para editar métodos y declinaciones.'}</span>
                    </div>
                    {selectedAreas.length > 0 && (
                      <>
                        <div className="editor-grid">
                          <span className="editor-head" aria-hidden="true" />
                          <span className="editor-head">Método</span>
                          <span className="editor-head"><HelpLabel label="Di" help="Declinación inicial anual, como fracción: 0,12 equivale a 12% por año. Usada por los métodos de declinación." /></span>
                          <span className="editor-head"><HelpLabel label="b" help="Exponente de curvatura de la declinación hiperbólica: mayor que 0 y hasta 2. Un b mayor suaviza la caída inicial." /></span>
                          {renderMethodRow('Bruta', 'grossMethod', GROSS_METHODS, 'grossDi', 'grossB')}
                          {renderMethodRow('Petróleo', 'oilMethod', OIL_METHODS, 'oilDi', 'oilB')}
                          {renderMethodRow('Gas', 'gasMethod', GAS_METHODS, 'gasDi', 'gasB')}
                        </div>
                        <label className="editor-initial">Valor inicial
                          <select
                            value={initialUniform === undefined ? MIXED : initialUniform ? 'history' : 'manual'}
                            disabled={buildBusy}
                            onChange={(event) => changeSelectionInitial(event.target.value)}
                          >
                            {initialUniform === undefined && <option value={MIXED}>— Varios —</option>}
                            <option value="history">Desde histórico</option>
                            <option value="manual">Manual en Excel</option>
                          </select>
                        </label>
                        <div className="selection-actions">
                          <button type="button" onClick={() => setAssetFormOpen((open) => !open)} disabled={buildBusy}>Agrupar en activo…</button>
                          {selectionHasAssigned && <button type="button" onClick={removeSelectionFromAssets} disabled={buildBusy}>Sacar del activo</button>}
                        </div>
                        {assetFormOpen && (
                          <div className="asset-name-form">
                            <label>Nombre del activo<input value={assetNameDraft} onChange={(event) => setAssetNameDraft(event.target.value)} placeholder="Ej. Mendoza, CLME o EFO" /></label>
                            <button type="button" className="asset-create" onClick={createAssetFromSelection}>Crear activo ({selectedAreas.length})</button>
                            <p className="helper-text">Cada concesión mantiene su pronóstico. El activo agrega una hoja con la suma y cuatro gráficos propios.</p>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                </>
              )}
            </section>

            <section className="panel">
              <SectionHeading step="2" title="Generación" description="Escribe en este libro; no descarga datos" />
              <div className="field-grid two-columns">
                <label>Horizonte (años)<input type="number" min="1" max="40" value={defaults.horizonYears} onChange={(event) => setDefaults({ ...defaults, horizonYears: boundedNumber(event.target.value, 1, 40, defaults.horizonYears) })} /></label>
                <label>Datos<input value={savedInfo ? `${savedInfo.areaCount} áreas` : 'Sin cargar'} disabled /></label>
              </div>
              <ModeSelector mode={forecastMode} onChange={setForecastMode} updateText="Conserva los supuestos editados en Prono y Pozos." regenerateText="Reconstruye pronósticos, gráficos y resumen desde cero." />
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
          <button type="button" className="primary forecast-action" disabled={busy || includedAreas.length === 0} onClick={runForecast}>{buildBusy ? 'Generando…' : `Generar pronósticos ${includedAreas.length ? `(${includedAreas.length})` : ''}`}</button>
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

function HelpLabel({ label, help }: { label: string; help: string }) {
  return <span className="field-label">{label}<span className="info-tooltip" tabIndex={0} role="note" aria-label={`${label}: ${help}`} data-tooltip={help} title={help}>?</span></span>;
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

export function paramSummary(areaParams?: AreaForecastParams): string {
  if (!areaParams) return 'Parámetros por definir';
  const chips = [
    methodChip('B', areaParams.grossMethod, areaParams.grossDi, areaParams.grossB),
    methodChip('P', areaParams.oilMethod, areaParams.oilDi, areaParams.oilB),
    methodChip('G', areaParams.gasMethod, areaParams.gasDi, areaParams.gasB),
  ];
  if (!areaParams.takeInitialFromHistory) chips.push('inicial manual');
  return chips.join(' · ');
}

function methodChip(prefix: string, method: string, di: number, b: number): string {
  let chip = `${prefix} ${METHOD_SHORT[method] ?? method}`;
  if (USES_DI.has(method)) chip += ` ${formatDecimal(di)}`;
  if (USES_B.has(method)) chip += `/${formatDecimal(b)}`;
  return chip;
}

function formatDecimal(value: number): string {
  return value.toLocaleString('es-AR', { maximumFractionDigits: 4 });
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
