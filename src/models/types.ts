export type ForecastMethod =
  | 'Constante'
  | 'HypMod'
  | 'Declinación Hip.'
  | 'Declinación Exp.'
  | 'Rap Np'
  | 'RGP';

export interface AreaCatalogItem {
  province: string;
  areaId: string;
  areaName: string;
  basin?: string;
  companies: string[];
}

export interface AreaSelection extends AreaCatalogItem {
  startYearOverride?: number;
}

export interface ForecastDefaults {
  startYear: number;
  horizonYears: number;
  grossMethod: Extract<ForecastMethod, 'Constante' | 'HypMod' | 'Declinación Hip.' | 'Declinación Exp.'>;
  oilMethod: Extract<ForecastMethod, 'Constante' | 'HypMod' | 'Declinación Hip.' | 'Declinación Exp.' | 'Rap Np'>;
  gasMethod: Extract<ForecastMethod, 'Constante' | 'HypMod' | 'Declinación Hip.' | 'Declinación Exp.' | 'RGP'>;
  takeInitialFromHistory: boolean;
}

export interface AreaForecastOverride {
  areaId: string;
  startYear?: number;
  grossMethod?: ForecastDefaults['grossMethod'];
  oilMethod?: ForecastDefaults['oilMethod'];
  gasMethod?: ForecastDefaults['gasMethod'];
  takeInitialFromHistory?: boolean;
}

export interface ProductionRecord {
  areaId: string;
  areaName: string;
  wellId: string;
  wellName: string;
  year: number;
  month: number;
  oil: number;
  gas: number;
  water: number;
  waterInjection: number;
  raw: Record<string, unknown>;
}

export type CapituloIvSource = 'convencional' | 'no convencional';

export type CapituloIvDownloadEvent =
  | {
      type: 'resource_started';
      areaId: string;
      source: CapituloIvSource;
      year?: number;
    }
  | {
      type: 'resource_completed';
      areaId: string;
      source: CapituloIvSource;
      year?: number;
      scannedRows: number;
      matchedRows: number;
      records: ProductionRecord[];
    }
  | {
      type: 'completed';
      areaId: string;
      records: ProductionRecord[];
    };

export type CapituloIvDownloadEventHandler = (event: CapituloIvDownloadEvent) => Promise<void> | void;

export interface MonthlyAggregate {
  date: string;
  year: number;
  month: number;
  oil: number;
  gas: number;
  water: number;
  gross: number;
  waterInjection: number;
  oilWells: number;
  gasWells: number;
  injectorWells: number;
  missing: boolean;
  missingKind: 'none' | 'leading' | 'middle';
}

export interface AreaWorkbookPlan {
  selection: AreaSelection;
  defaults: ForecastDefaults;
  override?: AreaForecastOverride;
  mode: 'update' | 'regenerate';
}

export interface BuildProgress {
  areaId?: string;
  areaIndex?: number;
  areaTotal?: number;
  completed: number;
  total: number;
  percent: number;
  message: string;
}

export type BuildProgressHandler = (progress: BuildProgress) => void;

export interface MissingMonthsDecision {
  areaId: string;
  areaName: string;
  months: string[];
}

export type MissingMonthsDecisionHandler = (decision: MissingMonthsDecision) => Promise<'blank' | 'zero'>;

export interface DebugEntry {
  timestamp: string;
  step: string;
  status: 'info' | 'ok' | 'warning' | 'error';
  detail: string;
}
