export function safeSheetPrefix(areaId: string): string {
  return areaId.replace(/[^A-Za-z0-9]/g, '').slice(0, 12) || 'AREA';
}

export function areaSheetNames(areaId: string) {
  const prefix = safeSheetPrefix(areaId);
  return {
    hdp: `${prefix}_HDP`,
    prono: `${prefix}_Prono`,
    pozos: `${prefix}_Pozos`,
    graficos: `${prefix}_Graficos`,
    detalle: `${prefix}_Detalle`,
  };
}

export const SUMMARY_SHEET = 'Resumen_Areas';
export const STATE_SHEET = '_CapIV_State';
