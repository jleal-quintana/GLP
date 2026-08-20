import type { AssetGroup } from '../models/types';

export function createAssetId(name: string, existing: AssetGroup[]): string {
  const base = name
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('es-AR')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '') || 'activo';
  const occupied = new Set(existing.map((group) => group.id));
  if (!occupied.has(base)) return base;
  let suffix = 2;
  while (occupied.has(`${base}-${suffix}`)) suffix++;
  return `${base}-${suffix}`;
}

export function normalizeAssetGroups(groups: AssetGroup[], availableAreaIds?: Iterable<string>): AssetGroup[] {
  const available = availableAreaIds ? new Set(availableAreaIds) : undefined;
  const assigned = new Set<string>();
  const ids = new Set<string>();
  const output: AssetGroup[] = [];
  for (const group of groups) {
    const name = group.name.trim();
    if (!name || !group.id || ids.has(group.id)) continue;
    const areaIds = group.areaIds.filter((areaId) => {
      if (!areaId || assigned.has(areaId) || (available && !available.has(areaId))) return false;
      assigned.add(areaId);
      return true;
    });
    if (areaIds.length === 0) continue;
    ids.add(group.id);
    output.push({ id: group.id, name, areaIds });
  }
  return output;
}
