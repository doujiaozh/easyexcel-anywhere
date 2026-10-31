/**
 * Header->property binding for mapped reading, plus value conversion.
 */

import {
  ensureSchema,
  type CellType,
  type PropertyMeta,
} from '../annotations/ExcelProperty.js';

export interface ReadBinding {
  col: number;
  key: string;
  convert: (raw: string | null) => unknown;
}

function convertByType(raw: string | null, type?: CellType, key?: string): unknown {
  if (raw === null || raw === '') return null;
  switch (type) {
    case 'number': {
      const n = Number(raw);
      return Number.isFinite(n) ? n : raw;
    }
    case 'boolean':
      return raw === 'TRUE' || raw === 'true' || raw === '1';
    case 'date': {
      const d = new Date(raw);
      return Number.isNaN(d.getTime()) ? raw : d;
    }
    case 'string':
      return raw;
    default:
      return raw;
  }
}

export function buildReadBindings(
  cls: new (...args: unknown[]) => unknown,
  headerRow: (string | null)[]
): ReadBinding[] {
  const metas = ensureSchema(cls);
  if (metas.length === 0) {
    throw new Error(`No @ExcelProperty decorated properties found on ${cls.name}`);
  }
  const byHeader = new Map<string, number>();
  headerRow.forEach((h, i) => {
    if (h != null) byHeader.set(String(h).trim(), i);
  });

  const bindings: ReadBinding[] = [];
  const missing: string[] = [];
  for (const m of metas) {
    let col = byHeader.get(m.header.trim());
    if (col === undefined && typeof m.index === 'number') {
      col = m.index;
    }
    if (col === undefined) {
      missing.push(`${m.key} (header "${m.header}")`);
      continue;
    }
    bindings.push({
      col,
      key: m.key,
      convert: m.converter?.fromCell
        ? (raw) => m.converter!.fromCell!(raw)
        : (raw) => convertByType(raw, m.type, m.key),
    });
  }
  if (missing.length > 0) {
    throw new Error(`Header row does not match class ${cls.name}; missing columns for: ${missing.join(', ')}`);
  }
  return bindings;
}

export function instantiate(cls: new (...args: unknown[]) => unknown): Record<string, unknown> {
  try {
    return new (cls as new (...args: unknown[]) => Record<string, unknown>)();
  } catch {
    return Object.create(cls.prototype) as Record<string, unknown>;
  }
}

export interface WriteBinding {
  meta: PropertyMeta;
}

export function buildWriteBindings(cls: new (...args: unknown[]) => unknown): PropertyMeta[] {
  const metas = ensureSchema(cls);
  if (metas.length === 0) {
    throw new Error(`No @ExcelProperty decorated properties found on ${cls.name}`);
  }
  return metas;
}

/** Property value -> cell value for the XML writer. */
export function toCellValue(value: unknown, meta: PropertyMeta): string | number | boolean | null {
  if (value === null || value === undefined) return null;
  if (meta.converter?.toCell) {
    const v = meta.converter.toCell(value);
    return v === undefined ? null : v;
  }
  switch (typeof value) {
    case 'number':
    case 'bigint':
    case 'boolean':
      return value as number | boolean;
    case 'string':
      return value;
    case 'object':
      if (value instanceof Date) return value.toISOString();
      return String(value);
    default:
      return String(value);
  }
}
