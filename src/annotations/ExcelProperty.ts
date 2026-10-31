/**
 * Annotation-driven row mapping, EasyExcel style.
 *
 * @example
 * ```ts
 * class User {
 *   @ExcelProperty('姓名')
 *   name = '';
 *
 *   @ExcelProperty('年龄', { type: 'number' })
 *   age = 0;
 * }
 * ```
 *
 * Property decorators execute in declaration order while the class is
 * being defined, so the registry naturally captures column order.
 */

export type CellType = 'string' | 'number' | 'boolean' | 'date';

/** Custom conversion hooks, for dates, enums or anything exotic. */
export interface ValueConverter {
  /** cell text -> property value */
  fromCell?: (raw: string | null) => unknown;
  /** property value -> cell value (string | number | boolean | null) */
  toCell?: (value: unknown) => string | number | boolean | null;
}

export interface ExcelPropertyOptions {
  /** Column header (falls back to the property name). */
  header?: string;
  /** Explicit 0-based column index (overrides declaration order). */
  index?: number;
  /** Rough column width in characters (informational, kept for parity with EasyExcel). */
  width?: number;
  /** How cell text converts to a property value. Default: kept as text. */
  type?: CellType;
  /** Custom converters take precedence over `type`. */
  converter?: ValueConverter;
}

export interface PropertyMeta extends Required<Omit<ExcelPropertyOptions, 'index' | 'type' | 'converter'>> {
  key: string;
  header: string;
  width: number;
  index?: number;
  type?: CellType;
  converter?: ValueConverter;
}

const REGISTRY = new Map<Function, PropertyMeta[]>();
const IGNORED = new Map<Function, Set<string>>();

/**
 * Fall back to the TypeScript `design:type` metadata for properties whose
 * @ExcelProperty did not state a `type` explicitly. Requires
 * `emitDecoratorMetadata` and (for the metadata to actually be recorded) the
 * `reflect-metadata` polyfill, which this library imports for you.
 *
 * Note: TS emits `design:type: Object` for properties that have only an
 * initializer (`active = false`) and no explicit annotation, so a second
 * fallback infers the type from the property's default value on a freshly
 * constructed instance. Either way, when nothing can be inferred the value
 * keeps its raw text form - the mapping still works via headers alone.
 */
function designCellType(cls: Function, key: string): CellType | undefined {
  const reflect = globalThis as { Reflect?: { getMetadata?: (k: string, t: object, p: string | symbol) => unknown } };
  const getMetadata = reflect.Reflect?.getMetadata;
  if (typeof getMetadata === 'function') {
    try {
      const t = getMetadata.call(reflect.Reflect, 'design:type', (cls as { prototype: object }).prototype, key);
      switch (t) {
        case String:
          return 'string';
        case Number:
          return 'number';
        case Boolean:
          return 'boolean';
        case Date:
          return 'date';
        default:
          break; // Object / unknown -> try the default-value probe below
      }
    } catch {
      /* fall through to the probe */
    }
  }
  try {
    const probe = new (cls as new () => Record<string, unknown>)();
    const v = probe[key];
    switch (typeof v) {
      case 'string':
        return 'string';
      case 'number':
        return 'number';
      case 'boolean':
        return 'boolean';
      case 'object':
        return v instanceof Date ? 'date' : undefined;
      default:
        return undefined;
    }
  } catch {
    return undefined;
  }
}

/**
 * Mark a property as an Excel column.
 * Usage: `@ExcelProperty('姓名')` or `@ExcelProperty({ header: '年龄', type: 'number' })`.
 */
export function ExcelProperty(headerOrOptions?: string | ExcelPropertyOptions, opts?: ExcelPropertyOptions) {
  const options: ExcelPropertyOptions =
    typeof headerOrOptions === 'string'
      ? { ...opts, header: headerOrOptions }
      : (headerOrOptions ?? {});

  return (target: object, propertyKey: string | symbol): void => {
    const ctor = (target as { constructor: Function }).constructor;
    let list = REGISTRY.get(ctor);
    if (!list) {
      list = [];
      REGISTRY.set(ctor, list);
    }
    const meta: PropertyMeta = {
      key: String(propertyKey),
      header: options.header ?? String(propertyKey),
      width: options.width ?? 0,
      index: options.index,
      type: options.type,
      converter: options.converter,
    };
    list.push(meta);
  };
}

/** Skip a property when reading/writing. */
export function ExcelIgnore() {
  return (target: object, propertyKey: string | symbol): void => {
    const ctor = (target as { constructor: Function }).constructor;
    let set = IGNORED.get(ctor);
    if (!set) {
      set = new Set();
      IGNORED.set(ctor, set);
    }
    set.add(String(propertyKey));
  };
}

function inheritanceChain(cls: Function): Function[] {
  const chain: Function[] = [];
  let c: Function | null = cls;
  while (c && c !== Object.prototype && c !== Function.prototype) {
    chain.unshift(c);
    c = Object.getPrototypeOf(c);
  }
  return chain;
}

/** Collect the effective schema of a class, base classes included. */
export function ensureSchema(cls: new (...args: unknown[]) => unknown): PropertyMeta[] {
  const merged = new Map<string, PropertyMeta>();
  for (const ctor of inheritanceChain(cls)) {
    for (const meta of REGISTRY.get(ctor) ?? []) {
      merged.set(meta.key, meta);
    }
  }
  const ignoredKeys = new Set<string>();
  for (const ctor of inheritanceChain(cls)) {
    for (const k of IGNORED.get(ctor) ?? []) ignoredKeys.add(k);
  }
  return [...merged.values()]
    .filter((m) => !ignoredKeys.has(m.key))
    .map((m) => (m.type || m.converter ? m : { ...m, type: designCellType(cls, m.key) }))
    .sort((a, b) => (a.index ?? Number.MAX_SAFE_INTEGER) - (b.index ?? Number.MAX_SAFE_INTEGER));
}
