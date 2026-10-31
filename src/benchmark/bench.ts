/**
 * ee_cli_benchmark: the million-row gauntlet.
 *
 * Pipeline: generate(1M rows) -> write xlsx -> read xlsx -> map to class ->
 * transform -> write xlsx -> verify. Every phase reports wall time and
 * peak process heap, proving the library stays flat on an 8GB Win7 box.
 */

import fs from 'node:fs';
import { Readable } from 'node:stream';
import path from 'node:path';
import { ExcelProperty } from '../annotations/ExcelProperty.js';
import { EasyExcel } from '../index.js';

/** The row shape used by the benchmark. */
export class DemoRow {
  @ExcelProperty('ID', { type: 'number' })
  id = 0;

  @ExcelProperty('姓名')
  name = '';

  @ExcelProperty('年龄', { type: 'number' })
  age = 0;

  @ExcelProperty('城市')
  city = '';

  @ExcelProperty('评分', { type: 'number' })
  score = 0;

  @ExcelProperty('在职') // boolean round-trip (t="b")
  active = false;

  @ExcelProperty('备注')
  note = '';

  @ExcelProperty('日期')
  date = '';
}

export interface PhaseMetrics {
  name: string;
  ms: number;
  peakHeapMB: number;
  rows?: number;
}

export interface BenchResult {
  rows: number;
  phases: PhaseMetrics[];
  sizes: { input: number; output: number };
}

class HeapSampler {
  private timer: ReturnType<typeof setInterval> | null = null;
  private peak = 0;

  start(): void {
    this.peak = process.memoryUsage().heapUsed;
    this.timer = setInterval(() => {
      const h = process.memoryUsage().heapUsed;
      if (h > this.peak) this.peak = h;
    }, 50);
  }

  stop(): number {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
    }
    const final = process.memoryUsage().heapUsed;
    return Math.max(this.peak, final) / 1024 / 1024;
  }
}

function timed<T>(name: string, work: () => Promise<T>): Promise<PhaseMetrics & { value: T }> {
  const sampler = new HeapSampler();
  const t0 = performance.now();
  sampler.start();
  return Promise.resolve(work()).then((value) => {
    const ms = performance.now() - t0;
    const peakHeapMB = sampler.stop();
    return { name, ms, peakHeapMB, value };
  });
}

// --- deterministic data -----------------------------------------------------

function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const NAMES = ['张伟', '王芳', '李娜', '刘洋', '陈静', '杨帆', '赵磊', '孙丽'];
const CITIES = ['北京', '上海', '广州', '深圳', '杭州', '成都', '武汉', '西安', '南京', '重庆'];
const NOTES = ['一手客户', '待跟进', '已签约', '复购', '流失风险', 'VIP', '线下门店', '转介绍'];

async function* generateRows(total: number): AsyncGenerator<DemoRow> {
  const rnd = mulberry32(20261031);
  for (let i = 1; i <= total; i++) {
    const row = new DemoRow();
    row.id = i;
    row.name = NAMES[(rnd() * NAMES.length) | 0] + '_' + i;
    row.age = 18 + ((rnd() * 48) | 0);
    row.city = CITIES[(rnd() * CITIES.length) | 0];
    row.score = Math.round(rnd() * 5000) / 100;
    row.active = rnd() > 0.5;
    row.note = NOTES[(rnd() * NOTES.length) | 0];
    row.date = new Date(1577836800000 + ((rnd() * 94608000) | 0)).toISOString().slice(0, 10);
    yield row;
  }
}

function readAsStream(p: string): ReadableStream<Uint8Array> {
  return Readable.toWeb(fs.createReadStream(p)) as ReadableStream<Uint8Array>;
}

export interface BenchmarkOptions {
  rows: number;
  outDir: string;
}

/** Run the full gauntlet and return metrics. */
export async function runBenchmark(opts: BenchmarkOptions): Promise<BenchResult> {
  fs.mkdirSync(opts.outDir, { recursive: true });
  const inputPath = path.join(opts.outDir, 'input.xlsx');
  const outputPath = path.join(opts.outDir, 'output.xlsx');
  const phases: PhaseMetrics[] = [];

  // 1. generate + write
  const p1 = await timed('1.generate+write', async () => {
    const w = EasyExcel.write(inputPath, { class: DemoRow, sheet: 'data' });
    await w.writeRows(generateRows(opts.rows));
    await w.close();
    return opts.rows;
  });
  phases.push({ name: p1.name, ms: p1.ms, peakHeapMB: p1.peakHeapMB, rows: p1.value });

  // 2. read -> map -> transform -> write
  const p2 = await timed('2.read->map->transform->write', async () => {
    const reader = EasyExcel.read(readAsStream(inputPath), { class: DemoRow });
    const transformed = (async function* () {
      for await (const row of reader.rows()) {
        row.note = row.note + '(已复核)';
        yield row;
      }
    })();
    const w = EasyExcel.write(outputPath, { class: DemoRow, sheet: 'data' });
    await w.writeRows(transformed);
    await w.close();
  });
  phases.push({ name: p2.name, ms: p2.ms, peakHeapMB: p2.peakHeapMB, rows: opts.rows });

  // 3. verify output
  const p3 = await timed('3.verify', async () => {
    let count = 0;
    let sumId = 0n;
    const reader = EasyExcel.read(readAsStream(outputPath), { class: DemoRow });
    for await (const row of reader.rows()) {
      count++;
      sumId += BigInt(row.id);
    }
    const expectedSum = (BigInt(opts.rows) * (BigInt(opts.rows) + 1n)) / 2n;
    if (count !== opts.rows) throw new Error(`verify: expected ${opts.rows} rows, got ${count}`);
    if (sumId !== expectedSum) throw new Error(`verify: id sum mismatch: ${sumId} !== ${expectedSum}`);
    return count;
  });
  phases.push({ name: p3.name, ms: p3.ms, peakHeapMB: p3.peakHeapMB, rows: p3.value });

  return {
    rows: opts.rows,
    phases,
    sizes: {
      input: fs.statSync(inputPath).size,
      output: fs.statSync(outputPath).size,
    },
  };
}

export function formatResult(r: BenchResult): string {
  const lines = [
    `rows: ${r.rows.toLocaleString('en-US')}`,
    '',
    'phase                              time(ms)   peak heap(MB)',
    '----------------------------------------------------------',
    ...r.phases.map(
      (p) =>
        `${p.name.padEnd(34)} ${p.ms.toFixed(0).padStart(10)} ${p.peakHeapMB.toFixed(1).padStart(14)}`
    ),
    '',
    `xlsx written: ${(r.sizes.input / 1024 / 1024).toFixed(1)} MB`,
    `xlsx re-written: ${(r.sizes.output / 1024 / 1024).toFixed(1)} MB`,
  ];
  return lines.join('\n');
}
