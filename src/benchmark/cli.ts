/**
 * ee_cli_benchmark CLI entry: `node dist/benchmark/cli.js --rows 1000000 --out bench-out`
 */

import fs from 'node:fs';
import path from 'node:path';
import { formatResult, runBenchmark } from './bench.js';

function usage(): string {
  return [
    'EasyExcel-Anywhere benchmark',
    '',
    'Usage: node dist/benchmark/cli.js [options]',
    '',
    'Options:',
    '  --rows <n>     number of rows to generate (default: 1000000)',
    '  --out <dir>    output directory (default: bench-out)',
    '  --help         show this help',
  ].join('\n');
}

function parseArgs(argv: string[]): { rows: number; outDir: string } {
  const opts = { rows: 1_000_000, outDir: 'bench-out' };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--help' || a === '-h') {
      console.log(usage());
      process.exit(0);
    }
    if (a === '--rows' || a === '--out') {
      const v = argv[i + 1];
      if (!v) {
        console.error(`missing value for ${a}`);
        process.exit(2);
      }
      i++;
      if (a === '--rows') opts.rows = parseInt(v, 10);
      else opts.outDir = v;
    }
  }
  if (!Number.isFinite(opts.rows) || opts.rows <= 0) {
    console.error(`invalid --rows value: ${opts.rows}`);
    process.exit(2);
  }
  return opts;
}

async function main(): Promise<void> {
  const opts = parseArgs(process.argv.slice(2));
  console.log(`EasyExcel-Anywhere benchmark: ${opts.rows.toLocaleString('en-US')} rows`);
  console.log(`node ${process.version}, platform ${process.platform}\n`);
  const result = await runBenchmark(opts);
  console.log(formatResult(result));
  const metricsPath = path.join(opts.outDir, 'metrics.json');
  fs.writeFileSync(metricsPath, JSON.stringify(result, null, 2));
  console.log(`\nmetrics written to ${metricsPath}`);
}

main().catch((e) => {
  console.error('benchmark failed:', e);
  process.exit(1);
});
