/**
 * Round-trip and API tests (node:test). Run: `npm test`
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { Readable } from 'node:stream';

import { EasyExcel } from '../index.js';
import { ExcelIgnore, ExcelProperty } from '../annotations/ExcelProperty.js';
import type { RawRow } from '../read/reader.js';
import type { WriteOptions } from '../write/writer.js';

class User {
  @ExcelProperty('姓名')
  name = '';

  @ExcelProperty('年龄', { type: 'number' })
  age = 0;

  @ExcelProperty('在职')
  active = false;

  @ExcelProperty('城市')
  city = '';
}

class WithIgnore {
  @ExcelProperty('姓名')
  name = '';

  @ExcelProperty('内部备注')
  note = '';

  @ExcelIgnore()
  password = '';
}

async function writeTmp(
  rows: AsyncIterable<unknown> | Iterable<unknown>,
  options?: WriteOptions
): Promise<string> {
  const p = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'eea-')), 'out.xlsx');
  const w = EasyExcel.write(p, options);
  await w.writeRows(rows);
  await w.close();
  return p;
}

function readBack<T = RawRow>(
  p: string,
  opts?: { class?: new (...args: unknown[]) => T; max?: number }
) {
  const stream = Readable.toWeb(fs.createReadStream(p)) as ReadableStream<Uint8Array>;
  return EasyExcel.read<T>(stream, opts).rows();
}

async function drain<T>(gen: AsyncGenerator<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const r of gen) out.push(r);
  return out;
}

test('class round-trip: header mapping, types, trailing-null cells', async () => {
  const users = [
    Object.assign(new User(), { name: '张三', age: 30, active: true, city: '北京' }),
    Object.assign(new User(), { name: '李四', age: 25, active: false, city: '上海' }),
  ];
  const p = await writeTmp(users, { class: User });
  const out = await drain(readBack(p, { class: User }));
  assert.equal(out.length, 2);
  assert.deepEqual(out[0], users[0]);
  assert.deepEqual(out[1], users[1]);
});

test('raw round-trip: arrays with gaps, XML specials, numbers', async () => {
  const rows = [
    ['a & b <c> "d" \'e\'', 42, null],
    [null, -1.5, '中文'],
  ];
  const p = await writeTmp(rows, { headers: ['name', 'age', 'note'] });
  const out = (await drain(readBack(p))).map((r) => r.cells);
  assert.equal(out.length, 3);
  assert.equal(out[0][0], 'name');
  assert.equal(out[1][0], 'a & b <c> "d" \'e\'');
  assert.equal(out[1][1], '42');
  // trailing empty cells are not written (Excel convention), so the cell list
  // is shorter than the header and the missing slot reads as undefined
  assert.equal(out[1][2], undefined);
  assert.equal(out[2][0], null);
  assert.equal(out[2][1], '-1.5');
  assert.equal(out[2][2], '中文');
});

test('the produced file is a well-formed streaming ZIP (signatures present)', async () => {
  const p = await writeTmp([['hello'], ['world']], { headers: ['col'] });
  const buf = fs.readFileSync(p);
  // local file header at offset 0
  assert.ok(buf.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04])) === 0);
  // data descriptor used by every entry
  assert.ok(buf.indexOf(Buffer.from([0x50, 0x4b, 0x07, 0x08])) > 0);
  // central directory + EOCD
  assert.ok(buf.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02])) > 0);
  assert.ok(buf.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])) > 0);
});

test('mapped read throws a helpful error on missing header', async () => {
  const p = await writeTmp([['x', 'y'], [1, 2]], { headers: ['x', 'y'] });
  await assert.rejects(
    async () => {
      for await (const _ of readBack(p, { class: User })) void _;
    },
    /Header row does not match class User; missing columns for: name \(header "姓名"\)/
  );
});

test('onRow callback receives every mapped row; max limits rows', async () => {
  const rows = [
    Object.assign(new User(), { name: 'a', age: 1, active: true, city: 'c' }),
    Object.assign(new User(), { name: 'b', age: 2, active: false, city: 'c' }),
  ];
  const p = await writeTmp(rows, { class: User });
  let seen = 0;
  const stream = Readable.toWeb(fs.createReadStream(p)) as ReadableStream<Uint8Array>;
  const reader = EasyExcel.read(stream, {
    class: User,
    max: 1,
    onRow: () => {
      seen++;
    },
  });
  const out = await drain(reader.rows());
  assert.equal(out.length, 1);
  assert.equal(seen, 1);
});

test('ExcelIgnore skips a property on write', async () => {
  const rows = [
    Object.assign(new WithIgnore(), { name: 'x', note: 'n', password: 'secret' }),
  ];
  const p = await writeTmp(rows, { class: WithIgnore });
  const back = await drain(readBack(p, { class: WithIgnore }));
  assert.equal(back.length, 1);
  assert.equal(back[0].name, 'x');
  assert.equal(back[0].note, 'n');
  // the password column never made it into the file; the reader instantiates
  // a fresh object, so the field keeps its class default
  assert.equal(back[0].password, '');
});

test('WritableStream target instead of a path (pure streams all the way)', async () => {
  const chunks: Uint8Array[] = [];
  const target = new WritableStream<Uint8Array>({
    write(chunk) {
      chunks.push(chunk);
    },
  });
  const w = EasyExcel.write(target, { headers: ['a', 'b'], sheet: 'S' });
  await w.writeRows([
    [1, 2],
    [3, 4],
  ]);
  await w.close();
  const total = chunks.reduce((n, c) => n + c.length, 0);
  assert.ok(total > 0);
  // round-trip the collected bytes through the reader via a Blob
  const blob = new Blob(chunks as unknown as BlobPart[]);
  const out = (await drain(EasyExcel.read(blob, { sheet: 'S' }).rows())).map(
    (r: RawRow) => r.cells
  );
  assert.deepEqual(out, [
    ['a', 'b'],
    ['1', '2'],
    ['3', '4'],
  ]);
});

test('custom converter: dates as ISO strings in both directions', async () => {
  class EventRow {
    @ExcelProperty('时间', {
      converter: {
        fromCell: (raw: string | null) => (raw ? new Date(raw) : null),
        toCell: (v: unknown) => (v instanceof Date ? v.toISOString() : v == null ? null : String(v)),
      },
    })
    at: Date | null = null;

    @ExcelProperty('标题')
    title = '';
  }
  const events = [
    Object.assign(new EventRow(), { at: new Date('2026-10-31T12:00:00.000Z'), title: '发布会' }),
  ];
  const p = await writeTmp(events, { class: EventRow });
  const back = await drain(readBack(p, { class: EventRow }));
  assert.equal(back.length, 1);
  assert.equal(back[0].at?.toISOString(), '2026-10-31T12:00:00.000Z');
  assert.equal(back[0].title, '发布会');
});
