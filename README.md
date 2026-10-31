# EasyExcel-Anywhere

**用 Web Streams 降维打击 Java，让百万行表格处理跑在每一个浏览器里。**
*"Beating Java at its own game with Web Streams — million-row spreadsheet
processing that runs in every browser."*

A stream-based, OOM-free xlsx toolkit written in **pure TypeScript** — no JVM,
no native C++ modules, no `node-gyp` build step. It borrows the core idea of
Alibaba's [EasyExcel](https://github.com/alibaba/easyexcel) (annotation-driven
mapping, streaming read/write) and rebuilds it on Web Streams so the same code
runs on Node.js 18+ and in any modern browser.

---

## 为什么需要它 / Why

.classical EasyExcel needs a JVM; typical JS xlsx libraries buffer the whole
workbook in memory. On an 8 GB Windows 7 box a few hundred thousand rows can
already exhaust the heap. EasyExcel-Anywhere never materialises the whole
sheet: bytes are pulled from a `ReadableStream`, inflated chunk by chunk,
parsed by a SAX-style XML pull parser, and handed to your callback one row at
a time — each row is released as soon as you are done with it.

**Memory stays flat whether the file has 1 thousand rows or 1 million.**

传统 EasyExcel 依赖 JVM；常见的 JS xlsx 库会把整个工作簿缓冲进内存。在 8G
内存的 Win7 老电脑上，几十万行就可能撑爆堆。EasyExcel-Anywhere 从不把整张表
物化到内存：字节从 `ReadableStream` 按需拉取、逐块解压、由 SAX 风格的
pull parser 解析、一行一行交给你的回调——处理完立即释放。**无论文件是一千行
还是一百万行，内存曲线都是平的。**

---

## 特性 / Features

| 模块 / Module | 能力 / Capability |
|---|---|
| `ee_stream_read` | 流式读取 `.xlsx`：ZIP 增量解压 + SAX 逐行解析，注册 `onRow` 回调边读边处理，支持 `max` 提前终止 |
| `ee_annotations` | `@ExcelProperty("姓名")` / `@ExcelIgnore()` 装饰器，读取时表头自动映射到对象，写入时自动生成表头；类型可显式声明、也可由 TS 元数据/默认值自动推断 |
| `ee_stream_write` | 流式写入：行数据一边接收一边生成 XML，经 DEFLATE 直接压成 ZIP 字节流，落盘或推给任意 `WritableStream`（含浏览器下载） |
| `ee_cli_benchmark` | 内置百万行压测 CLI：生成 → 读取 → 转换 → 重写 → 校验，输出各阶段耗时与峰值堆内存 |

* Same data path on Node.js and in the browser (`File`/`Blob` →
  `ReadableStream`, or `WritableStream` → download).
* Pure JS dependencies only: [`pako`](https://github.com/nodeca/pako)
  (DEFLATE) + [`sax`](https://github.com/isaacs/sax) (XML) + `reflect-metadata`.
* Streaming ZIP writer uses data descriptors, so output is append-only and
  never seeks — ideal for pipes and HTTP responses.

---

## 安装 / Install

```bash
npm install dsh-easyexcel-anywhere
```

Node.js >= 18.14 (needs global Web Streams). Browser usage needs a bundler
that preserves `import 'reflect-metadata'` (Vite/webpack/esbuild all do).

---

## 快速上手 / Quick Start

### 1. 定义模型（注解映射）— Define your model

```ts
import { ExcelProperty, ExcelIgnore } from 'dsh-easyexcel-anywhere';

export class User {
  @ExcelProperty('姓名')
  name = '';

  @ExcelProperty('年龄')
  age = 0;                       // number inferred from the default

  @ExcelProperty('在职')
  active = false;                // boolean inferred from the default

  @ExcelProperty('城市')
  city = '';

  @ExcelIgnore()                 // never written / never mapped
  password = 'secret';
}
```

类型推断有两级回退：TS 的 `design:type` 元数据（显式标注时），以及属性默认值
的运行时类型（`= 0` → number）。显式写法永远优先：`@ExcelProperty({ header:
'生日', type: 'date' })`。复杂转换用 `converter`：

```ts
@ExcelProperty({
  header: '生日',
  converter: {
    fromCell: (raw) => (raw ? new Date(raw) : null),
    toCell: (v: Date | null) => v?.toISOString() ?? null,
  },
})
birthday: Date | null = null;
```

### 2. 流式读取（防 OOM）— Streaming read

```ts
import { EasyExcel } from 'dsh-easyexcel-anywhere';

// Node
const reader = EasyExcel.read<User>(fs.createReadStream('big.xlsx'), {
  class: User,
  sheet: 'Sheet1',      // name or 0-based index (default 0)
  onRow: (row) => {
    // process and forget — the row is garbage-collected right after
  },
});
const n = await reader.forEach((u) => sendToDb(u));

// Browser — identical API, just a different byte source
const file: File = input.files[0];
for await (const u of EasyExcel.read<User>(file.stream(), { class: User }).rows()) {
  render(u);
}
```

### 3. 流式写入（防 OOM）— Streaming write

```ts
import { EasyExcel } from 'dsh-easyexcel-anywhere';

// Node: streamed straight to disk
await EasyExcel.write('/tmp/out.xlsx', { class: User, sheet: 'Users' })
  .then((w) => w.writeRows(userGenerator()));   // async iterable, never buffered

// Browser: straight into a download
const rs = new ReadableStream({ /* produce rows here */ });
await EasyExcel.write(writableStreamForDownload, { class: User })
  .then((w) => w.writeRows(rs));
```

行可以是类实例，也可以是裸数组（`headers` 显式指定表头）。大数组的内存占用
同样恒定：行被逐个消费、立即编码、立即压缩、立即推走。

---

## 架构 / How it works

```
source (fs stream / Blob.stream / Request)
  └─► StreamReader  …sequential buffered reads, no random access
      └─► ZipReader …parse local file headers on the fly
          └─► Inflate (pako, per-entry, chunkSize 64 KiB)
              └─► sax pull parser → rows yielded one at a time
```

Reading is fully pull-based: the file stream is only pulled when your code
asks for the next row. That backpressure — not buffering — is what keeps the
heap flat. The writer mirrors the same pipeline in reverse: rows → row XML in
64 KiB batches → per-chunk DEFLATE → ZIP local header + data descriptor →
the output stream. Nothing is ever accumulated in full, and Excel,
LibreOffice, 7-Zip and the JDK all read the produced archives (data
descriptors are standard per APPNOTE 4.4.8 / 4.4.13).

---

## 百万行压测 / Million-row benchmark

```bash
npm run bench          # 1,000,000 rows (8 columns) — the real deal
npm run bench:small    # 100,000 rows — quick smoke test
```

The pipeline: generate deterministic data → write xlsx → read it back, append
`"(已复核)"` to a note column, rewrite to a second file → re-read and verify
row count and id checksum. Timings and peak `heapUsed` for every phase are
printed and written to `bench-out/metrics.json`.

Results on the target machine (Windows 7, 8 GB RAM, Node 26, DEFLATE level 6):

| 阶段 / Phase | 耗时 / Time | 峰值堆 / Peak heap |
|---|---|---|
| 1. 生成 + 写入 1,000,000 行 | 38.7 s | 103.6 MB |
| 2. 读取 → 映射 → 转换 → 重新写入 | 88.3 s | 102.7 MB |
| 3. 重读校验（行数 + id 求和） | 57.0 s | 47.1 MB |

Output: two 40.7 MB xlsx files. **Peak heap stays ~104 MB across the whole
million-row cycle** — three orders of magnitude below the naive approach and
flat with respect to row count (the 64 KiB inflate/batch buffers dominate it).
Lower `compressionLevel` if you prefer CPU time over file size.

*Heap is sampled every 50 ms during the whole run; the number reported is the
peak, not the average. The same process does all three phases, so the figure
covers the full read → transform → write cycle.*

---

## 项目布局 / Layout

```
src/
  annotations/ExcelProperty.ts   @ExcelProperty / @ExcelIgnore / schema
  core/bytes.ts                  byte-source normaliser + buffered reader
  core/crc32.ts                  streaming CRC-32
  core/zip-reader.ts             streaming ZIP reader (data descriptors)
  core/zip-writer.ts             streaming ZIP writer (data descriptors)
  core/xml-fragments.ts          OOXML part templates + cell reference math
  core/xml-reader.ts             SAX-based sheet/sharedStrings/workbook parser
  read/mapper.ts                 header↔property bindings, value conversion
  read/reader.ts                 ExcelReader + EasyExcel.read
  write/writer.ts                ExcelWriter + EasyExcel.write
  benchmark/                     million-row benchmark + CLI
tests/roundtrip.test.ts          8 streaming round-trip tests (node:test)
```

---

## 约束 / Constraints（Win7 友好）

* **No JVM, no native modules.** Pure TS/JS — `tsc` itself must also be the
  pure-JS build (TypeScript >= 7 ships a Go binary that does *not* run on
  Windows 7; pin `typescript@~5.9` for Win7 compatibility).
* Dependencies: `pako` (DEFLATE), `sax` (XML), `reflect-metadata`.
* Node >= 18.14 for global Web Streams; browsers need `ReadableStream` +
  `WritableStream` (all modern ones do).

## 开发 / Development

```bash
npm install
npm run build        # tsc → dist/
npm test             # node --test dist/tests/*.test.js
npm run demo         # small annotated end-to-end example
```

## License

MIT © 2026 DSH Contributors
