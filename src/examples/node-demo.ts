/**
 * A quick end-to-end demo. Run: `npm run demo`
 */

import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { Readable } from 'node:stream';
import { EasyExcel } from '../index.js';
import { ExcelProperty } from '../annotations/ExcelProperty.js';

class Product {
  @ExcelProperty('商品')
  name = '';

  @ExcelProperty('价格', { type: 'number' })
  price = 0;

  @ExcelProperty('库存', { type: 'number' })
  stock = 0;

  @ExcelProperty('上架')
  listed = false;
}

async function main(): Promise<void> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eea-demo-'));
  const file = path.join(dir, 'products.xlsx');

  // write: rows flow in, xlsx bytes flow out - nothing is buffered as a whole
  const writer = EasyExcel.write(file, { class: Product, sheet: '商品' });
  await writer.writeRows([
    Object.assign(new Product(), { name: '键盘', price: 199.5, stock: 42, listed: true }),
    Object.assign(new Product(), { name: '鼠标', price: 89.0, stock: 0, listed: false }),
    Object.assign(new Product(), { name: '显示器 & 音箱', price: 1299, stock: 7, listed: true }),
  ]);
  await writer.close();

  console.log(`written: ${file} (${fs.statSync(file).size} bytes)`);

  // read: pull rows one by one; the file stream advances only while you pull
  const stream = Readable.toWeb(fs.createReadStream(file)) as ReadableStream<Uint8Array>;
  for await (const p of EasyExcel.read(stream, { class: Product, sheet: '商品' }).rows()) {
    const product = p as Product;
    console.log(
      `${product.name.padEnd(12)} price=${product.price} stock=${product.stock} listed=${product.listed}`
    );
  }

  // browser equivalent (no Node streams involved):
  //
  //   const file = input.files[0];
  //   for await (const row of EasyExcel.read(file, { class: Product })) { ... }
  //
  //   const writable = await fileHandle.createWritable(); // File System Access API
  //   const w = EasyExcel.write(writable, { class: Product });
  //   await w.writeRows(rows); await w.close();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
