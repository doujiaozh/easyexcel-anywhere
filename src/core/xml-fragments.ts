/**
 * Static OOXML worksheet package parts + XML escaping helpers.
 * Written with inline strings (t="inlineStr") so the writer never needs
 * a shared-strings table and can stay fully streaming.
 */

export function xmlEscape(s: string): string {
  return s.replace(/[&<>"']/g, (c) => {
    switch (c) {
      case '&': return '&amp;';
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '"': return '&quot;';
      default: return '&apos;';
    }
  });
}

/** 0-based column index -> 'A', 'B', ..., 'Z', 'AA' ... */
export function colName(index0: number): string {
  let n = index0;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** Cell reference like 'C7' from 1-based row and 0-based column. */
export function cellRef(row1: number, col0: number): string {
  return colName(col0) + row1;
}

const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export const SHEET_XML_PROLOGUE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<worksheet xmlns="${SHEET_NS}" xmlns:r="${REL_NS}"><sheetData>\r\n`;

export const SHEET_XML_EPILOGUE = '</sheetData></worksheet>';

function xmlDecl(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
}

function relsXml(rels: Array<{ id: string; type: string; target: string }>): string {
  return (
    xmlDecl() +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    rels
      .map(
        (r) =>
          `<Relationship Id="${xmlEscape(r.id)}" Type="${xmlEscape(r.type)}" Target="${xmlEscape(r.target)}"/>`
      )
      .join('') +
    '</Relationships>'
  );
}

export function contentTypesXml(sheetNames: string[]): string {
  const overrides =
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    sheetNames
      .map(
        (_, i) =>
          `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      )
      .join('');
  return (
    xmlDecl() +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    overrides +
    '</Types>'
  );
}

export function rootRelsXml(): string {
  return relsXml([
    {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'xl/workbook.xml',
    },
  ]);
}

export function workbookXml(sheetNames: string[]): string {
  const sheets = sheetNames
    .map(
      (name, i) =>
        `<sheet name="${xmlEscape(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`
    )
    .join('');
  return (
    xmlDecl() +
    `<workbook xmlns="${SHEET_NS}" xmlns:r="${REL_NS}"><sheets>${sheets}</sheets></workbook>`
  );
}

export function workbookRelsXml(sheetNames: string[]): string {
  const rels: Array<{ id: string; type: string; target: string }> = sheetNames.map((_, i) => ({
    id: `rId${i + 1}`,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
    target: `worksheets/sheet${i + 1}.xml`,
  }));
  rels.push({
    id: `rId${sheetNames.length + 1}`,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    target: 'styles.xml',
  });
  return relsXml(rels);
}

/** Minimal but Excel-valid styles.xml (Calibri 11, no formatting). */
export function stylesXml(): string {
  return (
    xmlDecl() +
    '<styleSheet xmlns="' + SHEET_NS + '">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFontId="0" numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
    '</styleSheet>'
  );
}
