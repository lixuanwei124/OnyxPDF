import { zipStore } from './docx.js';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const esc = (value) => String(value ?? '')
  .replace(/\r\n?/g, '\n')
  .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

function safeSheetName(value, used) {
  let name = String(value || '').replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31);
  if (!name) name = 'Sheet';
  let unique = name;
  let index = 2;
  while (used.has(unique.toLowerCase())) {
    const suffix = `_${index}`;
    unique = name.slice(0, 31 - suffix.length) + suffix;
    index += 1;
  }
  used.add(unique.toLowerCase());
  return unique;
}

function colorHex(value) {
  const rgba = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(String(value || ''));
  const rgb = rgba
    ? rgba.slice(1).map(item => Number(item).toString(16).padStart(2, '0')).join('').toUpperCase()
    : String(value || '#000000').replace('#', '').padEnd(6, '0').slice(0, 6).toUpperCase();
  return `FF${rgb}`;
}

const FONT_MAP = {
  simsun: '宋体', nsimsun: '新宋体', simhei: '黑体', kaiti: '楷体',
  fangsong: '仿宋', dengxian: '等线', 'microsoft yahei': '微软雅黑',
  pingfangsc: '苹方', hiraginosansgb: '冬青黑体', stzhongsong: '华文中宋',
  stkaiti: '华文楷体', stxihei: '华文细黑', stfangsong: '华文仿宋'
};

function cleanFont(value) {
  let name = String(value || '').split(',')[0].replace(/^[A-Z]{6}\+/, '')
    .replace(/["']/g, '').trim();
  name = name.replace(/-(?:Regular|Bold|Italic|Light|Medium)$/i, '')
    .replace(/(?:Regular|Bold|Italic|Light|Medium)$/i, '').trim();
  return FONT_MAP[name.toLowerCase()] || name || 'Microsoft YaHei';
}

function fontFamily(run, page) {
  const exported = page?.embeddedFonts?.byAlias?.get(String(run.font || '').toLowerCase());
  return run?.__fontFamily || exported?.family || cleanFont(run.font);
}

function columnLabel(index) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label || 'A';
}

function cellReference(row, column) {
  return `${columnLabel(column)}${row}`;
}

function sourceLines(block) {
  if (Array.isArray(block.exportLines) && block.exportLines.length) return block.exportLines;
  const runs = block.runs || [];
  const size = Number(block.size) || Number(runs[0]?.size) || 12;
  const cell = {
    x: Number(block.x) || 0,
    baseline: Number(block.baseline) || 0,
    width: Math.max(Number(block.width) || 1, 1),
    size,
    text: runs.map(run => run.text || '').join(''),
    runs
  };
  return [{ x: cell.x, baseline: cell.baseline, width: cell.width, size, cells: [cell] }];
}

function styleKey(run, alignment) {
  return [
    fontFamily(run),
    Math.round((Number(run.size) || 12) * 100),
    run.bold ? 1 : 0,
    run.italic ? 1 : 0,
    run.underline ? 1 : 0,
    colorHex(run.color),
    alignment
  ].join('\u0000');
}

function cellStyle(cell, page) {
  const firstRun = cell.runs[0] || {};
  const run = { ...firstRun, __fontFamily: fontFamily(firstRun, page) };
  const alignment = cell.alignment || 'left';
  return {
    key: styleKey(run, alignment),
    font: run.__fontFamily,
    fontId: `${run.__fontFamily}\u0000${Math.round((Number(run.size) || 12) * 100)}\u0000${run.bold ? 1 : 0}\u0000${run.italic ? 1 : 0}\u0000${run.underline ? 1 : 0}\u0000${colorHex(run.color)}`,
    size: Math.max(1, Number(run.size) || 12),
    bold: !!run.bold,
    italic: !!run.italic,
    underline: !!run.underline,
    color: colorHex(run.color),
    alignment
  };
}

function workbookStyles(usedStyles) {
  const fontIds = new Map();
  const fonts = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  const cellXfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  for (const style of usedStyles.values()) {
    if (!fontIds.has(style.fontId)) {
      fontIds.set(style.fontId, fonts.length);
      fonts.push('<font>' +
        (style.bold ? '<b/>' : '') +
        (style.italic ? '<i/>' : '') +
        (style.underline ? '<u/>' : '') +
        `<sz val="${style.size}"/><color rgb="${style.color}"/>` +
        `<name val="${esc(style.font)}"/><charset val="134"/></font>`);
    }
    cellXfs.push('<xf numFmtId="0" fontId="' + fontIds.get(style.fontId) +
      '" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">' +
      `<alignment horizontal="${style.alignment}" vertical="center" wrapText="0" shrinkToFit="0"/>` +
      '</xf>');
  }
  const stylesXml = XML_HEADER +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FFB7C3D7"/></left><right style="thin"><color rgb="FFB7C3D7"/></right>' +
    '<top style="thin"><color rgb="FFB7C3D7"/></top><bottom style="thin"><color rgb="FFB7C3D7"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';
  return { stylesXml, styleIndex: key => usedStyles.has(key) ? [...usedStyles.keys()].indexOf(key) + 1 : 0 };
}

function inlineString(cell) {
  const runs = (cell.runs || []).map(run => {
    const font = fontFamily(run);
    return '<r><rPr>' +
      (run.bold ? '<b/>' : '') + (run.italic ? '<i/>' : '') +
      (run.underline ? '<u/>' : '') +
      `<sz val="${Math.max(1, Number(run.size) || 12)}"/>` +
      `<color rgb="${colorHex(run.color)}"/>` +
      `<rFont val="${esc(font)}"/><charset val="134"/></rPr>` +
      `<t xml:space="preserve">${esc(run.text ?? '')}</t></r>`;
  }).join('');
  return `<is>${runs}</is>`;
}

function exportedTextLines(page) {
  return (page.paragraphs || []).flatMap(block => sourceLines(block))
    .map(line => ({
      ...line,
      cells: (line.cells || []).filter(cell =>
        (cell.runs || []).some(run => String(run.text ?? '').trim() !== '')
      )
    }))
    .filter(line => line.cells.length);
}

function clusterRows(page, lines) {
  const prepared = lines.map(line => {
    const size = Math.max(...line.cells.map(cell => Number(cell.size) || 12));
    return {
      baseline: Number(line.baseline) || 0,
      top: Math.max(0, Number(page.height) - Number(line.baseline || 0) - 0.82 * size),
      height: Math.max(8, size * 1.38),
      cells: line.cells.map(cell => {
        const x = Math.max(0, Number(cell.x) || 0);
        const runs = cell.runs.map(item => ({ ...item, __fontFamily: fontFamily(item, page) }));
        return {
          x,
          width: Math.max(Number(cell.width) || 1, 1),
          size,
          text: runs.map(item => item.text || '').join(''),
          runs,
          style: cellStyle({ ...cell, runs, alignment: 'left' }, page)
        };
      })
    };
  });
  prepared.sort((left, right) => right.baseline - left.baseline);
  const rows = [];
  for (const line of prepared) {
    const previous = rows[rows.length - 1];
    const tolerance = Math.max(2.5, Math.max(previous?.cells[0]?.size || line.cells[0].size, line.cells[0].size) * 0.45);
    if (previous && Math.abs(previous.baseline - line.baseline) <= tolerance) {
      previous.cells.push(...line.cells);
      previous.size = Math.max(previous.size || 0, line.size);
    } else {
      rows.push({ baseline: line.baseline, top: line.top, height: line.height, size: line.size, cells: [...line.cells] });
    }
  }
  return rows;
}

function inferColumns(rows, pageWidth) {
  if (!rows.length) return null;
  const coordinates = rows.flatMap(row => row.cells.map(cell => cell.x)).sort((left, right) => left - right);
  const sizes = rows.flatMap(row => row.cells.map(cell => cell.size || 12));
  const mergeDistance = Math.max(10, pageWidth * 0.025, (sizes.sort((a, b) => a - b)[Math.floor(sizes.length / 2)] || 12) * 1.6);
  let candidates = [];
  for (const x of coordinates) {
    const cluster = candidates[candidates.length - 1];
    if (cluster && x - cluster.to <= mergeDistance) {
      cluster.values.push(x);
      cluster.to = x;
    } else {
      candidates.push({ values: [x], from: x, to: x });
    }
  }
  const minimumCount = Math.max(2, Math.floor(rows.length * 0.08));
  candidates = candidates.filter(cluster => cluster.values.length >= minimumCount);
  const columns = candidates.map((cluster, index) => ({
    index,
    center: cluster.values.reduce((sum, value) => sum + value, 0) / cluster.values.length,
    radius: Math.max(14, (cluster.to - cluster.from) / 2 + 6)
  }));
  if (columns.length < 2) return null;

  const coverage = new Array(columns.length).fill(0);
  for (const row of rows) {
    const used = new Set();
    for (const cell of row.cells) {
      const column = columns.reduce((best, item) =>
        Math.abs(item.center - cell.x) < Math.abs(best.center - cell.x) ? item : best);
      cell.column = column.index + 1;
      used.add(column.index);
    }
    used.forEach(index => coverage[index] += 1);
  }
  const rowThreshold = Math.max(3, Math.floor(rows.length * 0.3));
  const strongColumns = coverage.filter(count => count >= rowThreshold).length;
  const multiCellRows = rows.filter(row => new Set(row.cells.map(cell => cell.column)).size >= 2).length;
  if (strongColumns < 2 || multiCellRows < rows.length * 0.75) return null;
  return columns;
}

function splitSpanningCells(rows, columns) {
  for (const row of rows) {
    const nextCells = [];
    for (const cell of row.cells) {
      const covered = columns.filter(column =>
        column.center >= cell.x - 1 && column.center <= cell.x + cell.width + 1);
      const segments = cell.runs
        .filter(run => String(run.text || '').trim())
        .map(run => ({ ...run, text: String(run.text || '').trim() }));
      if (covered.length < 2 || segments.length !== covered.length) {
        nextCells.push(cell);
        continue;
      }
      covered.forEach((column, index) => {
        const width = Math.max(12, column.radius * 2);
        nextCells.push({
          ...cell,
          x: column.center - width / 2,
          width,
          text: segments[index].text,
          runs: [{ ...cell.runs[0], text: segments[index].text }],
          column: column.index + 1
        });
      });
    }
    row.cells = nextCells.sort((left, right) => left.x - right.x);
  }
}

function extractTablePage(page, pageIndex) {
  const rows = clusterRows(page, exportedTextLines(page));
  if (rows.length < 5) return null;
  const columns = inferColumns(rows, Number(page.width) || 595);
  if (!columns) return null;
  splitSpanningCells(rows, columns);
  rows.forEach(row => row.cells.sort((left, right) => left.x - right.x));

  const gridRows = rows.map(row => {
    const grid = new Array(columns.length).fill(null);
    for (const cell of row.cells) {
      const column = cell.column - 1;
      if (!grid[column] || cell.width > grid[column].width) grid[column] = cell;
    }
    return {
      baseline: row.baseline,
      top: row.top,
      height: row.height,
      cells: grid
    };
  });
  return {
    pageIndex,
    pageWidth: Number(page.width) || 595,
    columns: columns.length,
    columnCenters: columns.map(column => column.center),
    rows: gridRows
  };
}

function rowDistance(left, right) {
  return Math.abs(left.baseline - right.baseline);
}

function matchTableRows(leftRows, rightRows) {
  const tolerance = Math.max(3, ((leftRows[0]?.cells.find(Boolean)?.size || 12) * 0.4));
  const pairs = [];
  rightRows.forEach((row, index) => {
    let best = -1;
    let distance = Infinity;
    leftRows.forEach((candidate, candidateIndex) => {
      const value = rowDistance(candidate, row);
      if (value < distance && !pairs.some(pair => pair.left === candidateIndex)) {
        best = candidateIndex;
        distance = value;
      }
    });
    if (best >= 0 && distance <= tolerance) pairs.push({ left: best, right: index, distance });
  });
  return pairs;
}

function combineTablePages(pages) {
  const tables = [];
  for (const page of pages) {
    const table = tables[tables.length - 1];
    if (!table) {
      tables.push({ ...page, rows: page.rows.map(row => ({ ...row, cells: [...row.cells] })) });
      continue;
    }
    const overlapLimit = Math.max(12, (Number(page.pageWidth) || 595) * 0.02);
    const overlappingColumns = Array.from({ length: Math.min(table.columnCenters.length, page.columnCenters.length) }, (_, index) =>
      Math.abs(table.columnCenters[index] - page.columnCenters[index]) <= overlapLimit
    ).filter(Boolean).length;
    const columnOverlapRatio = overlappingColumns / Math.min(table.columnCenters.length, page.columnCenters.length);

    if (columnOverlapRatio >= 0.6) {
      table.rows.push(...page.rows.map(row => ({ ...row, cells: [...row.cells] })));
      continue;
    }

    const pairs = matchTableRows(table.rows, page.rows);
    const rowMatchRatio = pairs.length / Math.min(table.rows.length, page.rows.length);
    if (overlappingColumns === 0 && rowMatchRatio >= 0.7) {
      const rightOffset = table.columns;
      const matchedRight = new Set();
      for (const pair of pairs) {
        const baseRow = table.rows[pair.left];
        const rightRow = page.rows[pair.right];
        matchedRight.add(pair.right);
        while (baseRow.cells.length < rightOffset) baseRow.cells.push(null);
        baseRow.cells.push(...rightRow.cells);
      }
      for (let index = 0; index < page.rows.length; index++) {
        if (matchedRight.has(index)) continue;
        const row = page.rows[index];
        table.rows.push({
          baseline: row.baseline,
          top: row.top,
          height: row.height,
          cells: Array(rightOffset).fill(null).concat(row.cells)
        });
      }
      for (const row of table.rows) {
        while (row.cells.length < table.columns + page.columns) row.cells.push(null);
      }
      table.columns += page.columns;
      table.columnCenters.push(...page.columnCenters);
      continue;
    }

    tables.push({ ...page, rows: page.rows.map(row => ({ ...row, cells: [...row.cells] })) });
  }
  return tables;
}

function spreadsheetTable(table, pageIndex, usedNames) {
  const usedStyles = new Map();
  table.rows.forEach(row => row.cells.forEach(cell => {
    if (cell && !usedStyles.has(cell.style.key)) usedStyles.set(cell.style.key, cell.style);
  }));
  const blankStyle = usedStyles.size
    ? [...usedStyles.values()].sort((left, right) => left.font.localeCompare(right.font))[0]
    : cellStyle({ runs: [{ font: 'Calibri', size: 12, color: '#000000' }], alignment: 'left' }, {});
  const rows = table.rows.map(row => ({
    top: row.top,
    height: row.height,
    cells: Array.from({ length: table.columns }, (_, index) => {
      const cell = row.cells[index];
      return cell ? { ...cell, column: index + 1 } : { blank: true, column: index + 1, style: blankStyle };
    })
  }));
  return {
    name: safeSheetName(`第${pageIndex + 1}页`, usedNames),
    rows,
    maxColumn: table.columns,
    blankStyle
  };
}

function spreadsheetPage(page, pageIndex, usedNames) {
  const rows = clusterRows(page, exportedTextLines(page)).map(row => {
    const cell = row.cells.find(item => item && (item.text || '').trim());
    if (!cell) {
      return { top: row.top, height: row.height, cells: [] };
    }
    const missingRuns = row.cells
      .filter(item => item !== cell)
      .flatMap(item => item.runs || []);
    const mergedCell = missingRuns.length
      ? { ...cell, runs: [...(cell.runs || []), ...missingRuns], text: row.cells.map(item => item.text || '').join('') }
      : cell;
    return {
      top: row.top,
      height: row.height,
      cells: [{ ...mergedCell, column: 1 }]
    };
  });
  return {
    name: safeSheetName(`第${pageIndex + 1}页`, usedNames),
    rows,
    maxColumn: 1
  };
}

function columnXml(widths, maxColumn) {
  const parts = [];
  for (let column = 1; column <= maxColumn; column += 1) {
    const width = widths.get(column) || 10;
    parts.push(`<col min="${column}" max="${column}" width="${width.toFixed(2)}" customWidth="1"/>`);
  }
  return parts.join('');
}

function worksheetXml(sheet, styleIndex) {
  const widths = new Map();
  const rowsXml = sheet.rows.map((row, rowIndex) => {
    const rowNumber = rowIndex + 1;
    const cellsXml = row.cells.map(cell => {
      const styleAttribute = `s="${styleIndex(cell.style.key)}"`;
      if (cell.blank) {
        return `<c r="${cellReference(rowNumber, cell.column)}" ${styleAttribute}/>`;
      }
      const text = String(cell.text || '').trim();
      const isNumber = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(text) && text.length < 15;
      const current = widths.get(cell.column) || 10;
      widths.set(cell.column, Math.max(current, Math.min(180, Math.max(10, cell.width / 5.8))));
      return isNumber
        ? `<c r="${cellReference(rowNumber, cell.column)}" ${styleAttribute}><v>${text}</v></c>`
        : `<c r="${cellReference(rowNumber, cell.column)}" ${styleAttribute} t="inlineStr">${inlineString(cell)}</c>`;
    }).join('');
    return `<row r="${rowNumber}" ht="${row.height.toFixed(2)}" customHeight="1">${cellsXml}</row>`;
  }).join('');
  const maxColumn = Math.max(1, ...widths.keys(), sheet.maxColumn);
  return XML_HEADER +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="A1:${cellReference(Math.max(1, sheet.rows.length), maxColumn)}"/>` +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    `<cols>${columnXml(widths, maxColumn)}</cols>` +
    `<sheetData>${rowsXml}</sheetData></worksheet>`;
}

export function buildXlsxLayout(pages) {
  const usedNames = new Set();
  const inputPages = (pages || []).filter(page => exportedTextLines(page).length);
  const tablePages = inputPages
    .map((page, index) => extractTablePage(page, index))
    .filter(Boolean);
  const tables = combineTablePages(tablePages);
  const tablePageIndexSet = new Set(tablePages.map(table => table.pageIndex));
  const sheets = [
    ...tables.map((table, index) => spreadsheetTable(table, index + 1, usedNames)),
    ...inputPages
      .map((page, index) => ({ page, index }))
      .filter(({ index }) => !tablePageIndexSet.has(index))
      .map(({ page, index }) => spreadsheetPage(page, index + 1, usedNames))
  ];
  const usedStyles = new Map();
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (!usedStyles.has(cell.style.key)) usedStyles.set(cell.style.key, cell.style);
      }
    }
  }
  const { stylesXml, styleIndex } = workbookStyles(usedStyles);
  const contentTypes = XML_HEADER +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    sheets.map((_, index) =>
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join('') + '</Types>';
  const workbookXml = XML_HEADER +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets>' + sheets.map((sheet, index) =>
      `<sheet name="${esc(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
    ).join('') + '</sheets></workbook>';
  const workbookRels = XML_HEADER +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    sheets.map((_, index) =>
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    ).join('') +
    `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    '</Relationships>';
  const files = [
    { name: '[Content_Types].xml', data: new TextEncoder().encode(contentTypes) },
    { name: '_rels/.rels', data: new TextEncoder().encode(XML_HEADER +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
    { name: 'xl/workbook.xml', data: new TextEncoder().encode(workbookXml) },
    { name: 'xl/_rels/workbook.xml.rels', data: new TextEncoder().encode(workbookRels) },
    { name: 'xl/styles.xml', data: new TextEncoder().encode(stylesXml) }
  ];
  sheets.forEach((sheet, index) => {
    files.push({ name: `xl/worksheets/sheet${index + 1}.xml`, data: new TextEncoder().encode(worksheetXml(sheet, styleIndex)) });
  });
  return zipStore(files);
}
