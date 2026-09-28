/**
 * PDF 内容流文本对象编辑。
 *
 * PDF.js 的 sourceOps 映射到内容流中的原子 Tj/TJ/'/" 显示操作。
 * 这里按原子序号移除命中的原文本操作，替代文字由导出层写入同一份 PDF；
 * 保存结果不再依赖画布遮盖。
 */

const TEXT_OPERATORS = new Set(['Tj', 'TJ', "'", '"']);
const OPERATORS = new Set([
  'q', 'Q', 'cm', 'm', 'l', 'c', 'v', 'y', 'h', 're', 'W', 'W*', 'n', 'S', 's',
  'f', 'F', 'f*', 'B', 'B*', 'b', 'b*', 'sh', 'BT', 'ET', 'Tc', 'Tw', 'Tz',
  'TL', 'Tf', 'Tr', 'Ts', 'Td', 'TD', 'Tm', 'T*', 'gs', 'Do', 'MP', 'DP',
  'BMC', 'BDC', 'EMC', 'd0', 'd1', 'CS', 'cs', 'SC', 'sc', 'SCN', 'scn',
  'G', 'g', 'RG', 'rg', 'K', 'k', 'M', 'i', 'j', 'J', 'w', 'ri', 'd',
  'BI', 'ID', 'EI'
]);
for (const operator of TEXT_OPERATORS) OPERATORS.add(operator);

function bytesToString(bytes) {
  let text = '';
  for (let i = 0; i < bytes.length; i += 8192) {
    text += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
  }
  return text;
}

function stringToBytes(text) {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 255;
  return bytes;
}

function tokenizeContent(bytes) {
  const source = bytesToString(bytes);
  const tokens = [];
  let pos = 0;
  const delimiters = new Set(['(', ')', '<', '>', '[', ']', '{', '}', '/', '%']);

  const pushToken = (type, start, end) => {
    if (end > start) tokens.push({ type, text: source.slice(start, end), start, end });
  };

  while (pos < source.length) {
    const ch = source[pos];
    if (/[\s\r\n]/.test(ch)) { pos++; continue; }
    if (ch === '%') {
      const end = source.indexOf('\n', pos);
      pos = end < 0 ? source.length : end + 1;
      continue;
    }
    const start = pos;
    if (ch === '(') {
      let depth = 1;
      pos++;
      while (pos < source.length && depth > 0) {
        const c = source[pos];
        if (c === '\\') pos += 2;
        else {
          if (c === '(') depth++;
          else if (c === ')') depth--;
          pos++;
        }
      }
      pushToken('string', start, pos);
      continue;
    }
    if (source.startsWith('<<', pos)) { pos += 2; pushToken('operand', start, pos); continue; }
    if (source.startsWith('>>', pos)) { pos += 2; pushToken('operand', start, pos); continue; }
    if (ch === '<') {
      pos = source.indexOf('>', pos + 1);
      pos = pos < 0 ? source.length : pos + 1;
      pushToken('string', start, pos);
      continue;
    }
    if (delimiters.has(ch)) {
      pos++;
      if (ch === '/') {
        while (pos < source.length && !/[\s\r\n]/.test(source[pos]) && !delimiters.has(source[pos])) pos++;
      }
      pushToken('operand', start, pos);
      continue;
    }
    while (pos < source.length && !/[\s\r\n]/.test(source[pos]) && !delimiters.has(source[pos])) pos++;
    pushToken(OPERATORS.has(source.slice(start, pos)) ? 'operator' : 'operand', start, pos);
  }
  return { source, tokens };
}

function splitOperations(bytes) {
  const { source, tokens } = tokenizeContent(bytes);
  const operations = [];
  let operands = [];
  let start = tokens.length ? tokens[0].start : 0;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type === 'operator' && token.text === 'BI') {
      const idIndex = tokens.findIndex((item, index) => index > i && item.text === 'ID');
      if (idIndex < 0) throw new Error('PDF inline image is missing ID');
      const tail = source.slice(tokens[idIndex].end);
      const match = /(?:^|[\s\r\n])EI(?=[\s\r\n]|$)/.exec(tail);
      if (!match) throw new Error('PDF inline image is missing EI');
      const end = tokens[idIndex].end + match.index + match[0].length;
      operations.push({ operator: 'INLINE_IMAGE', start, end });
      operations.push({ operator: 'EI', start: end - 2, end });
      operands = [];
      i = tokens.findIndex((item, index) => index > i && item.start >= end);
      if (i < 0) i = tokens.length;
      start = i < tokens.length ? tokens[i].start : end;
      continue;
    }
    if (token.type === 'operator') {
      operations.push({ operator: token.text, start, end: token.end, operands });
      operands = [];
      start = tokens[i + 1]?.start ?? token.end;
    } else {
      operands.push(token);
    }
  }
  return { source, operations };
}

function streamBytes(pdfLib, stream) {
  const isRawStream = stream instanceof pdfLib.PDFRawStream || stream.constructor?.name === 'PDFRawStream';
  if (isRawStream && pdfLib.decodePDFRawStream) {
    return pdfLib.decodePDFRawStream(stream).decode();
  }
  if (typeof stream.getContentsString === 'function') return stringToBytes(stream.getContentsString());
  return stream.asUint8Array();
}

function lookupDict(pdfLib, context, value) {
  if (!value) return null;
  const resolved = value.context && value.dict ? value : context.lookup(value);
  return resolved?.dict ? resolved : null;
}

function processStream(pdfLib, ref, stream, resources, removeOps, state, seen) {
  const cacheKey = ref?.objectNumber;
  if (cacheKey != null && seen.has(cacheKey)) return seen.get(cacheKey);
  const { source, operations } = splitOperations(streamBytes(pdfLib, stream));
  const pieces = [];
  let cursor = 0;
  let changed = false;
  const textOpCount = operations.filter(op => TEXT_OPERATORS.has(op.operator)).length;

  for (const op of operations) {
    if (TEXT_OPERATORS.has(op.operator)) {
      const index = state.showIndex++;
      pieces.push(source.slice(cursor, op.start));
      cursor = op.end;
      if (!removeOps.has(index)) pieces.push(source.slice(op.start, op.end));
      else {
        changed = true;
        state.removed.add(index);
      }
      continue;
    }

    if (op.operator === 'Do' && resources) {
      const nameOperand = (op.operands || []).find(item => item.text.startsWith('/'));
      const formName = nameOperand ? nameOperand.text.slice(1) : null;
      const xobjects = lookupDict(pdfLib, pdfLib ? pdfDocContext(resources) : null, resources.get(pdfLib.PDFName.of('XObject')));
      const formRef = xobjects?.get(formName);
      const form = formRef ? pdfDocContext(resources).lookup(formRef) : null;
      const subtype = form?.dict?.get?.(pdfLib.PDFName.of('Subtype'))?.asString?.() || '';
      if (form && String(subtype).includes('Form')) {
        const formResources = lookupDict(pdfLib, pdfDocContext(resources), form.dict.get(pdfLib.PDFName.of('Resources')) || resources);
        const result = processStream(pdfLib, formRef, form, formResources, removeOps, state, seen);
        if (result.changed) {
          pieces.push(source.slice(cursor, op.start));
          pieces.push('q\n' + result.text + '\nQ');
          cursor = op.end;
          changed = true;
        }
      }
    }
  }

  pieces.push(source.slice(cursor));
  const result = {
    changed,
    text: pieces.join(''),
    textOpCount
  };
  if (changed && stream instanceof pdfLib.PDFRawStream) {
    stream.contents = stringToBytes(result.text);
    stream.dict.delete(pdfLib.PDFName.of('Filter'));
    stream.dict.delete(pdfLib.PDFName.of('DecodeParms'));
    stream.updateDict();
  }
  if (cacheKey != null) seen.set(cacheKey, result);
  return result;
}

function pdfDocContext(anyPdfObject) {
  return anyPdfObject?.context || anyPdfObject?.dict?.context;
}

export function removeTextOperators(pdfDoc, pageIndex, sourceOps) {
  const pdfLib = window.PDFLib;
  if (!pdfLib?.PDFRawStream || !pdfLib?.decodePDFRawStream) {
    throw new Error('pdf-lib 内部对象编辑接口不可用');
  }
  const page = pdfDoc.getPage(pageIndex);
  const removeOps = new Set(sourceOps || []);
  const state = { showIndex: 0, removed: new Set(), debug: [] };
  const context = pdfDoc.context;
  const contents = page.node.normalizedEntries().Contents;
  const container = context.lookup(contents);
  const refs = container?.asArray ? container.asArray() : [contents];
  const seen = new Map();

  for (const ref of refs) {
    const stream = context.lookup(ref);
    if (!stream || !stream.dict) continue;
    const resources = lookupDict(pdfLib, context, page.node.Resources);
    const result = processStream(pdfLib, ref, stream, resources, removeOps, state, seen);
    state.debug.push({
      objectNumber: ref?.objectNumber ?? 'direct',
      type: stream.constructor?.name || stream.constructor?.toString?.(),
      textOps: result.textOpCount || 0,
      changed: !!result.changed
    });
  }
  if (state.removed.size !== removeOps.size) {
    throw new Error(`Content stream edit missed ${removeOps.size - state.removed.size} text operation(s); showIndex=${state.showIndex}; streams=${JSON.stringify(state.debug)}`);
  }
  return true;
}
