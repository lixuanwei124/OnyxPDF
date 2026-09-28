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
  for (let index = 0; index < bytes.length; index += 8192) {
    text += String.fromCharCode.apply(null, bytes.subarray(index, index + 8192));
  }
  return text;
}

function stringToBytes(text) {
  const bytes = new Uint8Array(text.length);
  for (let index = 0; index < text.length; index++) {
    bytes[index] = text.charCodeAt(index) & 255;
  }
  return bytes;
}

function tokenize(bytes) {
  const source = bytesToString(bytes);
  const tokens = [];
  const delimiters = new Set(['(', ')', '<', '>', '[', ']', '{', '}', '/', '%']);
  let position = 0;

  const push = (type, start, end) => {
    if (end > start) tokens.push({ type, text: source.slice(start, end), start, end });
  };

  while (position < source.length) {
    const character = source[position];
    if (/[\s\r\n]/.test(character)) {
      position++;
      continue;
    }
    if (character === '%') {
      const end = source.indexOf('\n', position);
      position = end < 0 ? source.length : end + 1;
      continue;
    }
    const start = position;
    if (character === '(') {
      let depth = 1;
      position++;
      while (position < source.length && depth > 0) {
        if (source[position] === '\\') position += 2;
        else {
          if (source[position] === '(') depth++;
          else if (source[position] === ')') depth--;
          position++;
        }
      }
      push('string', start, position);
      continue;
    }
    if (source.startsWith('<<', position) || source.startsWith('>>', position)) {
      position += 2;
      push('operand', start, position);
      continue;
    }
    if (character === '<') {
      position = source.indexOf('>', position + 1);
      position = position < 0 ? source.length : position + 1;
      push('string', start, position);
      continue;
    }
    if (delimiters.has(character)) {
      position++;
      if (character === '/') {
        while (position < source.length && !/[\s\r\n]/.test(source[position]) &&
          !delimiters.has(source[position])) position++;
      }
      push('operand', start, position);
      continue;
    }
    while (position < source.length && !/[\s\r\n]/.test(source[position]) &&
      !delimiters.has(source[position])) position++;
    push(OPERATORS.has(source.slice(start, position)) ? 'operator' : 'operand', start, position);
  }
  return { source, tokens };
}

function operations(bytes) {
  const { source, tokens } = tokenize(bytes);
  const result = [];
  let operands = [];
  let start = tokens[0]?.start || 0;
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index];
    if (token.type === 'operator' && token.text === 'BI') {
      const idIndex = tokens.findIndex((item, itemIndex) => itemIndex > index && item.text === 'ID');
      if (idIndex < 0) throw new Error('PDF inline image is missing ID');
      const tail = source.slice(tokens[idIndex].end);
      const match = /(?:^|[\s\r\n])EI(?=[\s\r\n]|$)/.exec(tail);
      if (!match) throw new Error('PDF inline image is missing EI');
      const end = tokens[idIndex].end + match.index + match[0].length;
      result.push({ operator: 'INLINE_IMAGE', start, end });
      result.push({ operator: 'EI', start: end - 2, end });
      operands = [];
      index = tokens.findIndex((item, itemIndex) => itemIndex > index && item.start >= end);
      if (index < 0) index = tokens.length;
      start = index < tokens.length ? tokens[index].start : end;
      continue;
    }
    if (token.type === 'operator') {
      result.push({ operator: token.text, start, end: token.end, operands });
      operands = [];
      start = tokens[index + 1]?.start ?? token.end;
    } else {
      operands.push(token);
    }
  }
  return { source, operations: result };
}

function streamBytes(pdfLib, stream) {
  if (stream instanceof pdfLib.PDFRawStream && pdfLib.decodePDFRawStream) {
    return pdfLib.decodePDFRawStream(stream).decode();
  }
  if (typeof stream.getContentsString === 'function') return stringToBytes(stream.getContentsString());
  return stream.asUint8Array();
}

function dictionaryOf(anyObject) {
  const context = anyObject?.context || anyObject?.dict?.context;
  const value = anyObject?.context && anyObject?.dict ? anyObject : context?.lookup(anyObject);
  return value?.dict ? value : null;
}

function rewriteStream(pdfLib, reference, stream, resources, removeSet, state, visiting, mutate = false) {
  const guardKey = reference?.objectNumber;
  if (guardKey != null && visiting.has(guardKey)) return { changed: false, text: '' };
  if (guardKey != null) visiting.add(guardKey);
  const { source, operations: streamOperations } = operations(streamBytes(pdfLib, stream));
  const pieces = [];
  let cursor = 0;
  let changed = false;

  for (const operation of streamOperations) {
    if (TEXT_OPERATORS.has(operation.operator)) {
      const index = state.showIndex++;
      pieces.push(source.slice(cursor, operation.start));
      cursor = operation.end;
      if (state.removeAll || removeSet.has(index)) {
        changed = true;
        state.removed.add(index);
      } else {
        pieces.push(source.slice(operation.start, operation.end));
      }
      continue;
    }

    if (operation.operator === 'Do' && resources) {
      const name = operation.operands?.find(item => item.text.startsWith('/'))?.text.slice(1);
      const resourceDict = dictionaryOf(resources);
      const formReference = resourceDict?.get(pdfLib.PDFName.of('XObject'))?.get(name);
      const form = formReference ? dictionaryOf(resources).context.lookup(formReference) : null;
      const subtype = form?.dict?.get?.(pdfLib.PDFName.of('Subtype'))?.asString?.() || '';
      if (form && subtype.includes('Form')) {
        const inherited = form.dict.get(pdfLib.PDFName.of('Resources')) || resources;
        const formResources = dictionaryOf(inherited) || dictionaryOf(resources);
        const result = rewriteStream(pdfLib, formReference, form, formResources, removeSet, state, visiting);
        if (result.changed) {
          pieces.push(source.slice(cursor, operation.start));
          pieces.push(`q\n${result.text}\nQ`);
          cursor = operation.end;
          changed = true;
        }
      }
    }
  }

  pieces.push(source.slice(cursor));
  const result = { changed, text: pieces.join('') };
  if (mutate && changed && stream instanceof pdfLib.PDFRawStream) {
    stream.contents = stringToBytes(result.text);
    stream.dict.delete(pdfLib.PDFName.of('Filter'));
    stream.dict.delete(pdfLib.PDFName.of('DecodeParms'));
    stream.updateDict();
  }
  if (guardKey != null) visiting.delete(guardKey);
  return result;
}

export function removeTextOperations(pdfDocument, pageIndex, sourceOperations, options = {}) {
  const pdfLib = window.PDFLib;
  if (!pdfLib?.PDFRawStream || !pdfLib?.decodePDFRawStream) {
    throw new Error('PDF 内容流编辑组件不可用');
  }
  const removeAll = options.all === true;
  const removeSet = new Set(sourceOperations || []);
  const page = pdfDocument.getPage(pageIndex);
  const state = { showIndex: 0, removed: new Set(), removeAll };
  const context = pdfDocument.context;
  const contents = page.node.normalizedEntries().Contents;
  const container = context.lookup(contents);
  const references = container?.asArray ? container.asArray() : [contents];
  const resources = dictionaryOf(context.lookup(page.node.Resources));
  const visiting = new Set();
  for (const reference of references) {
    const stream = context.lookup(reference);
    if (stream?.dict) rewriteStream(pdfLib, reference, stream, resources, removeSet, state, visiting, true);
  }
  if (!removeAll && state.removed.size !== removeSet.size) {
    throw new Error(`PDF 原文移除不完整：${state.removed.size}/${removeSet.size}`);
  }
  return true;
}
