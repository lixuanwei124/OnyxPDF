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
  if (typeof stream.getUnencodedContents === 'function') return stream.getUnencodedContents();
  if (typeof stream.getContentsString === 'function') return stringToBytes(stream.getContentsString());
  return stream.asUint8Array();
}

function dictionaryOf(anyObject) {
  const context = anyObject?.context || anyObject?.dict?.context;
  const value = anyObject?.context && anyObject?.dict ? anyObject : context?.lookup(anyObject);
  return value?.dict ? value : (value?.keys ? value : null);
}

function baseFontName(font) {
  const value = font?.get?.(pdfName('BaseFont'))?.decodeText?.() ||
    font?.get?.(pdfName('BaseFont'))?.asString?.() || '';
  return String(value).replace(/^\//, '').replace(/^[A-Z]{6}\+/, '').trim();
}

function pdfName(value) {
  return window.PDFLib.PDFName.of(value);
}

export function unreliableFontNames(pdfDocument, pageIndex) {
  const pdfLib = window.PDFLib;
  if (!pdfLib?.PDFName) return new Set();
  const page = pdfDocument.getPage(pageIndex);
  const context = pdfDocument.context;
  const names = new Set();
  const visiting = new Set();

  const visitResources = resources => {
    const resourceDict = dictionaryOf(resources);
    if (!resourceDict) return;
    const fonts = resourceDict.lookupMaybe(pdfName('Font'), pdfLib.PDFDict);
    if (!fonts) return;
    for (const [, reference] of fonts.entries()) {
      const key = reference?.objectNumber;
      if (key != null && visiting.has(key)) continue;
      if (key != null) visiting.add(key);
      const font = context.lookup(reference);
      const subtype = nameOf(font?.get?.(pdfName('Subtype')));
      const baseFont = baseFontName(font);
      if ((subtype === 'Type0' && !font.get(pdfName('ToUnicode'))) || subtype === 'Type3') {
        if (baseFont) names.add(baseFont);
      }
      if (subtype === 'Type0') {
        const descendant = context.lookup(font.get(pdfName('DescendantFonts'))?.get?.(0) ||
          font.get(pdfName('DescendantFonts')));
        const descendantName = baseFontName(descendant);
        if ((subtype === 'Type0' && !font.get(pdfName('ToUnicode'))) && descendantName) names.add(descendantName);
      }
      const formResources = font?.get?.(pdfName('Resources'));
      if (formResources) visitResources(formResources);
    }
  };

  visitResources(dictionaryOf(context.lookup(page.node.Resources())));
  return names;
}

function nameOf(value) {
  return String(value?.asString?.() ?? value?.decodeText?.() ?? '').replace(/^\//, '');
}

function isProtectedFont(resources, fontName, protectedNames) {
  const resourceDict = dictionaryOf(resources);
  const font = resourceDict?.lookupMaybe?.(pdfName('Font'), window.PDFLib.PDFDict)
    ?.lookupMaybe?.(fontName, window.PDFLib.PDFDict);
  if (!font) return false;
  const subtype = nameOf(font.get(pdfName('Subtype')));
  if (subtype === 'Type3') return true;
  if (subtype === 'Type0' && !font.get(pdfName('ToUnicode'))) return true;
  return protectedNames.has(baseFontName(font));
}

function rewriteStream(pdfLib, reference, stream, resources, removeSet, state, visiting, mutate = false) {
  const guardKey = reference?.objectNumber;
  if (guardKey != null && visiting.has(guardKey)) return { changed: false, text: '' };
  if (guardKey != null) visiting.add(guardKey);
  const { source, operations: streamOperations } = operations(streamBytes(pdfLib, stream));
  const pieces = [];
  let cursor = 0;
  let changed = false;
  let protectedFont = false;
  const fontStack = [];

  for (const operation of streamOperations) {
    if (operation.operator === 'q') {
      fontStack.push(protectedFont);
    } else if (operation.operator === 'Q') {
      protectedFont = fontStack.pop() || false;
    } else if (operation.operator === 'Tf') {
      const name = operation.operands?.find(item => item.text.startsWith('/'))?.text.slice(1);
      protectedFont = name ? isProtectedFont(resources, pdfLib.PDFName.of(name), state.protectedNames) : false;
    }
    if (TEXT_OPERATORS.has(operation.operator)) {
      const index = state.showIndex++;
      pieces.push(source.slice(cursor, operation.start));
      cursor = operation.end;
      if ((state.removeAll && !protectedFont) || removeSet.has(index)) {
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
      const formReference = name
        ? resourceDict?.get(pdfLib.PDFName.of('XObject'))?.get(pdfLib.PDFName.of(name))
        : null;
      const form = formReference ? state.context.lookup(formReference) : null;
      const subtype = form?.dict?.get?.(pdfLib.PDFName.of('Subtype'))?.asString?.() || '';
      if (form && subtype.includes('Form')) {
        const inherited = form.dict.get(pdfLib.PDFName.of('Resources')) || resources;
        const formResources = dictionaryOf(inherited) || dictionaryOf(resources);
        const result = rewriteStream(pdfLib, formReference, form, formResources, removeSet, state, visiting, true);
        changed = changed || result.changed;
      }
    }
  }

  pieces.push(source.slice(cursor));
  const result = { changed, text: pieces.join('') };
  if (mutate && changed && state.context) {
    const changedBytes = stringToBytes(result.text);
    if (stream instanceof pdfLib.PDFRawStream) {
      stream.contents = changedBytes;
      stream.dict.delete(pdfLib.PDFName.of('Filter'));
      stream.dict.delete(pdfLib.PDFName.of('DecodeParms'));
      stream.updateDict();
    } else {
      const cleanDict = stream.dict.clone(state.context);
      cleanDict.delete(pdfLib.PDFName.of('Filter'));
      cleanDict.delete(pdfLib.PDFName.of('DecodeParms'));
      state.context.assign(reference, pdfLib.PDFRawStream.of(cleanDict, changedBytes));
    }
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
  const context = pdfDocument.context;
  const state = { showIndex: 0, removed: new Set(), removeAll, context };
  state.protectedNames = unreliableFontNames(pdfDocument, pageIndex);
  const contents = page.node.normalizedEntries().Contents;
  const container = context.lookup(contents);
  const references = container?.asArray ? container.asArray() : [contents];
  const resources = dictionaryOf(context.lookup(page.node.Resources()));
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

export function removeMarkedContentOperations(pdfDocument, pageIndex, markedContentIds) {
  const wanted = new Set((markedContentIds || []).map(Number).filter(Number.isInteger));
  if (!wanted.size) return false;
  const pdfLib = window.PDFLib;
  if (!pdfLib?.PDFRawStream || !pdfLib?.decodePDFRawStream) {
    throw new Error('PDF 内容流编辑组件不可用');
  }

  const page = pdfDocument.getPage(pageIndex);
  const context = pdfDocument.context;
  const contents = page.node.normalizedEntries().Contents;
  const container = context.lookup(contents);
  const references = container?.asArray ? container.asArray() : [contents];
  const resources = dictionaryOf(context.lookup(page.node.Resources()));
  let changed = false;
  const visiting = new Set();

  const removeRanges = (reference, stream, streamResources) => {
    const guardKey = reference?.objectNumber;
    if (guardKey != null && visiting.has(guardKey)) return;
    if (guardKey != null) visiting.add(guardKey);
    const { source, operations: streamOperations } = operations(streamBytes(pdfLib, stream));
    const pieces = [];
    let cursor = 0;
    let skipDepth = 0;
    let removeStart = -1;

    for (let index = 0; index < streamOperations.length; index++) {
      const operation = streamOperations[index];
      const shouldRemove = skipDepth > 0;
      if (!shouldRemove && operation.operator === 'BDC') {
        const operandText = (operation.operands || []).map(item => item.text).join(' ');
        const match = /\/MCID\s+(\d+)/.exec(operandText);
        if (match && wanted.has(Number(match[1]))) {
          removeStart = operation.start;
          skipDepth = 1;
        }
      } else if (skipDepth > 0 && operation.operator === 'BDC') {
        skipDepth++;
      } else if (skipDepth > 0 && operation.operator === 'EMC') {
        skipDepth--;
        if (skipDepth === 0 && removeStart >= 0) {
          pieces.push(source.slice(cursor, removeStart));
          cursor = operation.end;
          removeStart = -1;
          changed = true;
        }
      }

      if (operation.operator === 'Do' && streamResources) {
        const name = operation.operands?.find(item => item.text.startsWith('/'))?.text.slice(1);
        const resourceDict = dictionaryOf(streamResources);
        const formReference = resourceDict?.get(pdfLib.PDFName.of('XObject'))?.get(name);
        const form = formReference ? context.lookup(formReference) : null;
        const subtype = form?.dict?.get?.('Subtype')?.asString?.() || '';
        if (form && subtype.includes('Form')) {
          const inherited = form.dict.get(pdfLib.PDFName.of('Resources')) || streamResources;
          removeRanges(formReference, form, dictionaryOf(inherited) || dictionaryOf(streamResources));
        }
      }
    }

    if (removeStart >= 0) return;
    pieces.push(source.slice(cursor));
    const result = pieces.join('');
    if (changed && stream instanceof pdfLib.PDFRawStream) {
      stream.contents = stringToBytes(result);
      stream.dict.delete(pdfLib.PDFName.of('Filter'));
      stream.dict.delete(pdfLib.PDFName.of('DecodeParms'));
      stream.updateDict();
    }
    if (guardKey != null) visiting.delete(guardKey);
  };

  for (const reference of references) {
    const stream = context.lookup(reference);
    if (stream?.dict) removeRanges(reference, stream, resources);
  }
  return changed;
}
