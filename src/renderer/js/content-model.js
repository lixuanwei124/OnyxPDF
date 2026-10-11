import * as pdfjsLib from '../vendor/pdf.min.mjs';
import { unreliableFontNames } from './pdf-content-stream.js';

function glyphText(args) {
  return (args || []).map(glyph => {
    if (typeof glyph === 'number') return ' ';
    return glyph?.unicode ?? glyph?.fontChar ?? '';
  }).join('');
}

function colorByte(value) {
  const number = Number(value) || 0;
  return Math.max(0, Math.min(255, Math.round(number > 1 ? number : number * 255)));
}

function rgbHex(color) {
  return '#' + color.map(colorByte).map(value => value.toString(16).padStart(2, '0')).join('');
}

function mapItemsToShowOperations(items, operations) {
  const mapped = items.map(() => []);
  const visibleTargets = [];
  items.forEach((item, index) => {
    if (String(item.str || '').replace(/\s+/g, '')) visibleTargets.push({ item, index });
  });

  let operationIndex = 0;
  let operationCursor = 0;
  for (const target of visibleTargets) {
    const wanted = String(target.item.str || '').replace(/\s+/g, '');
    let consumed = 0;
    while (consumed < wanted.length && operationIndex < operations.length) {
      const operation = operations[operationIndex];
      const available = String(operation.text || '').replace(/\s+/g, '');
      if (!available || operationCursor >= available.length) {
        operationIndex++;
        operationCursor = 0;
        continue;
      }
      const take = Math.min(available.length - operationCursor, wanted.length - consumed);
      if (take > 0) {
        mapped[target.index].push(operation.index);
        consumed += take;
        operationCursor += take;
      }
      if (operationCursor >= available.length) {
        operationIndex++;
        operationCursor = 0;
      }
    }
  }
  return mapped;
}

function colorDistance(left, right) {
  return Math.hypot(left[0] - right[0], left[1] - right[1], left[2] - right[2]);
}

function parseHexColor(value) {
  const match = /^#?([0-9a-f]{6})$/i.exec(String(value || ''));
  if (!match) return [0, 0, 0];
  return [0, 2, 4].map(offset => parseInt(match[1].slice(offset, offset + 2), 16));
}

function rgbValue(color) {
  return '#' + color.map(value => Math.max(0, Math.min(255, Math.round(value)))
    .toString(16).padStart(2, '0')).join('');
}

async function analyzeRenderedTextItems(page, items, desiredColors = []) {
  const appearance = items.map(item => item?.str?.trim() ? { visible: true, color: null } : null);
  const pageViewport = page.getViewport({ scale: 1 });
  const scale = Math.min(2, 2200 / Math.max(pageViewport.width, 1));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const context = canvas.getContext('2d', { willReadFrequently: true });
  try {
    await page.render({ canvasContext: context, viewport }).promise;
    const image = context.getImageData(0, 0, canvas.width, canvas.height);
    const pixels = image.data;
    items.forEach((item, itemIndex) => {
      if (!appearance[itemIndex] || !item.width) return;
      const size = Math.hypot(item.transform[2], item.transform[3]) ||
        Math.abs(item.transform[3]) || item.height || 10;
      const height = Math.max(item.height || size, size);
      const pdfTop = item.transform[5] + height * 0.98;
      const pdfBottom = item.transform[5] - height * 0.32;
      const x0 = Math.max(0, Math.floor(item.transform[4] * scale) - 2);
      const x1 = Math.min(canvas.width - 1,
        Math.ceil((item.transform[4] + item.width) * scale) + 2);
      const y0 = Math.max(0, Math.floor((pageViewport.height - pdfTop) * scale) - 2);
      const y1 = Math.min(canvas.height - 1,
        Math.ceil((pageViewport.height - pdfBottom) * scale) + 2);
      if (x1 <= x0 || y1 <= y0) {
        appearance[itemIndex].visible = false;
        return;
      }

      const buckets = new Map();
      for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
          const offset = (y * canvas.width + x) * 4;
          const color = [pixels[offset], pixels[offset + 1], pixels[offset + 2]];
          const key = ((color[0] >> 3) << 10) | ((color[1] >> 3) << 5) | (color[2] >> 3);
          const bucket = buckets.get(key);
          if (bucket) {
            bucket.count++;
            bucket.sum[0] += color[0];
            bucket.sum[1] += color[1];
            bucket.sum[2] += color[2];
          } else {
            buckets.set(key, { count: 1, sum: [...color] });
          }
        }
      }
      const samples = [...buckets.values()]
        .map(bucket => ({
          count: bucket.count,
          color: bucket.sum.map(value => value / bucket.count)
        }))
        .sort((left, right) => right.count - left.count);
      if (!samples.length) {
        appearance[itemIndex].visible = false;
        return;
      }

      const area = (x1 - x0 + 1) * (y1 - y0 + 1);
      const minimumCount = Math.max(2, Math.round(area * 0.0008));
      const desired = parseHexColor(desiredColors[itemIndex]);
      const backgroundSamples = desired
        ? samples.filter(sample => colorDistance(sample.color, desired) > 48)
        : samples;
      const background = backgroundSamples[0] || samples[0];

      let desiredPixelCount = 0;
      for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
          const offset = (y * canvas.width + x) * 4;
          const pixel = [pixels[offset], pixels[offset + 1], pixels[offset + 2]];
          if (colorDistance(pixel, desired) <= 72) desiredPixelCount++;
        }
      }

      const desiredContrast = colorDistance(desired, background.color);
      const sampledForeground = samples.slice(1)
        .filter(sample => sample.count >= minimumCount &&
          colorDistance(sample.color, background.color) >= 48)
        .sort((left, right) =>
          colorDistance(right.color, background.color) * Math.log(right.count + 1) -
          colorDistance(left.color, background.color) * Math.log(left.count + 1))[0];
      const operatorMatches = desiredPixelCount >= minimumCount && desiredContrast >= 48 &&
        (!sampledForeground ||
          desiredContrast + 12 >= colorDistance(sampledForeground.color, background.color));
      let foreground = operatorMatches
        ? { count: desiredPixelCount, color: desired }
        : sampledForeground;
      if (!operatorMatches && !foreground) {
        foreground = samples.slice(1)
          .filter(sample => sample.count >= minimumCount &&
            colorDistance(sample.color, background.color) >= 34)
          .sort((left, right) =>
            colorDistance(right.color, background.color) * Math.sqrt(right.count) -
            colorDistance(left.color, background.color) * Math.sqrt(left.count))[0];
      }
      appearance[itemIndex].visible = Boolean(foreground);
      appearance[itemIndex].color = foreground ? rgbValue(foreground.color) : null;
      appearance[itemIndex].operatorMatches = Boolean(foreground && operatorMatches);
    });
  } catch (error) {
    if (import.meta.env?.DEV) console.warn('PDF 文本外观采样失败', error);
  } finally {
    canvas.width = 1;
    canvas.height = 1;
  }
  return appearance;
}

function decodeLegacyFontName(value) {
  const text = String(value || '');
  if (!/^[\x20-\x7e\u00a0-\u00ff]+$/.test(text) || !/[\u00a0-\u00ff]/.test(text)) return text;
  try {
    const bytes = new Uint8Array([...text].map(character => character.charCodeAt(0) & 0xff));
    const decoded = new TextDecoder('gbk').decode(bytes).trim();
    return decoded || text;
  } catch (_) {
    return text;
  }
}

function canonicalFontName(value) {
  let name = decodeLegacyFontName(value)
    .replace(/^[A-Z]{6}\+/, '')
    .replace(/[-_ ]?(?:Bold|Heavy|Black|Semib?old|Light|Medium|Italic|Oblique)$/i, '')
    .trim();
  const aliases = new Map([
    ['MicrosoftYaHei', 'Microsoft YaHei'],
    ['MicrosoftYaHeiUI', 'Microsoft YaHei UI'],
    ['SimSun', '宋体'],
    ['NSimSun', '新宋体'],
    ['SimHei', '黑体'],
    ['KaiTi', '楷体'],
    ['FangSong', '仿宋'],
    ['DengXian', '等线'],
    ['SymbolMT', 'Symbol']
  ]);
  return aliases.get(name) || name;
}

function fontInfo(page, fontId) {
  try {
    const font = page.commonObjs.get(fontId);
    const rawName = decodeLegacyFontName(font?.name || font?.fallbackName || '');
    const bold = /(?:Bold|Heavy|Black|Semib?old)|(?:^|[-_ ])Bd(?: |$|-)/i.test(rawName);
    const name = canonicalFontName(rawName);
    return {
      name: name || null,
      bold,
      key: fontId,
      loadedName: font?.loadedName || null,
      data: font?.data instanceof Uint8Array ? font.data : null,
      mimetype: font?.mimetype || null
    };
  } catch (_) {
    return { name: null, bold: false, key: fontId, loadedName: null, data: null, mimetype: null };
  }
}

function textLooksUnreliable(value) {
  return /[\uE000-\uF8FF\uFFFD\uFFF9-\uFFFB]/u.test(String(value || ''));
}

function appendRuns(output, incoming) {
  for (const run of incoming) {
    if (!run.text) continue;
    const previous = output[output.length - 1];
    const same = previous &&
      previous.font === run.font &&
      previous.fontKey === run.fontKey &&
      previous.loadedName === run.loadedName &&
      Math.abs(previous.size - run.size) < 0.05 &&
      previous.color === run.color &&
      previous.bold === run.bold &&
      !previous.sourceOps.length !== !run.sourceOps.length;
    if (same && run.text !== ' ') {
      previous.text += run.text;
      previous.sourceOps.push(...run.sourceOps);
      if (Number.isFinite(run.x) && Number.isFinite(previous.x)) {
        previous.width = Math.max(
          Number(previous.width) || 0,
          run.x + (Number(run.width) || 0) - previous.x
        );
      }
      if (Number.isFinite(run.baseline)) previous.baseline = run.baseline;
      if (Number.isFinite(run.height)) {
        previous.height = Math.max(Number(previous.height) || 0, run.height);
      }
    } else {
      output.push({
        text: run.text,
        font: run.font || '微软雅黑',
        fontKey: run.fontKey || null,
        loadedName: run.loadedName || null,
        fontData: run.fontData || null,
        size: Math.round((run.size || 12) * 10) / 10,
        color: run.color || '#000000',
        bold: !!run.bold,
        x: Number.isFinite(run.x) ? run.x : undefined,
        baseline: Number.isFinite(run.baseline) ? run.baseline : undefined,
        width: Number.isFinite(run.width) ? run.width : undefined,
        height: Number.isFinite(run.height) ? run.height : undefined,
        sourceOps: [...(run.sourceOps || [])]
      });
    }
  }
  return output;
}

function weightedMedian(items) {
  if (!items.length) return 12;
  const sorted = [...items].sort((left, right) => left.value - right.value);
  const totalWeight = sorted.reduce((sum, item) => sum + item.weight, 0);
  let seen = 0;
  for (const item of sorted) {
    seen += item.weight;
    if (seen >= totalWeight / 2) return item.value;
  }
  return sorted[sorted.length - 1].value;
}

function dominantRun(line) {
  return [...line.runs].sort((left, right) => right.text.length - left.text.length)[0] || {};
}

function styleKey(line) {
  const run = dominantRun(line);
  return [
    run.font || '',
    Math.round((run.size || line.size) * 10) / 10,
    run.color || '#000000',
    run.bold ? 'bold' : 'regular'
  ].join('|');
}

function isHeadingLine(line, bodySize) {
  const run = dominantRun(line);
  const text = line.text.trim();
  return (run.size || line.size) >= bodySize * 1.12 ||
    /^[一二三四五六七八九十]+[、.．]\s*\S{1,28}$/.test(text) ||
    (run.bold && text.length <= 35 && !/[.。!！?？:：;；,，]$/.test(text));
}

function isItemStart(line) {
  return /^([•·▪◦\u2022\u25CF\u25AA\u2013\-—*]|[一二三四五六七八九十]+[、.．]|\d+[.、)])\s*/u.test(line.text.trimStart());
}

function lineAlignment(line, pageWidth) {
  const center = line.x + line.width / 2;
  const leftMargin = line.x;
  const rightMargin = Math.max(0, pageWidth - line.x - line.width);
  const marginDifference = Math.abs(leftMargin - rightMargin);
  const sizeUnit = Math.max(6, line.size);
  if (line.width <= pageWidth * 0.68 &&
    center > pageWidth * 0.32 && center < pageWidth * 0.68 && marginDifference <= sizeUnit * 1.6 &&
    Math.min(leftMargin, rightMargin) >= Math.max(sizeUnit * 2, pageWidth * 0.09)) {
    return 'center';
  }
  if (rightMargin <= sizeUnit && leftMargin >= Math.max(24, sizeUnit * 3)) return 'right';
  return 'left';
}

function dominantAlignment(lines, pageWidth) {
  const counts = { left: 0, center: 0, right: 0 };
  for (const line of lines) counts[line.alignment || lineAlignment(line, pageWidth)] += line.text.length || 1;
  return Object.entries(counts).sort((left, right) => right[1] - left[1])[0][0];
}

function canJoinLines(previous, current, bodySize, pageWidth) {
  const previousRun = dominantRun(previous);
  const currentRun = dominantRun(current);
  const sizeRatio = current.size / previous.size;
  if (sizeRatio < 0.86 || sizeRatio > 1.16) return false;
  const changesDominantFont = (previousRun.font || '') !== (currentRun.font || '');
  if ((previousRun.color || '') !== (currentRun.color || '')) return false;
  if (!!previousRun.bold !== !!currentRun.bold) return false;

  const gap = Math.abs(current.baseline - previous.baseline);
  const expectedGap = Math.max(previous.lineHeight, current.lineHeight);
  if (gap <= expectedGap * 0.55) return false;

  if (lineAlignment(previous, pageWidth) !== lineAlignment(current, pageWidth)) return false;
  const leftShift = current.x - previous.x;
  const maxIndent = Math.max(current.size * 3.1, 18);
  if (leftShift > maxIndent || leftShift < -current.size * 1.2) return false;

  const previousText = previous.text.trimEnd();
  const currentText = current.text.trimStart();
  if (!previousText || !currentText) return false;
  const paragraphEnd = /[.。!！?？]$/.test(previousText);
  const bulletStart = /^([•·▪◦\u2022\u25CF\u25AA\u2013\-—*]|[一二三四五六七八九十]+[、.．]|\d+[.、)])\s*/u.test(currentText);
  const wrappedEvidence = previous.width >= current.width * 0.72 &&
    previous.width >= Math.max(bodySize * 18, 72);
  if (changesDominantFont && !wrappedEvidence) return false;
  const hasContinuationPunctuation = /[，、；,;:]$/.test(previousText);
  const strongWrappedEvidence = wrappedEvidence && hasContinuationPunctuation;
  const maximumGap = strongWrappedEvidence
    ? Math.max(expectedGap * 3.2, Math.max(previous.size, current.size) * 4)
    : Math.max(expectedGap * 1.75, Math.max(previous.size, current.size) * 2.5);
  if (gap <= expectedGap * 0.55 || gap > maximumGap) return false;
  if (paragraphEnd && gap > expectedGap * 1.18 && !strongWrappedEvidence) return false;
  if (bulletStart && (current.x - previous.x > current.size * 0.8 || gap > expectedGap * 1.18)) return false;
  if (gap > expectedGap * 1.32 && !strongWrappedEvidence && /^[\p{Lu}]/u.test(currentText)) return false;
  const shortStandaloneText = Math.max(previous.width, current.width) < bodySize * 23;
  if (!wrappedEvidence && (!hasContinuationPunctuation || shortStandaloneText)) return false;
  return true;
}

function joinLineRuns(lines) {
  const runs = [];
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    const copied = line.runs.map(run => ({
      ...run,
      sourceOps: [...(run.sourceOps || [])]
    }));
    if (lineIndex > 0 && copied.length) {
      const previousLine = lines[lineIndex - 1];
      const expectedGap = Math.max(previousLine.lineHeight, line.lineHeight);
      const preserveLineBreak = isItemStart(line) ||
        line.baseline - previousLine.baseline > expectedGap * 1.32;
      if (preserveLineBreak) {
        appendRuns(runs, [{
          ...copied[0],
          text: '\n',
          sourceOps: []
        }]);
      } else {
        const left = previousLine.text.trimEnd().slice(-1) || '';
        const right = copied[0].text.trimStart().slice(0, 1) || '';
        const needsSpace = /[\p{L}\p{N},;:)\]}>%"'”’]/u.test(left) &&
          /[\p{L}\p{N}(\[<{“‘]/u.test(right) &&
          !/[\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/u.test(left + right);
        if (needsSpace) appendRuns(runs, [{ ...copied[0], text: ' ' }]);
      }
    } else if (runs.length && copied.length) {
      const previousLine = lines[lineIndex - 1];
      const left = previousLine?.text.trimEnd().slice(-1) || '';
      const right = copied[0].text.trimStart().slice(0, 1) || '';
      const needsSpace = /[\p{L}\p{N},;:)\]}>%"'”’]/u.test(left) &&
        /[\p{L}\p{N}(\[<{“‘]/u.test(right) &&
        !/[\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/u.test(left + right);
      if (needsSpace) appendRuns(runs, [{ ...copied[0], text: ' ' }]);
    }
    appendRuns(runs, copied);
  }
  return runs.map(run => ({
    ...run,
    sourceOps: [...new Set(run.sourceOps || [])].sort((left, right) => left - right)
  }));
}

function buildParagraphBlocks(lines, pageWidth) {
  const bodySize = weightedMedian(lines.flatMap(line =>
    line.runs.map(run => ({
      value: run.size,
      weight: Math.min(Math.max(run.text.replace(/\s+/g, '').length, 1), 120)
    }))
  ));
  const groups = [];
  let current = [];

  const flush = () => {
    if (current.length) groups.push(current);
    current = [];
  };

  for (const line of lines) {
    if (!current.length) {
      current = [line];
      continue;
    }
    const previous = current[current.length - 1];
    if (!isHeadingLine(previous, bodySize) && !isHeadingLine(line, bodySize) &&
      canJoinLines(previous, line, bodySize, pageWidth)) {
      current.push(line);
    } else {
      flush();
      current = [line];
    }
  }
  flush();

  return groups.map((group, index) => {
    const heading = group.length === 1 && isHeadingLine(group[0], bodySize);
    const runs = joinLineRuns(group);
    const left = Math.min(...group.map(line => line.x));
    const right = Math.max(...group.map(line => line.x + line.width));
    const gaps = group.slice(1).map((line, lineIndex) =>
      Math.abs(line.baseline - group[lineIndex].baseline)
    );
    const measuredLineHeight = weightedMedian(gaps.map(gap => ({ value: gap, weight: 1 })));
    const maxHeight = Math.max(...group.map(line => Math.max(...line.runs.map(run => run.size), line.size)));
    const lineHeight = Math.min(Math.max(
      group.length > 1 && measuredLineHeight > maxHeight * 0.9 ? measuredLineHeight : maxHeight * 1.22,
      maxHeight * 0.95
    ), maxHeight * 1.85);
    const sourceOps = [...new Set(group.flatMap(line => line.sourceOps))]
      .filter(Number.isInteger).sort((leftValue, rightValue) => leftValue - rightValue);

    return {
      id: `P${index}`,
      kind: heading ? 'heading' : 'paragraph',
      lineCount: group.length,
      x: group[0].x,
      baseline: group[0].baseline,
      width: Math.max(right - group[0].x, Math.max(...group.map(line => line.width)), 1),
      size: Math.round(maxHeight * 10) / 10,
      lineHeight: Math.round(lineHeight * 10) / 10,
      originalLineHeight: Math.round(lineHeight * 10) / 10,
      height: Math.max((group[0].baseline - group[group.length - 1].baseline) + maxHeight * 1.22, maxHeight * 1.22),
      ascent: 0.82,
      runs,
      exportLines: group.map(line => ({
        x: line.x,
        baseline: line.baseline,
        width: line.width,
        size: line.size,
        runs: line.runs.map(run => ({ ...run, sourceOps: [...(run.sourceOps || [])] })),
        cells: (line.cells || []).map(cell => ({
          ...cell,
          runs: cell.runs.map(run => ({ ...run, sourceOps: [...(run.sourceOps || [])] }))
        }))
      })),
      sourceLineRuns: group.map(line => line.runs.map(run => ({
        ...run,
        sourceOps: [...run.sourceOps]
      }))),
      sourceOps,
      alignment: dominantAlignment(group, pageWidth),
      text: runs.map(run => run.text).join('').replace(/\n/g, ''),
      sourceLines: group.map(line => ({ id: line.id, text: line.text }))
    };
  });
}

function normalizedText(value) {
  return String(value || '').replace(/[\s\-—–·•.,。、;；:：!！?？'"()（）[\]{}]/gu, '');
}

function walkStructureNodes(nodes, callback) {
  for (const node of nodes || []) {
    if (node?.bbox) callback(node);
    walkStructureNodes(node?.children, callback);
  }
}

function structureSemanticText(node) {
  const alt = String(node?.alt || '');
  const separator = alt.lastIndexOf('\r\n');
  const text = separator >= 0 ? alt.slice(separator + 2).trim() : '';
  return text && !/^\([^()]{1,30}\)$/.test(text) ? text.replace(/\s+/g, ' ') : '';
}

function markedContentId(node) {
  const ids = [];
  const visit = (current) => {
    const match = /(?:^|[-_])mc(\d+)$/.exec(String(current?.id || ''));
    if (match) ids.push(Number(match[1]));
    for (const child of current?.children || []) visit(child);
  };
  visit(node);
  return ids;
}

function fragmentMatchesStructureText(fragment, text, rectangle) {
  const overlapWidth = Math.min(fragment.x + fragment.width, rectangle[2]) -
    Math.max(fragment.x, rectangle[0]);
  const overlapHeight = Math.min(fragment.baseline + fragment.height, rectangle[3]) -
    Math.max(fragment.baseline, rectangle[1]);
  const overlapArea = Math.max(0, overlapWidth) * Math.max(0, overlapHeight);
  const structureArea = Math.max(1, (rectangle[2] - rectangle[0]) * (rectangle[3] - rectangle[1]));
  const fragmentArea = Math.max(1, fragment.width * fragment.height);
  const sourceText = normalizedText(fragment.text).slice(0, 32);
  return overlapArea / Math.min(structureArea, fragmentArea) > 0.32 &&
    normalizedText(text).includes(sourceText);
}

async function collectStructureText(page, fragments) {
  const nodes = [];
  const structureTree = await page.getStructTree?.();
  walkStructureNodes(structureTree?.children, node => {
  const text = structureSemanticText(node);
    if (text && Array.isArray(node.bbox) && node.bbox.length === 4) nodes.push({ node, text });
  });
  const missing = nodes.filter(({ node, text }) =>
    !fragments.some(fragment => fragmentMatchesStructureText(fragment, text, node.bbox)));
  const items = missing.map(({ node, text }) => {
    const [x0, y0, x1, y1] = node.bbox;
    const width = Math.max(1, x1 - x0);
    const height = Math.max(1, y1 - y0);
    const size = Math.min(Math.max(height * 0.76, 5), 300);
    const isCJK = /[\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/.test(text);
    return {
      str: text,
      width,
      height,
      transform: [size, 0, 0, size, x0, y0 + height * 0.18],
      fontName: null,
      structureFigure: true,
      structureFontName: isCJK ? '微软雅黑' : 'Arial'
    };
  });
  return {
    items,
    ids: [...new Set(nodes.flatMap(({ node }) => markedContentId(node)))]
  };
}

export function splitRawLine(rawLine) {
  const segments = [];
  let current = [];
  for (const fragment of rawLine) {
    if (!current.length) {
      current = [fragment];
      continue;
    }
    const previous = current[current.length - 1];
    const gap = fragment.x - (previous.x + previous.width);
    const size = Math.min(previous.size, fragment.size);
    const sizeRatio = Math.max(previous.size, fragment.size) / size;
    const mixedSizeGap = size * 1.45;
    const structuralSizeBreak = sizeRatio > 1.16 && gap >= size * 0.24;
    const structuralColorBreak = previous.color !== fragment.color &&
      sizeRatio > 1.16 && gap >= size * 0.16;
    const previousEndsCJK = /[\u2E80-\u9FFF\uF900-\uFAFF]/u.test(
      String(previous.text || '').trimEnd().slice(-1)
    );
    const fragmentStartsSpacedCJK = /^\s+[\u2E80-\u9FFF\uF900-\uFAFF]/u.test(
      String(fragment.text || '')
    );
    const tableCellBreak = previousEndsCJK && fragmentStartsSpacedCJK;
    if (gap > Math.max(mixedSizeGap, 10) || structuralSizeBreak ||
      structuralColorBreak || tableCellBreak) {
      segments.push(current);
      current = [fragment];
    } else {
      current.push(fragment);
    }
  }
  if (current.length) segments.push(current);
  return segments;
}

export async function extractContentPage(pdfDocument, pageNumber, options = {}) {
  const page = await pdfDocument.getPage(pageNumber);
  const pageViewport = page.getViewport({ scale: 1 });
  const textContent = await page.getTextContent();

  const operations = [];
  let stroke = [0, 0, 0];
  try {
    const operatorList = await page.getOperatorList();
    const OPS = pdfjsLib.OPS;
    let fill = [0, 0, 0];
    let renderMode = 0;
    const graphicsStateStack = [];
    for (let index = 0; index < operatorList.fnArray.length; index++) {
      const fn = operatorList.fnArray[index];
      const args = operatorList.argsArray[index];
      if (fn === OPS.save) {
        graphicsStateStack.push({ fill: [...fill], stroke: [...stroke], renderMode });
      } else if (fn === OPS.restore) {
        const savedState = graphicsStateStack.pop();
        if (savedState) ({ fill, stroke, renderMode } = savedState);
      } else if (fn === OPS.setFillRGBColor) fill = Array.from(args, colorByte);
      else if (fn === OPS.setStrokeRGBColor) stroke = Array.from(args, colorByte);
      else if (fn === OPS.setFillColor) fill = Array.from(args, colorByte);
      else if (fn === OPS.setStrokeColor) stroke = Array.from(args, colorByte);
      else if (fn === OPS.setFillCMYKColor) {
        const [cyan, magenta, yellow, black] = args;
        fill = [
          (1 - cyan) * (1 - black),
          (1 - magenta) * (1 - black),
          (1 - yellow) * (1 - black)
        ].map(value => Math.round(value * 255));
      } else if (fn === OPS.setFillGray) {
        const gray = Math.round(args[0] * 255);
        fill = [gray, gray, gray];
      } else if (fn === OPS.setStrokeGray) {
        const gray = Math.round(args[0] * 255);
        stroke = [gray, gray, gray];
      } else if (fn === OPS.setStrokeCMYKColor) {
        const [cyan, magenta, yellow, black] = args;
        stroke = [
          (1 - cyan) * (1 - black),
          (1 - magenta) * (1 - black),
          (1 - yellow) * (1 - black)
        ].map(value => Math.round(value * 255));
      } else if (fn === OPS.setTextRenderingMode) {
        renderMode = Number(args?.[0]) || 0;
      } else if (
        fn === OPS.showText || fn === OPS.showSpacedText ||
        fn === OPS.nextLineShowText || fn === OPS.nextLineSetSpacingShowText
      ) {
        operations.push({
          index: operations.length,
          text: glyphText(fn === OPS.nextLineSetSpacingShowText ? args?.[2] : args?.[0]),
          color: rgbHex(fill),
          stroke: rgbHex(stroke),
          renderMode
        });
      }
    }
  } catch (error) {
    if (import.meta.env?.DEV) console.warn('PDF 颜色解析失败', error);
  }

  const itemOperations = mapItemsToShowOperations(textContent.items, operations);
  const effectiveColor = (operation, appearance) => {
    if (!operation) return '#000000';
    const outlineModes = new Set([1, 2, 5, 6]);
    const whiteFill = operation.color === '#FFFFFF';
    const visibleStroke = operation.stroke && operation.stroke !== '#FFFFFF';
    let color = outlineModes.has(operation.renderMode) || (whiteFill && visibleStroke)
      ? operation.stroke
      : operation.color;
    const renderedColor = appearance?.color;
    if (renderedColor && appearance?.operatorMatches === false && renderedColor !== color) {
      color = renderedColor;
    }
    return color;
  };
  const isHorizontal = item =>
    Math.hypot(item.transform[1], item.transform[2]) <=
    Math.hypot(item.transform[0], item.transform[3]) * 0.035;
  const horizontalItemIndices = new Set(textContent.items
    .map((item, index) => isHorizontal(item) ? index : -1)
    .filter(index => index >= 0));
  const preliminaryFragments = [];
  textContent.items.forEach((item, itemIndex) => {
    if (!item.str || !item.str.trim() || !horizontalItemIndices.has(itemIndex)) return;
    const size = Math.hypot(item.transform[2], item.transform[3]) ||
      Math.abs(item.transform[3]) || 10;
    preliminaryFragments.push({
      x: item.transform[4],
      baseline: item.transform[5],
      width: item.width || 0,
      height: item.height || size,
      text: item.str
    });
  });
  const structureText = await collectStructureText(page, preliminaryFragments);
  const unreliableFonts = unreliableFontNames(options.libDocument, pageNumber - 1);
  const allAppearanceItems = textContent.items.concat(structureText.items);
  const renderedAppearance = await analyzeRenderedTextItems(
    page, allAppearanceItems, itemOperations.map(indices => operations[indices?.[0]]?.color));
  const fragments = [];
  let unreliableTextCount = 0;
  textContent.items.forEach((item, itemIndex) => {
    if (!item.str || !item.str.trim() || !horizontalItemIndices.has(itemIndex)) return;
    const size = Math.hypot(item.transform[2], item.transform[3]) ||
      Math.abs(item.transform[3]) || 10;
    const sourceOperations = [...new Set(itemOperations[itemIndex] || [])]
      .filter(Number.isInteger).sort((left, right) => left - right);
    const sourceRenderModes = sourceOperations.map(operationIndex => operations[operationIndex]?.renderMode);
    const syntheticBold = sourceOperations.some(operationIndex => {
      const operation = operations[operationIndex];
      return [2, 5].includes(operation?.renderMode) && operation?.stroke === operation?.color;
    });
    if (sourceRenderModes.length && sourceRenderModes.every(mode => mode === 3 || mode === 7)) return;
    const appearance = renderedAppearance[itemIndex];
    if (appearance && !appearance.visible) return;
    const font = fontInfo(page, item.fontName);
    if (textLooksUnreliable(item.str) || (font.name && unreliableFonts.has(font.name))) {
      unreliableTextCount++;
      return;
    }
    fragments.push({
      x: item.transform[4],
      baseline: item.transform[5],
      width: item.width || 0,
      height: item.height || size,
      size,
      rotated: !isHorizontal(item),
      text: item.str,
      font: { ...font, bold: font.bold || syntheticBold },
      color: effectiveColor(operations[itemOperations[itemIndex]?.[0]], appearance) ||
        appearance?.color || '#000000',
      sourceOps: sourceOperations
    });
  });

  structureText.items.forEach((item, offset) => {
    const appearance = renderedAppearance[textContent.items.length + offset];
    if (!appearance?.visible) return;
    const size = Math.hypot(item.transform[2], item.transform[3]);
    const fontName = item.structureFontName;
    fragments.push({
      x: item.transform[4],
      baseline: item.transform[5],
      width: item.width,
      height: item.height,
      size,
      rotated: false,
      text: item.str,
      font: { name: fontName, bold: false, key: null, data: null },
      color: appearance.color || '#161616',
      sourceOps: []
    });
  });

  fragments.sort((left, right) => left.baseline - right.baseline || left.x - right.x);
  const rawLines = [];
  let currentLine = [];
  for (const fragment of fragments) {
    const reference = currentLine[0];
    const tolerance = Math.max(reference?.size || fragment.size, fragment.size) * 0.45;
    if (reference && Math.abs(fragment.baseline - reference.baseline) <= tolerance) {
      currentLine.push(fragment);
    } else {
      if (currentLine.length) rawLines.push(currentLine);
      currentLine = [fragment];
    }
  }
  if (currentLine.length) rawLines.push(currentLine);

  const lines = [];
  for (const rawLine of rawLines) {
    rawLine.sort((left, right) => left.x - right.x);
    for (const lineFragments of splitRawLine(rawLine)) {
      const runs = [];
      let previousRight = null;
      for (const fragment of lineFragments) {
        if (previousRight !== null && fragment.x - previousRight > fragment.size * 0.18) {
          appendRuns(runs, [{
            text: ' ',
            font: fragment.font.name || '微软雅黑',
            fontKey: fragment.font.key,
            loadedName: fragment.font.loadedName,
            fontData: fragment.font.data,
            size: fragment.size,
            color: fragment.color,
            bold: fragment.font.bold,
            x: previousRight,
            baseline: fragment.baseline,
            width: Math.max(fragment.x - previousRight, 0),
            height: fragment.height,
            sourceOps: []
          }]);
        }
        appendRuns(runs, [{
          text: fragment.text,
          font: fragment.font.name || '微软雅黑',
          fontKey: fragment.font.key,
          loadedName: fragment.font.loadedName,
          fontData: fragment.font.data,
          size: fragment.size,
          color: fragment.color,
          bold: fragment.font.bold,
          x: fragment.x,
          baseline: fragment.baseline,
          width: fragment.width || 0,
          height: fragment.height,
          sourceOps: fragment.sourceOps
        }]);
        previousRight = fragment.x + fragment.width;
      }
      while (runs.length && !runs[runs.length - 1].text.trim()) runs.pop();
      if (!runs.length) continue;

      const left = Math.min(...lineFragments.map(item => item.x));
      const right = Math.max(...lineFragments.map(item => item.x + item.width));
      const size = Math.max(...lineFragments.map(item => item.size));
      const baseline = lineFragments[0].baseline;
      const sourceOps = [...new Set(lineFragments.flatMap(item => item.sourceOps))]
        .filter(Number.isInteger).sort((leftValue, rightValue) => leftValue - rightValue);
      const id = `L${lines.length}`;
      const cells = [lineFragments].map(lineFragmentGroup => {
        const cellRuns = [];
        let cellPreviousRight = null;
        for (const fragment of lineFragmentGroup) {
          if (cellPreviousRight !== null && fragment.x - cellPreviousRight > fragment.size * 0.18) {
            cellRuns.push({
              text: ' ',
              font: fragment.font.name || '微软雅黑',
              fontKey: fragment.font.key,
              loadedName: fragment.font.loadedName,
              fontData: fragment.font.data,
              size: fragment.size,
              color: fragment.color,
              bold: fragment.font.bold,
              x: cellPreviousRight,
              baseline: fragment.baseline,
              width: Math.max(fragment.x - cellPreviousRight, 0),
              height: fragment.height,
              sourceOps: []
            });
          }
          cellRuns.push({
            text: fragment.text,
            font: fragment.font.name || '微软雅黑',
            fontKey: fragment.font.key,
            loadedName: fragment.font.loadedName,
            fontData: fragment.font.data,
            size: fragment.size,
            color: fragment.color,
            bold: fragment.font.bold,
            x: fragment.x,
            baseline: fragment.baseline,
            width: fragment.width || 0,
            height: fragment.height,
            sourceOps: [...fragment.sourceOps]
          });
          cellPreviousRight = fragment.x + fragment.width;
        }
        while (cellRuns.length && !cellRuns[cellRuns.length - 1].text.trim()) cellRuns.pop();
        if (!cellRuns.length) return null;
        const cellLeft = Math.min(...lineFragmentGroup.map(item => item.x));
        const cellRight = Math.max(...lineFragmentGroup.map(item => item.x + item.width));
        return {
          x: cellLeft,
          baseline: lineFragmentGroup[0].baseline,
          width: Math.max(cellRight - cellLeft, 1),
          size: Math.max(...lineFragmentGroup.map(item => item.size)),
          text: cellRuns.map(run => run.text).join(''),
          runs: cellRuns
        };
      }).filter(Boolean);
      lines.push({
        id,
        kind: 'line',
        x: left,
        baseline,
        width: Math.max(right - left, 1),
        size: Math.round(size * 10) / 10,
        lineHeight: Math.round(size * 1.22 * 10) / 10,
        ascent: 0.82,
        runs,
        cells,
        sourceLineRuns: [runs.map(run => ({ ...run, sourceOps: [...run.sourceOps] }))],
        sourceOps,
        text: runs.map(run => run.text).join('')
      });
    }
  }
  lines.sort((left, right) => right.baseline - left.baseline || left.x - right.x);

  const blocks = buildParagraphBlocks(lines, pageViewport.width);
  return {
    pageNumber,
    width: pageViewport.width,
    height: pageViewport.height,
    sourceTextCount: textContent.items.length,
    unreliableTextCount,
    structureFigureIds: structureText.ids,
    nativeBlocks: blocks,
    nativeLines: lines,
    blocks
  };
}

export function applySavedEdits(pageModel, savedEdits) {
  const editMap = new Map((savedEdits || []).map(edit => [edit.blockKey, edit]));
  for (const block of [...(pageModel.nativeBlocks || pageModel.blocks || []), ...(pageModel.nativeLines || [])]) {
    const saved = editMap.get(block.id);
    if (saved?.engine !== 2 || !Array.isArray(saved.runs)) continue;
    if (Number.isFinite(Number(saved.x))) block.x = Number(saved.x);
    if (Number.isFinite(Number(saved.baseline))) block.baseline = Number(saved.baseline);
    if (Number(saved.lineHeight) > 0) block.lineHeight = Number(saved.lineHeight);
    if (Number(saved.size) > 0) block.size = Number(saved.size);
    if (Number(saved.height) > 0) block.height = Number(saved.height);
    if (Number(saved.maxWidth) > 0) block.width = Number(saved.maxWidth);
    if (saved.alignment) block.alignment = saved.alignment;
    if (Number.isInteger(saved.lineCount)) block.lineCount = saved.lineCount;
    block.manualWidth = saved.manualWidth === true;
    block.runs = saved.runs.map(run => ({ ...run, sourceOps: [...block.sourceOps] }));
    block.text = block.runs.map(run => run.text).join('');
    block.sourceLineRuns = [block.runs.map(run => ({ ...run, sourceOps: [...run.sourceOps] }))];
    const blockSize = Number(block.size) || Number(block.runs?.[0]?.size) || 12;
    block.exportLines = [{
      x: Number(block.x) || 0,
      baseline: Number(block.baseline) || 0,
      width: Math.max(Number(block.width) || 1, 1),
      size: blockSize,
      runs: block.runs.map(run => ({ ...run, sourceOps: [...(run.sourceOps || [])] })),
      cells: [{
        x: Number(block.x) || 0,
        baseline: Number(block.baseline) || 0,
        width: Math.max(Number(block.width) || 1, 1),
        size: blockSize,
        text: block.text,
        runs: block.runs.map(run => ({ ...run, sourceOps: [...(run.sourceOps || [])] }))
      }]
    }];
    block.savedEdit = true;
  }
  return buildContentViews(pageModel);
}

export function buildContentViews(pageModel) {
  const lines = pageModel.nativeLines || [];
  const paragraphs = pageModel.nativeBlocks || pageModel.blocks || [];
  const allRuns = joinLineRuns(lines);
  const sourceOps = [...new Set(lines.flatMap(line => line.sourceOps))]
    .filter(Number.isInteger).sort((left, right) => left - right);
  const left = Math.min(...lines.map(line => line.x));
  const right = Math.max(...lines.map(line => line.x + line.width));
  const size = Math.max(...lines.map(line => Math.max(...line.runs.map(run => run.size), line.size)));
  const gaps = lines.slice(1).map((line, index) => line.baseline - lines[index].baseline)
    .filter(gap => gap > 0);
  const measuredLineHeight = weightedMedian(gaps.map(gap => ({ value: gap, weight: 1 })));
  const word = {
    id: 'WORD',
    kind: 'word',
    lineCount: lines.length,
    x: left,
    baseline: Math.max(...lines.map(line => line.baseline)),
    width: Math.max(right - left, 1),
    size: Math.round(size * 10) / 10,
    lineHeight: Math.round((measuredLineHeight || size * 1.22) * 10) / 10,
    height: Math.max(Math.max(...lines.map(line => line.baseline)) - Math.min(...lines.map(line => line.baseline)) + size * 1.22, size * 1.22),
    ascent: 0.82,
    runs: allRuns,
    sourceLineRuns: lines.map(line => line.runs.map(run => ({ ...run, sourceOps: [...run.sourceOps] }))),
    sourceOps,
    alignment: 'left',
    text: allRuns.map(run => run.text).join('')
  };
  pageModel.blocks = paragraphs;
  pageModel.views = {
    paragraph: paragraphs,
    line: lines,
    word: [word]
  };
  pageModel.mode = pageModel.mode || 'paragraph';
  pageModel.activeBlocks = pageModel.views[pageModel.mode] || paragraphs;
  return pageModel;
}
