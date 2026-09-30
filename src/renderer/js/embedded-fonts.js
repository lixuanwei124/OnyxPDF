function cleanBaseFont(value) {
  return String(value || '')
    .replace(/^[A-Z]{6}\+/, '')
    .replace(/["']/g, '')
    .trim();
}

function styleFlags(value) {
  const name = cleanBaseFont(value);
  return {
    bold: /(?:Bold|Heavy|Black|Semib?old)|(?:^|[-_ ])Bd(?: |$|-)/i.test(name),
    italic: /(?:Italic|Oblique)|(?:^|[-_ ])It(?: |$|-)/i.test(name)
  };
}

function fallbackFamily(value) {
  return cleanBaseFont(value)
    .replace(/-(?:Regular|Bold|Italic|Light|Medium)$/i, '')
    .replace(/(?:Regular|Bold|Italic|Light|Medium)$/i, '')
    .trim();
}

function fontSignature(bytes) {
  let hash = 2166136261;
  const step = Math.max(1, Math.floor(bytes.length / 4096));
  for (let index = 0; index < bytes.length; index += step) {
    hash ^= bytes[index];
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return `${bytes.length}-${hash}`;
}

export async function extractEmbeddedFonts(bytes, PDFLib, fontkit) {
  if (!PDFLib?.PDFDocument || !PDFLib.decodePDFRawStream) {
    throw new Error('PDF 字体提取组件不可用');
  }

  const document = await PDFLib.PDFDocument.load(bytes, {
    ignoreEncryption: true,
    throwOnInvalidObject: false,
    updateMetadata: false
  });
  const fonts = new Map();
  const byAlias = new Map();

  const addAlias = (alias, font) => {
    const key = cleanBaseFont(alias).toLowerCase();
    if (key && !byAlias.has(key)) byAlias.set(key, font);
  };

  for (const [, object] of document.context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFLib.PDFDict)) continue;
    const descriptor = object.lookupMaybe(PDFLib.PDFName.of('FontDescriptor'), PDFLib.PDFDict);
    if (!descriptor) continue;

    const baseFont = cleanBaseFont(object.get(PDFLib.PDFName.of('BaseFont'))?.decodeText());
    if (!baseFont) continue;

    for (const fileKey of ['FontFile2', 'FontFile3', 'FontFile']) {
      const stream = descriptor.lookupMaybe(PDFLib.PDFName.of(fileKey), PDFLib.PDFRawStream);
      if (!stream) continue;
      try {
        const data = PDFLib.decodePDFRawStream(stream).decode();
        if (!data?.length) continue;

        const styles = styleFlags(baseFont);
        let family = fallbackFamily(baseFont);
        let internalPostScriptName = baseFont;
        try {
          const parsed = fontkit?.create ? fontkit.create(data) : null;
          family = parsed?.familyName || family;
          internalPostScriptName = parsed?.postscriptName || internalPostScriptName;
          if (typeof parsed?.bold === 'boolean') styles.bold = parsed.bold;
          if (typeof parsed?.italic === 'boolean') styles.italic = parsed.italic;
        } catch (_) {
          // Subsetted or malformed name tables still contain usable font bytes.
        }

        const identity = `${family}|${styles.bold ? 'bold' : 'regular'}|${styles.italic ? 'italic' : 'regular'}`;
        const signature = fontSignature(data);
        const existing = fonts.get(identity);
        if (existing) {
          existing.aliases.add(baseFont);
          addAlias(baseFont, existing);
          if (data.length > existing.bytes.length && existing.signature !== signature) {
            existing.bytes = data;
            existing.signature = signature;
          }
          continue;
        }

        const font = {
          family,
          bold: !!styles.bold,
          italic: !!styles.italic,
          baseFont,
          postscriptName: internalPostScriptName,
          format: fileKey === 'FontFile2' ? 'truetype' : fileKey === 'FontFile3' ? 'cff' : 'raw',
          bytes: data,
          signature,
          aliases: new Set([baseFont])
        };
        fonts.set(identity, font);
        addAlias(baseFont, font);
      } catch (_) {
        // A damaged optional font must not block document export.
      }
    }
  }

  for (const font of fonts.values()) {
    for (const alias of font.aliases) addAlias(alias, font);
  }

  return {
    fonts: [...fonts.values()].sort((left, right) =>
      left.family.localeCompare(right.family) ||
      Number(right.bold) - Number(left.bold)
    ),
    byAlias
  };
}

export function exportFontForRun(font, embeddedFonts) {
  const key = cleanBaseFont(font).toLowerCase();
  const candidate = embeddedFonts?.byAlias?.get(key) || null;
  return hasUnicodeCharacterMap(candidate?.bytes) ? candidate : null;
}

function hasUnicodeCharacterMap(bytes) {
  if (!bytes || bytes.length < 12) return false;
  const view = bytes instanceof ArrayBuffer ? new DataView(bytes) : new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sfntVersion = view.getUint32(0);
  if (sfntVersion !== 0x00010000 &&
    sfntVersion !== 0x4F54544F &&
    sfntVersion !== 0x74727565) return false;
  const tableCount = view.getUint16(4);
  if (tableCount <= 0 || tableCount > 512 || bytes.length < 12 + tableCount * 16) return false;

  let hasName = false;
  let hasCmap = false;
  for (let index = 0; index < tableCount; index++) {
    const offset = 12 + index * 16;
    const tag = view.getUint32(offset);
    if (tag === 0x636D6170) hasCmap = true;
    if (tag === 0x6E616D65) hasName = true;
  }
  return hasName && hasCmap;
}
