import { zipSync } from '../../../node_modules/fflate/esm/browser.js';

export function zipStore(files) {
  const entries = Object.fromEntries(files.map(file => [file.name, file.data]));
  return zipSync(entries, { level: 6 });
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function buildDocxParagraphLayout(pages) {
  const esc = (value) => String(value ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const FONT_MAP = {
    simsun: '宋体', nsimsun: '新宋体', simhei: '黑体', kaiti: '楷体',
    fangsong: '仿宋', dengxian: '等线', 'microsoft yahei': '微软雅黑',
    pingfangsc: '苹方', hiraginosansgb: '冬青黑体', stzhongsong: '华文中宋',
    stkaiti: '华文楷体', stxihei: '华文细黑', stfangsong: '华文仿宋'
  };
  const fontName = (value) => {
    let clean = String(value || '').split(',')[0].replace(/^[A-Z]{6}\+/, '')
      .replace(/["']/g, '').trim();
    clean = clean.replace(/-(?:Regular|Bold|Italic|Light|Medium)$/i, '')
      .replace(/(?:Regular|Bold|Italic|Light|Medium)$/i, '').trim();
    return FONT_MAP[clean.toLowerCase()] || clean || 'Microsoft YaHei';
  };
  const colorValue = (value) => {
    const rgb = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(String(value || ''));
    if (rgb) return rgb.slice(1).map(item => Number(item).toString(16).padStart(2, '0')).join('').toUpperCase();
    return String(value || '#000000').replace('#', '').padEnd(6, '0').slice(0, 6).toUpperCase();
  };
  const measureTextWidthPt = (runs) => {
    try {
      const context = document.createElement('canvas').getContext('2d');
      let width = 0;
      for (const run of runs || []) {
        const size = Math.max(1, Number(run.size) || 12);
        const font = fontName(run.font);
        context.font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${(size / 0.75).toFixed(2)}px "${font}"`;
        width += context.measureText(String(run.text ?? '')).width * 0.75;
      }
      return width;
    } catch (_) {
      return (runs || []).reduce((total, run) => {
        const size = Math.max(1, Number(run.size) || 12);
        return total + String(run.text ?? '').split('').reduce((width, char) =>
          width + (/[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u3000-\u303F]/.test(char) ? size : size * 0.56), 0);
      }, 0);
    }
  };
  const imageBytes = (dataUrl, name) => {
    const binary = atob(String(dataUrl).split(',')[1] || '');
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return { name, data: bytes };
  };
  const EMU_PER_PT = 12700;
  const emu = (value) => Math.max(0, Math.round((Number(value) || 0) * EMU_PER_PT));
  let drawingSequence = 0;

  const anchoredDrawing = ({ x = 0, y = 0, width, height, behindDoc = false, graphic }) => {
    drawingSequence += 1;
    const relativeHeight = behindDoc ? drawingSequence : 100000 + drawingSequence;
    return '<w:r><w:drawing>' +
      `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${relativeHeight}" ` +
      `behindDoc="${behindDoc ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
      '<wp:simplePos x="0" y="0"/>' +
      `<wp:positionH relativeFrom="page"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${emu(width)}" cy="${emu(height)}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
      `<wp:docPr id="${drawingSequence}" name="Onyx Object ${drawingSequence}"/>` +
      '<wp:cNvGraphicFramePr/>' +
      `<a:graphic><a:graphicData uri="${graphic.uri}">${graphic.xml}</a:graphicData></a:graphic>` +
      '</wp:anchor></w:drawing></w:r>';
  };

  const pictureDrawing = ({ x, y, width, height, relationshipId, name }) => {
    return anchoredDrawing({
      x,
      y,
      width,
      height,
      behindDoc: name.startsWith('PDF Background'),
      graphic: {
        uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        xml: '<pic:pic>' +
          `<pic:nvPicPr><pic:cNvPr id="${drawingSequence + 1}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
          `<pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
          '<pic:spPr><a:xfrm><a:off x="0" y="0"/>' +
          `<a:ext cx="${emu(width)}" cy="${emu(height)}"/></a:xfrm>` +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
      }
    });
  };

  const textboxDrawing = ({ x, y, width, height, alignment, lineHeightTwips, content, name }) => {
    return anchoredDrawing({
      x,
      y,
      width,
      height,
      graphic: {
        uri: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        xml: '<wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>' +
          '<wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
          `<a:ext cx="${emu(width)}" cy="${emu(height)}"/></a:xfrm>` +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
          '<a:ln><a:noFill/></a:ln></wps:spPr>' +
          `<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${lineHeightTwips}" w:lineRule="exact"/>` +
          `<w:jc w:val="${alignment}"/></w:pPr>${content}</w:p></w:txbxContent></wps:txbx>` +
          '<wps:bodyPr rot="0" spcFirstLastPara="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" ' +
          'numCol="1" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>'
      }
    });
  };

  const bodyParts = [];
  const mediaFiles = [];
  const relationships = [];
  const fontFamilies = new Set();
  let mediaSequence = 0;
  const pageWidthTwips = Math.round((Number(pages[0]?.width) || 595) * 20);
  const pageHeightTwips = Math.round((Number(pages[0]?.height) || 842) * 20);

  pages.forEach((page, pageIndex) => {
    const shapes = [];
    if (page.bgDataUrl) {
      mediaSequence++;
      const relationshipId = `rIdBg${pageIndex + 1}`;
      const mediaName = `media/bg${pageIndex + 1}.png`;
      relationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${mediaName}"/>`);
      mediaFiles.push(imageBytes(page.bgDataUrl, `word/${mediaName}`));
      shapes.push(pictureDrawing({
        x: 0,
        y: 0,
        width: Number(page.width),
        height: Number(page.height),
        relationshipId,
        name: `PDF Background ${pageIndex + 1}`
      }));
    }

    for (const image of page.images || []) {
      mediaSequence++;
      const relationshipId = `rIdImage${pageIndex + 1}_${mediaSequence}`;
      const isJpeg = /^data:image\/jpe?g/i.test(image.dataUrl || '');
      const mediaName = `media/content-image-${pageIndex + 1}-${mediaSequence}.${isJpeg ? 'jpg' : 'png'}`;
      relationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${mediaName}"/>`);
      mediaFiles.push(imageBytes(image.dataUrl, `word/${mediaName}`));
      const top = Number(page.height) - Number(image.y) - Number(image.h);
      shapes.push(pictureDrawing({
        x: Number(image.x),
        y: top,
        width: Number(image.w),
        height: Number(image.h),
        relationshipId,
        name: `PDF Image ${pageIndex + 1}-${mediaSequence}`
      }));
    }

    let textSequence = 0;
    for (const block of page.paragraphs || []) {
      textSequence++;
      const left = Math.max(0, Math.min(Number(block.x) || 0, page.width - 12));
      const top = Math.max(0, Math.min(Number(block.top) || 0, page.height - 12));
      const originalWidth = Math.max(12, Number(block.width) || 48);
      const alignment = block.alignment === 'center' ? 'center' : block.alignment === 'right' ? 'right' : 'left';
      const sourceLines = block.savedEdit || !Array.isArray(block.sourceLineRuns) || !block.sourceLineRuns.length
        ? [block.runs]
        : block.sourceLineRuns;
      const measuredWidth = Math.max(...sourceLines.map(lineRuns => measureTextWidthPt(lineRuns)), 0);
      const pageRightMargin = Math.max(18, page.width * 0.03);
      const requestedWidth = Math.max(originalWidth, measuredWidth + 2);
      const availableWidth = Math.max(12, page.width - left - pageRightMargin);
      const width = Math.min(requestedWidth, availableWidth);
      const wrappedLineCount = Math.max(
        sourceLines.length,
        sourceLines.reduce((total, lineRuns) => {
          const lineWidth = measureTextWidthPt(lineRuns);
          return total + Math.max(1, Math.ceil(lineWidth / Math.max(12, width)));
        }, 0)
      );
      const height = Math.max(
        Number(block.frameHeight) || Number(block.lineHeight) || 14,
        wrappedLineCount * (Number(block.lineHeight) || Number(block.size) * 1.22 || 14),
        10
      );
      const lineHeightTwips = Math.round((Number(block.lineHeight) || Number(block.size) * 1.22 || 14) * 20);

      let runXml = '';
      for (const lineRuns of sourceLines) {
        for (const run of lineRuns || []) {
          const parts = String(run.text ?? '').split(/\r\n|\n|\r/);
          parts.forEach((text, partIndex) => {
            if (partIndex) runXml += '<w:br/>';
            if (!text) return;
            const font = fontName(run.font);
            fontFamilies.add(font);
            runXml += '<w:r><w:rPr>' +
              `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:eastAsia="${esc(font)}"/>` +
              (run.bold ? '<w:b/>' : '') + (run.italic ? '<w:i/>' : '') +
              `<w:color w:val="${colorValue(run.color)}"/>` +
              `<w:sz w:val="${Math.max(10, Math.round((Number(run.size) || 12) * 2))}"/>` +
              (run.underline ? '<w:u w:val="single"/>' : '') +
              '</w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
          });
        }
        runXml += '<w:br/>';
      }
      runXml = runXml.replace(/<w:br\/>$/, '');

      shapes.push(textboxDrawing({
        x: left,
        y: top,
        width,
        height,
        alignment,
        lineHeightTwips,
        content: runXml,
        name: `PDF Text ${pageIndex + 1}-${textSequence}`
      }));
    }

    bodyParts.push(...shapes.map(shape =>
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr>' +
      shape + '</w:p>'
    ));
    if (pageIndex < pages.length - 1) bodyParts.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  });

  const documentXml = XML_HEADER +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<w:body>' + bodyParts.join('') +
    `<w:sectPr><w:pgSz w:w="${pageWidthTwips}" w:h="${pageHeightTwips}"/>` +
    '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>';

  const fontDeclarations = [...fontFamilies].map(family =>
    `<w:font w:name="${esc(family)}"><w:charset w:val="86"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>`
  ).join('');

  const fontTableXml = XML_HEADER +
    '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    fontDeclarations + '</w:fonts>';
  const settingsXml = XML_HEADER +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '</w:settings>';
  const contentTypes = XML_HEADER +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/>' +
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const rootRels = XML_HEADER +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const documentRels = XML_HEADER +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + relationships.join('') +
    '<Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>' +
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
    '</Relationships>';
  const encoder = new TextEncoder();
  return zipStore([
    { name: '[Content_Types].xml', data: encoder.encode(contentTypes) },
    { name: '_rels/.rels', data: encoder.encode(rootRels) },
    { name: 'word/document.xml', data: encoder.encode(documentXml) },
    { name: 'word/_rels/document.xml.rels', data: encoder.encode(documentRels) },
    { name: 'word/fontTable.xml', data: encoder.encode(fontTableXml) },
    { name: 'word/settings.xml', data: encoder.encode(settingsXml) },
    ...mediaFiles
  ]);
}
