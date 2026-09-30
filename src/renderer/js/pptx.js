/**
 * OnyxPDF - 最小 PPTX 生成器（ZIP store，每页 = 整页图片幻灯片，版式保留）
 */
import { zipStore } from './docx.js';
import { exportFontForRun } from './embedded-fonts.js';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** pagePngs = [{dataUrl,w,h}] → Uint8Array (.pptx) */
export function buildPptx(pagePngs) {
  const first = pagePngs[0];
  const CX = 12192000; // 16:9 宽
  const CY = Math.round(CX * first.h / first.w);

  // presentation.xml
  const sldIds = pagePngs.map((_, i) =>
    `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('');
  const presentationXml = XML_HEADER +
`<p:presentation ${NS}>
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>${sldIds}</p:sldIdLst>
<p:sldSz cx="${CX}" cy="${CY}"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

  // presentation rels: rId1=master, rId2..=slides
  const presRels = XML_HEADER +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${pagePngs.map((_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('\n')}
</Relationships>`;

  // slideMaster
  const masterXml = XML_HEADER +
`<p:sldMaster ${NS}>
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

  const masterRels = XML_HEADER +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

  // slideLayout（空白）
  const layoutXml = XML_HEADER +
`<p:sldLayout ${NS} type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

  const layoutRels = XML_HEADER +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

  // theme（最小 Office 主题）
  const themeXml = XML_HEADER +
`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
<a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements></a:theme>`;

  const files = [
    { name: '[Content_Types].xml', data: null }, // 占位，最后填
    { name: '_rels/.rels', data: new TextEncoder().encode(XML_HEADER +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`) },
    { name: 'ppt/presentation.xml', data: new TextEncoder().encode(presentationXml) },
    { name: 'ppt/_rels/presentation.xml.rels', data: new TextEncoder().encode(presRels) },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: new TextEncoder().encode(masterXml) },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: new TextEncoder().encode(masterRels) },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: new TextEncoder().encode(layoutXml) },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: new TextEncoder().encode(layoutRels) },
    { name: 'ppt/theme/theme1.xml', data: new TextEncoder().encode(themeXml) }
  ];

  // 幻灯片与图片
  pagePngs.forEach((pg, i) => {
    const n = i + 1;
    const slideXml = XML_HEADER +
`<p:sld ${NS}>
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Page${n}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
<p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${CX}" cy="${CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic></p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
    files.push({ name: `ppt/slides/slide${n}.xml`, data: new TextEncoder().encode(slideXml) });
    files.push({ name: `ppt/slides/_rels/slide${n}.xml.rels`, data: new TextEncoder().encode(XML_HEADER +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${n}.png"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`) });

    const b64 = pg.dataUrl.split(',')[1];
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let k = 0; k < bin.length; k++) u8[k] = bin.charCodeAt(k);
    files.push({ name: `ppt/media/image${n}.png`, data: u8 });
  });

  // [Content_Types].xml
  const slideOverrides = pagePngs.map((_, i) =>
    `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('\n');
  const contentTypes = XML_HEADER +
`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${slideOverrides}
</Types>`;
  files[0] = { name: '[Content_Types].xml', data: new TextEncoder().encode(contentTypes) };

  return zipStore(files);
}


/* ==================== 可编辑文字版 PPTX ==================== */
/**
 * pageData = [{ lines: [{x,y,w,h,size,text}], width, height }]（PDF 顶部原点坐标）
 * 幻灯片尺寸 = 页面 pt 尺寸 * 12700 EMU；文本框按行定位（可编辑）。
 */
export function buildPptxEditable(pageData) {
  const EMU = 12700; // 1pt
  const first = pageData[0];
  const CX = Math.round(first.width * EMU);
  const CY = Math.round(first.height * EMU);
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fontGroups = new Map();
  for (const page of pageData || []) {
    for (const block of page.paragraphs || []) {
      for (const run of block.runs || []) {
        const font = exportFontForRun(run.font, page.embeddedFonts);
        if (!font?.bytes?.length) continue;
        const family = displayFontName(font.family);
        const group = fontGroups.get(family) || { regular: null, bold: null, italic: null };
        const slot = font.italic ? 'italic' : font.bold ? 'bold' : 'regular';
        if (!group[slot] || font.bytes.length > group[slot].bytes.length) group[slot] = font;
        fontGroups.set(family, group);
      }
    }
  }
  const fontRelationships = [];
  const fontDeclarations = [];
  const fontFiles = [];
  let fontSequence = 0;
  for (const [family, group] of fontGroups) {
    fontSequence += 1;
    const variants = [
      ['regular', group.regular, 'regular'],
      ['bold', group.bold, 'bold'],
      ['italic', group.italic, 'italic']
    ].filter(([, font]) => font?.bytes?.length);
    const variantXml = variants.map(([slot, font, element], index) => {
      const relationshipId = `rIdEmbeddedFont${fontSequence}_${slot}`;
      const fileName = `font${fontSequence}-${slot}.fntdata`;
      fontRelationships.push(
        `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${fileName}"/>`
      );
      fontFiles.push({ name: `ppt/fonts/${fileName}`, data: font.bytes });
      void index;
      return `<p:${element} r:id="${relationshipId}"/>`;
    }).join('');
    fontDeclarations.push(`<p:embeddedFont><p:font typeface="${esc(family)}"/>${variantXml}</p:embeddedFont>`);
  }
  const embeddedFontXml = fontDeclarations.length
    ? `<p:embeddedFontLst>${fontDeclarations.join('')}</p:embeddedFontLst>`
    : '';

  const sldIds = pageData.map((_, i) => '<p:sldId id="' + (256 + i) + '" r:id="rId' + (i + 2) + '"/>').join('');
  const presentationXml = XML_HEADER +
'<p:presentation ' + NS + ' embedTrueTypeFonts="1" saveSubsetFonts="1">' +
'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
'<p:sldIdLst>' + sldIds + '</p:sldIdLst>' +
'<p:sldSz cx="' + CX + '" cy="' + CY + '"/><p:notesSz cx="6858000" cy="9144000"/>' +
embeddedFontXml +
'</p:presentation>';

  const presRels = XML_HEADER +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
pageData.map((_, i) => '<Relationship Id="rId' + (i + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + (i + 1) + '.xml"/>').join('\n') +
fontRelationships.join('') +
'</Relationships>';

  const masterXml = XML_HEADER +
'<p:sldMaster ' + NS + '>' +
'<p:cSld><p:spTree>' +
'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
'</p:spTree></p:cSld>' +
'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
'<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
'</p:sldMaster>';

  const masterRels = XML_HEADER +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
'</Relationships>';

  const layoutXml = XML_HEADER +
'<p:sldLayout ' + NS + ' type="blank" preserve="1">' +
'<p:cSld name="Blank"><p:spTree>' +
'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
'</p:spTree></p:cSld>' +
'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
'</p:sldLayout>';

  const layoutRels = XML_HEADER +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
'</Relationships>';

  const themeXml = XML_HEADER +
'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
'<a:themeElements>' +
'<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
'<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
'<a:fmtScheme name="Office">' +
'<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
'<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
'<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
'<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
'</a:fmtScheme></a:themeElements></a:theme>';

  const xmlText = (value) => String(value ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const pptxColor = (value) => {
    const rgb = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(String(value || ''));
    if (rgb) return rgb.slice(1).map(item => Number(item).toString(16).padStart(2, '0')).join('').toUpperCase();
    return String(value || '#000000').replace('#', '').padEnd(6, '0').slice(0, 6).toUpperCase();
  };
  function displayFontName(value) {
    let font = String(value || '')
      .replace(/^[A-Z]{6}\+/, '')
      .replace(/["']/g, '')
      .replace(/[-_ ]?(?:Regular|Bold|Italic|Oblique|Light|Medium)$/i, '')
      .replace(/(?:Regular|Bold|Italic|Oblique|Light|Medium)$/i, '')
      .trim();
    const aliases = {
      'MicrosoftYaHei': 'Microsoft YaHei',
      'MicrosoftYaHeiUI': 'Microsoft YaHei UI',
      'KaiTi_GB2312': 'KaiTi',
      '楷体_GB2312': 'KaiTi',
      'Helvetica': 'Arial',
      'ArialMT': 'Arial',
      'Arial-BoldMT': 'Arial',
      'Arial-BoldItalicMT': 'Arial',
      'Arial-ItalicMT': 'Arial',
      'TimesNewRomanPSMT': 'Times New Roman',
      'TimesNewRomanPS-BoldMT': 'Times New Roman',
      'TimesNewRomanPS-ItalicMT': 'Times New Roman',
      'TimesNewRomanPS-BoldItalicMT': 'Times New Roman',
      'TimesNewRoman-BoldMT': 'Times New Roman',
      'TimesNewRoman-ItalicMT': 'Times New Roman',
      'TimesNewRoman-BoldItalicMT': 'Times New Roman',
      'CourierNewPSMT': 'Courier New',
      'CourierNewPS-BoldMT': 'Courier New',
      'CourierNewPS-ItalicMT': 'Courier New',
      'CourierNewPS-BoldItalicMT': 'Courier New',
      'CourierNew-BoldMT': 'Courier New',
      'CourierNew-ItalicMT': 'Courier New',
      'Wingdings-Regular': 'Wingdings'
    };
    return aliases[font] || aliases[font.replace(/[-_ ]/g, '')] || font || 'Microsoft YaHei';
  }
  const textLines = (runs) => {
    const lines = [[]];
    for (const run of runs || []) {
      const parts = String(run.text ?? '').split(/\r\n|\n|\r/);
      parts.forEach((text, index) => {
        if (index) lines.push([]);
        if (text) lines[lines.length - 1].push({ ...run, text });
      });
    }
    return lines;
  };
  const substitutionMeasureContext = typeof document === 'undefined'
    ? null
    : document.createElement('canvas').getContext('2d');
  const substitutionSpacingXml = (run, fontName, size) => {
    if (!substitutionMeasureContext || !run.text || !Number.isFinite(Number(run.width))) return '';
    const declaredWidth = Number(run.width);
    if (declaredWidth <= 0) return '';
    substitutionMeasureContext.font =
      `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${size}px "${fontName}"`;
    const substitutedWidth = substitutionMeasureContext.measureText(run.text).width;
    const tolerance = Math.max(0.5, declaredWidth * 0.04);
    if (Math.abs(declaredWidth - substitutedWidth) <= tolerance) return '';
    const spacing = Math.max(-3000, Math.min(5000, Math.round(
      (declaredWidth - substitutedWidth) / Math.max(1, run.text.length) * 100
    )));
    return ` spc="${spacing}"`;
  };
  const richParagraphXml = (block, page) => {
    const alignment = block.alignment === 'center' ? 'ctr' : block.alignment === 'right' ? 'r' : 'l';
    const spacing = Math.round((Number(block.lineHeight) || Number(block.size) * 1.22 || 14) * 100);
    let sourceLines = block.savedEdit || !Array.isArray(block.sourceLineRuns) || !block.sourceLineRuns.length
      ? textLines(block.runs)
      : block.sourceLineRuns;
    normalizeSeparatorRunColors(sourceLines);
    if (isVerticalTextBlock(block)) {
      const verticalLines = [];
      for (const lineRuns of sourceLines) {
        for (const run of lineRuns) {
          for (const character of Array.from(String(run.text || ''))) {
            if (!/\S/.test(character)) continue;
            verticalLines.push([{ ...run, text: character }]);
          }
        }
      }
      sourceLines = verticalLines;
    }
    const lineSpacingOverrides = sourceLines.map((lineRuns, lineIndex) => {
      if (!substitutionMeasureContext || lineRuns.length < 2 || isVerticalTextBlock(block)) return null;
      const declaredWidth = lineRuns.reduce((total, run) => total + (Number(run.width) || 0), 0);
      const substitutedWidth = lineRuns.reduce((total, run) => {
        const font = displayFontName(String(run.font || '').split(',')[0]
          .replace(/^[A-Z]{6}\+/, '').replace(/["']/g, '').trim());
        substitutionMeasureContext.font =
          `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${Number(run.size) || Number(block.size) || 12}px "${font}"`;
        return total + substitutionMeasureContext.measureText(run.text || '').width;
      }, 0);
      const characterCount = Math.max(1, lineRuns.reduce((total, run) =>
        total + String(run.text || '').length, 0));
      if (declaredWidth <= 0 || substitutedWidth <= declaredWidth * 1.04) return null;
      const overflowRatio = substitutedWidth / declaredWidth;
      if (overflowRatio < 1.08 || Number(block.x) + declaredWidth < Number(page.width) * 0.92) return null;
      return ` spc="${Math.max(-3000, Math.round((declaredWidth - substitutedWidth) / characterCount * 100))}"`;
    });
    const runXml = (run, runLineIndex) => {
      const embeddedFont = exportFontForRun(run.font, page?.embeddedFonts);
      const font = embeddedFont?.family ||
        String(run.font || '').split(',')[0].replace(/^[A-Z]{6}\+/, '').replace(/["']/g, '').trim() || 'Microsoft YaHei';
      const displayFont = displayFontName(font);
      const size = Math.max(100, Math.min(400000, Math.round((Number(run.size) || Number(block.size) || 12) * 100)));
      const substitutionSpacing = embeddedFont || isVerticalTextBlock(block) ? '' :
        lineSpacingOverrides[runLineIndex] ?? substitutionSpacingXml(run, displayFont, size / 100);
      return '<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="' + size + '"' +
        substitutionSpacing +
        (run.bold ? ' b="1"' : '') + (run.italic ? ' i="1"' : '') + ' dirty="0">' +
        '<a:solidFill><a:srgbClr val="' + pptxColor(run.color) + '"/></a:solidFill>' +
        '<a:latin typeface="' + xmlText(displayFont) + '"/><a:ea typeface="' + xmlText(displayFont) + '"/>' +
        (run.underline ? '<a:uSng/>' : '') + '</a:rPr><a:t>' + xmlText(run.text) + '</a:t></a:r>';
    };
    const content = sourceLines.map((lineRuns, lineIndex) =>
      (lineIndex ? '<a:br/>' : '') + lineRuns.map(run => runXml(run, lineIndex)).join('')
    ).join('');
    return '<a:p><a:pPr algn="' + alignment + '"><a:lnSpc><a:spcPts val="' + spacing + '"/></a:lnSpc></a:pPr>' +
      content + '<a:endParaRPr lang="zh-CN" dirty="0"/></a:p>';
  };
  const normalizeSeparatorRunColors = (lines) => {
    const flatRuns = lines.flat();
    flatRuns.forEach((run, index) => {
      const text = String(run.text || '');
      if (run.color?.toUpperCase() !== '#FFFFFF' || !/^[\s—–-]+$/.test(text)) return;
      const neighbor = [flatRuns[index - 1], flatRuns[index + 1]]
        .find(item => String(item?.text || '').trim() && item.color?.toUpperCase() !== '#FFFFFF');
      if (neighbor) run.color = neighbor.color;
    });
  };
  function isVerticalTextBlock(block) {
    const characterCount = String(block.text || '').replace(/\s+/g, '').length;
    const size = Number(block.size) || 12;
    const runHeight = Number(block.runs?.[0]?.height) || Number(block.frameHeight) || 0;
    return characterCount > 1 && Number(block.width) <= size * 1.05 &&
      runHeight >= characterCount * size * 0.72;
  }
  const pictureXml = (shapeId, name, relationshipId, x, y, width, height) =>
    '<p:pic><p:nvPicPr><p:cNvPr id="' + shapeId + '" name="' + xmlText(name) + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="' + relationshipId + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="' + Math.round(x * EMU) + '" y="' + Math.round(y * EMU) + '"/>' +
    '<a:ext cx="' + Math.max(1, Math.round(width * EMU)) + '" cy="' + Math.max(1, Math.round(height * EMU)) + '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  const dataUrlBytes = (dataUrl) => {
    const binary = atob(String(dataUrl || '').split(',')[1] || '');
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
    return bytes;
  };

  const files = [
    { name: '[Content_Types].xml', data: null },
    { name: '_rels/.rels', data: new TextEncoder().encode(XML_HEADER +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
'</Relationships>') },
    { name: 'ppt/presentation.xml', data: new TextEncoder().encode(presentationXml) },
    { name: 'ppt/_rels/presentation.xml.rels', data: new TextEncoder().encode(presRels) },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: new TextEncoder().encode(masterXml) },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: new TextEncoder().encode(masterRels) },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: new TextEncoder().encode(layoutXml) },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: new TextEncoder().encode(layoutRels) },
    { name: 'ppt/theme/theme1.xml', data: new TextEncoder().encode(themeXml) }
  ];

  pageData.forEach((pg, i) => {
    const n = i + 1;
    // 每行一个文本框（合并相近行为段落行）
    let shapes = '';
    const slideMedia = [];
    const slideRelationships = [];
    let shapeSequence = 1;

    if (Array.isArray(pg.paragraphs)) {
      if (pg.bgDataUrl) {
        const relationshipId = `rIdBg${n}`;
        const isJpeg = /^data:image\/jpe?g/i.test(pg.bgDataUrl || '');
        const mediaName = `background-${n}.${isJpeg ? 'jpg' : 'png'}`;
        slideRelationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`);
        slideMedia.push({ name: mediaName, data: dataUrlBytes(pg.bgDataUrl) });
        shapes += pictureXml(++shapeSequence, `PDF Background ${n}`, relationshipId, 0, 0, pg.width, pg.height);
      }
      if (pg.patchDataUrl) {
        const relationshipId = `rIdPatch${n}`;
        const mediaName = `text-patch-${n}.png`;
        slideRelationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`);
        slideMedia.push({ name: mediaName, data: dataUrlBytes(pg.patchDataUrl) });
        shapes += pictureXml(++shapeSequence, `PDF Text Patch ${n}`, relationshipId, 0, 0, pg.width, pg.height);
      }
      for (const cover of pg.coverShapes || []) {
        shapeSequence++;
        shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + shapeSequence + '" name="Text Cover ' + shapeSequence + '"/>' +
          '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
          '<p:spPr><a:xfrm><a:off x="' + Math.round(Number(cover.x) * EMU) + '" y="' + Math.round(Number(cover.y) * EMU) + '"/>' +
          '<a:ext cx="' + Math.round(Number(cover.width) * EMU) + '" cy="' + Math.round(Number(cover.height) * EMU) + '"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="' + String(cover.color || '#FFFFFF').replace('#', '').toUpperCase() + '"/></a:solidFill>' +
          '<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>';
      }
      for (const image of pg.images || []) {
        shapeSequence++;
        const relationshipId = `rIdContentImage${n}_${shapeSequence}`;
        const isJpeg = /^data:image\/jpe?g/i.test(image.dataUrl || '');
        const mediaName = `content-image-${n}-${shapeSequence}.${isJpeg ? 'jpg' : 'png'}`;
        slideRelationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`);
        slideMedia.push({ name: mediaName, data: dataUrlBytes(image.dataUrl) });
        shapes += pictureXml(shapeSequence, `Content Image ${n}-${shapeSequence}`, relationshipId,
          Number(image.x), Number(pg.height) - Number(image.y) - Number(image.h), Number(image.w), Number(image.h));
      }
      for (const block of pg.paragraphs || []) {
      shapeSequence++;
      const top = Math.max(0, Number(block.top) || 0);
      const originalWidth = Math.max(12, Number(block.width) || 48);
      const isVertical = isVerticalTextBlock(block);
      const verticalShapeWidth = isVertical
        ? Math.max(originalWidth, (Number(block.size) || 12) * 1.12)
        : 0;
      const left = Math.max(0, isVertical
        ? Math.min(
            Math.max(0, Number(pg.width) - verticalShapeWidth),
            (Number(block.x) || 0) - (verticalShapeWidth - originalWidth) / 2
          )
        : Number(block.x) || 0);
        const tolerance = Math.max(2, originalWidth * 0.015, (Number(block.size) || 12) * 0.25);
        const width = Math.max(12, Math.min(
          Number(pg.width) - left,
          (Number(block.lineCount) || 1) === 1 ? originalWidth : originalWidth + tolerance
        ));
        const isSingleLine = !isVertical && (Number(block.lineCount) || 1) === 1;
        const singleLineAscent = Number(block.size) * 0.18;
        const shapeTop = isSingleLine ? Math.max(0, top - singleLineAscent) : top;
        const verticalCharacterCount = isVertical
          ? String(block.text || '').replace(/\s+/g, '').length : 0;
        const height = Math.max(10,
          verticalCharacterCount
            ? verticalCharacterCount * Number(block.size) * 1.2
            : Number(block.frameHeight) || Number(block.lineHeight) || Number(block.size) * 1.22 || 14);
        const shapeWidth = isVertical ? verticalShapeWidth : width;
        shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + shapeSequence + '" name="Paragraph ' + shapeSequence + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
          '<p:spPr><a:xfrm><a:off x="' + Math.round(left * EMU) + '" y="' + Math.round(shapeTop * EMU) + '"/>' +
          '<a:ext cx="' + Math.round(shapeWidth * EMU) + '" cy="' + Math.round(height * EMU) + '"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
          '<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" rtlCol="0"' +
          '><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
          richParagraphXml(block, pg) + '</p:txBody></p:sp>';
      }
    } else {
    // 矢量色块 → PPT 原生形状（solidFill，可单独选中改色）
    (pg.rects || []).forEach((rc, ri) => {
      const px = Math.round(rc.x * EMU), py = Math.round(rc.y * EMU);
      const pw = Math.max(1, Math.round(rc.w * EMU)), ph = Math.max(1, Math.round(rc.h * EMU));
      shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + (200 + ri) + '" name="Shape' + ri + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
'<p:spPr><a:xfrm><a:off x="' + px + '" y="' + py + '"/><a:ext cx="' + pw + '" cy="' + ph + '"/></a:xfrm>' +
'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
'<a:solidFill><a:srgbClr val="' + rc.fill.replace('#', '').toUpperCase() + '"/></a:solidFill>' +
'<a:ln><a:noFill/></a:ln></p:spPr>' +
'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>';
    });
    // 内嵌图片（版式还原）
    (pg.images || []).forEach((im, ii) => {
      const rid = 'rIdImg' + ii;
      const px = Math.round(im.x * EMU), py = Math.round(im.y * EMU);
      const pw = Math.round(im.w * EMU), ph = Math.round(im.h * EMU);
      shapes += '<p:pic><p:nvPicPr><p:cNvPr id="' + (100 + ii) + '" name="Img' + ii + '"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
'<p:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
'<p:spPr><a:xfrm><a:off x="' + px + '" y="' + py + '"/><a:ext cx="' + pw + '" cy="' + ph + '"/></a:xfrm>' +
'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
    });
    pg.lines.forEach((line, li) => {
      const x = Math.max(0, Math.round((line.x - 2) * EMU));
      const y = Math.max(0, Math.round((line.y - 2) * EMU));
      const w = Math.round(Math.max(line.w + 6, 40) * EMU);
      const h = Math.round(Math.max(line.h + 4, line.size * 1.4) * EMU);
      const szC = Math.max(100, Math.min(4000, Math.round(line.size * 100)));
      const txtColor = (line.color || '#000000').replace('#', '').toUpperCase();
      const fMain = (line.font || 'Microsoft YaHei').replace(/["']/g, '');
      shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + (li + 2) + '" name="Line' + li + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
'<p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm>' +
'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
'<p:txBody><a:bodyPr wrap="none"/><a:lstStyle/>' +
'<a:p><a:r><a:rPr lang="zh-CN" sz="' + szC + '" dirty="0"><a:solidFill><a:srgbClr val="' + txtColor + '"/></a:solidFill>' +
'<a:latin typeface="' + esc(fMain) + '"/><a:ea typeface="' + esc(fMain) + '"/></a:rPr><a:t>' + esc(line.text) + '</a:t></a:r></a:p>' +
'</p:txBody></p:sp>';
    });
    if (!pg.lines.length) {
      shapes = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Empty"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
'<p:spPr><a:xfrm><a:off x="4572000" y="3200400"/><a:ext cx="3048000" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>';
    }
    }
    const slideXml = XML_HEADER +
'<p:sld ' + NS + '><p:cSld><p:spTree>' +
'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
shapes +
'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
    files.push({ name: 'ppt/slides/slide' + n + '.xml', data: new TextEncoder().encode(slideXml) });
    const imgRels = slideRelationships.join('');
    files.push({ name: 'ppt/slides/_rels/slide' + n + '.xml.rels', data: new TextEncoder().encode(XML_HEADER +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
imgRels +
'</Relationships>') });
    for (const media of slideMedia) {
      files.push({ name: 'ppt/media/' + media.name, data: media.data });
    }
  });

  const slideOverrides = pageData.map((_, i) =>
    '<Override PartName="/ppt/slides/slide' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>').join('\n');
  const contentTypes = XML_HEADER +
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
'<Default Extension="xml" ContentType="application/xml"/>' +
'<Default Extension="png" ContentType="image/png"/>' +
'<Default Extension="jpg" ContentType="image/jpeg"/>' +
'<Default Extension="fntdata" ContentType="application/x-fontdata"/>' +
'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
slideOverrides +
'</Types>';
  files[0] = { name: '[Content_Types].xml', data: new TextEncoder().encode(contentTypes) };

  return zipStore(files.concat(fontFiles));
}
