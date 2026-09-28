/**
 * OnyxPDF - annotation and native content-edit exporter.
 */
import { ANNOT_COLORS } from './utils.js';
import { applyContentEdits } from './pdf-text-writer.js';

export async function exportAnnotatedPdf(originalBytes, annotationsList, suggestedName, contentEdits = []) {
  if (!window.PDFLib) throw new Error('pdf-lib 未加载');
  const { PDFDocument, rgb, StandardFonts } = window.PDFLib;
  const pdfDoc = await PDFDocument.load(originalBytes);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const pages = pdfDoc.getPages();
  let drawn = 0;

  const contentChanged = await applyContentEdits(pdfDoc, contentEdits);
  if (contentChanged) drawn += contentEdits.length;

  for (const annotation of annotationsList || []) {
    const page = pages[(annotation.page || 1) - 1];
    if (!page) continue;
    const { width, height } = page.getSize();
    const put = (x, y) => [clampN(x, 0, width), clampN(y, 0, height)];
    const colorOf = key => {
      const color = ANNOT_COLORS[key] || ANNOT_COLORS.default;
      return rgb(color.rgb[0], color.rgb[1], color.rgb[2]);
    };

    if (annotation.type === 'highlight') {
      const color = colorOf(annotation.color);
      for (const rect of annotation.rects) {
        page.drawRectangle({
          x: Math.min(rect.x, rect.x2), y: Math.min(rect.y, rect.y2),
          width: Math.abs(rect.x2 - rect.x), height: Math.abs(rect.y2 - rect.y),
          color, opacity: 0.38
        });
        drawn++;
      }
    } else if (annotation.type === 'underline') {
      const color = colorOf(annotation.color);
      for (const rect of annotation.rects) {
        page.drawLine({
          start: { x: rect.x, y: rect.y },
          end: { x: rect.x2, y: rect.y },
          thickness: 1.1, color, opacity: 0.9
        });
        drawn++;
      }
    } else if (annotation.type === 'rect') {
      page.drawRectangle({
        x: Math.min(annotation.rect.x, annotation.rect.x2),
        y: Math.min(annotation.rect.y, annotation.rect.y2),
        width: Math.abs(annotation.rect.x2 - annotation.rect.x),
        height: Math.abs(annotation.rect.y2 - annotation.rect.y),
        borderColor: rgb(1, 0.18, 0.576), borderWidth: 1.4,
        color: rgb(1, 0.18, 0.576), opacity: 0.05, borderOpacity: 1
      });
      drawn++;
    } else if (annotation.type === 'note') {
      const [nx, ny] = put(annotation.x, annotation.y);
      page.drawRectangle({ x: nx, y: ny, width: 12, height: 12, color: rgb(1, 0.69, 0.125), opacity: 0.95 });
      if (annotation.content) {
        page.drawText(annotation.content.replace(/\s+/g, ' ').slice(0, 46), {
          x: nx + 16, y: ny + 3, size: 8, font,
          color: rgb(0.15, 0.15, 0.18), opacity: 0.85
        });
      }
      drawn++;
    } else if (annotation.type === 'fillrect') {
      const [fx1, fy1] = put(annotation.rect.x, annotation.rect.y);
      const [fx2, fy2] = put(annotation.rect.x2, annotation.rect.y2);
      const color = annotation.color || '#000000';
      page.drawRectangle({
        x: Math.min(fx1, fx2), y: Math.min(fy1, fy2),
        width: Math.abs(fx2 - fx1), height: Math.abs(fy2 - fy1),
        color: rgb(parseInt(color.slice(1, 3), 16) / 255, parseInt(color.slice(3, 5), 16) / 255, parseInt(color.slice(5, 7), 16) / 255)
      });
      drawn++;
    } else if (annotation.type === 'image') {
      const image = /^data:image\/jpe?g/i.test(annotation.dataUrl)
        ? await page.doc.embedJpg(annotation.dataUrl)
        : await page.doc.embedPng(annotation.dataUrl);
      const [nx, ny] = put(annotation.x, annotation.y);
      page.drawImage(image, { x: nx, y: ny, width: annotation.w, height: annotation.h });
      drawn++;
    } else if (annotation.type === 'text') {
      if (annotation.text && annotation.text.trim()) {
        await drawNativeTextRuns(page, annotation.x, annotation.y, 100000, (annotation.size || 14) * 1.45, [{
          text: annotation.text, size: annotation.size || 14, color: annotation.color || '#d64545',
          font: annotation.font || 'Arial', bold: !!annotation.bold
        }]);
        drawn++;
      }
    }
  }

  return { bytes: await pdfDoc.save(), drawn };
}

function clampN(value, min, max) {
  return Math.min(max, Math.max(min, Number.isFinite(value) ? value : min));
}
