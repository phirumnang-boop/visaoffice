// Helper utilities for exporting DOM elements to PDF using html-to-image, html2canvas & jsPDF
// Converts modern CSS colors (oklab, oklch, color(srgb ...)) to standard RGB/RGBA for compatibility
import { toPng } from 'html-to-image';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

export function oklabToRgb(L: number, aVal: number, bVal: number, alpha: number = 1): string {
  // OKLAB -> LMS
  const l_ = L + 0.3963377774 * aVal + 0.2158037573 * bVal;
  const m_ = L - 0.1055613458 * aVal - 0.0638541728 * bVal;
  const s_ = L - 0.0894841775 * aVal - 0.1291986507 * bVal;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  // LMS -> Linear RGB
  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

  // Gamma correction for standard sRGB
  const gamma = (c: number) => {
    const clamped = Math.max(0, Math.min(1, c));
    return clamped <= 0.0031308
      ? 12.92 * clamped
      : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
  };

  const r = Math.round(gamma(rLin) * 255);
  const g = Math.round(gamma(gLin) * 255);
  const b = Math.round(gamma(bLin) * 255);

  if (alpha < 1) {
    return `rgba(${r}, ${g}, ${b}, ${Number(alpha.toFixed(3))})`;
  }
  return `rgb(${r}, ${g}, ${b})`;
}

export function convertModernCssColors(input: string): string {
  if (!input || typeof input !== 'string') return input;

  let result = input;

  // Convert oklch(...) to rgb/rgba
  if (result.includes('oklch')) {
    const oklchRegex = /oklch\(\s*([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:[\s,]+)([-\d.deg%radturn]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(oklchRegex, (fullMatch, lStr, cStr, hStr, aStr) => {
      try {
        let L = parseFloat(lStr);
        if (lStr.endsWith('%')) L = L / 100;

        let C = parseFloat(cStr);
        if (cStr.endsWith('%')) C = C / 100;

        let H = parseFloat(hStr.replace(/(deg|rad|turn)/gi, ''));
        if (hStr.toLowerCase().endsWith('rad')) H = (H * 180) / Math.PI;
        else if (hStr.toLowerCase().endsWith('turn')) H = H * 360;

        let alpha = 1;
        if (aStr !== undefined && aStr !== null) {
          alpha = parseFloat(aStr);
          if (aStr.endsWith('%')) alpha = alpha / 100;
        }

        if (isNaN(L) || isNaN(C) || isNaN(H)) return 'rgb(0, 0, 0)';

        const hRad = (H * Math.PI) / 180;
        const aVal = C * Math.cos(hRad);
        const bVal = C * Math.sin(hRad);

        return oklabToRgb(L, aVal, bVal, alpha);
      } catch {
        return 'rgb(0, 0, 0)';
      }
    });
  }

  // Convert oklab(...) to rgb/rgba
  if (result.includes('oklab')) {
    const oklabRegex = /oklab\(\s*([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(oklabRegex, (fullMatch, lStr, aStr, bStr, alphaStr) => {
      try {
        let L = parseFloat(lStr);
        if (lStr.endsWith('%')) L = L / 100;

        let aVal = parseFloat(aStr);
        if (aStr.endsWith('%')) aVal = aVal / 100;

        let bVal = parseFloat(bStr);
        if (bStr.endsWith('%')) bVal = bVal / 100;

        let alpha = 1;
        if (alphaStr !== undefined && alphaStr !== null) {
          alpha = parseFloat(alphaStr);
          if (alphaStr.endsWith('%')) alpha = alpha / 100;
        }

        if (isNaN(L) || isNaN(aVal) || isNaN(bVal)) return 'rgb(0, 0, 0)';

        return oklabToRgb(L, aVal, bVal, alpha);
      } catch {
        return 'rgb(0, 0, 0)';
      }
    });
  }

  // Convert color(srgb ...) / color(display-p3 ...)
  if (result.includes('color(')) {
    const colorRegex = /color\(\s*(?:srgb|srgb-linear|display-p3)\s+([-\d.%]+)\s+([-\d.%]+)\s+([-\d.%]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(colorRegex, (_full, rStr, gStr, bStr, aStr) => {
      try {
        let r = parseFloat(rStr);
        if (rStr.endsWith('%')) r = r / 100;
        let g = parseFloat(gStr);
        if (gStr.endsWith('%')) g = g / 100;
        let b = parseFloat(bStr);
        if (bStr.endsWith('%')) b = b / 100;

        let a = 1;
        if (aStr !== undefined && aStr !== null) {
          a = parseFloat(aStr);
          if (aStr.endsWith('%')) a = a / 100;
        }

        const r255 = Math.round(Math.max(0, Math.min(1, r)) * 255);
        const g255 = Math.round(Math.max(0, Math.min(1, g)) * 255);
        const b255 = Math.round(Math.max(0, Math.min(1, b)) * 255);

        if (a < 1) {
          return `rgba(${r255}, ${g255}, ${b255}, ${Number(a.toFixed(3))})`;
        }
        return `rgb(${r255}, ${g255}, ${b255})`;
      } catch {
        return 'rgb(0, 0, 0)';
      }
    });
  }

  // Final emergency cleanup for any leftover unparsed oklab / oklch functions
  if (result.includes('oklch(') || result.includes('oklab(') || result.includes('color(')) {
    result = result
      .replace(/oklch\([^)]+\)/gi, 'rgb(0, 0, 0)')
      .replace(/oklab\([^)]+\)/gi, 'rgb(0, 0, 0)')
      .replace(/color\([^)]+\)/gi, 'rgb(0, 0, 0)');
  }

  return result;
}

export const COLOR_CSS_PROPERTIES = [
  'color',
  'background-color',
  'border-color',
  'border-top-color',
  'border-right-color',
  'border-bottom-color',
  'border-left-color',
  'outline-color',
  'fill',
  'stroke',
  'box-shadow',
  'text-decoration-color',
  'column-rule-color',
  'caret-color',
];

export function sanitizeDocumentForHtml2Canvas(clonedDoc: Document, targetElementId?: string): void {
  if (!clonedDoc) return;

  if (clonedDoc.body) {
    clonedDoc.body.style.setProperty('-webkit-font-smoothing', 'antialiased');
    clonedDoc.body.style.textRendering = 'geometricPrecision';
  }

  const win = clonedDoc.defaultView || window;

  // 1. Sanitize all <style> tags in the cloned document
  const styleTags = Array.from(clonedDoc.getElementsByTagName('style'));
  styleTags.forEach((style) => {
    if (
      style.textContent &&
      (style.textContent.includes('oklch') ||
        style.textContent.includes('oklab') ||
        style.textContent.includes('color('))
    ) {
      style.textContent = convertModernCssColors(style.textContent);
    }
  });

  // 2. Sanitize all elements in the cloned document
  const allElements = Array.from(clonedDoc.querySelectorAll('*'));
  allElements.forEach((el) => {
    const htmlEl = el as HTMLElement;

    // Inline style attribute
    const styleAttr = htmlEl.getAttribute('style');
    if (
      styleAttr &&
      (styleAttr.includes('oklch') || styleAttr.includes('oklab') || styleAttr.includes('color('))
    ) {
      htmlEl.setAttribute('style', convertModernCssColors(styleAttr));
    }

    // Computed styles
    try {
      const computed = win.getComputedStyle(htmlEl);
      COLOR_CSS_PROPERTIES.forEach((prop) => {
        const val = computed.getPropertyValue(prop);
        if (
          val &&
          (val.includes('oklab') || val.includes('oklch') || val.includes('color('))
        ) {
          const converted = convertModernCssColors(val);
          htmlEl.style.setProperty(prop, converted, 'important');
        }
      });
    } catch {
      // ignore
    }
  });

  // 3. Clean any legacy (ដុំ) or outdated wording in text nodes
  const textWalk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE && node.nodeValue) {
      if (node.nodeValue.includes('(ដុំ)') || node.nodeValue.includes('( ដុំ )') || node.nodeValue.includes('（ដុំ）')) {
        node.nodeValue = node.nodeValue
          .replace(/\s*\(\s*ដុំ\s*\)/g, '')
          .replace(/（ដុំ）/g, '');
      }
      if (node.nodeValue.includes('មានកិត្តិយសសូមគោរព ជូន') || node.nodeValue.includes('មានកិត្តិយសសូមគោរពជូន')) {
        node.nodeValue = node.nodeValue.replace(/មានកិត្តិយសសូមគោរព\s*ជូន/g, 'មានកិត្តិយសសូមគោរពជម្រាបជូន');
      }
    } else {
      for (let i = 0; i < node.childNodes.length; i++) {
        textWalk(node.childNodes[i]);
      }
    }
  };
  textWalk(clonedDoc.body);

  // 3.1 Remove any stray row that only contains 'ដុំ' or '(ដុំ)'
  const rows = Array.from(clonedDoc.querySelectorAll('tr'));
  rows.forEach((tr) => {
    const text = (tr.textContent || '').trim();
    if (
      text === 'ដុំ' ||
      text === 'ដុំដុំដុំដុំ' ||
      text === '(ដុំ)(ដុំ)(ដុំ)(ដុំ)' ||
      /^(\s*\(?\s*ដុំ\s*\)?\s*)+$/.test(text)
    ) {
      tr.remove();
    }
  });

  // 4. Sanitize tables to prevent html2canvas collapsed border artifacts (horizontal lines through text)
  const tables = Array.from(clonedDoc.querySelectorAll('table'));
  tables.forEach((table) => {
    table.style.setProperty('border-collapse', 'separate', 'important');
    table.style.setProperty('border-spacing', '0px', 'important');
    table.style.setProperty('border-top', '1px solid #000000', 'important');
    table.style.setProperty('border-left', '1px solid #000000', 'important');
    table.style.setProperty('border-right', 'none', 'important');
    table.style.setProperty('border-bottom', 'none', 'important');

    const cells = Array.from(table.querySelectorAll('th, td'));
    cells.forEach((cell) => {
      const htmlCell = cell as HTMLElement;
      htmlCell.style.setProperty('border-right', '1px solid #000000', 'important');
      htmlCell.style.setProperty('border-bottom', '1px solid #000000', 'important');
      htmlCell.style.setProperty('border-top', 'none', 'important');
      htmlCell.style.setProperty('border-left', 'none', 'important');
      htmlCell.style.setProperty('box-sizing', 'border-box', 'important');
      htmlCell.style.setProperty('background-color', '#ffffff', 'important');
      htmlCell.style.setProperty('overflow', 'visible', 'important');
      htmlCell.style.setProperty('line-height', '1.25', 'important');
    });
  });

  // 5. If target element id provided, ensure parents don't clip bounds and text is crisp
  if (targetElementId) {
    const target = clonedDoc.getElementById(targetElementId);
    if (target) {
      target.style.textRendering = 'geometricPrecision';
      target.style.setProperty('-webkit-font-smoothing', 'antialiased');
      target.style.setProperty('-moz-osx-font-smoothing', 'grayscale');
      target.style.transform = 'none';

      let parent = target.parentElement;
      while (parent && parent !== clonedDoc.body) {
        parent.style.width = 'auto';
        parent.style.maxWidth = 'none';
        parent.style.overflow = 'visible';
        parent = parent.parentElement;
      }
    }
  }
}

export interface ExportPdfOptions {
  pixelRatio?: number;
  fitSinglePage?: boolean;
  orientation?: 'portrait' | 'landscape' | 'p' | 'l';
  adjustPagePercent?: number; // e.g. 97 for 97% normal size
  scale?: number; // e.g. 0.97
}

/**
 * Universal high-definition PDF exporter using browser native SVG foreignObject rasterization (html-to-image).
 * Completely eliminates html2canvas table strikethrough lines, misaligned borders, and Khmer font distortions.
 */
export async function exportElementToPdf(
  element: HTMLElement,
  filename: string,
  options?: ExportPdfOptions
): Promise<string> {
  if (document.fonts && document.fonts.ready) {
    await document.fonts.ready;
  }
  await new Promise((r) => setTimeout(r, 120));

  const isLandscape =
    options?.orientation === 'landscape' ||
    options?.orientation === 'l' ||
    (options?.orientation === undefined &&
      (element.classList.contains('w-[297mm]') ||
        element.style.width === '297mm' ||
        element.getAttribute('data-orientation') === 'landscape' ||
        (element.offsetWidth > 0 && element.offsetHeight > 0 && element.offsetWidth > element.offsetHeight * 1.15)));

  const targetWidth = isLandscape ? '297mm' : '210mm';
  const pdfOrientation: 'landscape' | 'portrait' = isLandscape ? 'landscape' : 'portrait';

  const attrAdjust = element.getAttribute('data-adjust-page');
  const scaleFactor = options?.adjustPagePercent !== undefined
    ? options.adjustPagePercent / 100
    : (options?.scale !== undefined
        ? options.scale
        : (attrAdjust ? parseFloat(attrAdjust) / 100 : 1.0));

  try {
    const dataUrl = await toPng(element, {
      quality: 1.0,
      pixelRatio: options?.pixelRatio || 3,
      backgroundColor: '#ffffff',
      cacheBust: true,
      skipFonts: true,
      fontEmbedCSS: '',
      style: {
        width: targetWidth,
        minWidth: targetWidth,
        maxWidth: targetWidth,
        transform: 'none',
        margin: '0',
        boxSizing: 'border-box',
      },
      filter: (node) => {
        if (node instanceof HTMLElement) {
          if (
            node.classList.contains('no-print') ||
            node.classList.contains('print:hidden') ||
            node.getAttribute('data-no-print') === 'true'
          ) {
            return false;
          }
        }
        return true;
      },
    });

    const pdf = new jsPDF({
      orientation: pdfOrientation,
      unit: 'mm',
      format: 'a4',
      compress: true,
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    const img = new Image();
    img.src = dataUrl;
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = reject;
    });

    const imgWidth = img.naturalWidth || img.width;
    const imgHeight = img.naturalHeight || img.height;
    const calculatedHeight = (imgHeight * pageWidth) / imgWidth;

    const scaledWidth = pageWidth * scaleFactor;
    const scaledHeight = calculatedHeight * scaleFactor;
    const offsetX = (pageWidth - scaledWidth) / 2;
    const offsetY = 0;

    if (scaledHeight <= pageHeight + 25 || options?.fitSinglePage !== false) {
      pdf.addImage(dataUrl, 'PNG', offsetX, offsetY, scaledWidth, Math.min(pageHeight, scaledHeight), undefined, 'FAST');
    } else {
      let heightLeft = scaledHeight;
      let position = offsetY;
      pdf.addImage(dataUrl, 'PNG', offsetX, position, scaledWidth, scaledHeight, undefined, 'FAST');
      heightLeft -= pageHeight;
      while (heightLeft > 15) {
        position -= pageHeight;
        pdf.addPage();
        pdf.addImage(dataUrl, 'PNG', offsetX, position, scaledWidth, scaledHeight, undefined, 'FAST');
        heightLeft -= pageHeight;
      }
    }

    pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
    const pdfBlob = pdf.output('blob');
    return URL.createObjectURL(pdfBlob);
  } catch (err) {
    console.warn('html-to-image PDF export failed, falling back to html2canvas:', err);
    try {
      const canvas = await html2canvas(element, {
        scale: options?.pixelRatio || 3,
        useCORS: true,
        allowTaint: true,
        logging: false,
        backgroundColor: '#ffffff',
        windowWidth: isLandscape ? 1600 : 1200,
        windowHeight: isLandscape ? 1200 : 1600,
        imageTimeout: 0,
        onclone: (clonedDoc) => {
          sanitizeDocumentForHtml2Canvas(clonedDoc, element.id);
          const clonedElement = element.id ? clonedDoc.getElementById(element.id) : null;
          if (clonedElement) {
            clonedElement.style.width = targetWidth;
            clonedElement.style.minWidth = targetWidth;
            clonedElement.style.maxWidth = targetWidth;
            clonedElement.style.margin = '0 auto';
            clonedElement.style.boxShadow = 'none';
            clonedElement.style.backgroundColor = '#ffffff';
            clonedElement.style.textRendering = 'geometricPrecision';
            clonedElement.style.setProperty('-webkit-font-smoothing', 'antialiased');
            clonedElement.style.setProperty('-moz-osx-font-smoothing', 'grayscale');
            clonedElement.style.transform = 'none';
          }
        },
      });

      const imgData = canvas.toDataURL('image/png', 1.0);
      const pdf = new jsPDF({
        orientation: pdfOrientation,
        unit: 'mm',
        format: 'a4',
        compress: true,
      });

      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const calculatedHeight = (canvas.height * pageWidth) / canvas.width;

      const scaledWidth = pageWidth * scaleFactor;
      const scaledHeight = calculatedHeight * scaleFactor;
      const offsetX = (pageWidth - scaledWidth) / 2;
      const offsetY = 0;

      pdf.addImage(imgData, 'PNG', offsetX, offsetY, scaledWidth, Math.min(pageHeight, scaledHeight), undefined, 'FAST');
      pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
      const pdfBlob = pdf.output('blob');
      return URL.createObjectURL(pdfBlob);
    } catch (fallbackErr) {
      console.error('All PDF export methods failed:', fallbackErr);
      throw fallbackErr;
    }
  }
}

/**
 * Universal multi-page high-definition PDF exporter using browser native SVG foreignObject rasterization (html-to-image).
 * Renders each page element individually onto an exact A4 page, with pristine typography and border rendering.
 */
export async function exportElementsToPdf(
  elements: HTMLElement[],
  filename: string,
  options?: {
    pixelRatio?: number;
    orientation?: 'portrait' | 'landscape' | 'p' | 'l';
  }
): Promise<void> {
  if (document.fonts && document.fonts.ready) {
    await document.fonts.ready;
  }
  await new Promise((r) => setTimeout(r, 150));

  const isLandscape =
    options?.orientation === 'landscape' ||
    options?.orientation === 'l';
  const pdfOrientation: 'landscape' | 'portrait' = isLandscape ? 'landscape' : 'portrait';
  const targetWidth = isLandscape ? '297mm' : '210mm';
  const targetHeight = isLandscape ? '210mm' : '297mm';

  try {
    const pdf = new jsPDF({
      orientation: pdfOrientation,
      unit: 'mm',
      format: 'a4',
      compress: true,
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    for (let i = 0; i < elements.length; i++) {
      const el = elements[i];
      let dataUrl: string;

      try {
        dataUrl = await toPng(el, {
          quality: 1.0,
          pixelRatio: options?.pixelRatio || 3,
          backgroundColor: '#ffffff',
          cacheBust: true,
          skipFonts: true,
          fontEmbedCSS: '',
          style: {
            width: targetWidth,
            minWidth: targetWidth,
            maxWidth: targetWidth,
            minHeight: targetHeight,
            transform: 'none',
            margin: '0',
            boxShadow: 'none',
            border: 'none',
            borderRadius: '0',
            boxSizing: 'border-box',
            backgroundColor: '#ffffff',
          },
          filter: (node) => {
            if (node instanceof HTMLElement) {
              if (
                node.classList.contains('no-print') ||
                node.classList.contains('print:hidden') ||
                node.getAttribute('data-no-print') === 'true'
              ) {
                return false;
              }
            }
            return true;
          },
        });
      } catch (toPngErr) {
        console.warn(`html-to-image failed for page ${i + 1}, attempting html2canvas fallback:`, toPngErr);
        const canvas = await html2canvas(el, {
          scale: options?.pixelRatio || 3,
          useCORS: true,
          allowTaint: true,
          logging: false,
          backgroundColor: '#ffffff',
          windowWidth: isLandscape ? 1600 : 1200,
          windowHeight: isLandscape ? 1200 : 1600,
          imageTimeout: 0,
          onclone: (clonedDoc) => {
            sanitizeDocumentForHtml2Canvas(clonedDoc, el.id);
            const clonedEl = el.id ? clonedDoc.getElementById(el.id) : null;
            if (clonedEl) {
              clonedEl.style.width = targetWidth;
              clonedEl.style.minWidth = targetWidth;
              clonedEl.style.maxWidth = targetWidth;
              clonedEl.style.minHeight = targetHeight;
              clonedEl.style.margin = '0 auto';
              clonedEl.style.boxShadow = 'none';
              clonedEl.style.border = 'none';
              clonedEl.style.borderRadius = '0';
              clonedEl.style.transform = 'none';
              clonedEl.style.backgroundColor = '#ffffff';
            }
          },
        });
        dataUrl = canvas.toDataURL('image/png', 1.0);
      }

      const img = new Image();
      img.src = dataUrl;
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
      });

      const imgWidth = img.naturalWidth || img.width;
      const imgHeight = img.naturalHeight || img.height;
      const calculatedHeight = (imgHeight * pageWidth) / imgWidth;

      if (i > 0) {
        pdf.addPage('a4', pdfOrientation);
      }

      // If calculatedHeight is close to pageHeight (within 6mm), snap exactly to pageHeight
      const renderHeight =
        Math.abs(calculatedHeight - pageHeight) <= 6
          ? pageHeight
          : Math.min(pageHeight, calculatedHeight);

      pdf.addImage(dataUrl, 'PNG', 0, 0, pageWidth, renderHeight, undefined, 'FAST');
    }

    pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
  } catch (err) {
    console.error('Multi-page PDF export failed:', err);
    throw err;
  }
}

/**
 * Single page A4 landscape PDF exporter specifically tuned for yearly tabular reports.
 * Employs html-to-image with fallback to html2canvas, perfectly scaling to fit A4 landscape (297mm x 210mm).
 */
export async function exportSinglePageA4LandscapePdf(
  element: HTMLElement,
  filename: string,
  targetElementId?: string
): Promise<string> {
  if (!element) return '';
  const target = targetElementId ? ((document.getElementById(targetElementId) as HTMLElement) || element) : element;
  return exportElementToPdf(target, filename, {
    orientation: 'landscape',
    fitSinglePage: true,
    pixelRatio: 3,
  });
}

