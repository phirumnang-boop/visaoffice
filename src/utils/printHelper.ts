/**
 * Universal A4 Document Printing Utility
 * 
 * Ensures that all Print buttons across the system:
 * 1. Correctly trigger the system/browser Print Dialog so the user can select their printer.
 * 2. Strictly format output on A4 paper (210mm x 297mm portrait or 297mm x 210mm landscape).
 * 3. Print ONLY the configured A4 paper document, omitting all headers, navigation bars,
 *    rulers, zoom toolbars, background shades, and action buttons.
 * 4. Strictly preserve the user's custom margins, tab stops, Khmer fonts, and page scaling.
 */

export interface PrintA4Options {
  /**
   * Page orientation: 'portrait' (default, 210x297mm) or 'landscape' (297x210mm)
   */
  orientation?: 'portrait' | 'landscape';
  /**
   * Scale factor for printing, e.g. 0.97 for 97%, 1.0 for 100%
   */
  scale?: number;
  /**
   * Title shown in the browser print header / destination PDF title
   */
  documentTitle?: string;
  /**
   * Page margin for @page rule. Defaults to '0' because official documents
   * already enforce their own exact padding (customMargins in cm).
   */
  pageMargin?: string;
  /**
   * Extra custom CSS to inject into the print context
   */
  extraCss?: string;
}

/**
 * Prints a specific document element strictly on A4 paper and opens the printer selection dialog.
 * 
 * @param elementOrId HTMLElement or ID string of the printable container
 * @param options Printing options (orientation, scale, title, margin)
 */
export function printA4Document(
  elementOrId: HTMLElement | string,
  options: PrintA4Options = {}
): boolean {
  const {
    orientation = 'portrait',
    scale,
    documentTitle,
    pageMargin = '0',
    extraCss = '',
  } = options;

  // Resolve target element
  const targetEl = typeof elementOrId === 'string'
    ? document.getElementById(elementOrId)
    : elementOrId;

  if (!targetEl) {
    console.warn(`[printA4Document] Target element not found: "${elementOrId}". Falling back to window.print()`);
    window.focus();
    window.print();
    return false;
  }

  // Update browser window title temporarily so the default PDF / Print job name is meaningful
  const originalTitle = document.title;
  if (documentTitle) {
    document.title = documentTitle;
  }

  try {
    // 1. Clone element to avoid touching the live active DOM
    const clone = targetEl.cloneNode(true) as HTMLElement;

    // 2. Strip interactive / screen-only controls in the clone
    clone.querySelectorAll('.print\\:hidden, button, aside, nav, [data-print-hide], .cursor-pointer').forEach((el) => {
      // Keep essential printable texts if any
      if (el.tagName.toLowerCase() === 'button' || el.classList.contains('print:hidden')) {
        el.remove();
      }
    });

    // 3. Remove contentEditable and screen-only transforms from the clone
    clone.setAttribute('contenteditable', 'false');
    clone.classList.remove('hidden');
    clone.style.transform = '';
    clone.style.boxShadow = 'none';

    // 4. Create an isolated hidden iframe
    const frameId = 'print-a4-isolated-iframe';
    const existing = document.getElementById(frameId);
    if (existing) {
      existing.remove();
    }

    const iframe = document.createElement('iframe');
    iframe.id = frameId;
    iframe.setAttribute(
      'style',
      'position:fixed;right:0;bottom:0;width:0;height:0;border:0;opacity:0;pointer-events:none;z-index:-9999;'
    );
    document.body.appendChild(iframe);

    const frameDoc = iframe.contentDocument || iframe.contentWindow?.document;
    if (!frameDoc) {
      throw new Error('Unable to access iframe document for printing');
    }

    // 5. Gather all style tags and stylesheets from the main document (Tailwind, Khmer fonts, etc.)
    const styleTags: string[] = [];
    document.querySelectorAll('link[rel="stylesheet"], style').forEach((node) => {
      styleTags.push(node.outerHTML);
    });

    const widthMm = orientation === 'landscape' ? '297mm' : '210mm';
    const minHeightMm = orientation === 'landscape' ? '210mm' : '297mm';

    // Calculate scale CSS
    const scaleValue = scale && scale > 0 ? scale : 1.0;
    const zoomRule = scaleValue !== 1.0 ? `zoom: ${scaleValue};` : '';
    const transformRule = scaleValue !== 1.0 ? `transform: scale(${scaleValue}); transform-origin: top center;` : '';

    const isolatedCss = `
      @page {
        size: A4 ${orientation};
        margin: ${pageMargin};
      }
      *, *::before, *::after {
        box-sizing: border-box !important;
      }
      html, body {
        margin: 0 !important;
        padding: 0 !important;
        width: 100% !important;
        min-width: ${widthMm} !important;
        height: auto !important;
        background: #ffffff !important;
        color: #000000 !important;
        overflow: visible !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        font-smooth: always;
        -webkit-font-smoothing: antialiased;
      }
      body {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: flex-start;
      }
      .print-a4-outer-container {
        width: ${widthMm} !important;
        min-height: ${minHeightMm} !important;
        margin: 0 auto !important;
        padding: 0 !important;
        background: #ffffff !important;
        box-sizing: border-box !important;
        ${zoomRule}
      }
      @supports not (zoom: 1) {
        .print-a4-outer-container {
          ${transformRule}
        }
      }
      /* Ensure no screen buttons or hidden UI show up in print */
      .print\\:hidden, button, aside, nav, header {
        display: none !important;
      }
      /* Official Khmer typography styles preservation */
      .font-moul {
        font-family: 'Khmer Mool1', 'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Moul', serif !important;
        font-weight: normal !important;
      }
      .font-siemreap {
        font-family: 'Khmer OS Siemreap', 'Siemreap', sans-serif !important;
      }
      .font-tacteing {
        font-family: 'Tacteing', 'Khmer OS Tacteing', serif !important;
      }
      ${extraCss}
    `;

    frameDoc.open();
    frameDoc.write(`
      <!DOCTYPE html>
      <html lang="km">
        <head>
          <meta charset="utf-8" />
          <meta name="viewport" content="width=device-width, initial-scale=1.0" />
          <title>${documentTitle || originalTitle || 'Print Document'}</title>
          ${styleTags.join('\n')}
          <style>${isolatedCss}</style>
        </head>
        <body>
          <div class="print-a4-outer-container">
            ${clone.outerHTML}
          </div>
        </body>
      </html>
    `);
    frameDoc.close();

    // 6. Execute Print on the isolated frame
    const executeFramePrint = () => {
      try {
        iframe.contentWindow?.focus();
        iframe.contentWindow?.print();
      } catch (printErr) {
        console.warn('Iframe print blocked, falling back to window.print():', printErr);
        performWindowPrintFallback(targetEl, scaleValue);
      } finally {
        // Restore document title
        if (documentTitle) {
          document.title = originalTitle;
        }
        // Cleanup iframe after spooling
        setTimeout(() => {
          try {
            if (iframe.parentNode) {
              iframe.parentNode.removeChild(iframe);
            }
          } catch {}
        }, 3000);
      }
    };

    // Wait for fonts to be ready
    if ((document as any).fonts && (document as any).fonts.ready) {
      (document as any).fonts.ready
        .then(() => setTimeout(executeFramePrint, 150))
        .catch(() => setTimeout(executeFramePrint, 250));
    } else {
      setTimeout(executeFramePrint, 250);
    }

    return true;
  } catch (err) {
    console.warn('Isolated print failed, executing direct window print fallback:', err);
    performWindowPrintFallback(targetEl, scale);
    if (documentTitle) {
      document.title = originalTitle;
    }
    return false;
  }
}

/**
 * Fallback to direct window printing if iframe printing is denied by sandboxing.
 */
function performWindowPrintFallback(targetEl: HTMLElement, scale?: number) {
  if (scale) {
    document.documentElement.style.setProperty('--robok-print-scale', `${scale}`);
  }
  document.body.classList.add('direct-print-mode');
  targetEl.classList.add('direct-print-target');

  window.focus();
  window.print();

  setTimeout(() => {
    document.body.classList.remove('direct-print-mode');
    targetEl.classList.remove('direct-print-target');
  }, 1000);
}
