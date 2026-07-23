import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, ImageRun, ExternalHyperlink, WidthType, BorderStyle,
  ShadingType, VerticalAlign, PageNumber,
} from 'docx';
import { createNumberingEngine } from './numberingEngine.js';
import { loadImageData } from './imageUtils.js';
import {
  mapAlignment, halfPointsFromPx, dxaFromPx, isBold, isItalic, isUnderline,
  isStrike, isAllCaps, verticalAlignRun, rgbToHex, fontFamilyPrimary,
} from './styleUtils.js';

// Using the spec-fixed numeric values instead of the ambient `Node` global:
// nodes here can come from an iframe's own realm, whose Node constructor is
// a different object than this module's global one.
const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

const BLOCK_TAGS = new Set([
  'DIV', 'P', 'TABLE', 'UL', 'OL', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
  'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'BLOCKQUOTE', 'HR', 'FIGURE',
]);
const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT']);

function hasBlockChild(el) {
  return Array.from(el.children).some((child) => BLOCK_TAGS.has(child.tagName));
}

function isVisible(win, el) {
  const style = win.getComputedStyle(el);
  return style.display !== 'none' && style.visibility !== 'hidden';
}

// --- header / footer detection -------------------------------------------

// SEI's letterhead is always the first meaningful block of the document: a
// centered wrapper whose only content is an <img>. SEI's own audit-trail
// footer ("Criado por ... em ...") is rendered as a non-selectable div — we
// key off that fingerprint rather than the (translatable) label text.
function detectHeaderImage(bodyEl, win) {
  for (const child of Array.from(bodyEl.children)) {
    if (!isVisible(win, child)) continue;
    const text = child.textContent.trim();
    const img = child.querySelector('img');
    if (img && text.length === 0) {
      return { container: child, img };
    }
    break; // only the very first meaningful block qualifies
  }
  return null;
}

function detectAuditFooter(bodyEl) {
  return bodyEl.querySelector('div[style*="user-select:none" i], div[style*="user-select: none" i]');
}

// SEI stamps its own process number ("00169.000874/2026-91") as plain text
// right after this document's id+version ("1045627v2") at the very end of
// the body — the same pair the native SEI-generated PDF turns into its
// running footer ("<título> SEI <processo> / pg. N"). We mimic that instead
// of leaving the raw pair sitting in the body.
//
// A process-number-shaped string can also show up earlier as ordinary
// content (e.g. a table row referencing a related process), so we can't
// just take the first match — we take the *last* one in document order,
// since SEI always appends the real identification block after everything
// else, including the signature.
const PROCESS_NUMBER_RE = /^\d{5}\.\d{6}\/\d{4}-\d{2}$/;

function detectIdentificationRow(bodyEl) {
  const leaves = Array.from(bodyEl.querySelectorAll('td, p, div, span'))
    .filter((el) => el.children.length === 0);

  let processEl = null;
  for (const el of leaves) {
    if (PROCESS_NUMBER_RE.test(el.textContent.trim())) processEl = el;
  }
  if (!processEl) return null;

  const wrapper = processEl.closest('table') || processEl.closest('tr') || processEl;
  return { processNumber: processEl.textContent.trim(), elementsToSkip: [wrapper] };
}

// --- inline (run) building -------------------------------------------------

function buildRunProps(win, el) {
  const style = win.getComputedStyle(el);
  const props = {
    bold: isBold(style),
    italics: isItalic(style),
    font: fontFamilyPrimary(style.fontFamily),
    size: halfPointsFromPx(parseFloat(style.fontSize) || 16),
    allCaps: isAllCaps(style),
  };
  if (isUnderline(style)) props.underline = {};
  if (isStrike(style)) props.strike = true;
  const vertAlign = verticalAlignRun(style);
  if (vertAlign === 'superscript') props.superScript = true;
  if (vertAlign === 'subscript') props.subScript = true;
  const color = rgbToHex(style.color);
  if (color && color !== '000000') props.color = color;
  return props;
}

function collectInlineRuns(win, node, runsOut) {
  if (node.nodeType === TEXT_NODE) {
    const text = node.textContent;
    if (text === '') return;
    const parentProps = buildRunProps(win, node.parentElement);
    runsOut.push(new TextRun({ text, ...parentProps }));
    return;
  }
  if (node.nodeType !== ELEMENT_NODE) return;
  if (SKIP_TAGS.has(node.tagName)) return;
  if (!isVisible(win, node)) return;

  if (node.tagName === 'BR') {
    runsOut.push(new TextRun({ text: '', break: 1 }));
    return;
  }

  if (node.tagName === 'A' && node.getAttribute('href')) {
    const hyperlinkRuns = [];
    Array.from(node.childNodes).forEach((child) => collectInlineRuns(win, child, hyperlinkRuns));
    if (hyperlinkRuns.length) {
      runsOut.push(new ExternalHyperlink({ link: node.getAttribute('href'), children: hyperlinkRuns }));
    }
    return;
  }

  Array.from(node.childNodes).forEach((child) => collectInlineRuns(win, child, runsOut));
}

// --- paragraph / table building --------------------------------------------

function buildParagraph(win, el, numberingEngine) {
  const style = win.getComputedStyle(el);
  const runs = [];

  const prefix = numberingEngine.resolvePrefix(Array.from(el.classList));
  if (prefix) {
    runs.push(new TextRun({
      text: `${prefix}\t`,
      font: fontFamilyPrimary(style.fontFamily),
      size: halfPointsFromPx(parseFloat(style.fontSize) || 16),
    }));
  }

  Array.from(el.childNodes).forEach((child) => collectInlineRuns(win, child, runs));

  if (runs.length === 0) {
    runs.push(new TextRun({ text: '' }));
  }

  const shading = style.backgroundColor && !style.backgroundColor.includes('0, 0, 0, 0') && style.backgroundColor !== 'transparent'
    ? { type: ShadingType.SOLID, color: rgbToHex(style.backgroundColor), fill: rgbToHex(style.backgroundColor) }
    : undefined;

  return new Paragraph({
    alignment: mapAlignment(style.textAlign),
    children: runs,
    shading,
  });
}

const CELL_BORDER = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const CELL_BORDERS = {
  top: CELL_BORDER, bottom: CELL_BORDER, left: CELL_BORDER, right: CELL_BORDER,
};

function buildTableCell(win, cellEl, numberingEngine) {
  const blocks = walkBlocks(win, cellEl, numberingEngine);
  const style = win.getComputedStyle(cellEl);
  const shading = style.backgroundColor && !style.backgroundColor.includes('0, 0, 0, 0') && style.backgroundColor !== 'transparent'
    ? { type: ShadingType.SOLID, color: rgbToHex(style.backgroundColor), fill: rgbToHex(style.backgroundColor) }
    : undefined;

  const width = cellEl.getBoundingClientRect().width;

  return new TableCell({
    children: blocks.length ? blocks : [new Paragraph({})],
    columnSpan: cellEl.colSpan > 1 ? cellEl.colSpan : undefined,
    rowSpan: cellEl.rowSpan > 1 ? cellEl.rowSpan : undefined,
    borders: CELL_BORDERS,
    shading,
    verticalAlign: VerticalAlign.CENTER,
    width: width ? { size: dxaFromPx(width), type: WidthType.DXA } : undefined,
  });
}

function buildTable(win, tableEl, numberingEngine) {
  const rows = Array.from(tableEl.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr'));
  const tableRows = rows.map((rowEl) => new TableRow({
    children: Array.from(rowEl.children)
      .filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
      .map((cellEl) => buildTableCell(win, cellEl, numberingEngine)),
  }));

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

function walkBlocks(win, container, numberingEngine, skipSet = new Set()) {
  const blocks = [];
  let inlineBuffer = [];

  const flushInlineBuffer = () => {
    if (inlineBuffer.length === 0) return;
    const runs = [];
    inlineBuffer.forEach((node) => collectInlineRuns(win, node, runs));
    if (runs.some((r) => r)) blocks.push(new Paragraph({ children: runs }));
    inlineBuffer = [];
  };

  Array.from(container.childNodes).forEach((child) => {
    if (child.nodeType === TEXT_NODE) {
      if (child.textContent.trim() !== '') inlineBuffer.push(child);
      return;
    }
    if (child.nodeType !== ELEMENT_NODE) return;
    if (SKIP_TAGS.has(child.tagName)) return;
    if (skipSet.has(child)) return;
    if (!isVisible(win, child)) return;

    if (!BLOCK_TAGS.has(child.tagName)) {
      inlineBuffer.push(child);
      return;
    }

    flushInlineBuffer();

    if (child.tagName === 'TABLE') {
      blocks.push(buildTable(win, child, numberingEngine));
    } else if (child.tagName === 'HR') {
      blocks.push(new Paragraph({ text: '' }));
    } else if (hasBlockChild(child)) {
      blocks.push(...walkBlocks(win, child, numberingEngine, skipSet));
    } else {
      blocks.push(buildParagraph(win, child, numberingEngine));
    }
  });
  flushInlineBuffer();

  return blocks;
}

// --- header / footer sections ----------------------------------------------

// A4, 1" margins on every side — matches what SEI documents are laid out
// for and what docx.js defaults to when `page` isn't overridden below.
const PAGE_WIDTH_TWIPS = 11906;
const MARGIN_TWIPS = 1440;
const USABLE_WIDTH_PX = Math.floor(((PAGE_WIDTH_TWIPS - 2 * MARGIN_TWIPS) / 1440) * 96);

async function buildHeader(win, headerInfo) {
  if (!headerInfo) return undefined;
  const imageData = await loadImageData(headerInfo.img);
  if (!imageData) return undefined;

  // The letterhead's on-page width in the SEI viewer reflects the browser
  // viewport, not the printable area of an A4 page — inserted verbatim it
  // overflows past the margins. Scale it down (never up) to fit between
  // the margins, the same way the SEI-generated PDF renders it.
  let { width, height } = imageData;
  if (width > USABLE_WIDTH_PX) {
    const scale = USABLE_WIDTH_PX / width;
    width = USABLE_WIDTH_PX;
    height = Math.round(height * scale);
  }

  return new Header({
    children: [new Paragraph({
      alignment: mapAlignment('center'),
      children: [new ImageRun({
        data: imageData.data,
        type: imageData.type,
        transformation: { width, height },
      })],
    })],
  });
}

function buildFooter(documentTitle, processNumber) {
  if (!documentTitle && !processNumber) return undefined;
  const label = [documentTitle, processNumber ? `SEI ${processNumber}` : null]
    .filter(Boolean)
    .join(' ');

  return new Footer({
    children: [new Paragraph({
      alignment: mapAlignment('center'),
      children: [
        new TextRun({
          text: `${label} / pg. `, font: 'Arial', size: 16, color: '595959',
        }),
        new TextRun({
          children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: '595959',
        }),
      ],
    })],
  });
}

// --- public entry point ------------------------------------------------------

export async function convertSeiDocumentToDocx(iframeDoc, iframeWin, documentTitle) {
  // Walk the *live* body (not a detached clone) so getComputedStyle /
  // getBoundingClientRect resolve real cascaded values — including
  // CSS-class based formatting SEI defines in the iframe's own <head>,
  // which a plain innerHTML copy would lose, and which a clone can't see
  // either since getComputedStyle only works on nodes attached to a
  // rendered document.
  const liveBody = iframeDoc.body;

  const headerInfo = detectHeaderImage(liveBody, iframeWin);
  const auditFooterEl = detectAuditFooter(liveBody);
  const identification = detectIdentificationRow(liveBody);

  const numberingEngine = createNumberingEngine();

  // Whatever we routed to the header/footer must be skipped while walking
  // the body, so it renders once, in the correct place, instead of twice.
  const skipSet = new Set();
  if (headerInfo) skipSet.add(headerInfo.container);
  if (auditFooterEl) skipSet.add(auditFooterEl);
  identification?.elementsToSkip.forEach((el) => skipSet.add(el));

  const bodyBlocks = walkBlocks(iframeWin, liveBody, numberingEngine, skipSet);

  const header = await buildHeader(iframeWin, headerInfo);
  const footer = buildFooter(documentTitle, identification?.processNumber);

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_WIDTH_TWIPS, height: 16838 },
          margin: {
            top: MARGIN_TWIPS, bottom: MARGIN_TWIPS, left: MARGIN_TWIPS, right: MARGIN_TWIPS,
          },
        },
      },
      headers: header ? { default: header } : undefined,
      footers: footer ? { default: footer } : undefined,
      children: bodyBlocks.length ? bodyBlocks : [new Paragraph({})],
    }],
  });

  return Packer.toBlob(doc);
}
