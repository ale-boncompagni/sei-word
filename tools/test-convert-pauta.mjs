import { JSDOM } from 'jsdom';
import { readFileSync, writeFileSync } from 'node:fs';
import { convertSeiDocumentToDocx } from '../src/converter/htmlToDocx.js';

const html = readFileSync('/tmp/sei_docx_inspect/pauta_body.html', 'utf-8');
const dom = new JSDOM(html, { pretendToBeVisual: true, resources: 'usable' });
const iframeDoc = dom.window.document;
const iframeWin = dom.window;

const documentTitle = 'CAUPR - Despacho Pauta - Reunião #200 (1078195)';
const blob = await convertSeiDocumentToDocx(iframeDoc, iframeWin, documentTitle);
const buffer = Buffer.from(await blob.arrayBuffer());
writeFileSync('/tmp/test-output-pauta.docx', buffer);
console.log('wrote /tmp/test-output-pauta.docx', buffer.length, 'bytes');
