import { JSDOM } from 'jsdom';
import { writeFileSync } from 'node:fs';
import { convertSeiDocumentToDocx } from '../src/converter/htmlToDocx.js';

const TINY_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

const html = `<!DOCTYPE html><html><head><style>
p.I05_Item_Nivel1_11 { text-transform:uppercase; font-weight:bold; background-color:#e6e6e6; text-align:justify; }
p.I09_Item_Alinea_Letra_11 { text-align:justify; }
p.Manrope_11_Centralizado { text-align:center; font-family: Manrope, Arial, sans-serif; }
</style></head><body>
<div align="center" style="height:60px;width:200px;">
  <img alt="Timbre" src="data:image/png;base64,${TINY_PNG_BASE64}" width="200" height="60">
</div>
<p class="Manrope_11_Centralizado" style="font-size:13pt;">PAUTA DA 200.ª REUNIÃO PLENÁRIA ORDINÁRIA</p>
<p class="I05_Item_Nivel1_11">verificação de quórum</p>
<p class="I09_Item_Alinea_Letra_11">Primeiro item da alínea com <b>negrito</b> e <a href="https://example.com">um link</a>.</p>
<table>
  <tbody>
    <tr><td colspan="2" style="background-color:#f0f0f0;">Cabeçalho tabela</td></tr>
    <tr><td style="width:70px;">8.1</td><td>Assunto de teste</td></tr>
  </tbody>
</table>
<div style="-webkit-user-select:none;user-select:none;">
  <hr>Criado por <a onclick="alert('x')">fulano</a>, versão 2 por fulano em 22/07/2026 15:17:30.
</div>
</body></html>`;

const dom = new JSDOM(html, { pretendToBeVisual: true });
const iframeDoc = dom.window.document;
const iframeWin = dom.window;

const blob = await convertSeiDocumentToDocx(iframeDoc, iframeWin);
const buffer = Buffer.from(await blob.arrayBuffer());
writeFileSync('/tmp/test-output.docx', buffer);
console.log('wrote /tmp/test-output.docx', buffer.length, 'bytes');
