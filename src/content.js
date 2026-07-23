import { convertSeiDocumentToDocx } from './converter/htmlToDocx.js';
import { findSelectedDocumentTitle, sanitizeFileName } from './seiTreeLookup.js';

// Adiciona o botão de exportação para Word no menu do SEI
function adicionarBotaoExportarWord() {
  if (document.getElementById('btnExportarWordSEI')) return;
  const barraAcoes = document.querySelector('#divArvoreAcoes');
  if (!barraAcoes) return;
  const iconUrl = chrome.runtime.getURL('icons/export-word-icon.png');
  const link = document.createElement('a');
  link.id = 'btnExportarWordSEI';
  link.title = 'Exportar para Word';
  link.href = '#';
  link.tabIndex = 451;
  link.onmouseover = () => window.infraTooltipMostrar && window.infraTooltipMostrar('Exportar para Word');
  link.onmouseout = () => window.infraTooltipOcultar && window.infraTooltipOcultar();
  link.onclick = (e) => { e.preventDefault(); exportarParaWord(); };
  link.innerHTML = `<img src="${iconUrl}" alt="Exportar para Word" style="width:40px; height:40px;" />`;
  barraAcoes.appendChild(link);
}

// Exporta o conteúdo do iframe para um arquivo Word (.docx), preservando
// cabeçalho, rodapé, numeração e formatação tal como aparecem no SEI.
async function exportarParaWord() {
  try {
    const iframe = document.getElementById('ifrVisualizacao');
    if (!iframe) throw new Error('iframe de visualização não encontrado');
    const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
    const iframeWin = iframe.contentWindow;
    if (!iframeDoc || !iframeDoc.body) throw new Error('Conteúdo do iframe não encontrado');

    const documentTitle = findSelectedDocumentTitle();
    const blob = await convertSeiDocumentToDocx(iframeDoc, iframeWin, documentTitle);

    const now = new Date();
    const dateStr = now.toISOString().split('T')[0];
    const fileName = documentTitle
      ? sanitizeFileName(documentTitle)
      : `Documento_SEI_${dateStr}`;

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileName}.docx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error('Erro na exportação para Word:', error);
    throw error;
  }
}

// Observa a criação do menu para inserir o botão Word
const observer = new MutationObserver(adicionarBotaoExportarWord);
observer.observe(document.body, { childList: true, subtree: true });

// Garante a inserção do botão ao carregar a página
window.addEventListener('DOMContentLoaded', adicionarBotaoExportarWord);

// Listener para mensagens do popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'exportToWord' || request.action === 'convertToDOCX') {
    exportarParaWord().then(() => sendResponse({ success: true })).catch(() => sendResponse({ success: false }));
    return true;
  }
  return undefined;
});
