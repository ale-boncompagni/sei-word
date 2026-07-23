// The document's display name ("Deliberação Plenária de CAU/UF 000 (1045627)")
// only exists in the SEI process tree — a sibling frame from the one holding
// the viewer/toolbar our content script normally runs in — as the element
// carrying SEI's own "selected node" class. There's no message-passing API
// for this in the page, so we reach across frames directly; same-origin
// frames allow contentDocument access without needing chrome.runtime.

const SELECTED_NODE_SELECTOR = '.infraArvoreNoSelecionado';

function safeDoc(getter) {
  try {
    const doc = getter();
    return doc || null;
  } catch (error) {
    return null; // cross-origin frame, not reachable
  }
}

function collectReachableDocuments() {
  const roots = new Set([document]);
  safeDoc(() => window.parent?.document) && roots.add(window.parent.document);
  safeDoc(() => window.top?.document) && roots.add(window.top.document);

  // SEI's tree frame is usually a sibling one or two levels up from wherever
  // our button lives, so also look one level into every reachable root's
  // own iframes.
  const discovered = [];
  roots.forEach((doc) => {
    Array.from(doc.querySelectorAll('iframe')).forEach((frame) => {
      const frameDoc = safeDoc(() => frame.contentDocument);
      if (frameDoc) discovered.push(frameDoc);
    });
  });
  discovered.forEach((doc) => roots.add(doc));

  return roots;
}

export function findSelectedDocumentTitle() {
  for (const doc of collectReachableDocuments()) {
    const el = doc.querySelector(SELECTED_NODE_SELECTOR);
    const text = el?.textContent?.trim();
    if (text) return text;
  }
  return null;
}

export function sanitizeFileName(name) {
  return name
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}
