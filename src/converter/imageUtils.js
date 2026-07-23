// Resolves an <img> element to raw bytes + its rendered size, so it can be
// embedded as a real docx ImageRun (as opposed to a data-URI src string,
// which is all the previous MHTML altChunk export could carry).

function dataUriToBytes(dataUri) {
  const [, meta, base64] = dataUri.match(/^data:([^;]+);base64,(.*)$/s) || [];
  if (!base64) return null;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return { bytes, mime: meta };
}

async function fetchToBytes(url) {
  const response = await fetch(url);
  const buffer = await response.arrayBuffer();
  return { bytes: new Uint8Array(buffer), mime: response.headers.get('content-type') || 'image/png' };
}

function extensionFromMime(mime) {
  if (!mime) return 'png';
  if (mime.includes('jpeg') || mime.includes('jpg')) return 'jpg';
  if (mime.includes('gif')) return 'gif';
  if (mime.includes('svg')) return 'svg';
  if (mime.includes('bmp')) return 'bmp';
  return 'png';
}

export async function loadImageData(imgEl) {
  const src = imgEl.currentSrc || imgEl.src;
  if (!src) return null;

  let payload;
  try {
    payload = src.startsWith('data:') ? dataUriToBytes(src) : await fetchToBytes(src);
  } catch (error) {
    console.warn('SEI para Word: falha ao carregar imagem', src, error);
    return null;
  }
  if (!payload || !payload.bytes.length) return null;

  const naturalWidth = imgEl.naturalWidth || imgEl.width || 100;
  const naturalHeight = imgEl.naturalHeight || imgEl.height || 100;
  const renderedWidth = imgEl.getBoundingClientRect().width || naturalWidth;
  const aspect = naturalHeight / naturalWidth || 1;

  return {
    data: payload.bytes,
    type: extensionFromMime(payload.mime),
    width: Math.round(renderedWidth),
    height: Math.round(renderedWidth * aspect),
  };
}
