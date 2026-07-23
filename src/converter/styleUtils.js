import { AlignmentType } from 'docx';

// All style decisions are read from getComputedStyle() of the *live* iframe
// element rather than re-parsing inline styles or SEI's stylesheet. That
// means formatting coming from CSS classes (e.g. "Manrope_11_Justificado")
// is captured correctly even though we never copy the <style> block itself —
// the browser has already cascaded it for us.

export function computed(win, el) {
  return win.getComputedStyle(el);
}

export function mapAlignment(textAlign) {
  switch (textAlign) {
    case 'center':
      return AlignmentType.CENTER;
    case 'right':
    case 'end':
      return AlignmentType.RIGHT;
    case 'justify':
      return AlignmentType.JUSTIFIED;
    default:
      return AlignmentType.LEFT;
  }
}

export function ptFromPx(px) {
  return px * 0.75; // 96 CSS px per inch, 72 pt per inch
}

export function halfPointsFromPx(px) {
  return Math.round(ptFromPx(px) * 2);
}

export function dxaFromPx(px) {
  return Math.round(px * 15); // 1px (96dpi) = 0.75pt = 15 dxa (20 dxa/pt)
}

export function isBold(style) {
  const weight = style.fontWeight;
  return weight === 'bold' || Number(weight) >= 600;
}

export function isItalic(style) {
  return style.fontStyle === 'italic' || style.fontStyle === 'oblique';
}

export function isUnderline(style) {
  return style.textDecorationLine?.includes('underline')
    || style.textDecoration?.includes('underline');
}

export function isStrike(style) {
  return style.textDecorationLine?.includes('line-through')
    || style.textDecoration?.includes('line-through');
}

export function isAllCaps(style) {
  return style.textTransform === 'uppercase';
}

export function verticalAlignRun(style) {
  if (style.verticalAlign === 'super') return 'superscript';
  if (style.verticalAlign === 'sub') return 'subscript';
  return undefined;
}

export function rgbToHex(rgb) {
  if (!rgb) return undefined;
  const match = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!match) return undefined;
  const [, r, g, b] = match;
  const toHex = (n) => Number(n).toString(16).padStart(2, '0');
  const hex = `${toHex(r)}${toHex(g)}${toHex(b)}`.toUpperCase();
  // Word treats pure white/black text color as "auto"; leave everything
  // else explicit so highlighted headers (e.g. SEI's #E6E6E6 row shading)
  // survive the trip.
  return hex;
}

export function fontFamilyPrimary(fontFamilyList) {
  if (!fontFamilyList) return 'Calibri';
  const first = fontFamilyList.split(',')[0].trim().replace(/^["']|["']$/g, '');
  return first || 'Calibri';
}
