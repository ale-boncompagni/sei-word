import { NUMBERING_RULES } from './numberingRules.js';

// Word's HTML importer does not resolve CSS `content: counter(...)` on
// `:before`, so `innerHTML` alone silently drops every "1.", "1.1.", "a)",
// "I -" prefix that SEI generates this way. This engine replays the same
// counter-increment/counter-reset semantics SEI's own stylesheet defines,
// in document order, so we can prepend the resolved number as real text
// before handing the paragraph to the docx builder.
function toLowerLatin(n) {
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(97 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s || 'a';
}

function toUpperRoman(n) {
  const table = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let result = '';
  let remaining = n;
  for (const [value, symbol] of table) {
    while (remaining >= value) {
      result += symbol;
      remaining -= value;
    }
  }
  return result || '';
}

function formatCounter(value, style) {
  switch (style) {
    case 'lower-latin':
    case 'lower-alpha':
      return toLowerLatin(value);
    case 'upper-roman':
      return toUpperRoman(value);
    default:
      return String(value);
  }
}

export function createNumberingEngine() {
  const counters = Object.create(null);

  function findRule(classList) {
    for (const cls of classList) {
      const rule = NUMBERING_RULES[cls.toLowerCase()];
      if (rule) return rule;
    }
    return null;
  }

  // Returns the resolved prefix text (e.g. "8.1.") for this element, or
  // null if the element doesn't carry one of SEI's numbered-paragraph classes.
  function resolvePrefix(classList) {
    const rule = findRule(classList);
    if (!rule) return null;

    for (const name of rule.reset) {
      counters[name] = 0;
    }
    if (rule.increment) {
      counters[rule.increment] = (counters[rule.increment] || 0) + 1;
    }

    if (!rule.template) return '';
    return rule.template
      .map((token) => (token.type === 'literal'
        ? token.text
        : formatCounter(counters[token.name] || 0, token.style)))
      .join('');
  }

  return { resolvePrefix };
}
