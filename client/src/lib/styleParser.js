// Client-side port of the server-side style parsing logic in server/routes.js.
// Used by the Settings preview panel — no network call needed.

const STYLE_COLORS = [
  'red','blue','green','orange','purple','black','white','navy','maroon','teal',
  'darkred','darkblue','brown','pink','gold','gray','grey','silver','crimson',
  'coral','salmon','indigo','violet','cyan','magenta','lime','olive','turquoise',
];

const STYLE_SIZES = {
  'very large':  '1.6em',
  'extra large': '1.6em',
  'large text':  '1.3em',
  'large':       '1.3em',
  'small text':  '0.85em',
  'small':       '0.85em',
};

const _styleKeywords = `bold|italic|underline|very large|extra large|large|small|${STYLE_COLORS.join('|')}`;

function buildHtml(text, styleDesc) {
  const lower = styleDesc.toLowerCase();
  const isBold      = lower.includes('bold');
  const isItalic    = lower.includes('italic');
  const isUnderline = lower.includes('underline');

  let color = null;
  for (const c of STYLE_COLORS) {
    if (lower.includes(c)) { color = c; break; }
  }

  let fontSize = null;
  for (const [label, size] of Object.entries(STYLE_SIZES)) {
    if (lower.includes(label)) { fontSize = size; break; }
  }

  const styles = [];
  if (color)       styles.push(`color:${color}`);
  if (fontSize)    styles.push(`font-size:${fontSize}`);
  if (isUnderline) styles.push('text-decoration:underline');

  let html = text;
  if (styles.length) html = `<span style="${styles.join(';')}">${html}</span>`;
  if (isItalic)      html = `<em>${html}</em>`;
  if (isBold)        html = `<strong>${html}</strong>`;

  if (!styles.length && !isItalic && !isBold) return null;
  return html;
}

// Returns array of { html } for each styled quoted phrase found, or [].
export function extractStyledPhrases(instructions) {
  if (!instructions) return [];
  const results = [];
  const colorList = STYLE_COLORS.join('|');
  const styleKeywords = `bold|italic|underline|large|small|very large|extra large|${colorList}`;
  const re = new RegExp(
    `(?:["\u201C])([^"\u201D]+)(?:["\u201D])\\s+in\\s+((?:(?:${styleKeywords})\\s*)+(?:text)?)`,
    'gi'
  );
  let match;
  while ((match = re.exec(instructions)) !== null) {
    const html = buildHtml(match[1], match[2]);
    if (html) results.push({ html });
  }
  return results;
}

// Returns a CSS style object for global response body styling, or null.
export function extractGlobalResponseStyle(instructions) {
  if (!instructions) return null;
  const re = new RegExp(
    `\\b(?:respond|answer|write|reply|display|show)\\s+(?:all\\s+(?:answers?|responses?)\\s+)?in\\s+((?:(?:${_styleKeywords})\\s*)+(?:text)?)`,
    'gi'
  );
  let match;
  while ((match = re.exec(instructions)) !== null) {
    const desc = match[1].toLowerCase();
    const bold      = desc.includes('bold');
    const italic    = desc.includes('italic');
    const underline = desc.includes('underline');
    let color = null;
    for (const c of STYLE_COLORS) { if (desc.includes(c)) { color = c; break; } }
    let fontSize = null;
    for (const [label, size] of Object.entries(STYLE_SIZES)) { if (desc.includes(label)) { fontSize = size; break; } }
    if (bold || italic || underline || color || fontSize) {
      const style = {};
      if (color)      style.color = color;
      if (fontSize)   style.fontSize = fontSize;
      if (bold)       style.fontWeight = 'bold';
      if (italic)     style.fontStyle = 'italic';
      if (underline)  style.textDecoration = 'underline';
      return style;
    }
  }
  return null;
}
