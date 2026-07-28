'use strict';

// Serialize a value the way System.Text.Json's JsonSerializer does with
// { WriteIndented = true } and the default (HTML-safe) JavaScriptEncoder:
//  - 2-space indentation, "\n" newlines, ": " after keys
//  - empty arrays/objects render inline ([] / {})
//  - null values are emitted
//  - strings escape ", \, &, ', <, >, +, ` and all non-ASCII / control chars
//    as \uXXXX (uppercase), with \b \t \n \f \r short forms.
// Verified against the reference C# tool's output.

const ESCAPE_ASCII = new Set([0x22, 0x26, 0x27, 0x3c, 0x3e, 0x2b, 0x60]);

function escapeString(s) {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 0x5c) out += '\\\\';
    else if (c === 0x08) out += '\\b';
    else if (c === 0x09) out += '\\t';
    else if (c === 0x0a) out += '\\n';
    else if (c === 0x0c) out += '\\f';
    else if (c === 0x0d) out += '\\r';
    else if (c >= 0x20 && c <= 0x7e) {
      if (ESCAPE_ASCII.has(c)) out += '\\u' + c.toString(16).toUpperCase().padStart(4, '0');
      else out += s[i];
    } else {
      out += '\\u' + c.toString(16).toUpperCase().padStart(4, '0');
    }
  }
  return out;
}

function pad(level) {
  return '  '.repeat(level);
}

function serialize(value, level) {
  if (value === null || value === undefined) return 'null';
  const t = typeof value;
  if (t === 'string') return `"${escapeString(value)}"`;
  if (t === 'number') return String(value);
  if (t === 'boolean') return value ? 'true' : 'false';

  if (Array.isArray(value)) {
    if (value.length === 0) return '[]';
    const items = value.map((v) => pad(level + 1) + serialize(v, level + 1));
    return '[\n' + items.join(',\n') + '\n' + pad(level) + ']';
  }

  // object
  const keys = Object.keys(value);
  if (keys.length === 0) return '{}';
  const entries = keys.map((k) => `${pad(level + 1)}"${escapeString(k)}": ${serialize(value[k], level + 1)}`);
  return '{\n' + entries.join(',\n') + '\n' + pad(level) + '}';
}

function serializeIndented(value) {
  return serialize(value, 0);
}

module.exports = { serializeIndented, escapeString };
