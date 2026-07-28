'use strict';

// .NET-compatible formatting helpers so the generated HTML matches the C#
// output byte-for-byte.

// Truncate the magnitude `v` (>= 0) to `digits` fractional places using the
// EXACT decimal expansion of the double (no rounding).
function truncateMagnitude(v, digits) {
  const s = v.toFixed(digits + 25);
  const dot = s.indexOf('.');
  return digits === 0 ? s.slice(0, dot) : s.slice(0, dot + 1 + digits);
}

// True iff `v` (>= 0) is an exact half at `digits` — i.e. the exact decimal
// value of the double has a '5' at position digits+1 and only zeros after.
function isExactHalf(v, digits) {
  const s = v.toFixed(digits + 25);
  const frac = s.slice(s.indexOf('.') + 1);
  return frac[digits] === '5' && /^0*$/.test(frac.slice(digits + 1));
}

/**
 * Equivalent of C# `value.ToString("F{digits}", InvariantCulture)`.
 * .NET Core rounds half-to-even on the exact double value; JS `toFixed`
 * rounds half-away-from-zero. They agree except on true ties, which we
 * detect against the exact decimal expansion and round to even.
 */
function fmt(value, digits) {
  if (value == null || !Number.isFinite(value)) return String(value);
  let s = value.toFixed(digits);

  const v = Math.abs(value);
  if (isExactHalf(v, digits)) {
    const down = truncateMagnitude(v, digits);
    const lastDigit = down.charCodeAt(down.length - 1) - 48;
    if (lastDigit % 2 === 0) {
      // Even -> round to even means round down; toFixed rounded away (up).
      s = value < 0 && !/^0(\.0+)?$/.test(down) ? '-' + down : down;
    }
  }

  // .NET prints "0" / "0.0" (no negative sign) when the value rounds to zero.
  if (s[0] === '-' && /^-0(\.0+)?$/.test(s)) s = s.slice(1);
  return s;
}

/**
 * Equivalent of C# `Math.Round(x)` (MidpointRounding.ToEven / banker's rounding),
 * which differs from JS `Math.round` (which rounds halves toward +Infinity).
 */
function bankRound(x) {
  const floor = Math.floor(x);
  const diff = x - floor;
  if (diff < 0.5) return floor;
  if (diff > 0.5) return floor + 1;
  // exact half -> round to even
  return floor % 2 === 0 ? floor : floor + 1;
}

// Convenience aliases matching the format specifiers used throughout the port.
const f0 = (v) => fmt(v, 0);
const f1 = (v) => fmt(v, 1);
const f2 = (v) => fmt(v, 2);
const f3 = (v) => fmt(v, 3);
const f4 = (v) => fmt(v, 4);

/**
 * Equivalent of C# `value.ToString("0.###", InvariantCulture)` (up to N
 * fractional digits, trailing zeros trimmed).
 */
function fmtTrim(value, maxDigits) {
  if (value == null || !Number.isFinite(value)) return String(value);
  let s = fmt(value, maxDigits);
  if (s.indexOf('.') >= 0) s = s.replace(/0+$/, '').replace(/\.$/, '');
  if (s[0] === '-' && /^-0$/.test(s)) s = s.slice(1);
  return s;
}

/** C# `value.ToString("X2")` — uppercase two-digit hex. */
function hex2(value) {
  const v = Math.max(0, Math.min(255, value | 0));
  return v.toString(16).toUpperCase().padStart(2, '0');
}

/**
 * Equivalent of System.Net.WebUtility.HtmlEncode.
 * Encodes < > " ' & and Latin-1 supplement chars (160-255) as numeric
 * entities; astral chars become &#scalar;. Other BMP chars pass through
 * unchanged (e.g. the em dash U+2014 stays literal).
 */
function htmlEncode(text) {
  if (text == null || text === '') return '';
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    if (ch <= 62) {
      switch (ch) {
        case 60: out += '&lt;'; break;
        case 62: out += '&gt;'; break;
        case 34: out += '&quot;'; break;
        case 39: out += '&#39;'; break;
        case 38: out += '&amp;'; break;
        default: out += text[i];
      }
    } else if (ch >= 160 && ch < 256) {
      out += '&#' + ch + ';';
    } else if (ch >= 0xd800 && ch <= 0xdbff) {
      const low = text.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        const scalar = (ch - 0xd800) * 0x400 + (low - 0xdc00) + 0x10000;
        out += '&#' + scalar + ';';
        i++;
      } else {
        out += '�';
      }
    } else if (ch >= 0xdc00 && ch <= 0xdfff) {
      out += '�';
    } else {
      out += text[i];
    }
  }
  return out;
}

module.exports = { fmt, f0, f1, f2, f3, f4, fmtTrim, hex2, htmlEncode, bankRound };
