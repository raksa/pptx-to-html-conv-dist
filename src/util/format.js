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
 *
 * The sign follows the value, not the rounded digits: since .NET Core 3.0 a
 * negative value that rounds to zero keeps its minus, and so does -0 itself
 * (`(-0.3).ToString("F0")` is "-0", `(-0.0).ToString("F1")` is "-0.0").
 * JS `toFixed` never signs -0, so the sign is put back here by hand.
 */
function fmt(value, digits) {
  if (value == null || !Number.isFinite(value)) return String(value);
  const v = Math.abs(value);
  let s = v.toFixed(digits);

  if (isExactHalf(v, digits)) {
    const down = truncateMagnitude(v, digits);
    const lastDigit = down.charCodeAt(down.length - 1) - 48;
    // Even -> round to even means round down; toFixed rounded away (up).
    if (lastDigit % 2 === 0) s = down;
  }

  return value < 0 || Object.is(value, -0) ? '-' + s : s;
}

/**
 * Equivalent of C# `Math.Round(x)` (MidpointRounding.ToEven / banker's rounding),
 * which differs from JS `Math.round` (which rounds halves toward +Infinity).
 */
function bankRound(x) {
  const floor = Math.floor(x);
  const diff = x - floor;
  let r;
  if (diff < 0.5) r = floor;
  else if (diff > 0.5) r = floor + 1;
  // exact half -> round to even
  else r = floor % 2 === 0 ? floor : floor + 1;
  // Math.Round keeps the sign of zero: Math.Round(-0.3) is -0, which F0 prints "-0".
  return r === 0 && (x < 0 || Object.is(x, -0)) ? -0 : r;
}

// Convenience aliases matching the format specifiers used throughout the port.
const f0 = (v) => fmt(v, 0);
const f1 = (v) => fmt(v, 1);
const f2 = (v) => fmt(v, 2);
const f3 = (v) => fmt(v, 3);
const f4 = (v) => fmt(v, 4);

// Adds one unit in the last place to a string of decimal digits ("0999" -> "1000", "99" -> "100").
function incrementDigits(digits) {
  const d = digits.split('');
  let i = d.length - 1;
  while (i >= 0 && d[i] === '9') d[i--] = '0';
  if (i < 0) return '1' + d.join('');
  d[i] = String.fromCharCode(d[i].charCodeAt(0) + 1);
  return d.join('');
}

/**
 * Equivalent of C# `value.ToString("0.###", InvariantCulture)` (up to N
 * fractional digits, trailing zeros trimmed).
 *
 * .NET rounds twice for a custom format, unlike "F": it first rounds the exact
 * double to 15 significant digits (ties to even), then rounds those digits
 * half away from zero at the last '#'. So (0.0625).ToString("0.###") is
 * "0.063" where "F3" gives "0.062", and 8.5185 (really 8.51849999999999951…)
 * is "8.519" where "F3" gives "8.518". A negative value keeps its minus even
 * when it rounds to zero ("-0"), as in .NET Core.
 */
function fmtTrim(value, maxDigits) {
  if (value == null || !Number.isFinite(value)) return String(value);
  // toFixed(100) is the exact expansion for |value| >= 2^-48; anything smaller prints 0 here.
  const exact = Math.abs(value).toFixed(100);
  const dot = exact.indexOf('.');
  let digits = exact.slice(0, dot) + exact.slice(dot + 1);
  let point = dot; // how many of `digits` come before the decimal point

  // Round the exact value to 15 significant digits, ties to even.
  const first = digits.search(/[1-9]/);
  if (first >= 0 && first + 15 < digits.length) {
    const cut = first + 15;
    let kept = digits.slice(0, cut);
    const next = digits[cut];
    const beyond = /[1-9]/.test(digits.slice(cut + 1));
    const odd = (kept.charCodeAt(cut - 1) - 48) % 2 === 1;
    if (next > '5' || (next === '5' && (beyond || odd))) {
      const up = incrementDigits(kept);
      if (up.length > kept.length) point++;
      kept = up;
    }
    digits = kept.padEnd(point, '0');
  }

  // Round those digits half away from zero at the format's last place.
  const end = point + maxDigits;
  if (end < digits.length) {
    let kept = digits.slice(0, end);
    if (digits[end] >= '5') {
      const up = incrementDigits(kept);
      if (up.length > kept.length) point++;
      kept = up;
    }
    digits = kept;
  }

  const whole = digits.slice(0, point).replace(/^0+(?=\d)/, '');
  const fraction = digits.slice(point).replace(/0+$/, '');
  const s = fraction === '' ? whole : whole + '.' + fraction;
  return value < 0 || Object.is(value, -0) ? '-' + s : s;
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
