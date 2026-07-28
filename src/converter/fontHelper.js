'use strict';

// Port of PptxToHtml/pptx/FontHelper.cs

const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawnSync } = require('child_process');
const { firstChild, elements, attr, descendants, NS } = require('../ooxml/dom');
const { DefaultPptLineHeightRatio, DefaultLineHeightRatio } = require('./constants');

const isRelNs = (ns) => ns === NS.R || (ns || '').toLowerCase().includes('relationships');

function* enumerateFilesRecursive(dir, extLower) {
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return;
  }
  const files = [];
  const dirs = [];
  for (const e of entries) {
    if (e.isDirectory()) dirs.push(e);
    else files.push(e);
  }
  for (const f of files) {
    if (f.name.toLowerCase().endsWith(extLower)) yield path.join(dir, f.name);
  }
  for (const d of dirs) {
    yield* enumerateFilesRecursive(path.join(dir, d.name), extLower);
  }
}

module.exports = {
  extractEmbeddedFonts(presentationPart) {
    const root = presentationPart.root;
    const embFontList = descendants(root, 'embeddedFont', NS.P);
    if (!embFontList || embFontList.length === 0) return;

    for (const embFont of embFontList) {
      const font = firstChild(embFont, 'font');
      const fontName = font ? attr(font, 'typeface') : null;
      if (!fontName) continue;

      this._embeddedFontFamilyNames.add(fontName);

      const candidates = [
        firstChild(embFont, 'regular'),
        firstChild(embFont, 'bold'),
        firstChild(embFont, 'italic'),
        firstChild(embFont, 'boldItalic'),
      ];

      for (const candidate of candidates) {
        const id = candidate ? attr(candidate, 'id', isRelNs) : null;
        if (id == null) continue;

        try {
          const part = presentationPart.getPartById(id);
          const data = Buffer.from(part.getBytes());

          if (this.isEotFont(data)) {
            const ttfData = this.convertEotToTtf(data, fontName);
            if (ttfData != null) {
              this._embeddedPptxFonts.set(fontName, ttfData);
              console.error(`  [Font] Extracted embedded EOT font: ${fontName}`);
            } else {
              console.error(`  [Font] Failed to convert EOT font: ${fontName}`);
            }
          } else {
            this.deobfuscateFont(data, part.uri);
            this._embeddedPptxFonts.set(fontName, data);
            console.error(`  [Font] Extracted embedded font: ${fontName}`);
          }
          break;
        } catch (ex) {
          console.error(`  [Font] Failed to extract embedded font ${fontName}: ${ex.message}`);
        }
      }
    }
  },

  isEotFont(data) {
    if (data.length < 36) return false;
    return data[34] === 0x4c && data[35] === 0x50;
  },

  convertEotToTtf(eotData, fontName) {
    if (this._eot2ttfToolPath == null) {
      console.error('  [Font] eot2ttf tool not provided. Use --eot2ttf <path> to specify the tool location.');
      return null;
    }
    const toolPath = this._eot2ttfToolPath;

    const tempDir = path.join(os.tmpdir(), 'pptx_eot_convert');
    fs.mkdirSync(tempDir, { recursive: true });
    const eotPath = path.join(tempDir, `${crypto.randomUUID()}.eot`);
    const ttfPath = path.join(tempDir, `${crypto.randomUUID()}.ttf`);

    try {
      fs.writeFileSync(eotPath, eotData);

      const proc = spawnSync(toolPath, [eotPath, ttfPath], {
        timeout: 10000,
        encoding: 'buffer',
        windowsHide: true,
      });

      if (proc.status !== 0) {
        const stderr = proc.stderr ? proc.stderr.toString() : '';
        console.error(`  [Font] eot2ttf failed for ${fontName}: ${stderr.trim()}`);
        return null;
      }

      if (!fs.existsSync(ttfPath)) return null;
      return fs.readFileSync(ttfPath);
    } finally {
      try { fs.unlinkSync(eotPath); } catch { /* ignore */ }
      try { fs.unlinkSync(ttfPath); } catch { /* ignore */ }
    }
  },

  deobfuscateFont(data, partUri) {
    const match = /\{?([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\}?/.exec(partUri);
    if (!match || data.length < 32) return;

    const guidBytes = guidToDotNetBytes(match[1]);
    if (!guidBytes) return;
    for (let i = 0; i < 32; i++) data[i] ^= guidBytes[i % 16];
  },

  collectFontsFromSlide(slidePart) {
    // Returns an array of unique font names in insertion order (case-insensitive
    // dedup), mirroring .NET HashSet<string>(OrdinalIgnoreCase) enumeration.
    const seen = new Map();
    const root = slidePart.root;
    if (root == null) return [];

    for (const el of descendants(root)) {
      let typeface = null;
      if (el.localName === 'latin') typeface = attr(el, 'typeface');
      else if (el.localName === 'cs') typeface = attr(el, 'typeface');
      else if (el.localName === 'ea') typeface = attr(el, 'typeface');

      if (!typeface) continue;

      if (typeface === '+mj-lt') typeface = this._themeMajorFont;
      else if (typeface === '+mn-lt') typeface = this._themeMinorFont;
      else if (typeface === '+mj-cs' || typeface === '+mn-cs') typeface = this._themeMinorFont;

      if (typeface && !typeface.startsWith('+')) {
        const key = typeface.toLowerCase();
        if (!seen.has(key)) seen.set(key, typeface);
      }
    }

    return Array.from(seen.values());
  },

  getFontFaceCss(fontName) {
    if (this._fontFaceCache.has(fontName)) return this._fontFaceCache.get(fontName);

    let fontData = null;
    let format = 'truetype';

    if (this._embeddedPptxFonts.has(fontName)) {
      fontData = this._embeddedPptxFonts.get(fontName);
    }

    if (fontData == null) {
      const fontPath = this.findFontFile(fontName);
      if (fontPath != null) {
        try {
          fontData = fs.readFileSync(fontPath);
          const ext = path.extname(fontPath).toLowerCase();
          if (ext === '.otf') format = 'opentype';
        } catch {
          /* font file unreadable */
        }
      }
    }

    if (fontData == null) {
      this._fontFaceCache.set(fontName, null);
      this._missingFonts.add(fontName);
      return null;
    }

    const fontExt = format === 'opentype' ? '.otf' : '.ttf';
    // Sanitize font name for filename (Unix invalid chars: '\0' and '/').
    const safeName = fontName.replace(/[\0/]/g, '_');
    const fileName = `${safeName}${fontExt}`;
    const filePath = path.join(this._outputDir, 'fonts', fileName);
    fs.writeFileSync(filePath, fontData);

    const css = `@font-face { font-family: '${this.encode(fontName)}'; src: url('fonts/${fileName}') format('${format}'); }`;
    this._fontFaceCache.set(fontName, css);
    console.error(`  [Font] Exported @font-face for: ${fontName}`);
    return css;
  },

  buildCssFontFamily(...familyNames) {
    const orderedFamilies = [];
    const seenFamilies = new Set(); // lowercased keys

    const addFamily = (familyName) => {
      if (familyName == null || familyName.trim() === '') return;
      const key = familyName.toLowerCase();
      if (!seenFamilies.has(key)) {
        seenFamilies.add(key);
        orderedFamilies.push(familyName);
      }
    };

    for (const rawFamilyName of familyNames) {
      const familyName = (rawFamilyName || '').trim();
      if (familyName === '' || familyName.startsWith('+')) continue;

      addFamily(familyName);
      for (const fallback of this.getFallbackFontFamilies(familyName)) addFamily(fallback);
    }

    if (orderedFamilies.length === 0) return '';

    const genericFamily = this.getGenericCssFontFamily(orderedFamilies[0]);
    if (genericFamily !== '' && !seenFamilies.has(genericFamily.toLowerCase())) {
      seenFamilies.add(genericFamily.toLowerCase());
      orderedFamilies.push(genericFamily);
    }

    return orderedFamilies.map((f) => this.formatCssFontFamily(f)).join(',');
  },

  getFallbackFontFamilies(familyName) {
    const eq = (a) => familyName.toLowerCase() === a.toLowerCase();
    if (eq('Calibri')) return ['Arial', 'Helvetica Neue'];
    if (eq('Calibri Light')) return ['Helvetica Neue', 'Arial'];
    if (eq('Arial Narrow')) return ['Arial'];
    if (eq('Cambria') || eq('Cambria Math') || eq('Constantia')) return ['Times New Roman', 'Georgia'];
    if (eq('Aptos') || eq('Aptos Display') || eq('Aptos Narrow')) return ['Arial', 'Helvetica Neue'];
    return [];
  },

  getGenericCssFontFamily(familyName) {
    const normalized = familyName.replace(/ /g, '').toLowerCase();

    if (
      normalized.includes('mono') ||
      normalized.includes('courier') ||
      normalized.includes('consolas') ||
      normalized.includes('menlo') ||
      normalized.includes('monaco')
    )
      return 'monospace';

    if (
      normalized.includes('serif') ||
      normalized.includes('times') ||
      normalized.includes('georgia') ||
      normalized.includes('cambria') ||
      normalized.includes('garamond') ||
      normalized.includes('baskerville') ||
      normalized.includes('palatino') ||
      normalized.includes('constantia')
    )
      return 'serif';

    return 'sans-serif';
  },

  formatCssFontFamily(familyName) {
    const lower = familyName.toLowerCase();
    if (lower === 'serif' || lower === 'sans-serif' || lower === 'monospace' || lower === 'cursive' || lower === 'fantasy') {
      return lower;
    }
    return `'${this.encode(familyName)}'`;
  },

  findFontFile(familyName) {
    const searchDirs = [];
    if (process.platform === 'win32') {
      const winDir = process.env.SystemRoot || process.env.windir || 'C:\\Windows';
      searchDirs.push(path.join(winDir, 'Fonts'));
      const local = process.env.LOCALAPPDATA;
      if (local) {
        const userFonts = path.join(local, 'Microsoft', 'Windows', 'Fonts');
        if (fs.existsSync(userFonts)) searchDirs.push(userFonts);
      }
    } else {
      searchDirs.push('/usr/share/fonts');
      searchDirs.push('/usr/local/share/fonts');
      const home = os.homedir();
      searchDirs.push(path.join(home, '.fonts'));
      if (process.platform === 'darwin') {
        searchDirs.push('/Library/Fonts');
        searchDirs.push('/System/Library/Fonts');
        searchDirs.push('/System/Library/Fonts/Supplemental');
        searchDirs.push(path.join(home, 'Library', 'Fonts'));
      }
    }

    const fontExtensions = ['.ttf', '.otf'];

    const normalized = familyName.replace(/ /g, '').toLowerCase();
    const normalizedNoDash = normalized.replace(/-/g, '');

    // First pass: exact filename match.
    for (const dir of searchDirs) {
      if (!fs.existsSync(dir)) continue;
      for (const ext of fontExtensions) {
        for (const file of enumerateFilesRecursive(dir, ext)) {
          const fname = path.basename(file, path.extname(file)).replace(/ /g, '').replace(/-/g, '').toLowerCase();
          if (fname === normalizedNoDash || fname === normalized) return file;
        }
      }
    }

    // Second pass: "regular" suffix.
    for (const dir of searchDirs) {
      if (!fs.existsSync(dir)) continue;
      for (const ext of fontExtensions) {
        for (const file of enumerateFilesRecursive(dir, ext)) {
          const fname = path.basename(file, path.extname(file)).replace(/ /g, '').replace(/-/g, '').toLowerCase();
          if (fname === normalizedNoDash + 'regular' || fname === normalized + 'regular') return file;
        }
      }
    }

    // Third pass: name-table family match, preferring Regular (400) weight.
    for (const dir of searchDirs) {
      if (!fs.existsSync(dir)) continue;
      let bestMatch = null;
      let bestWeightDist = Number.MAX_SAFE_INTEGER;
      for (const ext of fontExtensions) {
        for (const file of enumerateFilesRecursive(dir, ext)) {
          try {
            const nameInFont = this.readFontFamilyName(file);
            if (nameInFont != null && nameInFont.toLowerCase() === familyName.toLowerCase()) {
              const weight = this.readFontWeightClass(file);
              const dist = Math.abs(weight - 400);
              if (dist < bestWeightDist) {
                bestWeightDist = dist;
                bestMatch = file;
              }
            }
          } catch {
            /* ignore */
          }
        }
      }
      if (bestMatch != null) return bestMatch;
    }

    return null;
  },

  readFontWeightClass(filePath) {
    try {
      const buf = fs.readFileSync(filePath);
      const numTables = buf.readUInt16BE(4);
      let p = 12;
      for (let i = 0; i < numTables; i++) {
        const tag = buf.toString('latin1', p, p + 4);
        const offset = buf.readUInt32BE(p + 8);
        p += 16;
        if (tag === 'OS/2') {
          return buf.readUInt16BE(offset + 4);
        }
      }
    } catch {
      /* ignore */
    }
    return 400;
  },

  readFontFamilyName(filePath) {
    try {
      const buf = fs.readFileSync(filePath);
      const numTables = buf.readUInt16BE(4);
      let p = 12;
      let nameOffset = 0;
      for (let i = 0; i < numTables; i++) {
        const tag = buf.toString('latin1', p, p + 4);
        const offset = buf.readUInt32BE(p + 8);
        p += 16;
        if (tag === 'name') { nameOffset = offset; break; }
      }
      if (nameOffset === 0) return null;

      const count = buf.readUInt16BE(nameOffset + 2);
      const stringOffset = buf.readUInt16BE(nameOffset + 4);
      let rec = nameOffset + 6;
      for (let i = 0; i < count; i++) {
        const platformID = buf.readUInt16BE(rec);
        const nameID = buf.readUInt16BE(rec + 6);
        const length = buf.readUInt16BE(rec + 8);
        const offset = buf.readUInt16BE(rec + 10);
        rec += 12;

        if (nameID === 1) {
          const start = nameOffset + stringOffset + offset;
          const data = buf.subarray(start, start + length);
          let name;
          if (platformID === 3 || platformID === 0) {
            name = decodeUtf16BE(data);
          } else {
            name = data.toString('latin1');
          }
          if (name && name.trim() !== '') return name.trim();
        }
      }
    } catch {
      /* ignore */
    }
    return null;
  },

  // ── Font metrics helpers ──────────────────────────────────────────

  resolveFontName(rPr) {
    if (rPr == null) return '';
    const latin = firstChild(rPr, 'latin');
    const cs = firstChild(rPr, 'cs');
    const latinTypeface = this.getStringValue(attr(latin, 'typeface'));
    const csTypeface = this.getStringValue(attr(cs, 'typeface'));
    let typeface = latinTypeface !== '' ? latinTypeface : csTypeface;
    if (typeface === '+mj-lt') typeface = this._themeMajorFont;
    else if (typeface === '+mn-lt') typeface = this._themeMinorFont;
    else if (typeface === '+mj-cs' || typeface === '+mn-cs')
      typeface = csTypeface !== '' && csTypeface !== '+mn-cs' && csTypeface !== '+mj-cs' ? csTypeface : this._themeMinorFont;
    if (csTypeface !== '' && csTypeface !== '+mn-cs' && csTypeface !== '+mj-cs') typeface = csTypeface;
    return typeface;
  },

  getFontPptLineHeight(fontName) {
    if (!fontName) return DefaultPptLineHeightRatio;
    if (this._fontPptLineHeightCache.has(fontName)) return this._fontPptLineHeightCache.get(fontName);
    this.readTtfMetricsRatio(fontName);
    return this._fontPptLineHeightCache.getOrDefault(fontName, DefaultPptLineHeightRatio);
  },

  getParagraphPptLineHeight(para, inheritedRPr) {
    let maxRatio = 0;
    for (const run of elements(para, 'r')) {
      const font = this.resolveFontName(firstChild(run, 'rPr'));
      if (font !== '') maxRatio = Math.max(maxRatio, this.getFontPptLineHeight(font));
    }
    if (maxRatio === 0) {
      const font = this.resolveFontName(inheritedRPr);
      if (font !== '') maxRatio = this.getFontPptLineHeight(font);
    }
    return maxRatio > 0 ? maxRatio : DefaultPptLineHeightRatio;
  },

  getFontLineHeightRatio(fontName) {
    if (!fontName) return DefaultLineHeightRatio;
    if (this._fontMetricsCache.has(fontName)) return this._fontMetricsCache.get(fontName);
    this.readTtfMetricsRatio(fontName);
    return this._fontMetricsCache.getOrDefault(fontName, DefaultLineHeightRatio);
  },

  getParagraphLineHeightRatio(para, inheritedRPr) {
    let maxRatio = 0;
    for (const run of elements(para, 'r')) {
      const font = this.resolveFontName(firstChild(run, 'rPr'));
      if (font !== '') maxRatio = Math.max(maxRatio, this.getFontLineHeightRatio(font));
    }
    if (maxRatio === 0) {
      const font = this.resolveFontName(inheritedRPr);
      if (font !== '') maxRatio = this.getFontLineHeightRatio(font);
    }
    return maxRatio > 0 ? maxRatio : DefaultLineHeightRatio;
  },

  readTtfMetricsRatio(fontFamilyName) {
    if (this._embeddedPptxFonts.has(fontFamilyName)) {
      if (this._readTtfMetricsFromBuffer(fontFamilyName, this._embeddedPptxFonts.get(fontFamilyName), true)) return;
    }

    const p = this.findFontFile(fontFamilyName);
    if (p == null) {
      console.error(
        `  [FontMetrics] Font not found: ${fontFamilyName}, using fallback win=${DefaultLineHeightRatio.toFixed(4)} ppt=${DefaultPptLineHeightRatio.toFixed(4)}`
      );
      this._fontMetricsCache.set(fontFamilyName, DefaultLineHeightRatio);
      this._fontPptLineHeightCache.set(fontFamilyName, DefaultPptLineHeightRatio);
      return;
    }
    try {
      const buf = fs.readFileSync(p);
      if (!this._readTtfMetricsFromBuffer(fontFamilyName, buf, false)) {
        this._fontMetricsCache.set(fontFamilyName, DefaultLineHeightRatio);
        this._fontPptLineHeightCache.set(fontFamilyName, DefaultPptLineHeightRatio);
      }
    } catch {
      this._fontMetricsCache.set(fontFamilyName, DefaultLineHeightRatio);
      this._fontPptLineHeightCache.set(fontFamilyName, DefaultPptLineHeightRatio);
    }
  },

  _readTtfMetricsFromBuffer(fontFamilyName, data, embedded) {
    try {
      const buf = Buffer.isBuffer(data) ? data : Buffer.from(data);
      const numTables = buf.readUInt16BE(4);
      let p = 12;
      let headOffset = 0, os2Offset = 0, os2Len = 0;
      for (let i = 0; i < numTables; i++) {
        const tag = buf.toString('latin1', p, p + 4);
        const offset = buf.readUInt32BE(p + 8);
        const length = buf.readUInt32BE(p + 12);
        p += 16;
        if (tag === 'head') headOffset = offset;
        else if (tag === 'OS/2') { os2Offset = offset; os2Len = length; }
      }
      if (headOffset === 0 || os2Offset === 0) return false;

      const unitsPerEm = buf.readUInt16BE(headOffset + 18);
      if (unitsPerEm === 0) return false;
      if (os2Len < 78) return false;

      const base = os2Offset + 68;
      const sTypoAscender = buf.readInt16BE(base);
      const sTypoDescender = buf.readInt16BE(base + 2);
      const sTypoLineGap = buf.readInt16BE(base + 4);
      const winAscent = buf.readUInt16BE(base + 6);
      const winDescentRaw = buf.readInt16BE(base + 8);
      const winDescent = Math.abs(winDescentRaw);

      const winRatio = (winAscent + winDescent) / unitsPerEm;
      this._fontMetricsCache.set(fontFamilyName, winRatio > 0 ? winRatio : DefaultLineHeightRatio);

      const pptRatio = (sTypoAscender - sTypoDescender + sTypoLineGap) / unitsPerEm;
      this._fontPptLineHeightCache.set(fontFamilyName, pptRatio > 0 ? pptRatio : DefaultPptLineHeightRatio);

      const suffix = embedded ? ' (embedded)' : '';
      console.error(
        `  [FontMetrics] ${fontFamilyName}${suffix}: winAsc=${winAscent} winDesc=${winDescent} typoAsc=${sTypoAscender} typoDesc=${sTypoDescender} typoGap=${sTypoLineGap} UPM=${unitsPerEm} → win=${winRatio.toFixed(4)} ppt=${pptRatio.toFixed(4)}`
      );
      return true;
    } catch {
      return false;
    }
  },
};

// .NET Guid.ToByteArray() byte order (mixed-endian): first three groups are
// little-endian, the last two groups are big-endian (as written).
function guidToDotNetBytes(guidStr) {
  const hex = guidStr.replace(/-/g, '');
  if (hex.length !== 32) return null;
  const b = [];
  for (let i = 0; i < 16; i++) b.push(parseInt(hex.substr(i * 2, 2), 16));
  return Buffer.from([
    b[3], b[2], b[1], b[0],
    b[5], b[4],
    b[7], b[6],
    b[8], b[9], b[10], b[11], b[12], b[13], b[14], b[15],
  ]);
}

function decodeUtf16BE(data) {
  let s = '';
  for (let i = 0; i + 1 < data.length; i += 2) {
    s += String.fromCharCode((data[i] << 8) | data[i + 1]);
  }
  return s;
}
