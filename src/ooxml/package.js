'use strict';

// Open Packaging Conventions (OPC) reader. Replaces the pieces of
// DocumentFormat.OpenXml.Packaging that the converter relies on:
//   PresentationDocument.Open / PresentationPart / SlidePart /
//   SlideLayoutPart / SlideMasterPart / ThemePart / NotesSlidePart /
//   GetPartById / DataPartReferenceRelationships / part.Uri / ContentType /
//   GetStream.

const fs = require('fs');
const { unzipSync } = require('fflate');
const { DOMParser } = require('@xmldom/xmldom');
const { NS, elements, attr } = require('./dom');

const REL_TYPE = {
  OFFICE_DOCUMENT: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  SLIDE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  SLIDE_LAYOUT: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  SLIDE_MASTER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  THEME: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  NOTES_SLIDE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  CHART: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  VIDEO: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
  AUDIO: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
  MEDIA: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/media',
};

const _domParser = new DOMParser({
  onError: () => {
    // Ignore non-fatal parse warnings; malformed parts are handled by callers.
  },
});

function normalizePartName(name) {
  return name.startsWith('/') ? name : '/' + name;
}

function dirName(partName) {
  const idx = partName.lastIndexOf('/');
  return idx <= 0 ? '/' : partName.slice(0, idx);
}

function baseName(partName) {
  const idx = partName.lastIndexOf('/');
  return idx < 0 ? partName : partName.slice(idx + 1);
}

/** Resolve an OPC relationship target (possibly relative, with ../) to an absolute part name. */
function resolveRelativeTarget(baseDir, target) {
  if (target.startsWith('/')) return target;
  const stack = baseDir === '/' ? [] : baseDir.split('/').filter(Boolean);
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') stack.pop();
    else stack.push(segment);
  }
  return '/' + stack.join('/');
}

class Part {
  constructor(pkg, uri, contentType) {
    this.pkg = pkg;
    this.uri = uri; // e.g. "/ppt/slides/slide1.xml"
    this.contentType = contentType;
    this._root = null;
    this._rels = null;
  }

  /** Raw bytes of the part (Uint8Array). part.GetStream() -> bytes. */
  getBytes() {
    return this.pkg.files[this.uri.slice(1)];
  }

  /** Decoded UTF-8 text of the part. */
  getText() {
    const bytes = this.getBytes();
    if (!bytes) return '';
    let text = Buffer.from(bytes).toString('utf8');
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    return text;
  }

  /** Parsed document element (root) of an XML part. Cached. */
  get root() {
    if (this._root === null) {
      const text = this.getText();
      if (!text) return null;
      try {
        this._root = _domParser.parseFromString(text, 'text/xml').documentElement;
      } catch {
        this._root = undefined;
      }
    }
    return this._root || null;
  }

  /** Relationships defined for this part, keyed by Id. */
  relationships() {
    if (this._rels !== null) return this._rels;
    const rels = new Map();
    const relsPartName = dirName(this.uri) + '/_rels/' + baseName(this.uri) + '.rels';
    const bytes = this.pkg.files[relsPartName.slice(1)];
    if (bytes) {
      let text = Buffer.from(bytes).toString('utf8');
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
      try {
        const relsRoot = _domParser.parseFromString(text, 'text/xml').documentElement;
        for (const rel of elements(relsRoot, 'Relationship')) {
          const id = attr(rel, 'Id');
          const type = attr(rel, 'Type');
          const target = attr(rel, 'Target');
          const mode = attr(rel, 'TargetMode');
          if (!id) continue;
          const external = mode === 'External';
          const absolute = external ? target : resolveRelativeTarget(dirName(this.uri), target);
          rels.set(id, { id, type, target, external, absolute });
        }
      } catch {
        /* ignore malformed .rels */
      }
    }
    this._rels = rels;
    return rels;
  }

  /** Resolve a relationship id to the target Part. GetPartById(relId). */
  getPartById(relId) {
    const rel = this.relationships().get(relId);
    if (!rel || rel.external) return null;
    return this.pkg.getPart(rel.absolute);
  }

  /** First related internal part of a given relationship type. */
  relatedPart(relType) {
    for (const rel of this.relationships().values()) {
      if (rel.type === relType && !rel.external) return this.pkg.getPart(rel.absolute);
    }
    return null;
  }

  /**
   * Media data-part relationships: [{ id, dataPart }].
   * Mirrors SlidePart.DataPartReferenceRelationships.
   */
  dataPartReferenceRelationships() {
    const out = [];
    for (const rel of this.relationships().values()) {
      if (rel.external) continue;
      if (
        rel.type === REL_TYPE.VIDEO ||
        rel.type === REL_TYPE.AUDIO ||
        rel.type === REL_TYPE.MEDIA
      ) {
        const dataPart = this.pkg.getPart(rel.absolute);
        if (dataPart) out.push({ id: rel.id, dataPart });
      }
    }
    return out;
  }
}

class OpcPackage {
  constructor(files) {
    this.files = files; // { "ppt/presentation.xml": Uint8Array, ... } (no leading slash)
    this._parts = new Map();
    this._contentTypes = this._readContentTypes();
  }

  _readContentTypes() {
    const defaults = new Map(); // extension(lower) -> contentType
    const overrides = new Map(); // partName -> contentType
    const bytes = this.files['[Content_Types].xml'];
    if (bytes) {
      let text = Buffer.from(bytes).toString('utf8');
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
      try {
        const root = _domParser.parseFromString(text, 'text/xml').documentElement;
        for (const d of elements(root, 'Default')) {
          const ext = (attr(d, 'Extension') || '').toLowerCase();
          defaults.set(ext, attr(d, 'ContentType'));
        }
        for (const o of elements(root, 'Override')) {
          overrides.set(attr(o, 'PartName'), attr(o, 'ContentType'));
        }
      } catch {
        /* ignore */
      }
    }
    return { defaults, overrides };
  }

  contentTypeFor(partName) {
    const override = this._contentTypes.overrides.get(partName);
    if (override) return override;
    const dot = partName.lastIndexOf('.');
    if (dot >= 0) {
      const ext = partName.slice(dot + 1).toLowerCase();
      const def = this._contentTypes.defaults.get(ext);
      if (def) return def;
    }
    return null;
  }

  /** Get a Part by absolute part name (e.g. "/ppt/slides/slide1.xml"). */
  getPart(partName) {
    partName = normalizePartName(partName);
    if (this._parts.has(partName)) return this._parts.get(partName);
    if (!(partName.slice(1) in this.files)) {
      this._parts.set(partName, null);
      return null;
    }
    const part = new Part(this, partName, this.contentTypeFor(partName));
    this._parts.set(partName, part);
    return part;
  }

  /** Package-level relationships (from /_rels/.rels). */
  rootRelationships() {
    const rels = new Map();
    const bytes = this.files['_rels/.rels'];
    if (!bytes) return rels;
    let text = Buffer.from(bytes).toString('utf8');
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    try {
      const root = _domParser.parseFromString(text, 'text/xml').documentElement;
      for (const rel of elements(root, 'Relationship')) {
        const id = attr(rel, 'Id');
        const type = attr(rel, 'Type');
        const target = attr(rel, 'Target');
        const mode = attr(rel, 'TargetMode');
        if (!id) continue;
        const external = mode === 'External';
        rels.set(id, {
          id,
          type,
          target,
          external,
          absolute: external ? target : resolveRelativeTarget('/', target),
        });
      }
    } catch {
      /* ignore */
    }
    return rels;
  }

  /** The main document part (ppt/presentation.xml). PresentationPart. */
  mainDocumentPart() {
    for (const rel of this.rootRelationships().values()) {
      if (rel.type === REL_TYPE.OFFICE_DOCUMENT && !rel.external) {
        return this.getPart(rel.absolute);
      }
    }
    return this.getPart('/ppt/presentation.xml');
  }
}

/** Open a .pptx file from disk. Returns { pkg, presentationPart }. */
function openPptx(pptxPath) {
  const buffer = fs.readFileSync(pptxPath);
  const files = unzipSync(new Uint8Array(buffer));
  const pkg = new OpcPackage(files);
  return { pkg, presentationPart: pkg.mainDocumentPart() };
}

module.exports = {
  REL_TYPE,
  Part,
  OpcPackage,
  openPptx,
  NS,
};
