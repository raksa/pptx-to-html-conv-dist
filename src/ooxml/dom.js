'use strict';

// Thin DOM helper layer over @xmldom/xmldom that mirrors the navigation
// surface of the DocumentFormat.OpenXml SDK used by the C# converter:
//   GetFirstChild<T>()  -> firstChild(el, localName)
//   Elements<T>()       -> elements(el, localName)
//   Descendants<T>()    -> descendants(el, localName)
//   GetAttributes()     -> attrList(el) / attr(el, localName)
//   ChildElements       -> childElements(el)
//   InnerText           -> innerText(el)
//   OuterXml            -> outerXml(el)
//
// Every OOXML element is represented as a standard DOM Element. `localName`
// and `namespaceURI` are native DOM properties and are namespace-aware.

const { XMLSerializer } = require('@xmldom/xmldom');

const ELEMENT_NODE = 1;

const NS = {
  A: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  P: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  R: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  C: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  P14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  CT: 'http://schemas.openxmlformats.org/package/2006/content-types',
  REL: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

const _serializer = new XMLSerializer();

/** Iterate direct element children of `el` in document order. */
function* iterElementChildren(el) {
  if (!el) return;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === ELEMENT_NODE) yield n;
  }
}

/** All direct element children as an array (ChildElements). */
function childElements(el) {
  return Array.from(iterElementChildren(el));
}

function nameMatches(node, localName, ns) {
  if (localName != null && node.localName !== localName) return false;
  if (ns != null && node.namespaceURI !== ns) return false;
  return true;
}

/** First direct child element matching localName (and optional namespace). GetFirstChild<T>(). */
function firstChild(el, localName, ns) {
  if (!el) return null;
  for (const child of iterElementChildren(el)) {
    if (nameMatches(child, localName, ns)) return child;
  }
  return null;
}

/** Direct child elements matching localName (and optional namespace). Elements<T>(). */
function elements(el, localName, ns) {
  const out = [];
  if (!el) return out;
  for (const child of iterElementChildren(el)) {
    if (localName == null || nameMatches(child, localName, ns)) out.push(child);
  }
  return out;
}

/** Depth-first descendant elements (document order). Descendants(). */
function* iterDescendants(el) {
  if (!el) return;
  for (const child of iterElementChildren(el)) {
    yield child;
    yield* iterDescendants(child);
  }
}

/** Descendant elements matching localName (and optional namespace). Descendants<T>(). */
function descendants(el, localName, ns) {
  const out = [];
  for (const d of iterDescendants(el)) {
    if (localName == null || nameMatches(d, localName, ns)) out.push(d);
  }
  return out;
}

/** First descendant element matching a predicate. */
function firstDescendant(el, predicate) {
  for (const d of iterDescendants(el)) {
    if (predicate(d)) return d;
  }
  return null;
}

/** All descendants (no filter). */
function allDescendants(el) {
  return Array.from(iterDescendants(el));
}

/**
 * Attribute value by local name, matching irrespective of prefix.
 * Returns null when absent. Mirrors
 * GetAttributes().FirstOrDefault(a => a.LocalName == name).Value.
 * `nsPredicate` optionally constrains the attribute namespace.
 */
function attr(el, localName, nsPredicate) {
  if (!el || !el.attributes) return null;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.localName !== localName) continue;
    if (nsPredicate && !nsPredicate(a.namespaceURI || '')) continue;
    return a.value;
  }
  return null;
}

/** List of attributes as {localName, namespaceURI, value}. GetAttributes(). */
function attrList(el) {
  const out = [];
  if (!el || !el.attributes) return out;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    out.push({ localName: a.localName, namespaceURI: a.namespaceURI || '', value: a.value });
  }
  return out;
}

/** Concatenated text of all descendants. InnerText. */
function innerText(el) {
  return el ? (el.textContent || '') : '';
}

/** Serialized XML of the element and its subtree. OuterXml. */
function outerXml(el) {
  return el ? _serializer.serializeToString(el) : '';
}

// ── attribute value coercion helpers ─────────────────────────────────────

function attrInt(el, localName, fallback = null) {
  const v = attr(el, localName);
  if (v == null) return fallback;
  const n = parseInt(v, 10);
  return Number.isNaN(n) ? fallback : n;
}

function attrLong(el, localName, fallback = null) {
  // JS numbers safely represent EMU values well within Number.MAX_SAFE_INTEGER.
  const v = attr(el, localName);
  if (v == null) return fallback;
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

/** OOXML boolean attribute: absent -> null; "1"/"true" -> true; else false. */
function attrBool(el, localName) {
  const v = attr(el, localName);
  if (v == null) return null;
  return v === '1' || v.toLowerCase() === 'true';
}

module.exports = {
  NS,
  ELEMENT_NODE,
  iterElementChildren,
  childElements,
  firstChild,
  elements,
  descendants,
  iterDescendants,
  firstDescendant,
  allDescendants,
  attr,
  attrList,
  attrInt,
  attrLong,
  attrBool,
  innerText,
  outerXml,
};
