'use strict';

// Case-insensitive collections mirroring the C# StringComparer.OrdinalIgnoreCase
// dictionaries and sorted sets used by the converter.

/** Ordinal-ignore-case comparison of two strings (invariant uppercase, code-unit order). */
function ordinalIgnoreCaseCompare(a, b) {
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  if (ua < ub) return -1;
  if (ua > ub) return 1;
  // Tie-break by ordinal (matches .NET SortedSet stability for equal-ignoring-case keys).
  if (a < b) return -1;
  if (a > b) return 1;
  return 0;
}

/** Dictionary keyed case-insensitively; preserves the value only (keys not enumerated). */
class CIMap {
  constructor(entries) {
    this._m = new Map();
    if (entries) {
      for (const [k, v] of entries) this.set(k, v);
    }
  }
  _k(key) {
    return (key == null ? '' : String(key)).toLowerCase();
  }
  get(key) {
    return this._m.get(this._k(key));
  }
  set(key, value) {
    this._m.set(this._k(key), value);
    return this;
  }
  has(key) {
    return this._m.has(this._k(key));
  }
  delete(key) {
    return this._m.delete(this._k(key));
  }
  getOrDefault(key, fallback) {
    const lk = this._k(key);
    return this._m.has(lk) ? this._m.get(lk) : fallback;
  }
  clear() {
    this._m.clear();
  }
  get size() {
    return this._m.size;
  }
}

/**
 * Sorted set with OrdinalIgnoreCase semantics: case-insensitive dedup keeping
 * the first-added casing; enumeration/toArray yields the sorted order.
 */
class CISet {
  constructor() {
    this._m = new Map(); // lowercased -> original casing (first seen)
  }
  add(value) {
    const lk = value.toLowerCase();
    if (!this._m.has(lk)) this._m.set(lk, value);
    return this;
  }
  has(value) {
    return this._m.has(value.toLowerCase());
  }
  clear() {
    this._m.clear();
  }
  get size() {
    return this._m.size;
  }
  get count() {
    return this._m.size;
  }
  toArray() {
    return Array.from(this._m.values()).sort(ordinalIgnoreCaseCompare);
  }
  join(sep) {
    return this.toArray().join(sep);
  }
}

module.exports = { CIMap, CISet, ordinalIgnoreCaseCompare };
