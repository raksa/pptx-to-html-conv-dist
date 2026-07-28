'use strict';

// Minimal StringBuilder mirroring System.Text.StringBuilder usage in the C#
// code. AppendLine uses "\n" (Environment.NewLine on the darwin/linux targets
// the golden output was generated on).

class StringBuilder {
  constructor() {
    this._parts = [];
    this._len = 0;
  }
  append(s) {
    s = s == null ? '' : String(s);
    if (s.length) {
      this._parts.push(s);
      this._len += s.length;
    }
    return this;
  }
  appendLine(s = '') {
    return this.append(s + '\n');
  }
  clear() {
    this._parts.length = 0;
    this._len = 0;
    return this;
  }
  get length() {
    return this._len;
  }
  toString() {
    return this._parts.join('');
  }
}

module.exports = { StringBuilder };
