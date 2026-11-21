'use strict';

// Port of PptxToHtml/pptx/IndexPageHelper.cs.
//
// index.html is a self-contained viewer written next to the slides: slide show, presenter
// view, grid and everything in info.json. The page itself is a fixed template
// (indexPage.html, a byte-identical copy of the C# PptxToHtml/pptx/IndexPage.html); an export
// only fills in its info.json and whether the slides were exported without a background.

const fs = require('fs');
const path = require('path');

const IndexPageFileName = 'index.html';
const IndexPageInfoPlaceholder = '{{PPTX_INFO_JSON}}';
const IndexPageNoBackgroundPlaceholder = '{{PPTX_NO_BACKGROUND}}';

let indexPageTemplate = null;

function loadIndexPageTemplate() {
  if (indexPageTemplate == null) {
    indexPageTemplate = fs.readFileSync(path.join(__dirname, 'indexPage.html'), 'utf8').replace(/^﻿/, '');
  }
  return indexPageTemplate;
}

// split/join rather than String.replace, whose replacement string would expand `$&`, `$1`, …
// written in a note or file name.
function replaceOrdinal(text, placeholder, value) {
  return text.split(placeholder).join(value);
}

module.exports = {
  // The viewer page for one export. The info JSON goes into a <script type="application/json">
  // as it is: the serializer writes <, >, & and every non-ASCII character as \uXXXX, so it
  // cannot close the script element early. It is put in last, so a placeholder spelled out in
  // a note or file name is left alone.
  buildIndexPageHtml(infoJson, noBackground) {
    let html = loadIndexPageTemplate();
    html = replaceOrdinal(html, IndexPageNoBackgroundPlaceholder, noBackground ? 'true' : 'false');
    html = replaceOrdinal(html, IndexPageInfoPlaceholder, infoJson);
    return html;
  },

  writeIndexPage(outputDir, infoJson) {
    const indexPath = path.join(outputDir, IndexPageFileName);
    fs.writeFileSync(indexPath, '﻿' + this.buildIndexPageHtml(infoJson, this._noBackground), 'utf8');
    console.log(`Viewer written -> ${IndexPageFileName}`);
  },
};
