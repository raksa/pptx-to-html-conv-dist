'use strict';

// Port of PptxToHtml/PptxToHtml.cs — the shared entry point for converting a
// .pptx deck to HTML assets.

const fs = require('fs');
const path = require('path');
const { openPptx } = require('./ooxml/package');
const { firstChild, elements, attr, NS } = require('./ooxml/dom');
const { PptxConverter } = require('./converter/PptxConverter');
const { ToolName, ToolVersion } = require('./converter/constants');

const isRelNs = (ns) => ns === NS.R || (ns || '').toLowerCase().includes('relationships');

function pad2(n) {
  return String(n).padStart(2, '0');
}

function formatExportedAt(date) {
  return (
    `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())} ` +
    `${pad2(date.getHours())}:${pad2(date.getMinutes())}:${pad2(date.getSeconds())}`
  );
}

/**
 * Convert a .pptx deck to HTML assets.
 * @param {string} pptxPath Path to the input .pptx file.
 * @param {string} outputDirectory Folder to write HTML and assets (created if missing).
 * @param {string|null} eot2ttfPath Path to the eot2ttf tool, or null to skip that step.
 */
function exportPptx(pptxPath, outputDirectory, eot2ttfPath = null) {
  if (pptxPath == null || String(pptxPath).trim() === '') throw new Error('pptxPath is required.');
  if (outputDirectory == null || String(outputDirectory).trim() === '') throw new Error('outputDirectory is required.');

  if (!fs.existsSync(pptxPath)) throw new Error(`File not found: ${pptxPath}`);

  if (eot2ttfPath != null && !fs.existsSync(eot2ttfPath)) throw new Error(`eot2ttf tool not found: ${eot2ttfPath}`);

  fs.mkdirSync(outputDirectory, { recursive: true });

  const { presentationPart } = openPptx(pptxPath);
  const presentationRoot = presentationPart ? presentationPart.root : null;
  const sldIdLst = firstChild(presentationRoot, 'sldIdLst');
  if (sldIdLst == null) throw new Error('Invalid or empty presentation.');

  const slideRelIds = elements(sldIdLst, 'sldId', NS.P).map((s) => attr(s, 'id', isRelNs));

  console.log(`Input:   ${pptxPath}`);
  console.log(`Output:  ${outputDirectory}`);
  if (eot2ttfPath != null) console.log(`eot2ttf: ${eot2ttfPath}`);
  console.log(`Slides:  ${slideRelIds.length}`);
  console.log('Extracting embedded fonts...');

  const exportedAt = formatExportedAt(new Date());
  const converter = new PptxConverter();
  converter.run(presentationPart, slideRelIds, outputDirectory, pptxPath, ToolName, ToolVersion, exportedAt, eot2ttfPath);
}

/**
 * Count the slides in a .pptx deck.
 * @param {string} pptxPath Path to the input .pptx file.
 * @returns {number} The number of slides in the deck.
 */
function getSlidesCount(pptxPath) {
  if (pptxPath == null || String(pptxPath).trim() === '') throw new Error('pptxPath is required.');

  if (!fs.existsSync(pptxPath)) throw new Error(`File not found: ${pptxPath}`);

  const { presentationPart } = openPptx(pptxPath);
  const presentationRoot = presentationPart ? presentationPart.root : null;
  const sldIdLst = firstChild(presentationRoot, 'sldIdLst');
  if (sldIdLst == null) throw new Error('Invalid or empty presentation.');

  return elements(sldIdLst, 'sldId', NS.P).length;
}

function getVersion() {
  return ToolVersion;
}

module.exports = { exportPptx, getSlidesCount, getVersion, PptxConverter };
