#!/usr/bin/env node
'use strict';

// Port of PptxToHtml/Program.cs — the command-line entry point.

const path = require('path');
const { exportPptx } = require('../src/index');

function parseArgs(args) {
  const result = new Map();
  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith('--') && i + 1 < args.length && !args[i + 1].startsWith('--')) {
      result.set(args[i].toLowerCase(), args[i + 1]);
      i++;
    }
  }
  return result;
}

function main(argv) {
  const parsed = parseArgs(argv);

  const inputPath = parsed.get('--file');
  if (inputPath == null || inputPath === '') {
    console.error('Usage: pptx-to-html-conv --file <input.pptx> [--out <output-folder>] [--eot2ttf <path>]');
    return 1;
  }

  const outVal = parsed.get('--out');
  const outputDir =
    outVal != null && outVal !== ''
      ? outVal
      : path.join('out', path.basename(inputPath, path.extname(inputPath)));

  const eot2ttfPath = parsed.has('--eot2ttf') ? parsed.get('--eot2ttf') : null;

  try {
    exportPptx(inputPath, outputDir, eot2ttfPath);
  } catch (ex) {
    console.error(ex.message);
    return 1;
  }

  return 0;
}

process.exitCode = main(process.argv.slice(2));
