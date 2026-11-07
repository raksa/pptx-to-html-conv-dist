'use strict';

// Numeric constants shared across converter modules (from PptxConverter.cs /
// TextRenderer.cs).

const pkg = require('../../package.json');

module.exports = {
  // CSS line-height ratio fallbacks (see PptxConverter.cs comments).
  DefaultPptLineHeightRatio: 1.2,
  DefaultLineHeightRatio: 1.4,
  HalfLeadingCompensationFactor: 1.0,

  // Random-bars slide transition tuning.
  RandomBarsRevealDurationFraction: 0.55,
  RandomBarsVerticalChunkTargetPx: 56.0,
  RandomBarsHorizontalChunkTargetPx: 40.0,
  RandomBarsChunkMinScale: 0.65,
  RandomBarsChunkMaxScale: 1.35,

  // TextRenderer shadow calibration.
  TextShadowDistanceScale: 0.78,
  TextShadowBlurScale: 0.72,
  TextShadowOpacityScale: 0.68,

  // `pkg.name` is the npm id ("pptx-to-html-conv"); info.json must carry the same
  // tool name the C# port writes, so it stays a literal here.
  ToolName: 'PptxToHtml',
  ToolVersion: pkg.version,
};
