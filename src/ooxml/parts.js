'use strict';

// Named part-navigation helpers mirroring the typed part relationships the
// C# converter walks (SlidePart.SlideLayoutPart, SlideLayoutPart.SlideMasterPart,
// SlideMasterPart.ThemePart, SlidePart.NotesSlidePart).

const { REL_TYPE } = require('./package');

function slideLayoutPart(slidePart) {
  return slidePart ? slidePart.relatedPart(REL_TYPE.SLIDE_LAYOUT) : null;
}

function slideMasterPart(layoutPart) {
  return layoutPart ? layoutPart.relatedPart(REL_TYPE.SLIDE_MASTER) : null;
}

function themePart(masterPart) {
  return masterPart ? masterPart.relatedPart(REL_TYPE.THEME) : null;
}

function notesSlidePart(slidePart) {
  return slidePart ? slidePart.relatedPart(REL_TYPE.NOTES_SLIDE) : null;
}

module.exports = {
  slideLayoutPart,
  slideMasterPart,
  themePart,
  notesSlidePart,
};
