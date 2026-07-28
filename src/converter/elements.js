'use strict';

// Small shared element accessors used across converter modules, mirroring the
// OpenXML typed navigation for the common presentation shapes.

const { firstChild, childElements } = require('../ooxml/dom');

// P.Shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>()
function placeholderOf(sp) {
  return firstChild(firstChild(firstChild(sp, 'nvSpPr'), 'nvPr'), 'ph');
}

// shape.TextBody
function textBodyOf(shape) {
  return firstChild(shape, 'txBody');
}

// spPr child of a shape/picture/connection shape.
function shapePropsOf(shape) {
  return firstChild(shape, 'spPr');
}

// cSld.spTree for any layout/master/slide/notes root element.
function shapeTreeOf(root) {
  return firstChild(firstChild(root, 'cSld'), 'spTree');
}

function hasChildElements(el) {
  return childElements(el).length > 0;
}

module.exports = { placeholderOf, textBodyOf, shapePropsOf, shapeTreeOf, hasChildElements };
