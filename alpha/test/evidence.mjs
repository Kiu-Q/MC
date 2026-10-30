import assert from 'node:assert/strict';
import { classifyEvidence, measureBubble } from '../src/omr.mjs';
const labels = ['A', 'B', 'C', 'D'];
const ev = (fill, contrast) => ({
  fill,
  contrast
});
for (const [evidence, label] of [[[ev(0, 0), ev(0, 0), ev(0, 0), ev(0, 0)], 'BLANK'], [[ev(0, .52), ev(0, .03), ev(0, .03), ev(0, .03)], 'A'], [[ev(.9, 1), ev(.1, .6), ev(0, 0), ev(0, 0)], 'A'], [[ev(.9, 1), ev(.8, 1), ev(0, 0), ev(0, 0)], 'MULT'], [[ev(0, .6), ev(0, .7), ev(0, .03), ev(0, .03)], 'MULT']]) assert.equal(classifyEvidence(evidence, labels).label, label);
const gray = new Uint8Array(100 * 100).fill(253),
  b = {
    x: 20,
    y: 20,
    w: 24,
    h: 10
  };
for (let y = 20; y < 30; y++) for (let x = 20; x < 44; x++) gray[y * 100 + x] = y === 20 || y === 29 || x === 20 || x === 43 ? 30 : 224;
const e = measureBubble(gray, 100, b);
assert.equal(e.fill, 0);
assert.ok(e.contrast >= .4);
gray.fill(253);
assert.equal(measureBubble(gray, 100, b).contrast, 0);
console.log('7 evidence and pixel checks passed');
