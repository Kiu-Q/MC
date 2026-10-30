import assert from 'node:assert/strict';
import { detectSheet, detectBubbles, learnDocumentLayout, readDetectedRegions, hasIncompleteId } from '../src/omr.mjs';
const w = 1191,
  h = 1684,
  data = new Uint8Array(w * h * 4).fill(255);
const rect = (x, y, bw, bh, fill = 0) => {
  for (let py = y; py < y + bh; py++) for (let px = x; px < x + bw; px++) if (fill || py === y || py === y + bh - 1 || px === x || px === x + bw - 1) {
    const i = (py * w + px) * 4;
    data[i] = data[i + 1] = data[i + 2] = fill || 30;
  }
};
for (let c = 0; c < 3; c++) for (let r = 0; r < 25; r++) for (let option = 0; option < 4; option++) rect(413 + c * 255 + option * 48, 736 + r * 31 + Math.floor(r / 5) * 15, 24, 9);
for (const [column, x, count, selected] of [[0, 711, 6, 2], [1, 868, 10, 3], [2, 1033, 10, 1], [3, 1081, 10, 0]]) for (let row = 0; row < count; row++) {
  rect(x, 320 + row * 36, 24, 9, row === selected ? 70 : 0);
  if (row === selected && [1, 2].includes(column)) rect(x - 9, 323 + row * 36, 43, 3, 70);
}
const readId = sheet => sheet.regions.filter(r => r.type === 'id').map(r => sheet.results[r.id][0].label).join('');
const sheet = detectSheet(data, w, h);
assert.ok(sheet);
assert.equal(readId(sheet), '3D10');
assert.deepEqual(sheet.regions.filter(r => r.type === 'id').map(r => r.rows), [6, 10, 10, 10]);
assert.equal(sheet.regions.find(r => r.id.endsWith('_id_letter')).bubbles[3][0].inferred, true);
assert.equal(sheet.regions.find(r => r.id.endsWith('_id_n1')).bubbles[1][0].inferred, true);
const model = learnDocumentLayout([{
  width: w,
  height: h,
  regions: sheet.regions
}, {
  width: w,
  height: h,
  regions: sheet.regions
}]);
assert.equal(model.pageCount, 2);
assert.equal(JSON.stringify(model).includes('confidence'), false);
assert.equal(JSON.stringify(model).includes('selectedIndex'), false);
const erased = new Uint8Array(data);
for (let y = 300; y < 680; y++) for (let x = 680; x < 1120; x++) erased.fill(255, (y * w + x) * 4, (y * w + x) * 4 + 4);
const recovered = detectSheet(erased, w, h, 'p0', model);
assert.ok(recovered);
assert.equal(recovered.regions.filter(r => r.type === 'id').length, 4);
assert.equal(readId(recovered), 'BLANKBLANKBLANKBLANK');
assert.ok(recovered.regions.filter(r => r.type === 'id').every(r => r.alignmentReliable));
assert.equal(hasIncompleteId(recovered.regions, recovered.results), true);
assert.equal(hasIncompleteId(sheet.regions.filter(r => !r.id.endsWith('_id_n1')), sheet.results), true);
assert.equal(hasIncompleteId(sheet.regions.filter(r => !r.id.endsWith('_id_level')), sheet.results), true);
assert.equal(hasIncompleteId(sheet.regions, sheet.results), false);
const shifted = new Uint8Array(data.length).fill(255);
for (let y = 4; y < h; y++) for (let x = 0; x < w - 5; x++) shifted.set(data.subarray((y * w + x) * 4, (y * w + x) * 4 + 4), ((y - 4) * w + x + 5) * 4);
const moved = detectSheet(shifted, w, h, 'p0', model);
assert.equal(readId(moved), '3D10');
const missingAnswers = new Uint8Array(data);
for (let y = 720; y < h; y++) for (let x = 400; x < 600; x++) missingAnswers.fill(255, (y * w + x) * 4, (y * w + x) * 4 + 4);
const fallback = detectSheet(missingAnswers, w, h, 'p0', model);
assert.ok(fallback.recoveredFromDocument);
assert.equal(fallback.regions.filter(r => r.type === 'answer').reduce((s, r) => s + r.rows, 0), 75);
assert.ok(fallback.results.p0_b1.every(r => r.label === 'BLANK'));
assert.deepEqual(Object.values(readDetectedRegions(data, w, h, sheet.regions)).flat().map(r => r.label), Object.values(sheet.results).flat().map(r => r.label));
console.log('ID recovery, row-label stability, document registration, erased-ID safety, missing-field validation and missing-answer-column tests passed');
