import assert from 'node:assert/strict';
import { classifyInterior } from '../src/omr.mjs';
const labels = ['A', 'B', 'C', 'D'];
for (const [scores, label] of [[[0, 0, 0, 0], 'BLANK'], [[.25, 0, 0, 0], 'BLANK'], [[.31, 0, 0, 0], 'A'], [[0, .85, 0, 0], 'B'], [[.8, .8, 0, 0], 'MULT'], [[.9, .9, .9, .9], 'MULT']]) assert.equal(classifyInterior(scores, labels).label, label);
console.log('6 answer classification tests passed');
