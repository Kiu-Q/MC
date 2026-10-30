import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import { detectSheet, readDetectedRegions } from '../src/omr.mjs';
const reviewed = JSON.parse(fs.readFileSync(new URL('./reviewed-cases.json', import.meta.url), 'utf8'));
const expected = ['DAAD DDBCDBBCCBBACABD'.replaceAll(' ', ''), '_BADB', 'CADBACBBAC', 'DABDACDBBBDBADACABBD', 'BDCAC CDBB_'.replaceAll(' ', ''), 'DDCACADADBBACAB DABCC'.replaceAll(' ', '')];
const inputDir = process.argv[2] || '/Users/kiuq/Downloads';
const files = fs.readdirSync(inputDir).filter(f => f.endsWith('.pdf') && f.includes('Mathematics')).sort();
const report = [];
for (let fi = 0; fi < files.length; fi++) {
  const file = files[fi],
    path = inputDir + '/' + file;
  const info = spawnSync('pdfinfo', [path], {
    encoding: 'utf8'
  }).stdout;
  const pages = +info.match(/Pages:\s+(\d+)/)[1];
  let failures = [];
  for (let page = 1; page <= pages; page++) {
    const out = spawnSync('pdftoppm', ['-f', String(page), '-l', String(page), '-singlefile', '-scale-to', '1684', path], {
      maxBuffer: 30e6
    });
    if (out.status !== 0) throw Error(out.stderr.toString());
    const buf = out.stdout;
    const header = buf.toString('ascii', 0, 100).match(/^P6\s+(\d+)\s+(\d+)\s+255\s/);
    if (!header) throw Error('ppm');
    const w = +header[1],
      h = +header[2],
      offset = header[0].length,
      data = new Uint8Array(w * h * 4);
    for (let i = 0; i < w * h; i++) {
      data[i * 4] = buf[offset + i * 3];
      data[i * 4 + 1] = buf[offset + i * 3 + 1];
      data[i * 4 + 2] = buf[offset + i * 3 + 2];
      data[i * 4 + 3] = 255;
    }
    const s = detectSheet(data, w, h);
    const read = s ? readDetectedRegions(data, w, h, s.regions) : {};
    if (s && JSON.stringify(Object.values(read).flat().map(r => r.label)) !== JSON.stringify(Object.values(s.results).flat().map(r => r.label))) throw Error('reading mismatch');
    const answers = Object.values(s?.results || {}).flat().filter(r => r.qNum > 0 && r.detectedIndex !== undefined && s.regions.some(g => g.type === 'answer' && s.results[g.id].includes(r)));
    if (answers.length !== 75) failures.push({
      page,
      count: answers.length
    });
    for (const c of reviewed.filter(c => c.file === file && c.page === page)) {
      const actual = answers.find(r => r.qNum === c.q)?.label;
      if (actual !== c.label) failures.push({page, q: c.q, actual, expected:c.label});
    }
    if (page === 1) {
      const actual = answers.map(r => r.label === 'BLANK' ? '_' : r.label).join('');
      const exp = expected[fi].padEnd(75, '_');
      if (actual !== exp) failures.push({
        page,
        actual,
        expected: exp
      });
    }
  }
  const r = {
    file,
    pages,
    failures
  };
  report.push(r);
  console.log(JSON.stringify(r));
  if (process.argv[3]) fs.writeFileSync(process.argv[3], JSON.stringify(report, null, 2));
}
if (report.some(r => r.failures.length)) process.exitCode = 1;
