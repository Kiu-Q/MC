// Detect printed rectangular bubbles before measuring their interiors.
export function detectBubbles(data, width, height) {
  const gray = new Uint8Array(width * height);
  for (let i = 0; i < gray.length; i++) gray[i] = Math.round(.299 * data[i * 4] + .587 * data[i * 4 + 1] + .114 * data[i * 4 + 2]);
  const seen = new Uint8Array(gray.length),
    boxes = [],
    stack = [];
  for (let p = 0; p < gray.length; p++) {
    if (seen[p] || gray[p] > 180) continue;
    seen[p] = 1;
    stack.push(p);
    let x0 = width,
      y0 = height,
      x1 = 0,
      y1 = 0,
      count = 0;
    while (stack.length) {
      const q = stack.pop(),
        x = q % width,
        y = Math.floor(q / width);
      x0 = Math.min(x0, x);
      x1 = Math.max(x1, x);
      y0 = Math.min(y0, y);
      y1 = Math.max(y1, y);
      count++;
      for (const n of [x > 0 ? q - 1 : -1, x + 1 < width ? q + 1 : -1, y > 0 ? q - width : -1, y + 1 < height ? q + width : -1]) if (n >= 0 && !seen[n] && gray[n] <= 180) {
        seen[n] = 1;
        stack.push(n);
      }
    }
    const w = x1 - x0 + 1,
      h = y1 - y0 + 1;
    if (w > width * .012 && w < width * .032 && h > height * .002 && h < height * .012 && w / h > 1.6 && w / h < 7 && count > w) boxes.push({
      x: x0,
      y: y0,
      w,
      h,
      cx: (x0 + x1) / 2,
      cy: (y0 + y1) / 2
    });
  }
  return {
    gray,
    boxes
  };
}
export function classifyInterior(scores, labels) {
  const candidates = scores.map((s, i) => ({
    s,
    i
  })).filter(({
    s
  }) => s >= .40);
  return {
    label: candidates.length === 1 ? labels[candidates[0].i] : candidates.length ? 'MULT' : 'BLANK',
    selectedIndex: candidates.length === 1 ? candidates[0].i : -1,
    confidence: Math.max(...scores),
    averages: []
  };
}
// Continuous darkness keeps pale pencil marks that binary thresholding loses.
// A score of .40 requires about 24 gray levels of contrast against local paper.
export function measureBubble(gray, width, b) {
  const height = gray.length / width;
  const paper = [];
  for (let y = Math.max(0, Math.floor(b.y - b.h)); y < Math.min(height, b.y + b.h * 2); y++) {
    for (let x = Math.max(0, Math.floor(b.x - b.w * .3)); x < Math.min(width, b.x + b.w * 1.3); x++) {
      if (x < b.x || x >= b.x + b.w || y < b.y || y >= b.y + b.h) paper.push(gray[y * width + x]);
    }
  }
  paper.sort((a, b) => a - b);
  const background = paper.length ? paper[Math.floor((paper.length - 1) * .85)] : 255;
  const ix = Math.max(2, Math.round(b.w * .2));
  const iy = Math.max(1, Math.round(b.h * .35));
  let darkness = 0,
    count = 0,
    darkCount = 0;
  for (let y = Math.max(0, Math.round(b.y + iy)); y < Math.min(height, b.y + b.h - iy); y++) {
    for (let x = Math.max(0, Math.round(b.x + ix)); x < Math.min(width, b.x + b.w - ix); x++) {
      darkness += Math.max(0, background - gray[y * width + x]);
      if (gray[y * width + x] < 180) darkCount++;
      count++;
    }
  }
  return {
    contrast: count ? Math.min(1, darkness / count / 60) : 0,
    fill: count ? darkCount / count : 0
  };
}
export function bubbleScore(gray, width, b) {
  return measureBubble(gray, width, b).contrast;
}
// Dark fills are stronger evidence than pale erasure residue. When a row
// has no dark fill, use continuous contrast to retain uniformly light pencil.
export function classifyEvidence(evidence, labels) {
  const hasDarkMark = evidence.some(e => e.fill >= .40);
  return classifyInterior(evidence.map(e => hasDarkMark ? e.fill : e.contrast), labels);
}
export function detectSheet(data, width, height, prefix = 'p0', documentLayout = null) {
  const {
    gray,
    boxes
  } = detectBubbles(data, width, height);
  const rows = [];
  for (const b of boxes.filter(b => b.cy > height * .4 && b.cx > width * .31).sort((a, b) => a.cy - b.cy)) {
    let row = rows.find(r => Math.abs(r.cy - b.cy) < height * .004);
    if (!row) {
      row = {
        cy: b.cy,
        boxes: []
      };
      rows.push(row);
    }
    row.boxes.push(b);
  }
  const groups = [];
  for (const row of rows) {
    const bs = row.boxes.sort((a, b) => a.cx - b.cx);
    let g = [];
    for (const b of bs) {
      if (g.length && b.cx - g.at(-1).cx > width * .065) {
        if (g.length === 4) groups.push(g);
        g = [];
      }
      g.push(b);
    }
    if (g.length === 4) groups.push(g);
  }
  const columns = [];
  for (const g of groups) {
    let c = columns.find(c => Math.abs(c.x - g[0].cx) < width * .025);
    if (!c) {
      c = {
        x: g[0].cx,
        groups: []
      };
      columns.push(c);
    }
    c.groups.push(g);
  }
  const valid = columns.filter(c => c.groups.length >= 10).sort((a, b) => a.x - b.x);
  if (valid.length < 2) return recoverDocumentSheet(data, width, height, prefix, documentLayout, boxes, gray);
  // Use the row consensus across columns to preserve numbering when a
  // pencil stroke joins a printed outline and hides its component.
  const centers = [];
  for (const c of valid) for (const g of c.groups) {
    const y = g.reduce((s, b) => s + b.cy, 0) / 4;
    let r = centers.find(r => Math.abs(r.y - y) < height * .006);
    if (r) {
      r.ys.push(y);
      r.y = r.ys.reduce((a, b) => a + b, 0) / r.ys.length;
    } else centers.push({
      y,
      ys: [y]
    });
  }
  centers.sort((a, b) => a.y - b.y);
  if (centers.length !== 25 || valid.length !== 3) return recoverDocumentSheet(data, width, height, prefix, documentLayout, boxes, gray);
  for (const c of valid) {
    const original = c.groups;
    c.groups = centers.map(({
      y
    }) => {
      const g = original.find(g => Math.abs(g.reduce((s, b) => s + b.cy, 0) / 4 - y) < height * .006);
      if (g) return g;
      const near = original.reduce((a, b) => Math.abs(a[0].cy - y) < Math.abs(b[0].cy - y) ? a : b);
      // A damaged selected outline must not discard the three intact boxes.
      // Match those boxes first, then infer only missing cells from the local row.
      const referenceHeights = near.map(b => b.h).sort((a, b) => a - b);
      const referenceHeight = referenceHeights[Math.floor(referenceHeights.length / 2)];
      const matched = near.map(b => boxes.filter(candidate => candidate.h >= referenceHeight * .65 && candidate.h <= referenceHeight * 1.8 && Math.abs(candidate.cx - b.cx) < width * .012 && Math.abs(candidate.cy - y) < height * .008).sort((a, b) => Math.abs(a.cy - y) - Math.abs(b.cy - y))[0]);
      const anchors = matched.filter(Boolean);
      const median = values => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)];
      const typicalW = median(anchors.length ? anchors.map(b => b.w) : near.map(b => b.w));
      const typicalH = median(anchors.length ? anchors.map(b => b.h) : near.map(b => b.h));
      const mx = anchors.length ? anchors.reduce((s, b) => s + b.cx, 0) / anchors.length : 0;
      const my = anchors.length ? anchors.reduce((s, b) => s + b.cy, 0) / anchors.length : y;
      const variance = anchors.reduce((s, b) => s + (b.cx - mx) ** 2, 0);
      const slope = variance ? anchors.reduce((s, b) => s + (b.cx - mx) * (b.cy - my), 0) / variance : 0;
      return near.map((b, i) => {
        if (matched[i]) return matched[i];
        const cy = my + slope * (b.cx - mx);
        return {
          x: Math.round(b.cx - (typicalW - 1) / 2),
          y: Math.round(cy - (typicalH - 1) / 2),
          w: typicalW,
          h: typicalH,
          cx: b.cx,
          cy,
          inferred: true
        };
      });
    });
  }
  const regions = [],
    results = {};
  let q = 1;
  for (const c of valid) {
    c.groups.sort((a, b) => a[0].cy - b[0].cy);
    for (let i = 0; i < c.groups.length; i += 5) {
      const gs = c.groups.slice(i, i + 5),
        first = gs[0],
        last = gs.at(-1);
      const dx = (first[3].cx - first[0].cx) / 3,
        dy = gs.length > 1 ? (last[0].cy - first[0].cy) / (gs.length - 1) : height * .018;
      const id = `${prefix}_b${regions.length + 1}`,
        startQ = q;
      regions.push({
        id,
        type: 'answer',
        x: first[0].cx - dx / 2,
        y: first[0].cy - dy / 2,
        w: dx * 4,
        h: last[0].cy - first[0].cy + dy,
        rows: gs.length,
        cols: 4,
        startQ,
        labels: ['A', 'B', 'C', 'D'],
        hasGaps: false,
        gapHeightRatio: 1,
        bubbles: gs
      });
      results[id] = gs.map((g, rowIndex) => {
        const evidence = g.map(b => measureBubble(gray, width, b)),
          scores = evidence.map(e => e.contrast),
          cl = classifyEvidence(evidence, ['A', 'B', 'C', 'D']);
        return {
          ...cl,
          qNum: q++,
          rowIndex,
          detectedIndex: cl.selectedIndex,
          scores,
          evidence
        };
      });
    }
  }
  const idRegions = detectIdRegions(boxes, width, height, prefix, documentLayout, regions);
  for (const r of idRegions) {
    regions.push(r);
    const cl = readId(gray, width, r);
    results[r.id] = [{
      ...cl,
      qNum: cl.selectedIndex,
      rowIndex: cl.selectedIndex,
      detectedIndex: 0
    }];
  }
  return {
    regions,
    results,
    boxes
  };
}
export function readDetectedRegions(data, width, height, regions) {
  const gray = new Uint8Array(width * height);
  for (let i = 0; i < gray.length; i++) gray[i] = Math.round(.299 * data[i * 4] + .587 * data[i * 4 + 1] + .114 * data[i * 4 + 2]);
  const results = {};
  for (const r of regions) {
    if (r.type === 'id') {
      const cl = readId(gray, width, r);
      results[r.id] = [{
        ...cl,
        qNum: cl.selectedIndex,
        rowIndex: cl.selectedIndex,
        detectedIndex: 0
      }];
    } else results[r.id] = r.bubbles.map((g, rowIndex) => {
      const cl = classifyEvidence(g.map(b => measureBubble(gray, width, b)), r.labels);
      return {
        ...cl,
        qNum: r.startQ + rowIndex,
        rowIndex,
        detectedIndex: cl.selectedIndex
      };
    });
  }
  return results;
}
const median = values => {
  if (!values.length) return 0;
  const a = [...values].sort((a, b) => a - b);
  return (a[Math.floor((a.length - 1) / 2)] + a[Math.floor(a.length / 2)]) / 2;
};
export const ID_FIELDS = [{
  suffix: 'level',
  x: .61,
  labels: ['1', '2', '3', '4', '5', '6']
}, {
  suffix: 'letter',
  x: .745,
  labels: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
}, {
  suffix: 'n1',
  x: .885,
  labels: ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
}, {
  suffix: 'n2',
  x: .925,
  labels: ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
}];
const regionKey = r => r.type === 'id' ? `id:${r.id.split('_id_')[1]}` : `answer:${r.startQ}`;
const lineFit = points => {
  if (!points.length) return null;
  const slopes = [];
  for (let i = 0; i < points.length; i++) for (let j = i + 1; j < points.length; j++) if (points[j][0] !== points[i][0]) slopes.push((points[j][1] - points[i][1]) / (points[j][0] - points[i][0]));
  const slope = median(slopes);
  return {
    slope,
    intercept: median(points.map(([i, v]) => v - slope * i))
  };
};
// Learn geometry only. Student marks and recognized answers never enter this model.
export function learnDocumentLayout(pages) {
  const usable = pages.filter(p => p.regions?.some(r => r.type === 'answer' && r.bubbles));
  if (!usable.length) return null;
  const groups = new Map();
  for (const p of usable) for (const r of p.regions.filter(r => r.bubbles)) {
    const key = regionKey(r);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({
      p,
      r
    });
  }
  const regions = [];
  for (const group of groups.values()) {
    const base = group[0].r,
      agree = group.filter(({
        r
      }) => r.rows === base.rows && r.cols === base.cols);
    const norm = (b, p) => ({
      x: b.x / p.width,
      y: b.y / p.height,
      w: b.w / p.width,
      h: b.h / p.height,
      cx: b.cx / p.width,
      cy: b.cy / p.height
    });
    const bubbles = base.bubbles.map((row, i) => row.map((_, j) => {
      const cells = agree.map(({
        p,
        r
      }) => norm(r.bubbles[i][j], p));
      return Object.fromEntries(['x', 'y', 'w', 'h', 'cx', 'cy'].map(k => [k, median(cells.map(b => b[k]))]));
    }));
    regions.push({
      key: regionKey(base),
      type: base.type,
      suffix: base.id.split('_id_')[1],
      rows: base.rows,
      cols: base.cols,
      startQ: base.startQ,
      labels: [...base.labels],
      bubbles,
      support: agree.length
    });
  }
  return {
    version: 1,
    pageCount: usable.length,
    regions
  };
}
// Register the document pattern to each scan using intact answer boxes.
// Robust slopes and residual filtering tolerate scribbles and modest scan skew.
function documentTransform(model, regions, width, height) {
  const pairs = [];
  for (const r of regions.filter(r => r.type === 'answer' && r.bubbles)) {
    const ref = model.regions.find(t => t.key === regionKey(r));
    if (!ref) continue;
    r.bubbles.forEach((row, i) => row.forEach((b, j) => {
      const a = ref.bubbles[i]?.[j];
      if (a && !b.inferred) pairs.push({
        x: a.cx,
        y: a.cy,
        u: b.cx / width,
        v: b.cy / height
      });
    }));
  }
  if (pairs.length < 12) return {
    map: b => ({
      ...b
    }),
    reliable: false
  };
  // Answer columns give a wide x baseline; rows give a wide y baseline.
  const fit = (items, axis) => {
    const normal = [[0, 0, 0], [0, 0, 0], [0, 0, 0]],
      rhs = [0, 0, 0];
    for (const p of items) {
      const a = [1, p.x, p.y];
      for (let i = 0; i < 3; i++) {
        rhs[i] += a[i] * p[axis];
        for (let j = 0; j < 3; j++) normal[i][j] += a[i] * a[j];
      }
    }
    const m = normal.map((row, i) => [...row, rhs[i]]);
    for (let k = 0; k < 3; k++) {
      let best = k;
      for (let i = k + 1; i < 3; i++) if (Math.abs(m[i][k]) > Math.abs(m[best][k])) best = i;
      [m[k], m[best]] = [m[best], m[k]];
      if (Math.abs(m[k][k]) < 1e-10) return axis === 'u' ? [0, 1, 0] : [0, 0, 1];
      const d = m[k][k];
      for (let j = k; j < 4; j++) m[k][j] /= d;
      for (let i = 0; i < 3; i++) if (i !== k) {
        const f = m[i][k];
        for (let j = k; j < 4; j++) m[i][j] -= f * m[k][j];
      }
    }
    return m.map(row => row[3]);
  };
  let u = fit(pairs, 'u'),
    v = fit(pairs, 'v');
  const evalAt = (c, x, y) => c[0] + c[1] * x + c[2] * y;
  const clean = pairs.filter(p => Math.hypot(evalAt(u, p.x, p.y) - p.u, evalAt(v, p.x, p.y) - p.v) < .003);
  if (clean.length >= 12) {
    u = fit(clean, 'u');
    v = fit(clean, 'v');
  }
  return {
    reliable: clean.length >= 12,
    map: b => {
      const cx = evalAt(u, b.cx, b.cy),
        cy = evalAt(v, b.cx, b.cy),
        w = b.w * Math.hypot(u[1], v[1]),
        h = b.h * Math.hypot(u[2], v[2]);
      return {
        ...b,
        cx,
        cy,
        w,
        h,
        x: cx - w / 2,
        y: cy - h / 2
      };
    }
  };
}
function detectIdRegions(boxes, width, height, prefix, model, answers) {
  const idBoxes = boxes.filter(b => b.cy > height * .18 && b.cy < height * .4 && b.cx > width * .56);
  const typicalW = median(idBoxes.map(b => b.w)) || width * .019;
  const typicalH = median(idBoxes.map(b => b.h)) || height * .0055;
  const clean = idBoxes.filter(b => b.w < typicalW * 1.3 && b.h < typicalH * 1.45 && b.h > typicalH * .65);
  const groups = ID_FIELDS.map(f => clean.filter(b => Math.abs(b.cx / width - f.x) < .025).sort((a, b) => a.cy - b.cy));
  const steps = groups.flatMap(bs => bs.slice(1).map((b, i) => b.cy - bs[i].cy)).filter(d => d > height * .012 && d < height * .03);
  const step = median(steps) || height * .0214;
  const first = median(groups.filter(g => g.length).map(g => g[0].cy)) || height * .192;
  const transform = model ? documentTransform(model, answers, width, height) : null;
  return ID_FIELDS.map((field, fi) => {
    const reference = model?.regions.find(r => r.key === `id:${field.suffix}`);
    const predicted = reference?.bubbles.map(row => {
      const b = transform.map(row[0]);
      return {
        ...b,
        cx: b.cx * width,
        cy: b.cy * height,
        w: b.w * width,
        h: b.h * height
      };
    });
    const bs = groups[fi],
      x = median(bs.map(b => b.cx)) || width * field.x;
    const anchors = new Map();
    for (const b of bs) {
      let row = predicted ? predicted.reduce((best, p, i) => Math.abs(p.cy - b.cy) < Math.abs(predicted[best].cy - b.cy) ? i : best, 0) : Math.round((b.cy - first) / step);
      const expected = predicted?.[row]?.cy ?? first + row * step;
      if (row < 0 || row >= field.labels.length || Math.abs(b.cy - expected) > step * .32) continue;
      if (!anchors.has(row) || Math.abs(b.cy - expected) < Math.abs(anchors.get(row).cy - expected)) anchors.set(row, b);
    }
    const ax = [...anchors].map(([i, b]) => [i, b.cx]),
      ay = [...anchors].map(([i, b]) => [i, b.cy]);
    const xf = lineFit(ax),
      yf = lineFit(ay);
    const confident = anchors.size >= 3 || !!(predicted && transform.reliable && reference.support >= 2);
    const bubbles = field.labels.map((_, i) => {
      const p = predicted?.[i];
      const cx = anchors.size >= 3 ? xf.intercept + xf.slope * i : p?.cx ?? x;
      const cy = anchors.size >= 3 ? yf.intercept + yf.slope * i : p?.cy ?? first + i * step;
      const w = Math.max(3, Math.round(p?.w ?? typicalW)),
        h = Math.max(3, Math.round(p?.h ?? typicalH));
      const b = anchors.get(i);
      return [{
        x: Math.round(cx - (w - 1) / 2),
        y: Math.round(cy - (h - 1) / 2),
        w,
        h,
        cx,
        cy,
        inferred: !b,
        alignmentReliable: confident
      }];
    });
    const dy = anchors.size >= 3 ? yf.slope : step,
      id = `${prefix}_id_${field.suffix}`;
    return {
      id,
      type: 'id',
      x: bubbles[0][0].cx - typicalW,
      y: bubbles[0][0].cy - dy / 2,
      w: typicalW * 2,
      h: dy * field.labels.length,
      rows: field.labels.length,
      cols: 1,
      labels: [...field.labels],
      hasGaps: false,
      gapHeightRatio: 1,
      bubbles,
      alignmentReliable: confident,
      anchorCount: anchors.size
    };
  });
}
function readId(gray, width, region) {
  if (region.alignmentReliable === false) return {
    label: 'BLANK',
    selectedIndex: -1,
    confidence: 0,
    averages: [],
    needsReview: true
  };
  return classifyEvidence(region.bubbles.map(g => measureBubble(gray, width, g[0])), region.labels);
}
export function hasIncompleteId(regions, results) {
  const ids = regions.filter(r => r.type === 'id');
  if (!ids.length) return true;
  const standard = ids.some(r => ID_FIELDS.some(f => r.id.endsWith(`_id_${f.suffix}`)));
  if (standard && ID_FIELDS.some(f => !ids.some(r => r.id.endsWith(`_id_${f.suffix}`)))) return true;
  return ids.some(r => {
    const label = results?.[r.id]?.[0]?.label;
    return !label || ['BLANK', 'MULT', '?'].includes(label);
  });
}
// Whole-page fallback: borrow only geometry from the PDF consensus and
// register it to surviving printed boxes. Never borrow another pupil's marks.
function recoverDocumentSheet(data, width, height, prefix, model, boxes, gray) {
  if (!model) return null;
  const provisional = [];
  for (const ref of model.regions.filter(r => r.type === 'answer')) {
    const bubbles = ref.bubbles.map(row => row.map(a => {
      const candidates = boxes.filter(b => Math.abs(b.cx / width - a.cx) < .018 && Math.abs(b.cy / height - a.cy) < .007);
      const b = candidates.sort((b, c) => Math.hypot(b.cx / width - a.cx, b.cy / height - a.cy) - Math.hypot(c.cx / width - a.cx, c.cy / height - a.cy))[0];
      return b || {
        x: a.x * width,
        y: a.y * height,
        w: a.w * width,
        h: a.h * height,
        cx: a.cx * width,
        cy: a.cy * height,
        inferred: true
      };
    }));
    provisional.push({
      id: `${prefix}_b${provisional.length + 1}`,
      type: 'answer',
      startQ: ref.startQ,
      rows: ref.rows,
      cols: ref.cols,
      bubbles
    });
  }
  const observed = provisional.map(r => ({
    ...r,
    bubbles: r.bubbles.map(row => row.map(b => b.inferred ? null : b))
  }));
  // An entirely unrecognizable page cannot be safely registered.
  if (observed.flatMap(r => r.bubbles.flat()).filter(Boolean).length < 30) return null;
  const transform = documentTransform(model, provisional, width, height);
  if (!transform.reliable) return null;
  const regions = model.regions.filter(r => r.type === 'answer').map((ref, i) => {
    const bubbles = ref.bubbles.map(row => row.map(a => {
      const b = transform.map(a),
        w = Math.round(b.w * width),
        h = Math.round(b.h * height),
        cx = b.cx * width,
        cy = b.cy * height;
      return {
        x: Math.round(cx - (w - 1) / 2),
        y: Math.round(cy - (h - 1) / 2),
        w,
        h,
        cx,
        cy,
        inferred: true
      };
    }));
    const first = bubbles[0][0],
      last = bubbles.at(-1).at(-1),
      dx = (bubbles[0].at(-1).cx - first.cx) / (ref.cols - 1),
      dy = bubbles.length > 1 ? (bubbles.at(-1)[0].cy - first.cy) / (bubbles.length - 1) : height * .018;
    return {
      id: `${prefix}_b${i + 1}`,
      type: 'answer',
      startQ: ref.startQ,
      rows: ref.rows,
      cols: ref.cols,
      labels: [...ref.labels],
      bubbles,
      hasGaps: false,
      gapHeightRatio: 1,
      x: first.cx - dx / 2,
      y: first.cy - dy / 2,
      w: dx * ref.cols,
      h: last.cy - first.cy + dy
    };
  });
  regions.push(...detectIdRegions(boxes, width, height, prefix, model, provisional));
  return {
    regions,
    results: readDetectedRegions(data, width, height, regions),
    boxes,
    recoveredFromDocument: true
  };
}
