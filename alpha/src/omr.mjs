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
  }) => s >= .30);
  return {
    label: candidates.length === 1 ? labels[candidates[0].i] : candidates.length ? 'MULT' : 'BLANK',
    selectedIndex: candidates.length === 1 ? candidates[0].i : -1,
    confidence: Math.max(...scores),
    averages: []
  };
}
export function bubbleScore(gray, width, b) {
  let sum = 0,
    n = 0;
  // Exclude the printed outline, including its antialiased edge.
  const ix = Math.max(2, Math.round(b.w * .2)),
    iy = Math.max(1, Math.round(b.h * .35));
  for (let y = Math.max(0, Math.round(b.y + iy)); y < Math.min(gray.length / width, b.y + b.h - iy); y++) for (let x = Math.max(0, Math.round(b.x + ix)); x < Math.min(width, b.x + b.w - ix); x++) {
    sum += gray[y * width + x] < 180 ? 1 : 0;
    n++;
  }
  return n ? sum / n : 0;
}
export function detectSheet(data, width, height, prefix = 'p0') {
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
  if (valid.length < 2) return null;
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
  if (centers.length !== 25 || valid.length !== 3) return null;
  for (const c of valid) {
    const original = c.groups;
    c.groups = centers.map(({
      y
    }) => {
      const g = original.find(g => Math.abs(g.reduce((s, b) => s + b.cy, 0) / 4 - y) < height * .006);
      if (g) return g;
      const near = original.reduce((a, b) => Math.abs(a[0].cy - y) < Math.abs(b[0].cy - y) ? a : b);
      return near.map(b => ({
        ...b,
        y: Math.round(b.y + y - b.cy),
        cy: y
      }));
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
        const scores = g.map(b => bubbleScore(gray, width, b)),
          cl = classifyInterior(scores, ['A', 'B', 'C', 'D']);
        return {
          ...cl,
          qNum: q++,
          rowIndex,
          detectedIndex: cl.selectedIndex,
          scores
        };
      });
    }
  }
  // ID columns are recognized independently from the answer grid.
  const specs = [['level', .61, 6, ['1', '2', '3', '4', '5', '6']], ['letter', .745, 10, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']], ['n1', .885, 10, ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']], ['n2', .925, 10, ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']]];
  for (const [suffix, x, count, labels] of specs) {
    const bs = boxes.filter(b => Math.abs(b.cx / width - x) < .025 && b.cy / height > .18 && b.cy / height < .4).sort((a, b) => a.cy - b.cy);
    if (bs.length !== count) continue;
    const dy = (bs.at(-1).cy - bs[0].cy) / (count - 1),
      id = `${prefix}_id_${suffix}`;
    regions.push({
      id,
      type: 'id',
      x: bs[0].cx - bs[0].w,
      y: bs[0].cy - dy / 2,
      w: bs[0].w * 2,
      h: dy * count,
      rows: count,
      cols: 1,
      labels,
      hasGaps: false,
      gapHeightRatio: 1,
      bubbles: bs.map(b => [b])
    });
    const cl = classifyInterior(bs.map(b => bubbleScore(gray, width, b)), labels);
    results[id] = [{
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
      const cl = classifyInterior(r.bubbles.map(g => bubbleScore(gray, width, g[0])), r.labels);
      results[r.id] = [{
        ...cl,
        qNum: cl.selectedIndex,
        rowIndex: cl.selectedIndex,
        detectedIndex: 0
      }];
    } else results[r.id] = r.bubbles.map((g, rowIndex) => {
      const cl = classifyInterior(g.map(b => bubbleScore(gray, width, b)), r.labels);
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
