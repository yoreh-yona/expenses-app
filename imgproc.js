/* Receipt image processing: document detection, perspective correction, rotation, scan enhancement. */
(function (root) {
  'use strict';
  const mkCanvas = (w, h) => { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; };

  function solve(A, b) { // Gaussian elimination with partial pivoting
    const n = b.length; A = A.map((r, i) => [...r, b[i]]);
    for (let c = 0; c < n; c++) {
      let p = c; for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
      [A[c], A[p]] = [A[p], A[c]];
      for (let r = c + 1; r < n; r++) { const f = A[r][c] / A[c][c]; for (let k = c; k <= n; k++) A[r][k] -= f * A[c][k]; }
    }
    const x = new Array(n);
    for (let r = n - 1; r >= 0; r--) { let s = A[r][n]; for (let k = r + 1; k < n; k++) s -= A[r][k] * x[k]; x[r] = s / A[r][r]; }
    return x;
  }
  // H maps destination (x,y) -> source (u,v)
  function homography(src, dst) {
    const A = [], b = [];
    for (let i = 0; i < 4; i++) {
      const [x, y] = dst[i], [u, v] = src[i];
      A.push([x, y, 1, 0, 0, 0, -u * x, -u * y]); b.push(u);
      A.push([0, 0, 0, x, y, 1, -v * x, -v * y]); b.push(v);
    }
    return solve(A, b);
  }
  const fullQuad = (w, h) => [[0, 0], [w, 0], [w, h], [0, h]];
  const isFull = (q, w, h) => !q || fullQuad(w, h).every((p, i) => Math.abs(p[0] - q[i][0]) < 2 && Math.abs(p[1] - q[i][1]) < 2);

  /** quad: [tl, tr, br, bl] in source pixels */
  function warp(src, quad, maxSide = 2200) {
    const d = (p, q) => Math.hypot(p[0] - q[0], p[1] - q[1]);
    let W = Math.max(d(quad[0], quad[1]), d(quad[3], quad[2])), H = Math.max(d(quad[0], quad[3]), d(quad[1], quad[2]));
    const k = Math.min(1, maxSide / Math.max(W, H)); W = Math.max(2, Math.round(W * k)); H = Math.max(2, Math.round(H * k));
    const h = homography(quad, [[0, 0], [W - 1, 0], [W - 1, H - 1], [0, H - 1]]);
    const sw = src.width, sh = src.height;
    const sd = src.getContext('2d').getImageData(0, 0, sw, sh).data;
    const out = mkCanvas(W, H), octx = out.getContext('2d'), img = octx.createImageData(W, H), od = img.data;
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const den = h[6] * x + h[7] * y + 1;
        let u = (h[0] * x + h[1] * y + h[2]) / den, v = (h[3] * x + h[4] * y + h[5]) / den;
        u = u < 0 ? 0 : u > sw - 1.001 ? sw - 1.001 : u; v = v < 0 ? 0 : v > sh - 1.001 ? sh - 1.001 : v;
        const x0 = u | 0, y0 = v | 0, fx = u - x0, fy = v - y0;
        const i00 = (y0 * sw + x0) * 4, i10 = i00 + 4, i01 = i00 + sw * 4, i11 = i01 + 4;
        const o = (y * W + x) * 4;
        for (let c = 0; c < 3; c++) {
          const a = sd[i00 + c] + (sd[i10 + c] - sd[i00 + c]) * fx, b = sd[i01 + c] + (sd[i11 + c] - sd[i01 + c]) * fx;
          od[o + c] = a + (b - a) * fy;
        }
        od[o + 3] = 255;
      }
    }
    octx.putImageData(img, 0, 0);
    return out;
  }

  function rotate(src, deg) {
    deg = ((deg % 360) + 360) % 360; if (!deg) return src;
    const swap = deg === 90 || deg === 270;
    const c = mkCanvas(swap ? src.height : src.width, swap ? src.width : src.height), x = c.getContext('2d');
    x.translate(c.width / 2, c.height / 2); x.rotate(deg * Math.PI / 180); x.drawImage(src, -src.width / 2, -src.height / 2);
    return c;
  }

  // background (paper + lighting) estimate per channel: downscale, dilate (removes ink), blur, upscale
  function background(src) {
    const W = src.width, H = src.height;
    const sw = Math.max(16, Math.round(W / 24)), sh = Math.max(16, Math.round(H / 24));
    const s = mkCanvas(sw, sh), sx = s.getContext('2d');
    sx.imageSmoothingQuality = 'high'; sx.drawImage(src, 0, 0, sw, sh);
    let d = sx.getImageData(0, 0, sw, sh);
    const dilate = (img, r) => {
      const a = img.data, o = new Uint8ClampedArray(a.length);
      for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
        let m0 = 0, m1 = 0, m2 = 0;
        for (let dy = -r; dy <= r; dy++) { const yy = Math.min(sh - 1, Math.max(0, y + dy));
          for (let dx = -r; dx <= r; dx++) { const xx = Math.min(sw - 1, Math.max(0, x + dx)), i = (yy * sw + xx) * 4;
            if (a[i] > m0) m0 = a[i]; if (a[i + 1] > m1) m1 = a[i + 1]; if (a[i + 2] > m2) m2 = a[i + 2]; } }
        const i = (y * sw + x) * 4; o[i] = m0; o[i + 1] = m1; o[i + 2] = m2; o[i + 3] = 255;
      }
      return new ImageData(o, sw, sh);
    };
    d = dilate(d, 2); d = dilate(d, 1);
    sx.putImageData(d, 0, 0);
    const b = mkCanvas(W, H), bx = b.getContext('2d');
    bx.filter = 'blur(' + Math.max(4, Math.round(W / 60)) + 'px)';
    bx.imageSmoothingQuality = 'high'; bx.drawImage(s, 0, 0, W, H);
    return bx.getImageData(0, 0, W, H).data;
  }

  /** mode: 'none' | 'color' (white paper, keeps colour) | 'bw' (crisp grey document) */
  function enhance(src, mode) {
    if (!mode || mode === 'none') return src;
    const W = src.width, H = src.height, ctx = src.getContext('2d');
    const img = ctx.getImageData(0, 0, W, H), d = img.data, bg = background(src);
    const lut = new Uint8ClampedArray(256);
    if (mode === 'color') {
      for (let v = 0; v < 256; v++) { const t = Math.max(0, (v - 18) / (245 - 18)); lut[v] = 255 * Math.pow(Math.min(1, t), 1.25); }
      for (let i = 0; i < d.length; i += 4) {
        for (let c = 0; c < 3; c++) { const r = d[i + c] * 255 / Math.max(40, bg[i + c]); d[i + c] = lut[r > 255 ? 255 : r | 0]; }
      }
    } else {
      for (let v = 0; v < 256; v++) { const t = Math.max(0, (v - 40) / (225 - 40)); lut[v] = 255 * Math.pow(Math.min(1, t), 1.6); }
      for (let i = 0; i < d.length; i += 4) {
        const g = d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114;
        const gb = bg[i] * 0.299 + bg[i + 1] * 0.587 + bg[i + 2] * 0.114;
        const r = g * 255 / Math.max(40, gb), v = lut[r > 255 ? 255 : r | 0];
        d[i] = d[i + 1] = d[i + 2] = v;
      }
    }
    const out = mkCanvas(W, H); out.getContext('2d').putImageData(img, 0, 0);
    return out;
  }

  /** Find the document (largest bright region) – returns quad in source pixels, or null */
  function detect(src) {
    const k = Math.min(1, 360 / Math.max(src.width, src.height));
    const w = Math.max(8, Math.round(src.width * k)), h = Math.max(8, Math.round(src.height * k));
    const c = mkCanvas(w, h), x = c.getContext('2d'); x.filter = 'blur(1px)'; x.drawImage(src, 0, 0, w, h);
    const d = x.getImageData(0, 0, w, h).data, n = w * h, g = new Uint8Array(n), hist = new Uint32Array(256);
    for (let i = 0; i < n; i++) { const v = (d[i * 4] * 0.299 + d[i * 4 + 1] * 0.587 + d[i * 4 + 2] * 0.114) | 0; g[i] = v; hist[v]++; }
    // Otsu threshold
    let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i];
    let sB = 0, wB = 0, best = 0, th = 128;
    for (let t = 0; t < 256; t++) { wB += hist[t]; if (!wB) continue; const wF = n - wB; if (!wF) break; sB += t * hist[t];
      const mB = sB / wB, mF = (sum - sB) / wF, between = wB * wF * (mB - mF) ** 2; if (between > best) { best = between; th = t; } }
    const lab = new Int32Array(n).fill(-1), stack = new Int32Array(n);
    let bestId = -1, bestSize = 0, bestPts = null;
    for (let s = 0; s < n; s++) {
      if (g[s] <= th || lab[s] !== -1) continue;
      let sp = 0, size = 0; stack[sp++] = s; lab[s] = s;
      let a = [Infinity, 0], b = [-Infinity, 0], cc = [-Infinity, 0], dd = [Infinity, 0]; // min(x+y) max(x-y) max(x+y) min(x-y)
      let pa, pb, pc, pd;
      while (sp) {
        const p = stack[--sp]; size++;
        const px = p % w, py = (p / w) | 0;
        if (px + py < a[0]) { a[0] = px + py; pa = [px, py]; }
        if (px - py > b[0]) { b[0] = px - py; pb = [px, py]; }
        if (px + py > cc[0]) { cc[0] = px + py; pc = [px, py]; }
        if (px - py < dd[0]) { dd[0] = px - py; pd = [px, py]; }
        const nb = [p - 1, p + 1, p - w, p + w];
        for (let j = 0; j < 4; j++) { const q = nb[j]; if (q < 0 || q >= n) continue;
          if ((j === 0 && px === 0) || (j === 1 && px === w - 1)) continue;
          if (g[q] > th && lab[q] === -1) { lab[q] = s; stack[sp++] = q; } }
      }
      if (size > bestSize) { bestSize = size; bestId = s; bestPts = [pa, pb, pc, pd]; }
    }
    if (!bestPts || bestSize < n * 0.15 || bestSize > n * 0.97) return null;
    const q = bestPts.map(p => [(p[0] + 0.5) / k, (p[1] + 0.5) / k]);
    // reject degenerate shapes
    const area = Math.abs(q.reduce((s, p, i) => { const r = q[(i + 1) % 4]; return s + p[0] * r[1] - r[0] * p[1]; }, 0)) / 2;
    if (area < src.width * src.height * 0.12) return null;
    return q;
  }

  root.ImgProc = {warp, rotate, enhance, detect, fullQuad, isFull, mkCanvas};
})(window);
