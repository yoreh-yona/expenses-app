/* Receipt image processing: document detection, perspective correction, rotation, scan enhancement. */
(function (root) {
  'use strict';
  const mkCanvas = (w, h) => { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; };

  /** rotate by any angle (degrees, clockwise), canvas grows to fit, empty corners filled white */
  function rotateAny(src, deg) {
    deg = ((deg % 360) + 360) % 360;
    if (deg % 90 === 0) return rotate(src, deg);
    const t = deg * Math.PI / 180, c = Math.abs(Math.cos(t)), s = Math.abs(Math.sin(t));
    const W = Math.round(src.width * c + src.height * s), H = Math.round(src.width * s + src.height * c);
    const out = mkCanvas(W, H), x = out.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, W, H);
    x.translate(W / 2, H / 2); x.rotate(t); x.imageSmoothingQuality = 'high'; x.drawImage(src, -src.width / 2, -src.height / 2);
    return out;
  }
  /** where a source point lands after rotateAny(src, deg) */
  function rotPoint(p, w, h, deg) {
    const t = deg * Math.PI / 180, c = Math.cos(t), s = Math.sin(t);
    const W = w * Math.abs(c) + h * Math.abs(s), H = w * Math.abs(s) + h * Math.abs(c);
    const dx = p[0] - w / 2, dy = p[1] - h / 2;
    return [W / 2 + dx * c - dy * s, H / 2 + dx * s + dy * c, W, H];
  }
  /** crop with a rectangle normalised to 0..1 */
  function crop(src, r) {
    if (!r) return src;
    const x = Math.round(r.x * src.width), y = Math.round(r.y * src.height);
    const w = Math.max(2, Math.round(r.w * src.width)), h = Math.max(2, Math.round(r.h * src.height));
    const out = mkCanvas(w, h); out.getContext('2d').drawImage(src, x, y, w, h, 0, 0, w, h);
    return out;
  }
  /** detect the document and return the straightening angle + axis-aligned crop (in the rotated image) */
  function auto(src, baseDeg = 0) {
    const q = detect(src); if (!q) return null;
    const a1 = Math.atan2(q[1][1] - q[0][1], q[1][0] - q[0][0]), a2 = Math.atan2(q[2][1] - q[3][1], q[2][0] - q[3][0]);
    let angle = -((a1 + a2) / 2) * 180 / Math.PI;
    angle = Math.max(-45, Math.min(45, Math.round(angle * 2) / 2));
    const deg = baseDeg + angle;
    const pts = q.map(p => rotPoint(p, src.width, src.height, deg));
    const W = pts[0][2], H = pts[0][3];
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    // inner box (between the two outer extremes) avoids table edges after straightening
    xs.sort((a, b) => a - b); ys.sort((a, b) => a - b);
    const x0 = (xs[0] + xs[1]) / 2, x1 = (xs[2] + xs[3]) / 2, y0 = (ys[0] + ys[1]) / 2, y1 = (ys[2] + ys[3]) / 2;
    return {angle, crop: {x: Math.max(0, x0 / W), y: Math.max(0, y0 / H), w: Math.min(1, (x1 - x0) / W), h: Math.min(1, (y1 - y0) / H)}};
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

  root.ImgProc = {rotate, rotateAny, rotPoint, crop, auto, enhance, detect, mkCanvas};
})(window);
