/**
 * Offizin studio measurements: zoom steps, fitting a page to the pasteboard, ruler ticks in
 * inches or millimeters, paper names, and printer's marks.
 *
 * Everything here is plain math and string building with no browser dependencies, so the
 * tests run it in Node. The studio page (studio.js) puts the results on screen.
 */
import { PAPER } from "../core/util.js";

/** Screen pixels per point at 100 percent zoom (CSS: 96 pixels and 72 points per inch). */
export const PX_PER_PT = 96 / 72;

/** Zoom levels people can pick, as scale factors. "fit" fits the page to the pasteboard. */
export const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5];
export const ZOOM_CHOICES = ["fit", ...ZOOM_STEPS];

/** Turns a zoom choice ("fit", "0.75", 0.75) into "fit" or a number within the steps' range. */
export function normalizeZoom(value) {
  if (value === "fit" || value == null || value === "") return "fit";
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return "fit";
  return Math.min(ZOOM_STEPS[ZOOM_STEPS.length - 1], Math.max(ZOOM_STEPS[0], n));
}

/**
 * The scale that fits a sheet into the available space (in pixels). Fits the whole sheet when
 * that keeps it reasonably large; on narrow, tall screens it fits the width instead, so the page
 * stays readable and you scroll down it.
 */
export function fitScale({ availWidth, availHeight, sheetWidth, sheetHeight, max = 1.5, min = 0.2 }) {
  const w = Math.max(1, availWidth) / (Math.max(1, sheetWidth) * PX_PER_PT);
  const h = availHeight > 0 ? availHeight / (Math.max(1, sheetHeight) * PX_PER_PT) : Infinity;
  const s = h >= w * 0.5 ? Math.min(w, h) : w;
  return Math.round(Math.min(max, Math.max(min, s)) * 10000) / 10000;
}

/** The next zoom step above (dir 1) or below (dir -1) the current scale. */
export function stepZoom(scale, dir) {
  const s = Number(scale) || 1;
  if (dir > 0) return ZOOM_STEPS.find((z) => z > s + 0.001) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1];
  for (let i = ZOOM_STEPS.length - 1; i >= 0; i--) if (ZOOM_STEPS[i] < s - 0.001) return ZOOM_STEPS[i];
  return ZOOM_STEPS[0];
}

/** "75%" */
export function zoomPercent(scale) {
  return `${Math.round((Number(scale) || 0) * 100)}%`;
}

// ---------------------------------------------------------------------------
// Paper and units
// ---------------------------------------------------------------------------

const METRIC = new Set(["a3", "a4", "a5"]);
const SHORT_NAMES = { letter: "Letter", legal: "Legal", tabloid: "Tabloid", halfLetter: "Half letter", a3: "A3", a4: "A4", a5: "A5", card5x7: "Card 5 x 7" };

/** The PAPER key whose size matches width and height in points (either orientation), or null. */
export function matchPaper(width, height) {
  for (const [key, p] of Object.entries(PAPER)) {
    const same = (a, b) => Math.abs(a - b) < 1;
    if ((same(p.width, width) && same(p.height, height)) || (same(p.width, height) && same(p.height, width))) return key;
  }
  return null;
}

/** "mm" for metric paper (A3, A4, A5), "in" for the others; custom sizes follow the chosen paper. */
export function unitFor(page, paper) {
  const name = (page && page.sizeName) || (page ? matchPaper(page.width, page.height) : null);
  if (name) return METRIC.has(name) ? "mm" : "in";
  return paper === "a4" ? "mm" : "in";
}

function trim(n, places) {
  return String(Math.round(n * 10 ** places) / 10 ** places);
}

/** A length in points as text in a unit: "8.5 in", "210 mm". */
export function formatLength(pt, unit) {
  return unit === "mm" ? `${trim((pt * 25.4) / 72, 0)} mm` : `${trim(pt / 72, 2)} in`;
}

/** "8.5 x 11 in" or "210 x 297 mm". */
export function sizeText(page, unit) {
  const u = unit || unitFor(page);
  const a = formatLength(page.width, u).split(" ")[0];
  const b = formatLength(page.height, u).split(" ")[0];
  return `${a} x ${b} ${u}`;
}

/** The paper's short name ("Letter", "A4") or "Custom" with its size. */
export function paperName(page, unit) {
  const key = (page && page.sizeName) || matchPaper(page.width, page.height);
  if (key && SHORT_NAMES[key]) return SHORT_NAMES[key];
  return `Custom, ${sizeText(page, unit)}`;
}

/** "Portrait" or "Landscape". */
export function orientationName(page) {
  return page.width > page.height ? "Landscape" : "Portrait";
}

// ---------------------------------------------------------------------------
// Rulers
// ---------------------------------------------------------------------------

/**
 * How a ruler divides its unit at a given zoom.
 * in: one numbered inch, divided into 8, 4, 2, or 1 parts.
 * mm: one numbered centimeter (labeled in millimeters: 10, 20, 30), divided into 10, 2, or 1 parts.
 */
export function rulerScale(unit, pxPerPt) {
  if (unit === "mm") {
    const majorPt = (10 * 72) / 25.4;
    const majorPx = majorPt * pxPerPt;
    const divisions = majorPx >= 45 ? 10 : majorPx >= 18 ? 2 : 1;
    const labelEvery = majorPx >= 26 ? 1 : majorPx >= 13 ? 2 : 5;
    return { unit, majorPt, majorPx, divisions, labelEvery, labelStep: 10 };
  }
  const majorPt = 72;
  const majorPx = majorPt * pxPerPt;
  const divisions = majorPx >= 80 ? 8 : majorPx >= 40 ? 4 : majorPx >= 20 ? 2 : 1;
  const labelEvery = majorPx >= 24 ? 1 : 2;
  return { unit: "in", majorPt, majorPx, divisions, labelEvery, labelStep: 1 };
}

/**
 * Tick marks for a ruler `length` pixels long whose zero sits at `origin` pixels (the page
 * edge; it may be outside the ruler). Returns [{ pos, level, label? }]: level 0 is a numbered
 * major tick, 1 a half tick, 2 a small tick. Labels count away from zero in both directions,
 * the way rulers in desktop publishing programs do.
 */
export function rulerTicks({ length, origin, pxPerPt, unit }) {
  const sc = rulerScale(unit, pxPerPt);
  const minorPx = sc.majorPx / sc.divisions;
  if (!(minorPx > 0) || !(length > 0)) return [];
  const first = Math.ceil((0 - origin) / minorPx);
  const last = Math.floor((length - origin) / minorPx);
  const ticks = [];
  if (last - first > 5000) return ticks;
  const half = sc.divisions % 2 === 0 ? sc.divisions / 2 : 0;
  for (let i = first; i <= last; i++) {
    const pos = Math.round((origin + i * minorPx) * 100) / 100;
    const r = ((i % sc.divisions) + sc.divisions) % sc.divisions;
    if (r === 0) {
      const major = Math.abs(i / sc.divisions);
      const tick = { pos, level: 0 };
      if (major % sc.labelEvery === 0) tick.label = String(major * sc.labelStep);
      ticks.push(tick);
    } else {
      ticks.push({ pos, level: half && r % half === 0 ? 1 : 2 });
    }
  }
  return ticks;
}

/**
 * Where a sheet and its printable areas sit on a ruler, in ruler pixels.
 * axis "x" uses the left and right margins; "y" uses top and bottom. A sheet that holds two
 * pages side by side (a printer spread) gets one printable band per page.
 */
export function rulerBands({ axis, origin, pxPerPt, page, sheetPt }) {
  const m = page.margins || { top: 0, right: 0, bottom: 0, left: 0 };
  const pageLen = axis === "x" ? page.width : page.height;
  const total = sheetPt || pageLen;
  const across = axis === "x" ? Math.max(1, Math.round(total / pageLen)) : 1;
  const lo = axis === "x" ? m.left : m.top;
  const hi = axis === "x" ? m.right : m.bottom;
  const px = (pt) => Math.round((origin + pt * pxPerPt) * 100) / 100;
  const bands = [];
  for (let k = 0; k < across; k++) bands.push({ from: px(k * pageLen + lo), to: px((k + 1) * pageLen - hi) });
  return { sheet: { from: px(0), to: px(total) }, bands };
}

/**
 * The ruler as an SVG string. orientation "h" draws ticks up from the bottom edge, "v" draws
 * them in from the right edge with labels turned like a vertical ruler. Colors come from CSS
 * classes in studio.css: rl-bg, rl-sheet, rl-band, rl-tick, rl-label, rl-marker.
 */
export function rulerSvg({ orientation, length, thickness = 22, ticks, sheet, bands = [] }) {
  const L = Math.max(0, Math.round(length));
  const T = thickness;
  const h = orientation !== "v";
  const rect = (from, to, cls, inset = 0) => {
    const a = Math.max(0, Math.min(L, from));
    const b = Math.max(0, Math.min(L, to));
    if (b <= a) return "";
    return h
      ? `<rect class="${cls}" x="${a}" y="${inset}" width="${b - a}" height="${T - inset * 2}"/>`
      : `<rect class="${cls}" x="${inset}" y="${a}" width="${T - inset * 2}" height="${b - a}"/>`;
  };
  const sizes = [T * 0.42, T * 0.28, T * 0.16];
  let out = `<svg xmlns="http://www.w3.org/2000/svg" class="rl rl-${h ? "h" : "v"}" width="${h ? L : T}" height="${h ? T : L}" viewBox="0 0 ${h ? L : T} ${h ? T : L}" aria-hidden="true" focusable="false">`;
  out += rect(0, L, "rl-bg");
  if (sheet) out += rect(sheet.from, sheet.to, "rl-sheet", 3);
  for (const b of bands) out += rect(b.from, b.to, "rl-band", 3);
  let path = "";
  let labels = "";
  for (const t of ticks || []) {
    if (t.pos < 0 || t.pos > L) continue;
    const p = Math.round(t.pos) + 0.5;
    const s = sizes[t.level] ?? sizes[2];
    path += h ? `M${p} ${T}V${T - s}` : `M${T} ${p}H${T - s}`;
    if (t.label != null && t.pos > 2) {
      labels += h
        ? `<text class="rl-label" x="${p + 2.5}" y="${T * 0.48}">${t.label}</text>`
        : `<text class="rl-label" transform="translate(${T * 0.48} ${p + 2.5}) rotate(-90)" text-anchor="end">${t.label}</text>`;
    }
  }
  out += `<path class="rl-tick" d="${path}"/>${labels}`;
  // Margin markers: small triangles at each edge of each printable band.
  let markers = "";
  for (const b of bands) {
    for (const at of [b.from, b.to]) {
      if (at < 0 || at > L) continue;
      const p = Math.round(at) + 0.5;
      markers += h ? `M${p - 4} 0H${p + 4}L${p} 6Z` : `M0 ${p - 4}V${p + 4}L6 ${p}Z`;
    }
  }
  if (markers) out += `<path class="rl-marker" d="${markers}"/>`;
  return out + "</svg>";
}

// ---------------------------------------------------------------------------
// Pages, labels, and status
// ---------------------------------------------------------------------------

/**
 * The index of the sheet that fills most of the visible pasteboard. rects are the sheets'
 * { top, bottom } and the view is { top, bottom }, all in the same coordinates.
 */
export function currentSheetIndex(rects, view) {
  let best = 0;
  let bestSeen = -Infinity;
  rects.forEach((r, i) => {
    const seen = Math.min(r.bottom, view.bottom) - Math.max(r.top, view.top);
    if (seen > bestSeen + 0.5) {
      best = i;
      bestSeen = seen;
    }
  });
  return rects.length ? best : -1;
}

/** "Side" for folded pieces and landscape sheets, "Spread" for printer spreads, "Page" otherwise. */
export function sheetWord(model, spreads) {
  if (spreads) return "Spread";
  if (!model) return "Page";
  return model.folds || model.page.width > model.page.height ? "Side" : "Page";
}

/** Plain words for the overflow state: "Text fits", "Text overflows on side 2". */
export function overflowText(overflow, word = "page") {
  const over = [];
  (overflow || []).forEach((o, i) => o && over.push(i + 1));
  if (!over.length) return "Text fits";
  const w = word.toLowerCase();
  if (over.length === 1) return `Text overflows on ${w} ${over[0]}`;
  const list = over.slice(0, -1).join(", ") + " and " + over[over.length - 1];
  return `Text overflows on ${w}s ${list}`;
}

/** A file name from a document title: "Spring Newsletter!" becomes "spring-newsletter". */
export function fileSlug(title) {
  const s = String(title || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
  return s || "offizin-publication";
}

export const UNTITLED = "Untitled publication";

/**
 * The document title after a draw. `prev` is { lastTemplate, titleEdited, title } and `m` the
 * model on the page (null when it is empty). A different publication gets its own title, like
 * a new document (`reset` is true). The same publication keeps the title the person typed,
 * through an empty page and back (Undo, then Redo), and an untyped title follows the
 * publication's name.
 */
export function titleAfterRender(prev, m) {
  const next = { lastTemplate: prev.lastTemplate, titleEdited: prev.titleEdited, title: prev.title, reset: false };
  if (m && m.templateId !== prev.lastTemplate) {
    next.lastTemplate = m.templateId;
    next.titleEdited = false;
    next.reset = true;
  }
  if (!next.titleEdited) next.title = m ? m.name : UNTITLED;
  return next;
}

/** The studio fonts (IBM Plex Sans and Mono) added to the template fonts' Google Fonts link. */
export const STUDIO_FONT_FAMILIES = ["IBM Plex Sans:ital,wght@0,400;0,500;0,600;0,700;1,400", "IBM Plex Mono:wght@400;500"];

export function studioFontsHref(base) {
  const extra = STUDIO_FONT_FAMILIES.map((f) => "family=" + f.replace(/ /g, "+")).join("&");
  if (!base) return "https://fonts.googleapis.com/css2?" + extra + "&display=swap";
  const i = base.indexOf("&display=");
  return i < 0 ? base + "&" + extra : base.slice(0, i) + "&" + extra + base.slice(i);
}

/**
 * Crop marks and registration targets around a sheet, as an SVG string. The SVG is
 * `w + 2 * bleed` by `h + 2 * bleed` pixels and is placed `bleed` pixels outside the sheet.
 * Crop marks sit at the four corners, offset from the trim so they never print on the piece;
 * registration targets sit at the middle of each side.
 */
export function marksSvg(w, h, bleed = 30) {
  const W = Math.round(w + bleed * 2);
  const H = Math.round(h + bleed * 2);
  const x0 = bleed;
  const y0 = bleed;
  const x1 = bleed + w;
  const y1 = bleed + h;
  const off = 6;
  const len = Math.max(8, bleed - off - 4);
  const r = (n) => Math.round(n * 10) / 10;
  let crop = "";
  for (const [x, sx] of [[x0, -1], [x1, 1]]) {
    for (const [y, sy] of [[y0, -1], [y1, 1]]) {
      crop += `M${r(x + sx * off)} ${r(y)}h${r(sx * len)}M${r(x)} ${r(y + sy * off)}v${r(sy * len)}`;
    }
  }
  const target = (cx, cy) => {
    const R = Math.min(7, bleed / 3.6);
    return `<g class="mk-reg"><circle cx="${r(cx)}" cy="${r(cy)}" r="${r(R)}"/><circle cx="${r(cx)}" cy="${r(cy)}" r="${r(R / 2.2)}"/><path d="M${r(cx - R * 1.6)} ${r(cy)}h${r(R * 3.2)}M${r(cx)} ${r(cy - R * 1.6)}v${r(R * 3.2)}"/></g>`;
  };
  const mid = bleed / 2;
  const targets = target(x0 + w / 2, mid) + target(x0 + w / 2, H - mid) + target(mid, y0 + h / 2) + target(W - mid, y0 + h / 2);
  return `<svg xmlns="http://www.w3.org/2000/svg" class="mk" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" aria-hidden="true" focusable="false"><path class="mk-crop" d="${crop}"/>${targets}</svg>`;
}
