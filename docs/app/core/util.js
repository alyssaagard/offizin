/**
 * Offizin core utilities: units, paper sizes, and color math.
 * Shared by the web demo, the Word add-in, and the Google Docs sidebar.
 * This file has no browser or Office dependencies, so it also runs in Node for tests.
 */

/** Points are the unit everywhere in Offizin: 72 points = 1 inch. */
export const PT_PER_IN = 72;
export const inch = (n) => n * PT_PER_IN;
export const mm = (n) => (n * PT_PER_IN) / 25.4;

/** Portrait dimensions in points. */
export const PAPER = {
  letter: { width: 612, height: 792, label: "Letter (8.5 x 11 in)" },
  legal: { width: 612, height: 1008, label: "Legal (8.5 x 14 in)" },
  tabloid: { width: 792, height: 1224, label: "Tabloid (11 x 17 in)" },
  halfLetter: { width: 396, height: 612, label: "Half letter (5.5 x 8.5 in)" },
  a3: { width: mm(297), height: mm(420), label: "A3 (297 x 420 mm)" },
  a4: { width: mm(210), height: mm(297), label: "A4 (210 x 297 mm)" },
  a5: { width: mm(148), height: mm(210), label: "A5 (148 x 210 mm)" },
  card5x7: { width: inch(5), height: inch(7), label: "Card (5 x 7 in)" },
};

export function pageSize(name, orientation = "portrait") {
  const p = PAPER[name] || PAPER.letter;
  return orientation === "landscape" ? { width: p.height, height: p.width } : { width: p.width, height: p.height };
}

export function round(n, places = 2) {
  const f = 10 ** places;
  return Math.round(n * f) / f;
}

export function clamp(n, lo, hi) {
  return Math.min(hi, Math.max(lo, n));
}

// ---------------------------------------------------------------------------
// Color math
// ---------------------------------------------------------------------------

export function normalizeHex(hex) {
  let h = String(hex || "").trim().replace(/^#/, "");
  if (h.length === 3) h = h.split("").map((c) => c + c).join("");
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return "#000000";
  return "#" + h.toUpperCase();
}

export function hexToRgb(hex) {
  const h = normalizeHex(hex).slice(1);
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

export function rgbToHex([r, g, b]) {
  return "#" + [r, g, b].map((v) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** Mix color a toward color b by t (0 = a, 1 = b). */
export function mix(a, b, t) {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  return rgbToHex(x.map((v, i) => v + (y[i] - v) * t));
}

export function luminance(hex) {
  const lin = hexToRgb(hex).map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2];
}

export function contrastRatio(a, b) {
  const [hi, lo] = [luminance(a), luminance(b)].sort((m, n) => n - m);
  return (hi + 0.05) / (lo + 0.05);
}

/** White or near-black, whichever reads better on the given background. */
export function readableOn(bg, dark = "#1A1A1A") {
  return contrastRatio(bg, "#FFFFFF") >= contrastRatio(bg, dark) ? "#FFFFFF" : dark;
}

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------

export function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Removes characters that are not allowed in XML 1.0 documents. */
export function xmlSafe(s) {
  return String(s).replace(/[^\x09\x0A\x0D\x20-퟿-�\u{10000}-\u{10FFFF}]/gu, "");
}

export const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
