/**
 * Generates every Offizin logo, icon, and favicon from one font:
 * Cormorant Garamond Italic 500 (SIL Open Font License, see fonts/OFL.txt),
 * the same face Cachet uses, so the two products read as one family.
 *
 * The mark is a white sheet of paper with a pastel blue edge, a blue folded corner, and an
 * italic "O" in deep navy. Cachet's mark is a round seal; Offizin's is a printed page.
 *
 * Usage, from the repository root:
 *   npm install
 *   npm run brand
 *
 * Output:
 *   docs/assets/brand/  offizin-mark.svg, offizin-logo.svg/.png, offizin-logo-solid.png,
 *                       offizin-wordmark.svg/.png, offizin-wordmark-reverse.png, offizin-social.png
 *   docs/assets/icons/  icon-16 ... icon-512.png (Office ribbon, Google, home screens)
 *   docs/favicon.ico, docs/favicon.svg, docs/apple-touch-icon.png
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import opentype from "opentype.js";
import { Resvg } from "@resvg/resvg-js";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const FONT_FILE = path.join(ROOT, "brand/fonts/CormorantGaramond-500-italic.woff");
const BRAND_DIR = path.join(ROOT, "docs/assets/brand");
const ICON_DIR = path.join(ROOT, "docs/assets/icons");
const SITE_DIR = path.join(ROOT, "docs");

export const COLORS = {
  ink: "#1C2B3A", // deep navy for the "O" and the wordmark
  paper: "#FBFDFF", // the sheet
  edge: "#A7C7E7", // pastel blue edge, so the light sheet reads on light ribbons
  accent: "#6FA3D6", // the folded corner and icon accents
  dot: "#5B9BD5", // the dot of the final "i"
  wash: "#EAF3FB", // light blue backgrounds
};

const font = opentype.parse(fs.readFileSync(FONT_FILE).buffer);

/** Path data for one glyph drawn at the origin. Positioning is done with SVG transforms. */
function glyphData(glyph, size) {
  // A fresh lookup by index avoids an opentype.js quirk with glyphs returned by ligature substitution.
  return font.glyphs.get(glyph.index).getPath(0, 0, size);
}

/** Split a path into its closed contours, so the dot of an "i" can be colored on its own. */
function contours(p) {
  const out = [];
  let current = [];
  for (const cmd of p.commands) {
    // Each contour starts with a move; some fonts do not close every contour explicitly.
    if (cmd.type === "M" && current.length) {
      out.push(current);
      current = [];
    }
    current.push(cmd);
  }
  if (current.length) out.push(current);
  return out.map((commands) => {
    const c = new opentype.Path();
    c.commands = commands;
    return c;
  });
}

/**
 * The wordmark "Offizin": the font's own ffi ligature, with the dot of the final "i" in blue.
 * Returns SVG paths positioned on a baseline at y = 0, and the overall bounding box.
 */
function wordmarkPaths(size, letterColor, dotColor) {
  const glyphs = font.stringToGlyphs("Offizin", { features: { liga: true } });
  const scale = size / font.unitsPerEm;
  let x = 0;
  let box = { x1: Infinity, y1: Infinity, x2: -Infinity, y2: -Infinity };
  const parts = [];
  glyphs.forEach((g, index) => {
    const p = glyphData(g, size);
    const b = p.getBoundingBox();
    box = {
      x1: Math.min(box.x1, b.x1 + x),
      y1: Math.min(box.y1, b.y1),
      x2: Math.max(box.x2, b.x2 + x),
      y2: Math.max(box.y2, b.y2),
    };
    const isLastI = index === glyphs.length - 2 && g.name === "i";
    if (isLastI) {
      const pieces = contours(p).sort((a, c) => a.getBoundingBox().y1 - c.getBoundingBox().y1);
      // Enlarge the dot around its own center, as Cachet enlarges its period.
      const d = pieces[0].getBoundingBox();
      const cx = x + (d.x1 + d.x2) / 2;
      const cy = (d.y1 + d.y2) / 2;
      const grow = 1.4;
      parts.push(`<path transform="translate(${cx.toFixed(2)} ${cy.toFixed(2)}) scale(${grow}) translate(${(-(d.x1 + d.x2) / 2).toFixed(2)} ${(-cy).toFixed(2)})" d="${pieces[0].toPathData(2)}" fill="${dotColor}"/>`);
      pieces.slice(1).forEach((piece) => parts.push(`<path transform="translate(${x.toFixed(2)} 0)" d="${piece.toPathData(2)}" fill="${letterColor}"/>`));
    } else {
      parts.push(`<path transform="translate(${x.toFixed(2)} 0)" d="${p.toPathData(2)}" fill="${letterColor}"/>`);
    }
    x += g.advanceWidth * scale;
  });
  return { svg: parts.join(""), box };
}

/** The italic "O", centered in a square of side `boxSize`. `extra` thickens strokes for tiny icons. */
function monogram(boxSize, { height = 0.56, color = COLORS.ink, shiftX = 0, shiftY = 0, extra = 0 } = {}) {
  const p = font.getPath("O", 0, 0, 1000);
  const b = p.getBoundingBox();
  const s = (boxSize * height) / (b.y2 - b.y1);
  const tx = (boxSize - (b.x2 - b.x1) * s) / 2 - b.x1 * s + shiftX * boxSize;
  const ty = (boxSize - (b.y2 - b.y1) * s) / 2 - b.y1 * s + shiftY * boxSize;
  const stroke = extra ? ` stroke="${color}" stroke-width="${(extra / s).toFixed(1)}" stroke-linejoin="round"` : "";
  return `<path transform="translate(${tx.toFixed(2)} ${ty.toFixed(2)}) scale(${s.toFixed(5)})" d="${p.toPathData(1)}" fill="${color}"${stroke}/>`;
}

/**
 * The mark: a white sheet with a pastel blue edge, a blue folded corner, and the italic "O".
 * `pixels` is the final size, used to keep the edge at least one pixel wide and to thicken the "O".
 */
function markSvg(pixels = 512) {
  const B = 512;
  const fold = 128;
  const edge = Math.max(16, B / pixels);
  const r = 44;
  const extraBySize = { 16: 30, 20: 24, 24: 20, 32: 14, 40: 10, 48: 8, 64: 6, 80: 4 };
  const extra = extraBySize[pixels] || 0;
  const outer = `M${r} 0 H${B - fold} L${B} ${fold} V${B - r} Q${B} ${B} ${B - r} ${B} H${r} Q0 ${B} 0 ${B - r} V${r} Q0 0 ${r} 0 Z`;
  const e = edge;
  const inner = `M${r} ${e} H${B - fold - e * 0.41} L${B - e} ${fold + e * 0.41} V${B - r} Q${B - e} ${B - e} ${B - r} ${B - e} H${r} Q${e} ${B - e} ${e} ${B - r} V${r} Q${e} ${e} ${r} ${e} Z`;
  const flap = `M${B - fold} 0 V${fold - 30} Q${B - fold} ${fold} ${B - fold + 30} ${fold} H${B} Z`;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${B} ${B}" width="${pixels}" height="${pixels}">` +
    `<title>Offizin</title>` +
    `<path d="${outer}" fill="${COLORS.edge}"/>` +
    `<path d="${inner}" fill="${COLORS.paper}"/>` +
    `<path d="${flap}" fill="${COLORS.accent}"/>` +
    monogram(B, { height: 0.58, shiftX: -0.025, shiftY: 0.02, extra }) +
    `</svg>`;
}

/**
 * Ribbon button icons for Word: Publications, Design, Page Parts, Business Info, Check and Print.
 * Light shapes with navy outlines and a blue accent read on both light and dark ribbons.
 */
function ribbonIconSvg(name, pixels) {
  const sw = pixels <= 16 ? 2.4 : pixels <= 32 ? 1.8 : 1.4; // stroke width in a 32 unit box
  const ink = `stroke="${COLORS.ink}" stroke-width="${sw}" stroke-linejoin="round" stroke-linecap="round"`;
  const sheet = (x = 6, y = 3, w = 20, h = 26, f = 6) =>
    `<path d="M${x} ${y} H${x + w - f} L${x + w} ${y + f} V${y + h} H${x} Z" fill="${COLORS.paper}" ${ink}/>` +
    `<path d="M${x + w - f} ${y} V${y + f} H${x + w} Z" fill="${COLORS.accent}" ${ink}/>`;
  const shapes = {
    create: sheet() + `<path d="M16 11 V23 M10 17 H22" fill="none" ${ink}/>`,
    design:
      `<circle cx="12" cy="13" r="8" fill="${COLORS.paper}" ${ink}/>` +
      `<circle cx="20" cy="13" r="8" fill="${COLORS.accent}" fill-opacity="0.85" ${ink}/>` +
      `<circle cx="16" cy="20.5" r="8" fill="${COLORS.edge}" fill-opacity="0.9" ${ink}/>`,
    insert:
      sheet() +
      `<rect x="9.5" y="11" width="13" height="7" fill="${COLORS.edge}" ${ink}/>` +
      `<path d="M9.5 22 H22.5 M9.5 25.5 H18" fill="none" ${ink}/>`,
    info:
      `<rect x="3" y="7" width="26" height="18" rx="2" fill="${COLORS.paper}" ${ink}/>` +
      `<circle cx="10.5" cy="14" r="3" fill="${COLORS.accent}" ${ink}/>` +
      `<path d="M6.5 21 Q10.5 16.5 14.5 21 M18 12.5 H25 M18 16.5 H25 M18 20.5 H22.5" fill="none" ${ink}/>`,
    check: sheet() + `<path d="M10.5 17.5 L14.5 21.5 L22 12.5" fill="none" stroke="${COLORS.dot}" stroke-width="${sw * 1.6}" stroke-linecap="round" stroke-linejoin="round"/>`,
  };
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="${pixels}" height="${pixels}">${shapes[name]}</svg>`;
}

function wordmarkSvg(letterColor, dotColor, pad = 0.18) {
  const size = 400;
  const { svg, box } = wordmarkPaths(size, letterColor, dotColor);
  const p = size * pad;
  const w = box.x2 - box.x1 + 2 * p;
  const h = box.y2 - box.y1 + 2 * p;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${(box.x1 - p).toFixed(1)} ${(box.y1 - p).toFixed(1)} ${w.toFixed(1)} ${h.toFixed(1)}" width="${Math.round(w)}" height="${Math.round(h)}"><title>Offizin</title>${svg}</svg>`;
}

/** Primary logo, in the spirit of Cachet's: the mark above the wordmark, centered on a square. */
function logoSvg(background) {
  const S = 1000;
  const markSize = 380;
  const { svg, box } = wordmarkPaths(250, COLORS.ink, COLORS.dot);
  const wmWidth = box.x2 - box.x1;
  const wmX = (S - wmWidth) / 2 - box.x1;
  const wmBaseline = 800;
  const mark = markSvg(512).replace(/^<svg[^>]*>/, "").replace(/<\/svg>$/, "").replace(/<title>.*?<\/title>/, "");
  const bg = background ? `<rect width="${S}" height="${S}" fill="${background}"/>` : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${S} ${S}" width="${S}" height="${S}"><title>Offizin</title>${bg}` +
    `<g transform="translate(${(S - markSize) / 2} 150) scale(${markSize / 512})">${mark}</g>` +
    `<g transform="translate(${wmX.toFixed(1)} ${wmBaseline})">${svg}</g></svg>`;
}

/** 1280 x 640 social preview for GitHub and link previews. */
function socialSvg() {
  const W = 1280;
  const H = 640;
  const { svg, box } = wordmarkPaths(210, COLORS.ink, COLORS.dot);
  const mark = markSvg(512).replace(/^<svg[^>]*>/, "").replace(/<\/svg>$/, "").replace(/<title>.*?<\/title>/, "");
  const markSize = 230;
  const gap = 56;
  const total = markSize + gap + (box.x2 - box.x1);
  const left = (W - total) / 2;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">` +
    `<rect width="${W}" height="${H}" fill="${COLORS.wash}"/>` +
    `<rect x="24" y="24" width="${W - 48}" height="${H - 48}" fill="none" stroke="${COLORS.edge}" stroke-width="2"/>` +
    `<g transform="translate(${left} ${(H - markSize) / 2 - 40}) scale(${markSize / 512})">${mark}</g>` +
    `<g transform="translate(${(left + markSize + gap - box.x1).toFixed(1)} ${H / 2 + 30})">${svg}</g>` +
    `</svg>`;
}

function png(svg, width) {
  return new Resvg(svg, { fitTo: { mode: "width", value: width }, font: { loadSystemFonts: false } }).render().asPng();
}

/** A .ico file holding PNG images (supported by every current browser). */
function ico(images) {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(images.length, 4);
  const entries = [];
  let offset = 6 + 16 * images.length;
  for (const { size, data } of images) {
    const e = Buffer.alloc(16);
    e.writeUInt8(size >= 256 ? 0 : size, 0);
    e.writeUInt8(size >= 256 ? 0 : size, 1);
    e.writeUInt8(0, 2);
    e.writeUInt8(0, 3);
    e.writeUInt16LE(1, 4);
    e.writeUInt16LE(32, 6);
    e.writeUInt32LE(data.length, 8);
    e.writeUInt32LE(offset, 12);
    offset += data.length;
    entries.push(e);
  }
  return Buffer.concat([header, ...entries, ...images.map((i) => i.data)]);
}

function write(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, data);
}

function main() {
  // Icons for the Word ribbon (16 to 80), Google (32, 128), and home screens (180 to 512).
  for (const size of [16, 20, 24, 32, 40, 48, 64, 80, 96, 128, 192, 256, 512]) {
    write(path.join(ICON_DIR, `icon-${size}.png`), png(markSvg(size), size));
  }
  for (const name of ["create", "design", "insert", "info", "check"]) {
    for (const size of [16, 20, 24, 32, 40, 48, 64, 80]) {
      write(path.join(ICON_DIR, `ribbon-${name}-${size}.png`), png(ribbonIconSvg(name, size), size));
    }
  }
  write(path.join(BRAND_DIR, "offizin-mark.svg"), markSvg(512));
  // The panel shows the mark inline, so it also works inside the Google Docs sidebar.
  write(
    path.join(ROOT, "docs/app/ui/mark.js"),
    `/** The Offizin mark as inline SVG. Generated by brand/make-brand-assets.mjs; do not edit by hand. */\nexport const MARK_SVG = ${JSON.stringify(markSvg(48).replace(/<title>.*?<\/title>/, ""))};\n`
  );
  write(path.join(SITE_DIR, "favicon.svg"), markSvg(32));
  write(path.join(SITE_DIR, "favicon.ico"), ico([16, 32, 48].map((size) => ({ size, data: png(markSvg(size), size) }))));

  // Phones show transparency as black, so the home screen icon sits on a light blue square.
  const apple = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 180 180" width="180" height="180"><rect width="180" height="180" fill="${COLORS.wash}"/><g transform="translate(18 18) scale(${144 / 512})">${markSvg(512).replace(/^<svg[^>]*>/, "").replace(/<\/svg>$/, "")}</g></svg>`;
  write(path.join(SITE_DIR, "apple-touch-icon.png"), png(apple, 180));

  write(path.join(BRAND_DIR, "offizin-wordmark.svg"), wordmarkSvg(COLORS.ink, COLORS.dot));
  write(path.join(BRAND_DIR, "offizin-wordmark.png"), png(wordmarkSvg(COLORS.ink, COLORS.dot), 1600));
  write(path.join(BRAND_DIR, "offizin-wordmark-reverse.png"), png(wordmarkSvg("#FFFFFF", COLORS.edge), 1600));

  write(path.join(BRAND_DIR, "offizin-logo.svg"), logoSvg(null));
  write(path.join(BRAND_DIR, "offizin-logo.png"), png(logoSvg(null), 1000));
  write(path.join(BRAND_DIR, "offizin-logo-solid.png"), png(logoSvg(COLORS.wash), 1000));
  write(path.join(BRAND_DIR, "offizin-social.png"), png(socialSvg(), 1280));

  console.log("Brand assets written to docs/assets.");
}

main();
