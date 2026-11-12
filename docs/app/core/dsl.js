/**
 * Building blocks for writing Offizin templates.
 *
 * A template's build(ctx) returns an array of these blocks (or, for booklets, { pages: [...] }).
 * The composer (compose.js) resolves roles, color tokens, {fields}, spacing, and widths.
 *
 *   p(text, role, options)        a paragraph. Roles: display, title, headline, h1, h2, kicker,
 *                                 lead, body, small, caption, quote, label, button.
 *                                 Options: align, color, size, bold, italic, caps, before, after,
 *                                 line, fixed (not sample text), indent: { left, hanging }.
 *   list(items, options)          bulleted paragraphs (native bullets in Word and Google Docs)
 *   rule(options)                 a horizontal line: color, weight, width, align, before, after
 *   space(height)                 vertical space
 *   pageBreak()                   start a new page
 *   pic(w, h, options)            a picture placeholder ("full" width allowed)
 *   logoBox(w, h, options)        the user's logo, or a placeholder
 *   band(blocks, options)         a full-width colored box: fill, pad, width, height, align
 *   grid(widths, rows, options)   a layout table. Rows are arrays of cells or { height, exact, cells }.
 *                                 Options: gap, abs (widths in points), fixedGap (keep gap exact),
 *                                 width, align, border.
 *   cell(blocks, options)         a grid cell: fill, pad, valign, border
 *   table(header, rows, options)  a data table with a colored header row
 *   columns(count, blocks, opts)  text that flows from column to column: gap, rule
 *                                 (native columns in Word; balanced columns elsewhere)
 *   scope(vars, blocks)           blocks with extra {fields}, used for mail merge records
 */
import { round, mm } from "./util.js";

export const p = (text, role = "body", o = {}) => ({ t: "p", text, role, ...o });
export const list = (items, o = {}) => ({ t: "list", items, role: "body", ...o });
export const rule = (o = {}) => ({ t: "rule", ...o });
export const space = (height) => ({ t: "spacer", height });
export const pageBreak = () => ({ t: "break" });
export const pic = (width, height, o = {}) => ({ t: "image", kind: "photo", width, height, ...o });
export const logoBox = (width, height, o = {}) => ({ t: "image", kind: "logo", width, height, ...o });
export const band = (blocks, o = {}) => ({ t: "band", blocks, ...o });
export const grid = (widths, rows, o = {}) => ({ t: "grid", widths, rows, ...o });
export const cell = (blocks, o = {}) => ({ blocks, ...o });
export const table = (header, rows, o = {}) => ({ t: "table", header, rows, ...o });
export const columns = (count, blocks, o = {}) => ({ t: "columns", count, blocks, ...o });
export const scope = (vars, blocks) => ({ t: "scope", vars, blocks });

/** A label with the value under it, used for dates, times, and places. */
export const detail = (label, value, o = {}) => [
  p(label, "label", { align: o.align || "center", color: o.labelColor || "accent", fixed: true }),
  p(value, "h2", { align: o.align || "center", after: 0, color: o.color, size: o.size }),
];

/** Contact lines from business information. */
export const contactLines = (o = {}) => [
  p("{address}\n{city}", "small", { align: o.align, color: o.color, after: 4 }),
  p("{phone}\n{email}\n{web}", "small", { align: o.align, color: o.color, after: 0 }),
];

/** Logo when the user has one, otherwise the organization name set as a wordmark. */
export const brandMark = (ctx, w, h, o = {}) =>
  ctx.hasLogo
    ? logoBox(w, h, { align: o.align || "left", after: o.after ?? 8 })
    : p("{org}", "h2", { align: o.align || "left", color: o.color || "primary", after: o.after ?? 6, size: o.size });

/** Fold guide positions for a sheet folded into n equal panels. */
export const panelFolds = (n) => () => ({ vertical: Array.from({ length: n - 1 }, (_, i) => (i + 1) / n), horizontal: [] });

export const PANEL_PAD = 20;

/** Padding for panel i of n: the outer edges already have the page margin. */
export const panelPad = (i, n) => [PANEL_PAD, i === n - 1 ? 4 : PANEL_PAD, PANEL_PAD, i === 0 ? 4 : PANEL_PAD];

/** Page margins that center a grid of equal pieces (cards, labels, badges) on a sheet. */
export function sheetMargins(paper, pieceW, pieceH, across, down, gapX = 0, gapY = 0) {
  const page = paper === "a4" ? { width: mm(210), height: mm(297) } : { width: 612, height: 792 };
  const usedW = across * pieceW + (across - 1) * gapX;
  const usedH = down * pieceH + (down - 1) * gapY;
  const side = round((page.width - usedW) / 2);
  const top = round((page.height - usedH) / 2);
  return { top, bottom: top, left: side, right: side };
}

/** Light dashed cut guides for cards, badges, and tickets. */
export const CUT_GUIDE = { all: { color: "line", weight: 0.25, style: "dashed" } };

/** Splits merge records into groups of `size` (one group per printed sheet). */
export function chunk(records, size) {
  const out = [];
  for (let i = 0; i < records.length; i += size) out.push(records.slice(i, i + size));
  return out.length ? out : [[]];
}
