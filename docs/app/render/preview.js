/**
 * Draws a composed model as HTML pages: the live demo, gallery thumbnails, and the larger
 * preview inside the Word and Google Docs panels all use this renderer.
 *
 * Pages are laid out in points (CSS supports the pt unit) and scaled to fit. Margin and fold
 * guides appear on screen only, like the guides in a desktop publishing program.
 *
 * Options for renderPreview(container, model, options):
 *   width or scale    fit each page to this many pixels, or scale it directly
 *   guides            margin and fold lines (default true)
 *   pages             how many pages (or spreads) to draw
 *   gap, labels       pixels between pages, and "Side 1" or "Page 1" labels
 *   checkOverflow     measure text that runs past the page (default true); badges shows it
 *   spreads           booklets only: draw printer spreads (two pages per sheet, in print order)
 *   printPreview      hide guides and badges, and shade the edge most printers cannot reach
 *
 * Each sheet is an element with class "ozp-sheet" and data-page="<index>", after a label
 * (class "ozp-label") when labels are on. Guides use "ozp-margin", "ozp-fold", and "ozp-fold-h".
 */
import { escapeHtml } from "../core/util.js";

const PX_PER_PT = 96 / 72;

/** Single line spacing in Word and Docs is a little more than the font size. */
const LINE_FACTOR = 1.17;

/** Most home printers cannot print the outer quarter inch of the sheet. */
const UNPRINTABLE_PT = 18;

/** A simple picture pictogram (mountains and sun) used inside picture placeholders. */
export function pictogramSvg(color, size = 28) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 36" width="${size}" height="${Math.round(size * 0.75)}" aria-hidden="true"><rect x="1.5" y="1.5" width="45" height="33" rx="3" fill="none" stroke="${color}" stroke-width="2.5"/><circle cx="34" cy="11" r="4" fill="${color}"/><path d="M4 31 L17 16 L26 25 L31 20 L44 31 Z" fill="${color}"/></svg>`;
}

/** Splits a model's blocks into pages at each page break. */
export function splitPages(blocks) {
  const pages = [[]];
  for (const b of blocks) {
    if (b.t === "break") pages.push([]);
    else pages[pages.length - 1].push(b);
  }
  return pages;
}

function runHtml(r) {
  const text = escapeHtml(r.text).replace(/\n/g, "<br>");
  return r.bold ? `<strong>${text}</strong>` : text;
}

function paragraphHtml(b) {
  const style = [
    `font-family:${b.css}`,
    `font-size:${b.size}pt`,
    `color:${b.color}`,
    `font-weight:${b.bold ? 700 : 400}`,
    `font-style:${b.italic ? "italic" : "normal"}`,
    `text-align:${b.align}`,
    `margin:${b.before}pt 0 ${b.after}pt 0`,
    `line-height:${(b.line * LINE_FACTOR).toFixed(3)}`,
  ];
  if (b.tracking) style.push(`letter-spacing:${b.tracking}em`);
  let inner = b.runs.map(runHtml).join("");
  if (!inner.trim()) inner = "&nbsp;";
  if (b.bullet) {
    style.push(`padding-left:${b.bullet.indent}pt`, `text-indent:-${b.bullet.indent}pt`);
    inner = `<span class="ozp-bullet" style="color:${b.bullet.color};display:inline-block;width:${b.bullet.indent}pt;text-indent:0">${escapeHtml(b.bullet.char)}</span>${inner}`;
  } else if (b.indent) {
    style.push(`padding-left:${b.indent.left}pt`, `text-indent:${-b.indent.hanging}pt`);
  }
  const plain = escapeHtml(b.runs.map((r) => r.text).join(""));
  return `<p class="ozp-p" data-style="${escapeHtml(b.style || "body")}" data-text="${plain}" style="${style.join(";")}">${inner}</p>`;
}

function alignMargin(align) {
  if (align === "center") return "margin-left:auto;margin-right:auto";
  if (align === "right") return "margin-left:auto;margin-right:0";
  return "margin-left:0;margin-right:auto";
}

function imageHtml(b) {
  const box = `width:${b.width}pt;height:${b.height}pt;${alignMargin(b.align)};margin-top:${b.before}pt;margin-bottom:${b.after}pt`;
  if (b.src) return `<div class="ozp-img" style="${box}"><img src="${escapeHtml(b.src)}" alt="${escapeHtml(b.alt)}" style="width:100%;height:100%;object-fit:contain;display:block"></div>`;
  const icon = Math.max(14, Math.min(48, Math.min(b.width, b.height) * 0.28));
  return `<div class="ozp-img ozp-placeholder" data-placeholder="1" style="${box};background:${b.fill};display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4pt;overflow:hidden">${pictogramSvg(b.ink, icon)}${b.label && b.height > 40 ? `<span style="font:600 ${Math.max(6, Math.min(10, b.height / 14))}pt system-ui,sans-serif;color:${b.ink};opacity:.85">${escapeHtml(b.label)}</span>` : ""}</div>`;
}

function borderCss(side, s) {
  if (!s) return "";
  return `border-${side}:${Math.max(0.25, s.weight)}pt ${s.style === "dashed" ? "dashed" : "solid"} ${s.color}`;
}

function gridHtml(b) {
  const cols = b.cols.map((w) => `<col style="width:${w}pt">`).join("");
  const rows = b.rows
    .map((r) => {
      const cells = r.cells
        .map((c) => {
          const style = [`padding:${c.pad.map((v) => v + "pt").join(" ")}`, `vertical-align:${c.valign === "middle" ? "middle" : c.valign}`];
          if (c.fill) style.push(`background:${c.fill}`);
          if (c.border) for (const side of ["top", "right", "bottom", "left"]) if (c.border[side]) style.push(borderCss(side, c.border[side]));
          let inner = blocksHtml(c.blocks);
          if (r.exact && r.height) {
            const h = r.height - c.pad[0] - c.pad[2];
            inner = `<div style="max-height:${h}pt;overflow:hidden">${inner}</div>`;
          }
          return `<td style="${style.join(";")}">${inner}</td>`;
        })
        .join("");
      return `<tr${r.height ? ` style="height:${r.height}pt"` : ""}>${cells}</tr>`;
    })
    .join("");
  return `<table class="ozp-grid" style="width:${b.width}pt;${alignMargin(b.align)}"><colgroup>${cols}</colgroup><tbody>${rows}</tbody></table>`;
}

/** Flowing columns: the browser balances them, the way Word flows text between columns. */
function columnsHtml(b) {
  const rule = b.rule ? `;column-rule:${b.rule.weight}pt solid ${b.rule.color}` : "";
  return `<div class="ozp-columns" style="width:${b.width}pt;column-count:${b.count};column-gap:${b.gap}pt${rule}">${blocksHtml(b.blocks)}</div>`;
}

export function blocksHtml(blocks) {
  return blocks
    .map((b) => {
      switch (b.t) {
        case "p":
          return paragraphHtml(b);
        case "rule":
          return `<div class="ozp-rule" style="width:${b.width}pt;${alignMargin(b.align)};border-top:${b.weight}pt solid ${b.color};margin-top:${b.before}pt;margin-bottom:${b.after}pt"></div>`;
        case "spacer":
          return `<div class="ozp-spacer" style="height:${b.height}pt"></div>`;
        case "image":
          return imageHtml(b);
        case "grid":
          return gridHtml(b);
        case "columns":
          return columnsHtml(b);
        default:
          return "";
      }
    })
    .join("");
}

/** The footer for page `n` of `total`, with {page} and {pages} filled in. */
function footerHtml(model, page, n, total) {
  const f = model.footer;
  if (!f) return "";
  const text = f.runs
    .map((r) => runHtml({ ...r, text: r.text.replace(/\{page\}/g, String(n)).replace(/\{pages\}/g, String(total)) }))
    .join("");
  const m = page.margins;
  const bottom = Math.max(6, m.bottom / 2 - f.size * 0.6);
  return (
    `<div class="ozp-footer" style="position:absolute;left:${m.left}pt;right:${m.right}pt;bottom:${bottom.toFixed(1)}pt;` +
    `font-family:${f.css};font-size:${f.size}pt;color:${f.color};text-align:${f.align};line-height:1.2;white-space:nowrap;overflow:hidden">${text}</div>`
  );
}

const PREVIEW_CSS = `
.ozp-doc{display:flex;flex-direction:column;align-items:center}
.ozp-sheet{position:relative;flex:none;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.18),0 6px 18px rgba(0,0,0,.08)}
.ozp-page{position:absolute;top:0;left:0;transform-origin:0 0;background:#fff;box-sizing:border-box;overflow:hidden}
.ozp-half{position:absolute;top:0;box-sizing:border-box;overflow:hidden}
.ozp-content{position:relative;height:100%;overflow:hidden}
.ozp-content p{margin:0}
.ozp-grid{border-collapse:collapse;table-layout:fixed}
.ozp-grid td{box-sizing:border-box;overflow:hidden}
.ozp-columns .ozp-grid,.ozp-columns .ozp-img{break-inside:avoid}
.ozp-guides{position:absolute;inset:0;pointer-events:none}
.ozp-margin{position:absolute;outline:1px dotted rgba(70,140,230,.55)}
.ozp-fold{position:absolute;top:0;bottom:0;border-left:1px dashed rgba(214,64,140,.7)}
.ozp-fold-h{position:absolute;left:0;right:0;border-top:1px dashed rgba(214,64,140,.7)}
.ozp-unprintable{position:absolute;inset:0;pointer-events:none;border-style:solid;border-color:rgba(28,43,58,.06);box-sizing:border-box}
.ozp-overflow-badge{position:absolute;right:6px;bottom:6px;background:#B3261E;color:#fff;font:600 11px system-ui,sans-serif;padding:3px 7px;border-radius:4px}
.ozp-hit{outline:2px solid #E4A11B;outline-offset:2px;background:rgba(228,161,27,.12)}
.ozp-label{font:500 11px system-ui,sans-serif;color:#6b6b6b;margin:6px 0 2px;text-align:center}
`;

let cssInjected = false;
export function ensurePreviewCss(doc = document) {
  if (cssInjected && doc === document) return;
  const style = doc.createElement("style");
  style.textContent = PREVIEW_CSS;
  doc.head.appendChild(style);
  if (doc === document) cssInjected = true;
}

function guidesHtml(page, model, show) {
  if (!show) return "";
  const m = page.margins;
  let html = `<div class="ozp-margin" style="left:${m.left}pt;top:${m.top}pt;right:${m.right}pt;bottom:${m.bottom}pt"></div>`;
  if (model.folds) {
    for (const f of model.folds.vertical || []) html += `<div class="ozp-fold" style="left:${(f * 100).toFixed(4)}%"></div>`;
    for (const f of model.folds.horizontal || []) html += `<div class="ozp-fold-h" style="top:${(f * 100).toFixed(4)}%"></div>`;
  }
  return html;
}

/** One page: content inside the margins, the footer, and the guides. */
function pageHtml(model, page, blocks, n, total, opts) {
  return (
    `<div class="ozp-content">${blocksHtml(blocks)}</div>` +
    footerHtml(model, page, n, total) +
    `<div class="ozp-guides">${guidesHtml(page, model, opts.guides !== false && !opts.printPreview)}</div>`
  );
}

/** The sheets to draw: pages in reading order, or booklet printer spreads. */
function sheetsFor(model, opts) {
  const page = model.page;
  if (opts.spreads && model.booklet && model.pages) {
    const sheet = model.booklet.sheet;
    return model.booklet.spreads.map(([l, r], i) => ({
      width: sheet.width,
      height: sheet.height,
      label: `Sheet ${Math.floor(i / 2) + 1}, ${i % 2 ? "back" : "front"}`,
      halves: [
        { index: l, blocks: model.pages[l] },
        { index: r, blocks: model.pages[r] },
      ],
    }));
  }
  const pages = model.pages || splitPages(model.blocks);
  const word = model.folds || page.width > page.height ? "Side" : "Page";
  return pages.map((blocks, i) => ({ width: page.width, height: page.height, label: `${word} ${i + 1}`, blocks, index: i }));
}

/**
 * Renders a model into `container`. See the options at the top of this file.
 * Returns { scale, overflow: [per sheet], sheets: count }.
 */
export function renderPreview(container, model, opts = {}) {
  ensurePreviewCss(container.ownerDocument);
  const { page } = model;
  const sheets = sheetsFor(model, opts).slice(0, opts.pages || Infinity);
  const first = sheets[0] || { width: page.width, height: page.height };
  const scale = opts.width ? opts.width / (first.width * PX_PER_PT) : opts.scale ?? 1;
  const gap = opts.gap ?? 18;
  const total = model.pages ? model.pages.length : splitPages(model.blocks).length;
  const pad = `${page.margins.top}pt ${page.margins.right}pt ${page.margins.bottom}pt ${page.margins.left}pt`;
  const unprintable = opts.printPreview ? `<div class="ozp-unprintable" style="border-width:${UNPRINTABLE_PT}pt"></div>` : "";
  let html = `<div class="ozp-doc${opts.printPreview ? " ozp-print" : ""}" style="gap:${gap}px">`;
  sheets.forEach((s, i) => {
    const label = opts.labels && (sheets.length > 1 || s.halves) ? `<div class="ozp-label">${s.label}</div>` : "";
    let inner;
    if (s.halves) {
      inner = s.halves
        .map((h, k) => `<div class="ozp-half" style="left:${k ? page.width : 0}pt;width:${page.width}pt;height:${page.height}pt;padding:${pad}">${pageHtml(model, page, h.blocks || [], h.index + 1, total, opts)}</div>`)
        .join("");
      if (opts.guides !== false && !opts.printPreview) inner += `<div class="ozp-guides"><div class="ozp-fold" style="left:50%"></div></div>`;
    } else {
      inner = pageHtml(model, page, s.blocks, s.index + 1, total, opts);
    }
    const pagePad = s.halves ? "0" : pad;
    html +=
      `<div>${label}<div class="ozp-sheet${s.halves ? " ozp-spread" : ""}" data-page="${i}" style="width:${(s.width * PX_PER_PT * scale).toFixed(2)}px;height:${(s.height * PX_PER_PT * scale).toFixed(2)}px">` +
      `<div class="ozp-page" style="width:${s.width}pt;height:${s.height}pt;padding:${pagePad};transform:scale(${scale.toFixed(5)})">${inner}${unprintable}</div></div></div>`;
  });
  html += "</div>";
  container.innerHTML = html;
  const overflow = [];
  if (opts.checkOverflow !== false) {
    container.querySelectorAll(".ozp-sheet").forEach((sheet) => {
      const over = [...sheet.querySelectorAll(".ozp-content")].some((content) => content.scrollHeight > content.clientHeight + 2);
      overflow.push(over);
      if (over && opts.badges !== false && !opts.printPreview) sheet.insertAdjacentHTML("beforeend", `<div class="ozp-overflow-badge">Text overflows this page</div>`);
    });
  }
  return { scale, overflow, sheets: sheets.length };
}

/** Highlights the first paragraph whose text starts with `text` and scrolls it into view. */
export function highlightText(container, text) {
  container.querySelectorAll(".ozp-hit").forEach((el) => el.classList.remove("ozp-hit"));
  const needle = text.replace(/\s+/g, " ").trim().toLowerCase();
  for (const el of container.querySelectorAll(".ozp-p, .ozp-placeholder")) {
    const hay = (el.getAttribute("data-text") || (el.dataset.placeholder ? "picture" : "")).replace(/\s+/g, " ").toLowerCase();
    if (hay.includes(needle)) {
      el.classList.add("ozp-hit");
      if (el.scrollIntoView) el.scrollIntoView({ block: "center", behavior: "smooth" });
      return true;
    }
  }
  return false;
}
