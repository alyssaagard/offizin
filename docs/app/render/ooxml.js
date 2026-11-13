/**
 * Writes a composed model as WordprocessingML (the XML inside a .docx file).
 *
 * Two outputs:
 *   docxBytes(model, images, opts)   a complete .docx. It uses Word's own features, so the
 *                                    publication behaves like any well made Word document:
 *                                    native styles (Title, Heading 1, Heading 2, Normal, Quote,
 *                                    Caption, and Offizin styles), native bullets, flowing section
 *                                    columns, a footer with page number fields, and the print
 *                                    settings Word keeps in the file (book fold, mirror margins,
 *                                    gutter, two pages per sheet, embedded fonts, print layout).
 *   flatPackage(blocks, width, ...)  a single-XML package for Word's insertOoxml, used to insert
 *                                    page parts at the cursor. Parts keep full direct formatting,
 *                                    because the destination document may define its styles
 *                                    differently, and they are wrapped in an invisible content
 *                                    control so Offizin's own Undo can find them again.
 *
 * Units: twips (1/20 pt) for spacing and widths, half points for font sizes,
 * eighths of a point for borders, and EMUs (12,700 per point) for pictures.
 *
 * Element order follows the ECMA-376 schema sequences (CT_PPr, CT_RPr, CT_SectPr, CT_Style,
 * CT_Settings, CT_Lvl), which Word enforces when it opens a file.
 */
import { escapeXml, xmlSafe } from "../core/util.js";
import { wordSheet } from "../core/compose.js";
import { zip, base64ToBytes } from "./zip.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="${W_NS}" ` +
  `xmlns:r="${R_NS}" ` +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const tw = (pt) => Math.round(pt * 20);
const halfPt = (pt) => Math.round(pt * 2);
const emu = (pt) => Math.round(pt * 12700);
const hex = (c) => String(c || "000000").replace("#", "").toUpperCase();
const JC = { left: "left", center: "center", right: "right", justify: "both" };
const near = (a, b, eps = 0.01) => Math.abs((a || 0) - (b || 0)) < eps;

/** A paragraph that takes up almost no space: used after tables and for page breaks. */
const TINY_RPR = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>';
const TINY_SPACING = '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>';
const TINY_P = `<w:p><w:pPr>${TINY_SPACING}${TINY_RPR}</w:pPr></w:p>`;
const PAGE_BREAK_P = `<w:p><w:pPr>${TINY_SPACING}${TINY_RPR}</w:pPr><w:r>${TINY_RPR}<w:br w:type="page"/></w:r></w:p>`;

/** Which role defines each Word style, and the Word name of each style. */
const STYLE_SOURCES = [
  ["Normal", "body", "Normal"],
  ["Title", "title", "Title"],
  ["Heading1", "h1", "heading 1"],
  ["Heading2", "h2", "heading 2"],
  ["Quote", "quote", "Quote"],
  ["Caption", "caption", "caption"],
  ["OffizinDisplay", "display", "Offizin Display"],
  ["OffizinKicker", "kicker", "Offizin Kicker"],
  ["OffizinLead", "lead", "Offizin Lead"],
  ["OffizinSmall", "small", "Offizin Small"],
  ["OffizinLabel", "label", "Offizin Label"],
  ["OffizinButton", "button", "Offizin Button"],
];
const BUILT_IN = new Set(["Normal", "Title", "Heading1", "Heading2", "Quote", "Caption"]);
const OUTLINE = { Heading1: 0, Heading2: 1 };

// ---------------------------------------------------------------------------
// Writer state: pictures, bullets, and native styles
// ---------------------------------------------------------------------------

/** Collects pictures as relationships while the body is written. */
class Media {
  constructor(images) {
    this.images = images; // function(block) -> { base64, mime } | null
    this.items = [];
    this.nextDrawingId = 1;
  }
  add(block) {
    const img = this.images ? this.images(block) : null;
    if (!img) return null;
    const ext = /jpe?g/.test(img.mime) ? "jpeg" : "png";
    const n = this.items.length + 1;
    const item = { id: `rIdOzImg${n}`, name: `media/ozimage${n}.${ext}`, mime: ext === "jpeg" ? "image/jpeg" : "image/png", base64: img.base64 };
    this.items.push(item);
    return item;
  }
}

/** Native bullet lists: one numbering definition per distinct bullet look. */
class Numbering {
  constructor() {
    this.defs = new Map();
  }
  idFor(bullet, font) {
    const key = [bullet.char, hex(bullet.color), tw(bullet.indent), font].join("|");
    if (!this.defs.has(key)) this.defs.set(key, { id: this.defs.size + 1, bullet, font });
    return this.defs.get(key).id;
  }
  xml() {
    const defs = [...this.defs.values()];
    const abstracts = defs
      .map((d) => {
        const f = escapeXml(d.font);
        const ind = tw(d.bullet.indent);
        return (
          `<w:abstractNum w:abstractNumId="${d.id}"><w:multiLevelType w:val="singleLevel"/>` +
          `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${escapeXml(d.bullet.char)}"/><w:lvlJc w:val="left"/>` +
          `<w:pPr><w:ind w:left="${ind}" w:hanging="${ind}"/></w:pPr>` +
          `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/><w:color w:val="${hex(d.bullet.color)}"/></w:rPr></w:lvl></w:abstractNum>`
        );
      })
      .join("");
    const nums = defs.map((d) => `<w:num w:numId="${d.id}"><w:abstractNumId w:val="${d.id}"/></w:num>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W_NS}">${abstracts}${nums}</w:numbering>`;
  }
}

/**
 * Writes blocks. In "doc" mode paragraphs use native styles and carry only the formatting that
 * differs from their style. In "part" mode every paragraph carries full direct formatting.
 */
class Writer {
  constructor(media, opts = {}) {
    this.media = media;
    this.mode = opts.mode || "part";
    this.numbering = this.mode === "doc" ? new Numbering() : null;
    this.styles = {}; // Word style ID -> resolved role style
    if (this.mode === "doc" && opts.modelStyles) {
      for (const [id, role] of STYLE_SOURCES) if (opts.modelStyles[role]) this.styles[id] = opts.modelStyles[role];
    }
    this.modelStyles = opts.modelStyles || null;
  }

  styleIdFor(p) {
    const s = this.modelStyles && this.modelStyles[p.style];
    return s ? s.word : "Normal";
  }

  /** The style a paragraph is measured against, or null in part mode. */
  baseFor(p) {
    if (this.mode !== "doc") return null;
    return this.styles[this.styleIdFor(p)] || null;
  }

  runProps(p, run, base, colorOverride) {
    const parts = [];
    const font = escapeXml(p.font);
    const bold = run.bold ?? p.bold;
    const color = colorOverride || run.color || p.color;
    if (!base || base.font !== p.font) parts.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`);
    if (!base || !!base.bold !== !!bold) parts.push(bold ? "<w:b/><w:bCs/>" : base ? '<w:b w:val="0"/><w:bCs w:val="0"/>' : "");
    if (!base || !!base.italic !== !!p.italic) parts.push(p.italic ? "<w:i/><w:iCs/>" : base ? '<w:i w:val="0"/><w:iCs w:val="0"/>' : "");
    if (base && !!base.caps !== !!p.caps) parts.push(p.caps ? "<w:caps/>" : '<w:caps w:val="0"/>');
    if (!base || hex(base.color) !== hex(color)) parts.push(`<w:color w:val="${hex(color)}"/>`);
    const trackA = Math.round((p.tracking || 0) * p.size * 20);
    const trackB = base ? Math.round((base.tracking || 0) * base.size * 20) : 0;
    if (base ? trackA !== trackB : trackA) parts.push(`<w:spacing w:val="${trackA}"/>`);
    if (!base || !near(base.size, p.size)) parts.push(`<w:sz w:val="${halfPt(p.size)}"/><w:szCs w:val="${halfPt(p.size)}"/>`);
    const inner = parts.join("");
    return inner ? `<w:rPr>${inner}</w:rPr>` : "";
  }

  textRuns(p, run, base) {
    const rPr = this.runProps(p, run, base);
    const lines = xmlSafe(run.text).split("\n");
    return lines
      .map((line, i) => {
        const segs = line.split("\t").map((s) => `<w:t xml:space="preserve">${escapeXml(s)}</w:t>`).join("<w:tab/>");
        return `<w:r>${rPr}${i > 0 ? "<w:br/>" : ""}${segs}</w:r>`;
      })
      .join("");
  }

  paragraph(p) {
    const base = this.baseFor(p);
    const styleId = this.styleIdFor(p);
    const ppr = [];
    // pStyle: always in documents; in parts only built-in styles, which every document has.
    if (this.mode === "doc" ? styleId !== "Normal" : BUILT_IN.has(styleId) && styleId !== "Normal") ppr.push(`<w:pStyle w:val="${styleId}"/>`);
    const styleKeepsNext = base && OUTLINE[styleId] !== undefined;
    if (p.keepNext && !styleKeepsNext) ppr.push("<w:keepNext/>");
    const useNumbering = p.bullet && this.numbering;
    if (useNumbering) ppr.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${this.numbering.idFor(p.bullet, p.font)}"/></w:numPr>`);
    if (p.bullet && !useNumbering) ppr.push(`<w:tabs><w:tab w:val="left" w:pos="${tw(p.bullet.indent)}"/></w:tabs>`);
    if (!base || !near(base.before, p.before) || !near(base.after, p.after) || !near(base.line, p.line, 0.001)) {
      ppr.push(`<w:spacing w:before="${tw(p.before)}" w:after="${tw(p.after)}" w:line="${Math.round(p.line * 240)}" w:lineRule="auto"/>`);
    }
    if (p.bullet) ppr.push(`<w:ind w:left="${tw(p.bullet.indent)}" w:hanging="${tw(p.bullet.indent)}"/>`);
    else if (p.indent) ppr.push(`<w:ind w:left="${tw(p.indent.left)}" w:hanging="${tw(p.indent.hanging)}"/>`);
    if (!base || p.align !== "left") ppr.push(`<w:jc w:val="${JC[p.align] || "left"}"/>`);
    if (!base || !near(base.size, p.size)) ppr.push(`<w:rPr><w:sz w:val="${halfPt(p.size)}"/><w:szCs w:val="${halfPt(p.size)}"/></w:rPr>`);
    let runs = "";
    if (p.bullet && !useNumbering) runs += `<w:r>${this.runProps(p, { bold: false }, null, p.bullet.color)}<w:t xml:space="preserve">${escapeXml(p.bullet.char)}</w:t><w:tab/></w:r>`;
    runs += p.runs.map((r) => this.textRuns(p, r, base)).join("");
    return `<w:p>${ppr.length ? `<w:pPr>${ppr.join("")}</w:pPr>` : ""}${runs}</w:p>`;
  }

  rule(b, container) {
    const extra = Math.max(0, container - b.width);
    const [l, r] = b.align === "center" ? [extra / 2, extra / 2] : b.align === "right" ? [extra, 0] : [0, extra];
    const sz = Math.max(2, Math.min(96, Math.round(b.weight * 8)));
    return (
      `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="${sz}" w:space="0" w:color="${hex(b.color)}"/></w:pBdr>` +
      `<w:spacing w:before="${tw(b.before)}" w:after="${tw(b.after)}" w:line="20" w:lineRule="exact"/>` +
      `<w:ind w:left="${tw(l)}" w:right="${tw(r)}"/>${TINY_RPR}</w:pPr></w:p>`
    );
  }

  spacer(b) {
    return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.max(20, tw(b.height))}" w:lineRule="exact"/>${TINY_RPR}</w:pPr></w:p>`;
  }

  image(b) {
    const item = this.media.add(b);
    const jc = `<w:jc w:val="${JC[b.align] || "left"}"/>`;
    const spacing = `<w:spacing w:before="${tw(b.before)}" w:after="${tw(b.after)}" w:line="240" w:lineRule="auto"/>`;
    if (!item) {
      // No picture data available (for example in tests): keep a labeled paragraph in its place.
      return `<w:p><w:pPr>${spacing}${jc}</w:pPr><w:r><w:t xml:space="preserve">${escapeXml(b.label || "Picture")}</w:t></w:r></w:p>`;
    }
    const id = this.media.nextDrawingId++;
    const cx = emu(b.width);
    const cy = emu(b.height);
    const descr = escapeXml(b.alt || "Picture");
    return (
      `<w:p><w:pPr>${spacing}${jc}</w:pPr><w:r><w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}" descr="${descr}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${item.name.split("/").pop()}" descr="${descr}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${item.id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    );
  }

  grid(b) {
    const borderSide = (name, s) => {
      if (!s) return `<w:${name} w:val="nil"/>`;
      const sz = Math.max(2, Math.min(96, Math.round(s.weight * 8)));
      return `<w:${name} w:val="${s.style === "dashed" ? "dashed" : "single"}" w:sz="${sz}" w:space="0" w:color="${hex(s.color)}"/>`;
    };
    const tblPr =
      `<w:tblPr><w:tblW w:w="${tw(b.width)}" w:type="dxa"/>` +
      (b.align === "center" ? '<w:jc w:val="center"/>' : b.align === "right" ? '<w:jc w:val="right"/>' : "") +
      `<w:tblInd w:w="0" w:type="dxa"/>` +
      `<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>` +
      `<w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>` +
      `<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>`;
    const gridCols = `<w:tblGrid>${b.cols.map((w) => `<w:gridCol w:w="${tw(w)}"/>`).join("")}</w:tblGrid>`;
    const rows = b.rows
      .map((r) => {
        const trPr = r.height ? `<w:trPr><w:trHeight w:val="${tw(r.height)}" w:hRule="${r.exact ? "exact" : "atLeast"}"/></w:trPr>` : "";
        const cells = r.cells
          .map((c, i) => {
            const w = b.cols[i];
            const inner = Math.max(12, w - c.pad[1] - c.pad[3]);
            const tcPr = [`<w:tcW w:w="${tw(w)}" w:type="dxa"/>`];
            if (c.border) tcPr.push(`<w:tcBorders>${borderSide("top", c.border.top)}${borderSide("left", c.border.left)}${borderSide("bottom", c.border.bottom)}${borderSide("right", c.border.right)}</w:tcBorders>`);
            if (c.fill) tcPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hex(c.fill)}"/>`);
            tcPr.push(`<w:tcMar><w:top w:w="${tw(c.pad[0])}" w:type="dxa"/><w:left w:w="${tw(c.pad[3])}" w:type="dxa"/><w:bottom w:w="${tw(c.pad[2])}" w:type="dxa"/><w:right w:w="${tw(c.pad[1])}" w:type="dxa"/></w:tcMar>`);
            tcPr.push(`<w:vAlign w:val="${c.valign === "middle" ? "center" : c.valign}"/>`);
            let content = this.blocks(c.blocks, inner, { inCell: true });
            if (!content || /<\/w:tbl>$/.test(content)) content += TINY_P;
            return `<w:tc><w:tcPr>${tcPr.join("")}</w:tcPr>${content}</w:tc>`;
          })
          .join("");
        return `<w:tr>${trPr}${cells}</w:tr>`;
      })
      .join("");
    return `<w:tbl>${tblPr}${gridCols}${rows}</w:tbl>`;
  }

  /** Columns drawn as a borderless table, split where the composer balanced them. */
  columnsAsTable(b) {
    const starts = [0, ...(b.split || []), b.blocks.length];
    const cols = [];
    const cells = [];
    for (let i = 0; i < b.count; i++) {
      if (i) {
        cols.push(b.gap);
        cells.push({ fill: null, pad: [0, 0, 0, 0], valign: "top", border: null, blocks: [] });
      }
      cols.push(b.colWidth);
      const from = starts[i] ?? b.blocks.length;
      const to = starts[i + 1] ?? b.blocks.length;
      cells.push({ fill: null, pad: [0, 0, 0, 0], valign: "top", border: null, blocks: b.blocks.slice(from, to) });
    }
    return this.grid({ t: "grid", width: b.width, cols, rows: [{ height: null, exact: false, cells }], align: "left" });
  }

  /** Body XML for blocks inside a container `width` points wide. */
  blocks(list, width, opts = {}) {
    let out = "";
    list.forEach((b, i) => {
      switch (b.t) {
        case "p":
          out += this.paragraph(b);
          break;
        case "rule":
          out += this.rule(b, width);
          break;
        case "spacer":
          out += this.spacer(b);
          break;
        case "image":
          out += this.image(b);
          break;
        case "break":
          out += PAGE_BREAK_P;
          break;
        case "grid":
        case "columns": {
          out += b.t === "grid" ? this.grid(b) : this.columnsAsTable(b);
          // Word joins tables that touch, and needs a paragraph between a table and a page break.
          const next = list[i + 1];
          if (!next || next.t === "grid" || next.t === "columns" || next.t === "break") {
            if (!(opts.inCell && !next)) out += TINY_P;
          }
          break;
        }
        default:
          break;
      }
    });
    return out;
  }
}

// ---------------------------------------------------------------------------
// Document parts
// ---------------------------------------------------------------------------

/** The page each section is laid out on. Book fold uses the whole sheet; Word halves it. */
function sectionPage(model, bookFold) {
  const sheet = wordSheet(model, bookFold);
  if (sheet.width === model.page.width && sheet.height === model.page.height) return model.page;
  return { width: sheet.width, height: sheet.height, margins: model.page.margins };
}

function sectPrXml(page, { cols = 1, gap = 36, sep = false, continuous = false, footer = false, gutter = 0 } = {}) {
  const w = tw(page.width);
  const h = tw(page.height);
  const m = page.margins;
  const colsXml = cols > 1 ? `<w:cols w:num="${cols}" w:space="${tw(gap)}"${sep ? ' w:sep="1"' : ""}/>` : '<w:cols w:space="720"/>';
  return (
    `<w:sectPr>${footer ? '<w:footerReference w:type="default" r:id="rIdOzFooter"/>' : ""}` +
    (continuous ? '<w:type w:val="continuous"/>' : "") +
    `<w:pgSz w:w="${w}" w:h="${h}"${page.width > page.height ? ' w:orient="landscape"' : ""}/>` +
    `<w:pgMar w:top="${tw(m.top)}" w:right="${tw(m.right)}" w:bottom="${tw(m.bottom)}" w:left="${tw(m.left)}" w:header="360" w:footer="360" w:gutter="${tw(gutter)}"/>` +
    `${colsXml}<w:docGrid w:linePitch="360"/></w:sectPr>`
  );
}

/** Body XML for a whole publication, with native section columns where the model flows text. */
function bodyXml(model, writer, bookFold) {
  const page = sectionPage(model, bookFold);
  const width = model.page.width - model.page.margins.left - model.page.margins.right;
  const common = { footer: !!model.footer, gutter: (model.print && model.print.gutter) || 0 };
  let out = "";
  let pending = [];
  let sectionCount = 0;
  const flush = () => {
    out += writer.blocks(pending, width);
    pending = [];
  };
  for (const b of model.blocks) {
    if (b.t !== "columns") {
      pending.push(b);
      continue;
    }
    // Close the single column section before the columns, then the columns section itself.
    flush();
    out += `<w:p><w:pPr>${TINY_SPACING}${TINY_RPR}${sectPrXml(page, { ...common, continuous: sectionCount > 0 })}</w:pPr></w:p>`;
    sectionCount++;
    out += writer.blocks(b.blocks, b.colWidth);
    out += `<w:p><w:pPr>${TINY_SPACING}${TINY_RPR}${sectPrXml(page, { ...common, cols: b.count, gap: b.gap, sep: !!b.rule, continuous: true })}</w:pPr></w:p>`;
    sectionCount++;
  }
  flush();
  // End with a tiny paragraph so Word never adds a full-size one that spills onto a new page.
  if (!out || /<\/w:tbl>$/.test(out) || /w:type="page"\/><\/w:r><\/w:p>$/.test(out)) out += TINY_P;
  return out + sectPrXml(page, { ...common, continuous: sectionCount > 0 });
}

/** Styles for a full document, generated from the model's role styles. */
function stylesXml(model) {
  const s = model.styles || {};
  const body = s.body || { font: (model.fonts && model.fonts.body && model.fonts.body.name) || "Arial", size: 10, color: "#1A1A1A", line: 1.3, after: 6 };
  const f = escapeXml(body.font);
  const rpr = (st) => {
    const font = escapeXml(st.font);
    return (
      `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` +
      (st.bold ? "<w:b/><w:bCs/>" : "") +
      (st.italic ? "<w:i/><w:iCs/>" : "") +
      (st.caps ? "<w:caps/>" : "") +
      `<w:color w:val="${hex(st.color)}"/>` +
      (st.tracking ? `<w:spacing w:val="${Math.round(st.tracking * st.size * 20)}"/>` : "") +
      `<w:sz w:val="${halfPt(st.size)}"/><w:szCs w:val="${halfPt(st.size)}"/></w:rPr>`
    );
  };
  const ppr = (st, id) =>
    `<w:pPr>${OUTLINE[id] !== undefined ? "<w:keepNext/><w:keepLines/>" : ""}` +
    `<w:spacing w:before="${tw(st.before || 0)}" w:after="${tw(st.after || 0)}" w:line="${Math.round((st.line || 1) * 240)}" w:lineRule="auto"/>` +
    (OUTLINE[id] !== undefined ? `<w:outlineLvl w:val="${OUTLINE[id]}"/>` : "") +
    `</w:pPr>`;
  const defs = STYLE_SOURCES.filter(([, role]) => s[role] || role === "body")
    .map(([id, role, name]) => {
      const st = s[role] || body;
      const isNormal = id === "Normal";
      const based = id === "OffizinDisplay" ? "Title" : "Normal";
      return (
        `<w:style w:type="paragraph"${isNormal ? ' w:default="1"' : ""} w:styleId="${id}"><w:name w:val="${name}"/>` +
        (isNormal ? "" : `<w:basedOn w:val="${based}"/><w:next w:val="Normal"/>`) +
        (OUTLINE[id] !== undefined ? '<w:uiPriority w:val="9"/>' : "") +
        "<w:qFormat/>" +
        ppr(st, id) +
        rpr(st) +
        `</w:style>`
      );
    })
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/><w:sz w:val="${halfPt(body.size)}"/><w:szCs w:val="${halfPt(body.size)}"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    defs +
    `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>` +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
    `</w:styles>`
  );
}

/**
 * settings.xml. The children follow the CT_Settings sequence in ECMA-376 Part 1 (view, zoom,
 * embedTrueTypeFonts, saveSubsetFonts, mirrorMargins, defaultTabStop, bookFoldPrinting,
 * characterSpacingControl, printTwoOnOne, compat).
 */
export function settingsXml(print = {}, bookFold = false) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W_NS}">` +
    `<w:view w:val="print"/><w:zoom w:val="fullPage" w:percent="100"/>` +
    (print.embedFonts ? "<w:embedTrueTypeFonts/><w:saveSubsetFonts/>" : "") +
    (print.mirrorMargins ? "<w:mirrorMargins/>" : "") +
    `<w:defaultTabStop w:val="720"/>` +
    (bookFold ? "<w:bookFoldPrinting/>" : "") +
    `<w:characterSpacingControl w:val="doNotCompress"/>` +
    (print.twoPerSheet && !bookFold ? "<w:printTwoOnOne/>" : "") +
    `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
  );
}

/** The footer part, with PAGE and NUMPAGES fields in place of {page} and {pages}. */
function footerXml(footer, writer) {
  const base = null;
  const rPr = writer.runProps(footer, footer.runs[0] || { text: "" }, base);
  const text = footer.runs.map((r) => r.text).join("");
  const pieces = text.split(/(\{page\}|\{pages\})/);
  const runs = pieces
    .map((piece) => {
      if (piece === "{page}") return `<w:fldSimple w:instr=" PAGE "><w:r>${rPr}<w:t>1</w:t></w:r></w:fldSimple>`;
      if (piece === "{pages}") return `<w:fldSimple w:instr=" NUMPAGES "><w:r>${rPr}<w:t>1</w:t></w:r></w:fldSimple>`;
      if (!piece) return "";
      return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(xmlSafe(piece))}</w:t></w:r>`;
    })
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}>` +
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${JC[footer.align] || "center"}"/></w:pPr>${runs}</w:p></w:ftr>`
  );
}

function coreXml(title) {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${escapeXml(title)}</dc:title><dc:creator>Offizin</dc:creator>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`
  );
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Offizin</Application></Properties>`;

/**
 * The publication record as custom document properties, so a .docx opened as a new document
 * still knows which template and theme made it. Word for Windows and Mac keep at most 255
 * characters per property, so the theme is split across as many properties as it needs.
 */
export const PROPERTY_LIMIT = 255;
const PROP_TEMPLATE = "OffizinTemplate";
const PROP_THEME = "OffizinTheme";

export function publicationProperties(record) {
  if (!record || !record.templateId) return [];
  const props = [{ name: PROP_TEMPLATE, value: String(record.templateId).slice(0, PROPERTY_LIMIT) }];
  const theme = JSON.stringify(record.theme || {});
  const parts = [];
  for (let i = 0; i < theme.length; i += PROPERTY_LIMIT) parts.push(theme.slice(i, i + PROPERTY_LIMIT));
  props.push({ name: `${PROP_THEME}Parts`, value: String(parts.length) });
  parts.forEach((value, i) => props.push({ name: `${PROP_THEME}${i + 1}`, value }));
  return props;
}

/** Reads the record back from [{ key or name, value }], as Word's customProperties lists them. */
export function publicationFromProperties(items) {
  const map = {};
  for (const p of items || []) {
    const key = p && (p.key ?? p.name);
    if (typeof key === "string" && p.value !== undefined && p.value !== null) map[key.toLowerCase()] = String(p.value);
  }
  const templateId = map[PROP_TEMPLATE.toLowerCase()];
  if (!templateId) return null;
  const count = Number(map[`${PROP_THEME}Parts`.toLowerCase()]) || 0;
  let json = "";
  for (let i = 1; i <= count; i++) json += map[`${PROP_THEME}${i}`.toLowerCase()] || "";
  let theme = null;
  try {
    theme = json ? JSON.parse(json) : null;
  } catch (e) {
    theme = null;
  }
  return { templateId, theme: theme && typeof theme === "object" ? theme : null };
}

function customPropertiesXml(props) {
  const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    props.map((p, i) => `<property fmtid="${FMTID}" pid="${i + 2}" name="${escapeXml(p.name)}"><vt:lpwstr>${escapeXml(p.value)}</vt:lpwstr></property>`).join("") +
    `</Properties>`
  );
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
}

/** Book fold applies to booklets unless the caller asks for the page-by-page layout. */
function useBookFold(model, opts) {
  if (!model.booklet) return false;
  if (opts && typeof opts.bookFold === "boolean") return opts.bookFold;
  return !!(model.print && model.print.bookFold);
}

/** The complete word/document.xml for a model. `media` collects pictures. */
export function documentXml(model, media, opts = {}) {
  const writer = opts.writer || new Writer(media, { mode: "doc", modelStyles: model.styles });
  const body = bodyXml(model, writer, useBookFold(model, opts));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`;
}

/**
 * A complete .docx file as bytes.
 * `images(block)` returns { base64, mime } for each picture block, or null.
 * opts.bookFold: true prints booklets with Word's book fold (whole sheet page size); false lays
 * booklets out page by page. The default follows model.print.bookFold.
 * opts.record: a publication record ({ templateId, theme }) written as custom document properties.
 */
export function docxBytes(model, images, opts = {}) {
  const media = new Media(images);
  const writer = new Writer(media, { mode: "doc", modelStyles: model.styles });
  const bookFold = useBookFold(model, opts);
  const doc = documentXml(model, media, { writer, bookFold });
  const hasNumbering = writer.numbering.defs.size > 0;
  const hasFooter = !!model.footer;
  const customProps = publicationProperties(opts.record);
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
    (hasNumbering ? `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` : "") +
    (hasFooter ? `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` : "") +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    (customProps.length ? `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>` : "") +
    `</Types>`;
  const rootRels = relsXml([
    `<Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/>`,
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
    `<Relationship Id="rId3" Type="${REL}extended-properties" Target="docProps/app.xml"/>`,
    ...(customProps.length ? [`<Relationship Id="rId4" Type="${REL}custom-properties" Target="docProps/custom.xml"/>`] : []),
  ]);
  const docRels = relsXml([
    `<Relationship Id="rIdOzStyles" Type="${REL}styles" Target="styles.xml"/>`,
    `<Relationship Id="rIdOzSettings" Type="${REL}settings" Target="settings.xml"/>`,
    ...(hasNumbering ? [`<Relationship Id="rIdOzNumbering" Type="${REL}numbering" Target="numbering.xml"/>`] : []),
    ...(hasFooter ? [`<Relationship Id="rIdOzFooter" Type="${REL}footer" Target="footer1.xml"/>`] : []),
    ...media.items.map((m) => `<Relationship Id="${m.id}" Type="${REL}image" Target="${m.name}"/>`),
  ]);
  return zip([
    { name: "[Content_Types].xml", data: contentTypes },
    { name: "_rels/.rels", data: rootRels },
    { name: "docProps/core.xml", data: coreXml(model.name) },
    { name: "docProps/app.xml", data: APP_XML },
    ...(customProps.length ? [{ name: "docProps/custom.xml", data: customPropertiesXml(customProps) }] : []),
    { name: "word/document.xml", data: doc },
    { name: "word/styles.xml", data: stylesXml(model) },
    { name: "word/settings.xml", data: settingsXml(model.print || {}, bookFold) },
    ...(hasNumbering ? [{ name: "word/numbering.xml", data: writer.numbering.xml() }] : []),
    ...(hasFooter ? [{ name: "word/footer1.xml", data: footerXml(model.footer, new Writer(media, { mode: "part" })) }] : []),
    { name: "word/_rels/document.xml.rels", data: docRels },
    ...media.items.map((m) => ({ name: "word/" + m.name, data: base64ToBytes(m.base64) })),
  ]);
}

/**
 * A flat OPC package for Word's insertOoxml, holding just the given blocks.
 * opts.tag and opts.alias wrap the blocks in a block-level content control, so Offizin can find
 * the part again for its own Undo and Redo. Page setup is not included: Word ignores it on insertion.
 */
export function flatPackage(blocks, width, images, opts = {}) {
  const media = new Media(images);
  const writer = new Writer(media, { mode: "part", modelStyles: opts.modelStyles || null });
  let body = writer.blocks(blocks, width);
  if (/<\/w:tbl>$/.test(body) || !body) body += TINY_P;
  if (opts.tag) {
    // w15:appearance hidden: no frame or title tab when the person clicks inside the part.
    body =
      `<w:sdt><w:sdtPr><w:alias w:val="${escapeXml(opts.alias || "Offizin")}"/><w:tag w:val="${escapeXml(opts.tag)}"/><w15:appearance w15:val="hidden"/></w:sdtPr>` +
      `<w:sdtContent>${body}</w:sdtContent></w:sdt>`;
  }
  const doc = `<w:document ${NS}><w:body>${body}</w:body></w:document>`;
  const parts = [
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>`,
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${doc}</pkg:xmlData></pkg:part>`,
  ];
  if (media.items.length) {
    parts.push(
      `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${media.items.map((m) => `<Relationship Id="${m.id}" Type="${REL}image" Target="${m.name}"/>`).join("")}</Relationships></pkg:xmlData></pkg:part>`
    );
    for (const m of media.items) parts.push(`<pkg:part pkg:name="/word/${m.name}" pkg:contentType="${m.mime}" pkg:compression="store"><pkg:binaryData>${m.base64}</pkg:binaryData></pkg:part>`);
  }
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">${parts.join("")}</pkg:package>`;
}

/** Body XML for blocks in part mode (used by tests and by hosts that need raw body XML). */
export function blocksXml(blocks, width, media, opts = {}) {
  return new Writer(media, { mode: opts.mode || "part", modelStyles: opts.modelStyles || null }).blocks(blocks, width);
}

export { TINY_P };
