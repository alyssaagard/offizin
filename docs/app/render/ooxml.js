/**
 * Writes a composed model as WordprocessingML (the XML inside a .docx file).
 *
 * Two outputs:
 *   docxBytes(model, images)   a complete .docx with exact page size and margins. The Word
 *                              add-in inserts it with insertFileFromBase64, and the web demo
 *                              offers it as a download that also opens in Google Docs.
 *   flatPackage(blocks, ...)   a single-XML package for Word's insertOoxml, used to insert
 *                              page parts at the cursor.
 *
 * Units: twips (1/20 pt) for spacing and widths, half points for font sizes,
 * eighths of a point for borders, and EMUs (12,700 per point) for pictures.
 */
import { escapeXml, xmlSafe } from "../core/util.js";
import { zip, base64ToBytes } from "./zip.js";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const tw = (pt) => Math.round(pt * 20);
const halfPt = (pt) => Math.round(pt * 2);
const emu = (pt) => Math.round(pt * 12700);
const hex = (c) => String(c || "000000").replace("#", "").toUpperCase();
const JC = { left: "left", center: "center", right: "right", justify: "both" };

/** A paragraph that takes up almost no space: used after tables and for page breaks. */
const TINY_PPR = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>';
const TINY_P = `<w:p>${TINY_PPR}</w:p>`;
const PAGE_BREAK_P = `<w:p>${TINY_PPR}<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:br w:type="page"/></w:r></w:p>`;

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

function runProps(b, run, colorOverride) {
  const font = escapeXml(b.font);
  const parts = [`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`];
  if (run.bold ?? b.bold) parts.push("<w:b/><w:bCs/>");
  if (b.italic) parts.push("<w:i/><w:iCs/>");
  parts.push(`<w:color w:val="${hex(colorOverride || run.color || b.color)}"/>`);
  if (b.tracking) parts.push(`<w:spacing w:val="${Math.round(b.tracking * b.size * 20)}"/>`);
  parts.push(`<w:sz w:val="${halfPt(b.size)}"/><w:szCs w:val="${halfPt(b.size)}"/>`);
  return `<w:rPr>${parts.join("")}</w:rPr>`;
}

function textRuns(b, run) {
  const rPr = runProps(b, run);
  const lines = xmlSafe(run.text).split("\n");
  return lines
    .map((line, i) => {
      const segs = line.split("\t").map((s) => `<w:t xml:space="preserve">${escapeXml(s)}</w:t>`).join("<w:tab/>");
      return `<w:r>${rPr}${i > 0 ? "<w:br/>" : ""}${segs}</w:r>`;
    })
    .join("");
}

function paragraphXml(b) {
  const ppr = [`<w:spacing w:before="${tw(b.before)}" w:after="${tw(b.after)}" w:line="${Math.round(b.line * 240)}" w:lineRule="auto"/>`];
  if (b.bullet) {
    const ind = tw(b.bullet.indent);
    ppr.unshift(`<w:tabs><w:tab w:val="left" w:pos="${ind}"/></w:tabs>`);
    ppr.push(`<w:ind w:left="${ind}" w:hanging="${ind}"/>`);
  }
  ppr.push(`<w:jc w:val="${JC[b.align] || "left"}"/>`);
  ppr.push(`<w:rPr><w:sz w:val="${halfPt(b.size)}"/><w:szCs w:val="${halfPt(b.size)}"/></w:rPr>`);
  let runs = "";
  if (b.bullet) runs += `<w:r>${runProps(b, { bold: false }, b.bullet.color)}<w:t xml:space="preserve">${escapeXml(b.bullet.char)}</w:t><w:tab/></w:r>`;
  runs += b.runs.map((r) => textRuns(b, r)).join("");
  return `<w:p><w:pPr>${ppr.join("")}</w:pPr>${runs}</w:p>`;
}

function indentsFor(width, container, align) {
  const extra = Math.max(0, container - width);
  if (align === "center") return [extra / 2, extra / 2];
  if (align === "right") return [extra, 0];
  return [0, extra];
}

function ruleXml(b, container) {
  const [l, r] = indentsFor(b.width, container, b.align);
  const sz = Math.max(2, Math.min(96, Math.round(b.weight * 8)));
  return (
    `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="${sz}" w:space="0" w:color="${hex(b.color)}"/></w:pBdr>` +
    `<w:spacing w:before="${tw(b.before)}" w:after="${tw(b.after)}" w:line="20" w:lineRule="exact"/>` +
    `<w:ind w:left="${tw(l)}" w:right="${tw(r)}"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`
  );
}

function spacerXml(b) {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.max(20, tw(b.height))}" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`;
}

function imageXml(b, media) {
  const item = media.add(b);
  const jc = `<w:jc w:val="${JC[b.align] || "left"}"/>`;
  const spacing = `<w:spacing w:before="${tw(b.before)}" w:after="${tw(b.after)}" w:line="240" w:lineRule="auto"/>`;
  if (!item) {
    // No picture data available (for example in tests): keep the space with a shaded frame.
    return `<w:p><w:pPr>${spacing}${jc}</w:pPr><w:r><w:t xml:space="preserve">${escapeXml(b.label || "Picture")}</w:t></w:r></w:p>`;
  }
  const id = media.nextDrawingId++;
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

function borderSide(name, s) {
  if (!s) return `<w:${name} w:val="nil"/>`;
  const sz = Math.max(2, Math.min(96, Math.round(s.weight * 8)));
  return `<w:${name} w:val="${s.style === "dashed" ? "dashed" : "single"}" w:sz="${sz}" w:space="0" w:color="${hex(s.color)}"/>`;
}

function gridXml(b, media) {
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
          let content = blocksXml(c.blocks, inner, media, { inCell: true });
          if (!content || /<\/w:tbl>$/.test(content)) content += TINY_P;
          return `<w:tc><w:tcPr>${tcPr.join("")}</w:tcPr>${content}</w:tc>`;
        })
        .join("");
      return `<w:tr>${trPr}${cells}</w:tr>`;
    })
    .join("");
  return `<w:tbl>${tblPr}${gridCols}${rows}</w:tbl>`;
}

/** Body XML for a list of blocks inside a container `width` points wide. */
export function blocksXml(blocks, width, media, opts = {}) {
  let out = "";
  blocks.forEach((b, i) => {
    switch (b.t) {
      case "p":
        out += paragraphXml(b);
        break;
      case "rule":
        out += ruleXml(b, width);
        break;
      case "spacer":
        out += spacerXml(b);
        break;
      case "image":
        out += imageXml(b, media);
        break;
      case "break":
        out += PAGE_BREAK_P;
        break;
      case "grid": {
        out += gridXml(b, media);
        // Word joins tables that touch, and needs a paragraph between a table and a page break.
        const next = blocks[i + 1];
        if (!next || next.t === "grid" || next.t === "break") {
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

function sectPr(page) {
  const w = tw(page.width);
  const h = tw(page.height);
  const m = page.margins;
  return (
    `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${page.width > page.height ? ' w:orient="landscape"' : ""}/>` +
    `<w:pgMar w:top="${tw(m.top)}" w:right="${tw(m.right)}" w:bottom="${tw(m.bottom)}" w:left="${tw(m.left)}" w:header="360" w:footer="360" w:gutter="0"/>` +
    `<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`
  );
}

function contentWidth(page) {
  return page.width - page.margins.left - page.margins.right;
}

function relsXml(items, extra = []) {
  const rels = [
    ...extra,
    ...items.map((m) => `<Relationship Id="${m.id}" Type="${REL_IMAGE}" Target="${m.name}"/>`),
  ];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
}

function stylesXml(model) {
  const f = escapeXml(model.fonts.body.name);
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>` +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
    `</w:styles>`
  );
}

const SETTINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>` +
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

function coreXml(title) {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${escapeXml(title)}</dc:title><dc:creator>Offizin</dc:creator>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`
  );
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Offizin</Application></Properties>`;

/** The complete word/document.xml for a model. */
export function documentXml(model, media) {
  const body = blocksXml(model.blocks, contentWidth(model.page), media);
  // End with a tiny paragraph so Word never adds a full-size one that spills onto a new page.
  const tail = /<\/w:tbl>$/.test(body) || /w:type="page"\/><\/w:r><\/w:p>$/.test(body) ? TINY_P : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${tail}${sectPr(model.page)}</w:body></w:document>`;
}

/**
 * A complete .docx file as bytes.
 * `images(block)` returns { base64, mime } for each picture block, or null.
 */
export function docxBytes(model, images) {
  const media = new Media(images);
  const doc = documentXml(model, media);
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
  const rootRels = relsXml([], [
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`,
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`,
  ]);
  const docRels = relsXml(media.items, [
    `<Relationship Id="rIdOzStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    `<Relationship Id="rIdOzSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`,
  ]);
  return zip([
    { name: "[Content_Types].xml", data: contentTypes },
    { name: "_rels/.rels", data: rootRels },
    { name: "docProps/core.xml", data: coreXml(model.name) },
    { name: "docProps/app.xml", data: APP_XML },
    { name: "word/document.xml", data: doc },
    { name: "word/styles.xml", data: stylesXml(model) },
    { name: "word/settings.xml", data: SETTINGS_XML },
    { name: "word/_rels/document.xml.rels", data: docRels },
    ...media.items.map((m) => ({ name: "word/" + m.name, data: base64ToBytes(m.base64) })),
  ]);
}

/**
 * A flat OPC package for Word's insertOoxml, holding just the given blocks.
 * Page setup is not included: Word ignores it on insertion.
 */
export function flatPackage(blocks, width, images) {
  const media = new Media(images);
  const body = blocksXml(blocks, width, media);
  const tail = /<\/w:tbl>$/.test(body) ? TINY_P : "";
  const doc = `<w:document ${NS}><w:body>${body}${tail}</w:body></w:document>`;
  const parts = [
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>`,
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${doc}</pkg:xmlData></pkg:part>`,
  ];
  if (media.items.length) {
    parts.push(
      `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${media.items.map((m) => `<Relationship Id="${m.id}" Type="${REL_IMAGE}" Target="${m.name}"/>`).join("")}</Relationships></pkg:xmlData></pkg:part>`
    );
    for (const m of media.items) parts.push(`<pkg:part pkg:name="/word/${m.name}" pkg:contentType="${m.mime}" pkg:compression="store"><pkg:binaryData>${m.base64}</pkg:binaryData></pkg:part>`);
  }
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">${parts.join("")}</pkg:package>`;
}

