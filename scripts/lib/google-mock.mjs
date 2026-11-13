/**
 * A simulated Google Docs (DocumentApp, CacheService, Utilities) for testing Code.gs without Google.
 *
 * It implements only the methods Offizin uses, fails loudly on anything else, and enforces the
 * rules real Docs and Apps Script enforce: insert positions must be valid, the last paragraph of
 * a document or cell cannot be removed, every table cell and the document end with a paragraph,
 * page breaks cannot go inside table cells, and cache values must stay under 100 KB and six hours.
 * Used by scripts/check.mjs and scripts/tests/google.mjs.
 */
import fs from "node:fs";
import vm from "node:vm";

const TYPES = ["BODY_SECTION", "FOOTER_SECTION", "PARAGRAPH", "LIST_ITEM", "TABLE", "TABLE_ROW", "TABLE_CELL", "TEXT", "INLINE_IMAGE", "PAGE_BREAK", "HORIZONTAL_RULE", "INLINE_DRAWING", "EQUATION"];
const enumOf = (names) => Object.fromEntries(names.map((n) => [n, n]));
const DocumentApp = {
  ElementType: enumOf(TYPES),
  Attribute: enumOf(["FONT_FAMILY", "FONT_SIZE", "FOREGROUND_COLOR", "BACKGROUND_COLOR", "BOLD", "ITALIC", "UNDERLINE", "STRIKETHROUGH", "LINK_URL"]),
  HorizontalAlignment: enumOf(["LEFT", "CENTER", "RIGHT", "JUSTIFY"]),
  VerticalAlignment: enumOf(["TOP", "CENTER", "BOTTOM"]),
  ParagraphHeading: enumOf(["NORMAL", "TITLE", "SUBTITLE", "HEADING1", "HEADING2", "HEADING3", "HEADING4", "HEADING5", "HEADING6"]),
  GlyphType: enumOf(["BULLET", "HOLLOW_BULLET", "SQUARE_BULLET", "NUMBER"]),
};

/** Wraps a mock so any method the mock does not implement fails loudly. */
function strict(obj) {
  return new Proxy(obj, {
    get(target, prop) {
      if (typeof prop === "symbol" || prop in target || prop === "then" || prop === "toJSON") return target[prop];
      throw new Error(`DocumentApp mock: ${target.type}.${String(prop)} is not available`);
    },
  });
}

let nextListId = 1;

/** Empties a container the way Docs does: removed elements are no longer attached to anything. */
function detach(node) {
  for (const c of node.children) c.parent = null;
  node.children = [];
}

class Node {
  constructor(type) {
    this.type = type;
    this.children = [];
    this.parent = null;
    this.attrs = {};
    this.self = strict(this);
    return this.self;
  }
  getType() { return this.type; }
  getParent() { return this.parent; }
  getNumChildren() { return this.children.length; }
  getChild(i) {
    if (i < 0 || i >= this.children.length) throw new Error(`${this.type}.getChild(${i}) out of range`);
    return this.children[i];
  }
  getChildIndex(c) {
    const i = this.children.indexOf(c);
    if (i < 0) throw new Error("getChildIndex: not a child");
    return i;
  }
  removeFromParent() {
    const p = this.parent;
    if (!p) throw new Error("removeFromParent: element is not attached");
    const isPara = this.type === "PARAGRAPH" || this.type === "LIST_ITEM";
    if (isPara && ["TABLE_CELL", "BODY_SECTION", "FOOTER_SECTION"].includes(p.type)) {
      const paras = p.children.filter((c) => c.type === "PARAGRAPH" || c.type === "LIST_ITEM");
      if (paras.length === 1) throw new Error(`cannot remove the last paragraph of a ${p.type.toLowerCase()}`);
      if (p.children[p.children.length - 1] === this.self) throw new Error(`cannot remove the last paragraph of a ${p.type.toLowerCase()}`);
    }
    p.children.splice(p.children.indexOf(this.self), 1);
    this.parent = null;
    return this.self;
  }
  add(i, child) {
    if (i < 0 || i > this.children.length) throw new Error(`${this.type}: insert at ${i} out of range 0..${this.children.length}`);
    child.parent = this.self;
    this.children.splice(i, 0, child);
    return child;
  }
  setAttributes(a) { Object.assign(this.attrs, a); return this.self; }
  insideCell() {
    for (let n = this.parent; n; n = n.parent) if (n.type === "TABLE_CELL") return true;
    return false;
  }
}

/** The text of a paragraph, with attribute runs. */
class TextEl extends Node {
  constructor(para) { super("TEXT"); this.para = para; }
  asText() { return this.self; }
  getText() { return this.para.text; }
  check(s, e, what) { if (s < 0 || e >= this.para.text.length || e < s) throw new Error(`${what}(${s}, ${e}) out of range for "${this.para.text}"`); }
  at(i, key, fallback) {
    for (let r = this.para.runs.length - 1; r >= 0; r--) {
      const run = this.para.runs[r];
      if (i >= run.s && i <= run.e && key in run.a) return run.a[key];
    }
    return key in this.para.attrs ? this.para.attrs[key] : fallback;
  }
  setAttributes(s, e, a) {
    if (typeof s === "object") return Object.assign(this.para.attrs, s), this.self;
    this.check(s, e, "setAttributes");
    this.para.runs.push({ s, e, a: { ...a } });
    return this.self;
  }
  setBold(s, e, v) { this.check(s, e, "setBold"); this.para.runs.push({ s, e, a: { BOLD: v } }); return this.self; }
  setLinkUrl(s, e, url) { this.check(s, e, "setLinkUrl"); this.para.runs.push({ s, e, a: { LINK_URL: url } }); return this.self; }
  setForegroundColor(s, e, c) { this.check(s, e, "setForegroundColor"); this.para.runs.push({ s, e, a: { FOREGROUND_COLOR: c } }); return this.self; }
  setFontFamily(s, e, f) { this.check(s, e, "setFontFamily"); this.para.runs.push({ s, e, a: { FONT_FAMILY: f } }); return this.self; }
  /** Like Docs: the start of every distinct formatting run, including where a run ends. */
  getTextAttributeIndices() { return [...new Set([0, ...this.para.runs.flatMap((r) => [r.s, r.e + 1])])].filter((i) => i < Math.max(1, this.para.text.length)).sort((a, b) => a - b); }
  getForegroundColor(i) { const c = this.at(i, "FOREGROUND_COLOR", "#000000"); return c ? String(c).toLowerCase() : null; }
  getBackgroundColor(i) { return this.at(i, "BACKGROUND_COLOR", null); }
  getFontFamily(i) { return this.at(i, "FONT_FAMILY", "Arial"); }
  getFontSize(i) { return this.at(i, "FONT_SIZE", 11); }
  isBold(i) { return !!this.at(i, "BOLD", false); }
  isItalic(i) { return !!this.at(i, "ITALIC", false); }
  isUnderline(i) { return !!this.at(i, "UNDERLINE", false); }
  isStrikethrough(i) { return !!this.at(i, "STRIKETHROUGH", false); }
  getLinkUrl(i) { return this.at(i, "LINK_URL", null); }
}

class Para extends Node {
  constructor(text, type = "PARAGRAPH") {
    super(type);
    this.text = String(text);
    this.runs = [];
    this.heading = "NORMAL";
    this.alignment = null;
    this.indent = { start: 0, first: 0, end: 0 };
    this.spacing = { before: 0, after: 0, line: 1 };
    this.glyph = "BULLET";
    this.nesting = 0;
    this.listId = type === "LIST_ITEM" ? "list" + nextListId++ : null;
    if (/\r/.test(this.text)) throw new Error("paragraph text contains a carriage return");
    if (this.text.length) this.add(0, new TextEl(this));
  }
  getText() { return this.text; }
  editAsText() { return this.children.find((c) => c.type === "TEXT") || new TextEl(this); }
  asParagraph() { return this.self; }
  asListItem() { if (this.type !== "LIST_ITEM") throw new Error("not a list item"); return this.self; }
  setText(t) { this.clear(); return this.appendText(t), this.self; }
  appendText(t) {
    if (/\r/.test(t)) throw new Error("text contains a carriage return");
    let te = this.children.find((c) => c.type === "TEXT");
    if (!te) te = this.add(this.children.length, new TextEl(this));
    this.text += String(t);
    return te;
  }
  clear() { this.text = ""; this.runs = []; detach(this); return this.self; }
  getHeading() { return this.heading; }
  setHeading(h) { if (!h) throw new Error("setHeading(undefined)"); this.heading = h; return this.self; }
  getAlignment() { return this.alignment; }
  setAlignment(a) { if (!a) throw new Error("setAlignment(undefined)"); this.alignment = a; return this.self; }
  setSpacingBefore(v) { if (!(v >= 0)) throw new Error("bad spacing " + v); this.spacing.before = v; return this.self; }
  setSpacingAfter(v) { if (!(v >= 0)) throw new Error("bad spacing " + v); this.spacing.after = v; return this.self; }
  setLineSpacing(v) { if (!(v > 0)) throw new Error("bad line spacing " + v); this.spacing.line = v; return this.self; }
  getSpacingBefore() { return this.spacing.before; }
  getSpacingAfter() { return this.spacing.after; }
  getLineSpacing() { return this.spacing.line; }
  setIndentFirstLine(v) { if (!Number.isFinite(v)) throw new Error("bad indent"); this.indent.first = v; return this.self; }
  setIndentStart(v) { if (!Number.isFinite(v)) throw new Error("bad indent"); this.indent.start = v; return this.self; }
  setIndentEnd(v) { this.indent.end = v; return this.self; }
  getIndentFirstLine() { return this.indent.first; }
  getIndentStart() { return this.indent.start; }
  getIndentEnd() { return this.indent.end; }
  setGlyphType(g) { this.glyph = g; return this.self; }
  getGlyphType() { return this.glyph; }
  setNestingLevel(n) { this.nesting = n; return this.self; }
  getNestingLevel() { return this.nesting; }
  setListId(item) { this.listId = item.getListId(); return this.self; }
  getListId() { return this.listId; }
  appendInlineImage(blob) { const img = new Image(blob); this.add(this.children.length, img); return img; }
  appendPageBreak() { if (this.insideCell()) throw new Error("appendPageBreak inside a table cell"); return this.add(this.children.length, new Node("PAGE_BREAK")); }
  appendHorizontalRule() { return this.add(this.children.length, new Node("HORIZONTAL_RULE")); }
}

class Image extends Node {
  constructor(blob) {
    super("INLINE_IMAGE");
    if (!blob || !blob.bytes) throw new Error("an image needs a blob");
    this.blob = blob;
    this.alt = "";
    this.title = "";
    this.w = 100;
    this.h = 100;
  }
  asInlineImage() { return this.self; }
  getBlob() { return this.blob; }
  setWidth(v) { if (!(v > 0) || !Number.isInteger(v)) throw new Error("bad image width " + v); this.w = v; return this.self; }
  setHeight(v) { if (!(v > 0) || !Number.isInteger(v)) throw new Error("bad image height " + v); this.h = v; return this.self; }
  getWidth() { return this.w; }
  getHeight() { return this.h; }
  setAltDescription(v) { this.alt = v; return this.self; }
  setAltTitle(v) { this.title = v; return this.self; }
  getAltDescription() { return this.alt; }
  getAltTitle() { return this.title; }
}

const containerMethods = {
  insertParagraph(i, text) { return this.add(i, new Para(text)); },
  insertListItem(i, text) { return this.add(i, new Para(text, "LIST_ITEM")); },
  insertTable(i, cells) { return this.add(i, new Table(cells)); },
  insertImage(i, blob) {
    const p = this.add(i, new Para(""));
    const img = new Image(blob);
    p.add(0, img);
    return img;
  },
  appendParagraph(text) { return this.add(this.children.length, new Para(text)); },
  appendListItem(text) { return this.add(this.children.length, new Para(text, "LIST_ITEM")); },
  appendTable(cells) { return this.add(this.children.length, new Table(cells)); },
};

class Cell extends Node {
  constructor(text = "") {
    super("TABLE_CELL");
    this.bg = null;
    this.pad = [5, 5, 5, 5];
    this.valign = "TOP";
    this.add(0, new Para(text));
  }
  setPaddingTop(v) { if (!(v >= 0)) throw new Error("bad padding"); this.pad[0] = v; return this.self; }
  setPaddingRight(v) { if (!(v >= 0)) throw new Error("bad padding"); this.pad[1] = v; return this.self; }
  setPaddingBottom(v) { if (!(v >= 0)) throw new Error("bad padding"); this.pad[2] = v; return this.self; }
  setPaddingLeft(v) { if (!(v >= 0)) throw new Error("bad padding"); this.pad[3] = v; return this.self; }
  getPaddingTop() { return this.pad[0]; }
  getPaddingRight() { return this.pad[1]; }
  getPaddingBottom() { return this.pad[2]; }
  getPaddingLeft() { return this.pad[3]; }
  setBackgroundColor(c) { if (!/^#[0-9A-F]{6}$/i.test(c)) throw new Error("bad color " + c); this.bg = c; return this.self; }
  getBackgroundColor() { return this.bg || null; }
  setVerticalAlignment(v) { if (!v) throw new Error("setVerticalAlignment(undefined)"); this.valign = v; return this.self; }
  getVerticalAlignment() { return this.valign; }
  getWidth() {
    const row = this.parent;
    const table = row && row.parent;
    return table ? table.widths[row.children.indexOf(this.self)] ?? null : null;
  }
  asTableCell() { return this.self; }
}
Object.assign(Cell.prototype, containerMethods);

class Row extends Node {
  constructor(texts) { super("TABLE_ROW"); this.minHeight = 0; texts.forEach((t, i) => this.add(i, new Cell(String(t)))); }
  getCell(i) { return this.getChild(i); }
  getNumCells() { return this.children.length; }
  setMinimumHeight(v) { if (!(v > 0)) throw new Error("bad row height"); this.minHeight = v; return this.self; }
  getMinimumHeight() { return this.minHeight; }
}

class Table extends Node {
  constructor(cells) {
    super("TABLE");
    this.border = null;
    this.borderWidth = 1;
    this.cols = cells[0].length;
    this.widths = [];
    cells.forEach((r, i) => {
      if (r.length !== this.cols) throw new Error("ragged table");
      this.add(i, new Row(r));
    });
  }
  getRow(r) { return this.getChild(r); }
  getCell(r, c) { return this.getChild(r).getChild(c); }
  getNumRows() { return this.children.length; }
  setBorderWidth(v) { if (!(v >= 0)) throw new Error("bad border width"); this.borderWidth = v; return this.self; }
  getBorderWidth() { return this.borderWidth; }
  setBorderColor(c) { this.border = c; return this.self; }
  getBorderColor() { return this.border || null; }
  setColumnWidth(i, w) { if (i >= this.cols || !(w > 0)) throw new Error(`setColumnWidth(${i}, ${w})`); this.widths[i] = w; return this.self; }
  asTable() { return this.self; }
}

class Footer extends Node {
  constructor() { super("FOOTER_SECTION"); this.add(0, new Para("")); }
  clear() { detach(this); this.add(0, new Para("")); return this.self; }
  getText() { return this.children.map((c) => c.text || "").join("\n"); }
}
Object.assign(Footer.prototype, containerMethods);

class Body extends Node {
  constructor() {
    super("BODY_SECTION");
    this.add(0, new Para("Existing text"));
    this.page = { pw: 612, ph: 792, mt: 72, mb: 72, ml: 72, mr: 72 };
    this.headings = { NORMAL: { FONT_FAMILY: "Arial", FONT_SIZE: 11 }, TITLE: { FONT_SIZE: 26 }, SUBTITLE: { FONT_SIZE: 15 }, HEADING1: { FONT_SIZE: 20 }, HEADING2: { FONT_SIZE: 16 } };
  }
  clear() { detach(this); this.add(0, new Para("")); return this.self; }
  insertPageBreak(i) { const p = this.add(i, new Para("")); const pb = new Node("PAGE_BREAK"); p.add(0, pb); return pb; }
  setPageWidth(v) { this.page.pw = v; return this.self; }
  setPageHeight(v) { this.page.ph = v; return this.self; }
  setMarginTop(v) { this.page.mt = v; return this.self; }
  setMarginBottom(v) { this.page.mb = v; return this.self; }
  setMarginLeft(v) { this.page.ml = v; return this.self; }
  setMarginRight(v) { this.page.mr = v; return this.self; }
  getPageWidth() { return this.page.pw; }
  getPageHeight() { return this.page.ph; }
  getMarginTop() { return this.page.mt; }
  getMarginBottom() { return this.page.mb; }
  getMarginLeft() { return this.page.ml; }
  getMarginRight() { return this.page.mr; }
  /** Like Docs: every attribute is reported, and one the style does not set is null. */
  getHeadingAttributes(h) {
    if (!this.headings[h] && !DocumentApp.ParagraphHeading[h]) throw new Error("bad heading");
    const out = Object.fromEntries(Object.keys(DocumentApp.Attribute).map((k) => [k, null]));
    return { ...out, ...(this.headings[h] || {}) };
  }
  setHeadingAttributes(h, a) { if (!DocumentApp.ParagraphHeading[h]) throw new Error("bad heading " + h); this.headings[h] = { ...(this.headings[h] || {}), ...a }; return this.self; }
  getImages() {
    const out = [];
    const walk = (n) => { if (n.type === "INLINE_IMAGE") out.push(n); n.children.forEach(walk); };
    walk(this.self);
    return out;
  }
  findText(pattern) {
    const re = new RegExp(pattern);
    let hit = null;
    const walk = (n) => {
      if (hit) return;
      if ((n.type === "PARAGRAPH" || n.type === "LIST_ITEM") && re.test(n.text)) {
        const m = re.exec(n.text);
        hit = { getElement: () => n.editAsText(), getStartOffset: () => m.index, getEndOffsetInclusive: () => m.index + m[0].length - 1 };
      }
      n.children.forEach(walk);
    };
    walk(this.self);
    return hit;
  }
}
Object.assign(Body.prototype, containerMethods);

const attached = (el, root) => {
  for (let n = el; n; n = n.parent) if (n === root) return true;
  return false;
};

/** Plain text of a container, for tests: paragraphs joined by new lines, tables cell by cell. */
export function textOf(node) {
  if (node.type === "PARAGRAPH" || node.type === "LIST_ITEM") return node.text;
  return node.children.map(textOf).filter((t) => t !== "").join("\n");
}

export function makeGoogle(codePath) {
  const body = new Body();
  let footer = null;
  const named = [];
  const docProps = {};
  const userProps = {};
  const cacheStore = new Map();
  const props = (store) => ({
    getProperty: (k) => store[k] ?? null,
    setProperty: (k, v) => { if (String(v).length > 9216) throw new Error("property too large"); store[k] = v; },
    deleteProperty: (k) => { delete store[k]; },
  });
  const rangeOf = (elements) => ({ getRangeElements: () => elements.filter((el) => attached(el, body) || (footer && attached(el, footer))).map((el) => ({ getElement: () => el })) });
  const doc = {
    getBody: () => body,
    getCursor: () => null,
    getSelection: () => null,
    getFooter: () => footer,
    addFooter: () => { if (footer) throw new Error("the document already has a footer"); footer = new Footer(); return footer; },
    newRange: () => {
      const els = [];
      const b = { addElement: (el) => { if (!el || !el.getType) throw new Error("addElement needs an element"); els.push(el); return b; }, build: () => rangeOf(els) };
      return b;
    },
    addNamedRange: (name, range) => {
      if (!name || name.length > 256) throw new Error("bad named range name");
      const nr = { name, range };
      const obj = { getName: () => name, getRange: () => nr.range, remove: () => { const i = named.indexOf(nr); if (i >= 0) named.splice(i, 1); } };
      nr.obj = obj;
      named.push(nr);
      return obj;
    },
    getNamedRanges: (name) => named.filter((n) => name === undefined || n.name === name).map((n) => n.obj),
    setSelection: () => {},
  };
  const cache = {
    putAll(values, seconds) {
      if (!(seconds >= 1 && seconds <= 21600)) throw new Error("cache expiration must be between 1 second and 6 hours");
      for (const [k, v] of Object.entries(values)) {
        if (k.length > 250) throw new Error("cache key too long");
        if (Buffer.byteLength(String(v), "utf8") > 100 * 1024) throw new Error("cache value over 100 KB");
        cacheStore.set(k, String(v));
      }
      if (cacheStore.size > 1000) throw new Error("more than 1000 cached items");
    },
    getAll(keys) { const out = {}; keys.forEach((k) => cacheStore.has(k) && (out[k] = cacheStore.get(k))); return out; },
    removeAll(keys) { keys.forEach((k) => cacheStore.delete(k)); },
  };
  const blob = (bytes, mime, name) => ({ bytes, mime, name, getBytes: () => bytes, getContentType: () => mime });
  const context = {
    DocumentApp: { ...DocumentApp, getActiveDocument: () => doc, getUi: () => ({}) },
    PropertiesService: { getDocumentProperties: () => props(docProps), getUserProperties: () => props(userProps) },
    CacheService: { getDocumentCache: () => cache },
    Utilities: {
      base64Decode: (s) => [...Buffer.from(s, "base64")],
      base64Encode: (bytes) => Buffer.from(bytes).toString("base64"),
      newBlob: blob,
      getUuid: () => "uuid-" + Math.random().toString(36).slice(2),
    },
    HtmlService: {},
    console,
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(codePath, "utf8"), context);
  return { context, body, doc, cache: cacheStore, docProps, getFooter: () => footer, named, Para, Table, blob };
}

/** Returns a list of rule violations in a document body. */
export function invariants(body) {
  const problems = [];
  const fail = (m) => problems.push(m);
  const last = body.children[body.children.length - 1];
  if (!last || (last.type !== "PARAGRAPH" && last.type !== "LIST_ITEM")) fail("the document does not end with a paragraph");
  const walk = (n) => {
    if (n.type === "TABLE_CELL") {
      const l = n.children[n.children.length - 1];
      if (!l || (l.type !== "PARAGRAPH" && l.type !== "LIST_ITEM")) fail("a table cell does not end with a paragraph");
    }
    n.children.forEach(walk);
  };
  walk(body);
  return problems;
}
