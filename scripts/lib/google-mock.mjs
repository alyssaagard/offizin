/**
 * A simulated Google Docs (DocumentApp) for testing Code.gs without Google.
 *
 * It implements only the methods Offizin uses, fails loudly on anything else, and enforces
 * the rules real Docs enforces: insert positions must be valid, and every table cell and
 * the document itself must end with a paragraph. Used by scripts/check.mjs.
 */
import fs from "node:fs";
import vm from "node:vm";

const TYPES = ["BODY_SECTION", "PARAGRAPH", "LIST_ITEM", "TABLE", "TABLE_ROW", "TABLE_CELL", "TEXT", "INLINE_IMAGE", "PAGE_BREAK"];
const enumOf = (names) => Object.fromEntries(names.map((n) => [n, n]));
const DocumentApp = {
  ElementType: enumOf(TYPES),
  Attribute: enumOf(["FONT_FAMILY", "FONT_SIZE", "FOREGROUND_COLOR", "BOLD", "ITALIC", "UNDERLINE"]),
  HorizontalAlignment: enumOf(["LEFT", "CENTER", "RIGHT", "JUSTIFY"]),
  VerticalAlignment: enumOf(["TOP", "CENTER", "BOTTOM"]),
  ParagraphHeading: enumOf(["NORMAL"]),
  GlyphType: enumOf(["BULLET"]),
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
    if (p.type === "TABLE_CELL" && p.children.length === 1) throw new Error("cannot remove the last paragraph of a cell");
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
}

class TextEl extends Node {
  constructor(para) { super("TEXT"); this.para = para; }
  asText() { return this.self; }
  getText() { return this.para.text; }
  setBold(s, e, v) { if (s < 0 || e >= this.para.text.length || e < s) throw new Error(`setBold(${s}, ${e}) out of range for "${this.para.text}"`); return this.self; }
  getTextAttributeIndices() { return [0]; }
  getForegroundColor() { return (this.para.attrs.FOREGROUND_COLOR || "#000000").toLowerCase(); }
  setForegroundColor(s, e, c) { this.para.attrs.FOREGROUND_COLOR = c; return this.self; }
  getFontFamily() { return this.para.attrs.FONT_FAMILY || "Arial"; }
  setFontFamily(s, e, f) { this.para.attrs.FONT_FAMILY = f; return this.self; }
}

class Para extends Node {
  constructor(text, type = "PARAGRAPH") {
    super(type);
    this.text = String(text);
    if (/\r/.test(this.text)) throw new Error("paragraph text contains a carriage return");
    if (this.text.length) this.add(0, new TextEl(this));
  }
  getText() { return this.text; }
  editAsText() { return this.children.find((c) => c.type === "TEXT") || new TextEl(this); }
  asParagraph() { return this.self; }
  setHeading() { return this.self; }
  setAlignment(a) { if (!a) throw new Error("setAlignment(undefined)"); return this.self; }
  setSpacingBefore(v) { if (!(v >= 0)) throw new Error("bad spacing " + v); return this.self; }
  setSpacingAfter(v) { if (!(v >= 0)) throw new Error("bad spacing " + v); return this.self; }
  setLineSpacing(v) { if (!(v > 0)) throw new Error("bad line spacing " + v); return this.self; }
  setGlyphType() { return this.self; }
  setIndentFirstLine() { return this.self; }
  setIndentStart() { return this.self; }
  setListId() { return this.self; }
}

class Image extends Node {
  constructor() { super("INLINE_IMAGE"); this.alt = ""; }
  getParent() { return this.parent; }
  setWidth(v) { if (!(v > 0)) throw new Error("bad image width"); return this.self; }
  setHeight(v) { if (!(v > 0)) throw new Error("bad image height"); return this.self; }
  setAltDescription(v) { this.alt = v; return this.self; }
  setAltTitle() { return this.self; }
  getAltDescription() { return this.alt; }
}

const containerMethods = {
  insertParagraph(i, text) { return this.add(i, new Para(text)); },
  insertListItem(i, text) { return this.add(i, new Para(text, "LIST_ITEM")); },
  insertTable(i, cells) { return this.add(i, new Table(cells)); },
  insertImage(i, blob) {
    if (!blob || !blob.bytes) throw new Error("insertImage needs a blob");
    const p = this.add(i, new Para(""));
    const img = new Image();
    p.add(0, img);
    return img;
  },
};

class Cell extends Node {
  constructor() { super("TABLE_CELL"); this.bg = null; this.add(0, new Para("")); }
  setPaddingTop(v) { if (!(v >= 0)) throw new Error("bad padding"); return this.self; }
  setPaddingRight(v) { return this.setPaddingTop(v); }
  setPaddingBottom(v) { return this.setPaddingTop(v); }
  setPaddingLeft(v) { return this.setPaddingTop(v); }
  setBackgroundColor(c) { if (!/^#[0-9A-F]{6}$/i.test(c)) throw new Error("bad color " + c); this.bg = c; return this.self; }
  getBackgroundColor() { return this.bg || null; }
  setVerticalAlignment(v) { if (!v) throw new Error("setVerticalAlignment(undefined)"); return this.self; }
  asTableCell() { return this.self; }
}
Object.assign(Cell.prototype, containerMethods);

class Row extends Node {
  constructor(n) { super("TABLE_ROW"); for (let i = 0; i < n; i++) this.add(i, new Cell()); }
  getCell(i) { return this.getChild(i); }
  setMinimumHeight(v) { if (!(v > 0)) throw new Error("bad row height"); return this.self; }
}

class Table extends Node {
  constructor(cells) {
    super("TABLE");
    this.border = null;
    this.cols = cells[0].length;
    cells.forEach((r, i) => {
      if (r.length !== this.cols) throw new Error("ragged table");
      this.add(i, new Row(r.length));
    });
  }
  getRow(r) { return this.getChild(r); }
  getCell(r, c) { return this.getChild(r).getChild(c); }
  setBorderWidth(v) { if (!(v >= 0)) throw new Error("bad border width"); return this.self; }
  setBorderColor(c) { this.border = c; return this.self; }
  getBorderColor() { return this.border || null; }
  setColumnWidth(i, w) { if (i >= this.cols || !(w > 0)) throw new Error(`setColumnWidth(${i}, ${w})`); return this.self; }
  asTable() { return this.self; }
}

class Body extends Node {
  constructor() { super("BODY_SECTION"); this.add(0, new Para("Existing text")); }
  clear() { this.children = []; this.add(0, new Para("")); return this.self; }
  insertPageBreak(i) { const p = this.add(i, new Para("")); const pb = new Node("PAGE_BREAK"); p.add(0, pb); return pb; }
  setPageWidth(v) { this.pw = v; return this.self; }
  setPageHeight(v) { this.ph = v; return this.self; }
  setMarginTop() { return this.self; }
  setMarginBottom() { return this.self; }
  setMarginLeft(v) { this.ml = v; return this.self; }
  setMarginRight(v) { this.mr = v; return this.self; }
  getPageWidth() { return this.pw || 612; }
  getMarginLeft() { return this.ml || 72; }
  getMarginRight() { return this.mr || 72; }
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

export function makeGoogle(codePath) {
  const body = new Body();
  const docProps = {};
  const userProps = {};
  const props = (store) => ({ getProperty: (k) => store[k] ?? null, setProperty: (k, v) => { if (String(v).length > 9216) throw new Error("property too large"); store[k] = v; } });
  const doc = {
    getBody: () => body,
    getCursor: () => null,
    getSelection: () => null,
    newRange: () => { const r = { addElement: () => r, build: () => ({}) }; return r; },
    setSelection: () => {},
  };
  const context = {
    DocumentApp: { ...DocumentApp, getActiveDocument: () => doc, getUi: () => ({}) },
    PropertiesService: { getDocumentProperties: () => props(docProps), getUserProperties: () => props(userProps) },
    Utilities: { base64Decode: (s) => [...Buffer.from(s, "base64")], newBlob: (bytes, mime, name) => ({ bytes, mime, name }) },
    HtmlService: {},
    console,
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(codePath, "utf8"), context);
  return { context, body };
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

