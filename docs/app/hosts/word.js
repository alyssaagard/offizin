/**
 * The Microsoft Word host (Office.js). Loaded by word/taskpane.html.
 *
 * How Offizin builds a publication in Word:
 *   1. The panel composes the publication and this file writes it as a complete .docx in the
 *      browser (native styles, bullets, columns, footer, pictures). Nothing is sent to a server.
 *   2. The .docx replaces the open document (Document.insertFileFromBase64, WordApi 1.5, which
 *      also copies section properties such as page size), or opens as a new document
 *      (Application.createDocument, WordApi 1.3), which keeps the file's own print settings.
 *   3. Older versions fall back to Body.insertFileFromBase64 (WordApi 1.1), then insertOoxml.
 *   4. Where Word lets add-ins change page setup (WordApiDesktop 1.3: Word for Windows and Mac),
 *      Offizin applies the page size again, and the print settings Word keeps per section:
 *      book fold for booklets, two pages per sheet, mirror margins, and the gutter.
 *
 * Offizin's own history (contract version 2): every change returns an op that can be undone and
 * redone without touching Word's own undo list or the user's other edits.
 *   create     a snapshot of the document before (body OOXML, every section's page setup, headers
 *              and footers, and the paragraph style definitions) is restored on undo
 *   insert     each page part sits in a hidden content control tagged "offizin:op:<id>";
 *              undo reads what the part holds now, then empties it to a tiny anchor, and redo
 *              fills the anchor with exactly that, so the person's own typing in the part survives
 *   restyle    undo applies the reverse color and font map
 *   pageSetup  and columns restore the values read before the change, section by section (found
 *              by index in the document, not by where the cursor is at undo time)
 *
 * A publication opened as a new document carries a short record (template and theme) in its
 * custom document properties, so Offizin still knows the design when that file is opened later.
 *
 * Requirement sets are checked at run time, because an XML manifest cannot require
 * desktop-only sets. Sources: learn.microsoft.com/javascript/api/requirement-sets/word/.
 */
/* global Office, Word */
import { round, escapeXml } from "../core/util.js";
import { invertRestyleMap, wordSheet } from "../core/compose.js";
import { docxBytes, flatPackage, publicationFromProperties, TINY_P } from "../render/ooxml.js";
import { prepareImages } from "../render/pictures.js";
import { bytesToBase64 } from "../render/zip.js";

const STORE_KEY = "offizin.word.state";
const SETTINGS_KEY = "offizin";
const PLACEHOLDER_ALT = "Offizin picture placeholder";
const TAG_PREFIX = "offizin:op:";
const PAGE_FIELDS =
  "orientation,pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin,gutter,bookFoldPrinting,twoPagesOnOne,mirrorMargins,headerDistance,footerDistance,differentFirstPageHeaderFooter,oddAndEvenPagesHeaderFooter";
const HF_TYPES = ["Primary", "FirstPage", "EvenPages"];
const STYLE_FONT = "name,size,bold,italic,color";
const STYLE_PARA = "alignment,spaceBefore,spaceAfter,lineSpacing,keepWithNext,keepTogether,leftIndent,rightIndent,firstLineIndent,outlineLevel";
const BY_HAND = "turn on book fold by hand under Layout, Margins, Custom Margins, Multiple pages";

const supports = (set, version) => {
  try {
    return Office.context.requirements.isSetSupported(set, version);
  } catch (e) {
    return false;
  }
};

function sizeText(page) {
  const inches = (pt) => round(pt / 72, 2);
  const mm = (pt) => Math.round((pt / 72) * 25.4);
  return `${inches(page.width)} x ${inches(page.height)} in (${mm(page.width)} x ${mm(page.height)} mm)`;
}

function saveSettings(data) {
  return new Promise((resolve) => {
    try {
      if (data) Office.context.document.settings.set(SETTINGS_KEY, data);
      else Office.context.document.settings.remove(SETTINGS_KEY);
      Office.context.document.settings.saveAsync(() => resolve());
    } catch (e) {
      resolve();
    }
  });
}

function readSettings() {
  try {
    return Office.context.document.settings.get(SETTINGS_KEY) || null;
  } catch (e) {
    return null;
  }
}

/** A flat package holding one tiny paragraph: the anchor an undone page part leaves behind. */
const ANCHOR_PACKAGE =
  '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
  '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>' +
  `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${TINY_P}</w:body></w:document></pkg:xmlData></pkg:part>` +
  "</pkg:package>";

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Removes the content control tagged `tag` from a package read with getOoxml, keeping what it
 * holds. Redo puts that content back into the anchor control, so the wrapper must not come
 * along, or Word would nest a second control with the same tag.
 */
export function unwrapControl(xml, tag) {
  if (!xml) return "";
  const tagAt = xml.indexOf(`<w:tag w:val="${escapeXml(tag)}"/>`);
  if (tagAt < 0) return xml;
  const start = Math.max(xml.lastIndexOf("<w:sdt>", tagAt), xml.lastIndexOf("<w:sdt ", tagAt));
  if (start < 0) return xml;
  const re = /<w:sdt>|<w:sdt |<\/w:sdt>/g;
  re.lastIndex = start;
  let depth = 0;
  let end = -1;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0] === "</w:sdt>") {
      depth--;
      if (depth === 0) {
        end = m.index;
        break;
      }
    } else depth++;
  }
  if (end < 0) return xml;
  const open = "<w:sdtContent>";
  const contentStart = xml.indexOf(open, tagAt);
  const contentEnd = xml.lastIndexOf("</w:sdtContent>", end);
  if (contentStart < 0 || contentStart > end || contentEnd < contentStart) return xml;
  return xml.slice(0, start) + xml.slice(contentStart + open.length, contentEnd) + xml.slice(end + "</w:sdt>".length);
}

/** Copies the readable values of `keys` from an API object, skipping what this Word lacks. */
function pick(obj, keys) {
  const out = {};
  for (const k of keys.split(",")) {
    try {
      const v = obj[k];
      if (v !== undefined && v !== null && v !== "" && v !== "Mixed" && v !== "Unknown") out[k] = v;
    } catch (e) {
      // A property this version of Word does not have.
    }
  }
  return out;
}

/** Names in a sentence: "Quote", "Quote and Caption", or "Title, Quote, and Caption". */
function nameList(names) {
  if (names.length < 3) return names.join(" and ");
  return `${names.slice(0, -1).join(", ")}, and ${names[names.length - 1]}`;
}

let opCounter = 0;
const newId = () => `${Date.now().toString(36)}${(opCounter++).toString(36)}`;

export function createWordHost() {
  const api13 = supports("WordApi", "1.3");
  const api15 = supports("WordApi", "1.5");
  const desktop13 = supports("WordApiDesktop", "1.3");
  const isWeb = Office.context.platform === Office.PlatformType.OfficeOnline;
  const anchors = new Set(); // tags of page parts that are currently undone

  // -------------------------------------------------------------------------
  // Section helpers
  // -------------------------------------------------------------------------

  /** The document's sections, loaded (WordApi 1.1). */
  async function loadSections(context) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    return sections.items;
  }

  /**
   * Indexes, in context.document.sections, of the sections the selection touches. Page setup and
   * columns are recorded and restored by index, so Undo acts on the sections Offizin changed even
   * after the cursor has moved (Range.compareLocationWith and Body.getRange, WordApi 1.3).
   */
  async function selectedSectionIndexes(context) {
    const items = await loadSections(context);
    if (items.length <= 1) return [0];
    const selection = context.document.getSelection();
    const relations = items.map((s) => selection.compareLocationWith(s.body.getRange("Whole")));
    await context.sync();
    const touching = [];
    const adjacent = [];
    relations.forEach((r, i) => {
      const v = String(r.value || "");
      if (/^(Inside|Contains|Equal|Overlaps)/.test(v)) touching.push(i);
      else if (/^Adjacent/.test(v)) adjacent.push(i);
    });
    if (touching.length) return touching;
    return adjacent.length ? [adjacent[0]] : [0];
  }

  /** Every section's headers and footers as OOXML (Section.getHeader and getFooter, WordApi 1.1). */
  async function readHeadersFooters(context) {
    const items = await loadSections(context);
    const reads = items.map((s) => ({ headers: HF_TYPES.map((t) => s.getHeader(t).getOoxml()), footers: HF_TYPES.map((t) => s.getFooter(t).getOoxml()) }));
    await context.sync();
    return reads.map((r, index) => ({
      index,
      headers: Object.fromEntries(HF_TYPES.map((t, i) => [t, r.headers[i].value])),
      footers: Object.fromEntries(HF_TYPES.map((t, i) => [t, r.footers[i].value])),
    }));
  }

  async function writeHeadersFooters(context, saved) {
    const items = await loadSections(context);
    for (const entry of saved) {
      const s = items[entry.index];
      if (!s) continue;
      for (const t of HF_TYPES) {
        if (entry.headers[t]) s.getHeader(t).insertOoxml(entry.headers[t], "Replace");
        if (entry.footers[t]) s.getFooter(t).insertOoxml(entry.footers[t], "Replace");
      }
    }
    await context.sync();
  }

  /**
   * The document's paragraph styles (WordApi 1.5). Only the name, type, and built-in flag are
   * loaded for the whole collection. Character, table, and list styles have no paragraph
   * formatting, and one property Word rejects fails the whole sync, so callers load fonts and
   * paragraph formatting from the paragraph styles this returns, in a second step.
   */
  async function paragraphStyles(context) {
    const styles = context.document.getStyles();
    styles.load("items/nameLocal,items/type,items/builtIn");
    await context.sync();
    const paragraphType = Word.StyleType ? Word.StyleType.paragraph : "Paragraph";
    return styles.items.filter((st) => st.type === paragraphType);
  }

  /** The paragraph style definitions a publication's .docx can change (WordApi 1.5). */
  async function readStyles(context) {
    const items = await paragraphStyles(context);
    for (const st of items) {
      st.font.load(STYLE_FONT);
      st.paragraphFormat.load(STYLE_PARA);
    }
    await context.sync();
    return items.map((st) => ({ name: st.nameLocal, font: pick(st.font, STYLE_FONT), paragraph: pick(st.paragraphFormat, STYLE_PARA) }));
  }

  /**
   * Puts saved style definitions back, and deletes Offizin styles the publication added. `saved`
   * lists every paragraph style the document had before, so a built-in style that is not in it
   * (Quote or Caption, for example) came with the publication. Word cannot delete a built-in
   * style or reset it, so it keeps the publication's look; the names of those styles are returned.
   */
  async function writeStyles(context, saved) {
    const byName = new Map(saved.map((s) => [s.name, s]));
    const added = [];
    for (const st of await paragraphStyles(context)) {
      const def = byName.get(st.nameLocal);
      if (def) {
        for (const [k, v] of Object.entries(def.font)) st.font[k] = v;
        for (const [k, v] of Object.entries(def.paragraph)) st.paragraphFormat[k] = v;
      } else if (!st.builtIn && /^Offizin /.test(st.nameLocal || "")) {
        st.delete();
      } else if (st.builtIn && st.nameLocal) {
        added.push(st.nameLocal);
      }
    }
    await context.sync();
    return added;
  }

  // -------------------------------------------------------------------------
  // Page setup helpers (WordApiDesktop 1.3 only)
  // -------------------------------------------------------------------------

  /** Page setup of the sections at `indexes` (every section when null), each with its index. */
  async function readPageSetup(context, indexes) {
    const items = await loadSections(context);
    const picked = (indexes || items.map((s, i) => i)).filter((i) => items[i]);
    picked.forEach((i) => items[i].pageSetup.load(PAGE_FIELDS));
    await context.sync();
    return picked.map((i) => ({ index: i, values: pick(items[i].pageSetup, PAGE_FIELDS) }));
  }

  /**
   * Applies page size, margins, and per-section print settings to the sections at `indexes`
   * (every section when null). setup.raw restores values read earlier: [{ index, values }], each
   * written back to its own section.
   */
  async function writePageSetup(context, setup, indexes) {
    const items = await loadSections(context);
    const set = (ps, k, v) => {
      if (v === undefined || v === null) return;
      try {
        ps[k] = v;
      } catch (e) {
        // Not available in this version of Word.
      }
    };
    if (setup.raw) {
      for (const entry of setup.raw) {
        const s = items[entry.index];
        if (!s) continue;
        // Book fold goes last, after the sheet size is back.
        set(s.pageSetup, "bookFoldPrinting", false);
        for (const [k, v] of Object.entries(entry.values)) if (k !== "bookFoldPrinting") set(s.pageSetup, k, v);
        if (entry.values.bookFoldPrinting) set(s.pageSetup, "bookFoldPrinting", true);
      }
      await context.sync();
      return;
    }
    const targets = (indexes || items.map((s, i) => i)).filter((i) => items[i]);
    for (const i of targets) {
      const ps = items[i].pageSetup;
      const landscape = setup.width > setup.height;
      const PO = Word.PageOrientation;
      // A setup that does not ask for book fold or two pages per sheet turns them off.
      if (!setup.bookFoldPrinting) set(ps, "bookFoldPrinting", false);
      if (!setup.twoPagesOnOne) set(ps, "twoPagesOnOne", false);
      set(ps, "orientation", PO ? (landscape ? PO.landscape : PO.portrait) : landscape ? "Landscape" : "Portrait");
      set(ps, "pageWidth", setup.width);
      set(ps, "pageHeight", setup.height);
      set(ps, "topMargin", setup.margins.top);
      set(ps, "bottomMargin", setup.margins.bottom);
      set(ps, "leftMargin", setup.margins.left);
      set(ps, "rightMargin", setup.margins.right);
      set(ps, "gutter", setup.gutter);
      set(ps, "mirrorMargins", setup.mirrorMargins);
      if (setup.twoPagesOnOne) set(ps, "twoPagesOnOne", true);
      if (setup.bookFoldPrinting) set(ps, "bookFoldPrinting", true);
    }
    await context.sync();
  }

  /** The columns of the section at `index`: count, line between, and the widths when uneven. */
  async function readColumns(context, index) {
    const items = await loadSections(context);
    const s = items[index];
    if (!s) throw new Error("That section is no longer in the document.");
    const cols = s.pageSetup.textColumns;
    cols.load("items/width,items/spacing");
    const line = cols.getHasLineBetween();
    const even = cols.getIsEvenlySpaced();
    await context.sync();
    return { count: Math.max(1, cols.items.length), line: !!line.value, even: even.value !== false, widths: cols.items.map((c) => ({ width: c.width, spacing: c.spacing })) };
  }

  async function writeColumns(context, index, spec) {
    const items = await loadSections(context);
    const s = items[index];
    if (!s) throw new Error("That section is no longer in the document.");
    const cols = s.pageSetup.textColumns;
    cols.setCount(spec.count);
    cols.setIsEvenlySpaced(spec.even !== false);
    cols.setHasLineBetween(!!spec.line);
    await context.sync();
    if (spec.even === false && spec.widths && spec.widths.length === spec.count) {
      // Uneven columns: put each width and gap back.
      try {
        cols.load("items");
        await context.sync();
        cols.items.forEach((c, i) => {
          c.width = spec.widths[i].width;
          c.spacing = spec.widths[i].spacing;
        });
        await context.sync();
      } catch (e) {
        console.warn("Column widths could not be restored", e);
      }
    }
  }

  /** The page setup a model asks for, including booklet and print settings. */
  function pageSetupFor(model, bookFold) {
    const sheet = wordSheet(model, bookFold);
    const print = model.print || {};
    return {
      width: sheet.width,
      height: sheet.height,
      margins: model.page.margins,
      gutter: print.gutter || 0,
      mirrorMargins: !!print.mirrorMargins,
      twoPagesOnOne: !!print.twoPerSheet && !bookFold,
      bookFoldPrinting: !!bookFold,
    };
  }

  // -------------------------------------------------------------------------
  // Creating a publication
  // -------------------------------------------------------------------------

  /**
   * What Undo of a publication needs: the body, every section's page setup (Windows and Mac),
   * headers and footers, and the style definitions (WordApi 1.5), because Create imports the
   * publication's footer and styles along with its content. The style list names every paragraph
   * style present before, which is how Undo tells which built-in styles the publication added.
   */
  async function snapshotDocument() {
    const ooxml = await Word.run(async (context) => {
      const o = context.document.body.getOoxml();
      await context.sync();
      return o.value;
    });
    const attempt = async (label, fn) => {
      try {
        return await Word.run(fn);
      } catch (e) {
        console.warn(label + " could not be read", e);
        return null;
      }
    };
    const pageSetup = desktop13 ? await attempt("Page setup", (context) => readPageSetup(context, null)) : null;
    const headers = await attempt("Headers and footers", (context) => readHeadersFooters(context));
    const styles = api15 ? await attempt("Styles", (context) => readStyles(context)) : null;
    return { ooxml, pageSetup, headers, styles, record: readSettings() };
  }

  /** Puts a .docx into the open document, trying the best method this Word supports first. */
  async function insertDocx(base64, model, images) {
    const attempt = async (label, fn) => {
      try {
        await Word.run(async (context) => {
          fn(context);
          await context.sync();
        });
        return true;
      } catch (e) {
        console.warn(label + " failed", e);
        return false;
      }
    };
    if (api15 && (await attempt("Document.insertFileFromBase64", (c) => c.document.insertFileFromBase64(base64, "Replace", { importTheme: false, importStyles: true, importParagraphSpacing: true, importPageColor: false })))) return "document";
    if (await attempt("Body.insertFileFromBase64", (c) => c.document.body.insertFileFromBase64(base64, "Replace"))) return "body";
    // Last resort: content only. Page size must then be set by hand.
    const width = model.page.width - model.page.margins.left - model.page.margins.right;
    await Word.run(async (context) => {
      context.document.body.insertOoxml(flatPackage(model.blocks, width, images, { modelStyles: model.styles }), "Replace");
      await context.sync();
    });
    return "ooxml";
  }

  async function applyCreate(data) {
    const how = await insertDocx(data.base64, data.model, data.images);
    if (desktop13) {
      try {
        await Word.run((context) => writePageSetup(context, data.setup, null));
      } catch (e) {
        console.warn("Page setup could not be applied", e);
      }
    }
    await saveSettings(data.record);
    console.info(`Offizin built ${data.model.templateId} using ${how}`);
    return how;
  }

  /**
   * Book fold can only be turned on through the page setup API (WordApiDesktop 1.3: Microsoft 365
   * Word 2507 or later on Windows, 16.99.2 or later on Mac). Elsewhere the file opened as a new
   * document keeps book fold, or the person turns it on by hand in Word for Windows or Mac.
   */
  function createMessage(model, bookFold) {
    if (model.booklet && !bookFold) {
      const newDoc = api13 ? "create the booklet with Open in a new document, which keeps book fold in the file, or " : "";
      const byHand = isWeb ? `open this document in Word for Windows or Mac and ${BY_HAND}` : BY_HAND;
      return `${model.name} created page by page. Offizin turns on book fold only in Microsoft 365 Word version 2507 or later on Windows and 16.99.2 or later on Mac. Elsewhere, ${newDoc}${byHand}.`;
    }
    if (desktop13) return `${model.name} created.`;
    return `${model.name} created. The page should be ${sizeText(model.page)}. If it is not, set it in Layout, then Size.`;
  }

  /**
   * Switches colors and fonts. Paragraph and cell formatting is remapped through the body's
   * Office Open XML (Word keeps content controls, tables, and pictures), and paragraph styles
   * are remapped through the styles API where Word has it (WordApi 1.5).
   */
  async function applyRestyle(map) {
    await Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      await context.sync();
      // One pass per kind, so a swap such as Georgia to Arial and Arial to Georgia cannot chain.
      let xml = ooxml.value;
      const colors = {};
      for (const [from, to] of Object.entries(map.colors || {})) colors[from.replace("#", "").toUpperCase()] = to.replace("#", "").toUpperCase();
      if (Object.keys(colors).length) {
        xml = xml.replace(new RegExp(`"(${Object.keys(colors).join("|")})"`, "gi"), (m, hexValue) => `"${colors[hexValue.toUpperCase()]}"`);
      }
      const fonts = map.fonts || {};
      if (Object.keys(fonts).length) {
        const names = Object.keys(fonts).map(escapeRegExp).join("|");
        xml = xml.replace(new RegExp(`(w:(?:ascii|hAnsi|eastAsia|cs)=")(${names})(")`, "g"), (m, a, name, b) => `${a}${fonts[name]}${b}`);
      }
      body.insertOoxml(xml, "Replace");
      await context.sync();
    });
    if (!api15) return;
    try {
      await Word.run(async (context) => {
        const items = await paragraphStyles(context);
        for (const st of items) st.font.load("color,name");
        await context.sync();
        const colors = {};
        for (const [k, v] of Object.entries(map.colors || {})) colors[k.toUpperCase()] = v;
        for (const st of items) {
          const c = String(st.font.color || "").toUpperCase();
          if (colors[c]) st.font.color = colors[c];
          if (map.fonts && map.fonts[st.font.name]) st.font.name = map.fonts[st.font.name];
        }
        await context.sync();
      });
    } catch (e) {
      console.warn("Styles could not be restyled", e);
    }
  }

  // -------------------------------------------------------------------------
  // The host
  // -------------------------------------------------------------------------

  const createModes = [
    {
      id: "replace",
      label: "Replace this document",
      note: "Replaces everything in this document. Offizin's Undo brings it back.",
    },
  ];
  if (api13) {
    createModes.push({
      id: "new",
      label: "Open in a new document",
      note: isWeb ? "Word on the web opens new documents in a new tab. If nothing opens, allow pop-ups for this site." : "Opens in a new Word window and leaves this document as it is.",
    });
  }

  const host = {
    kind: "word",
    platform: "word",
    label: "for Word",
    capabilities: {
      createModes,
      pageSetup: desktop13 ? true : "manual",
      pageSetupHelp: desktop13
        ? ""
        : "Publications from the Create tab bring their own page size. To change it yourself, use the Layout tab in Word: Size, Orientation, and Margins. (This version of Word does not let add-ins change page size.)",
      columns: desktop13,
      columnsHelp: desktop13 ? "" : "To flow text in columns, use Layout, then Columns, in Word.",
      restyle: true,
      history: true,
      historyNote:
        "Offizin's Undo reverses only what Offizin did, and leaves your own typing and Word's Undo alone. Undoing an inserted page part keeps what you typed in it for Redo. Undoing a new publication brings back the document's content, headers, footers, and styles as they were before (and its page setup in Word for Windows and Mac), so changes made to the publication since then are lost. The history lasts while this panel is open.",
      finish: true,
      // Replace: the page setup API keeps these on Windows and Mac. New document: the file keeps them.
      printSettings: {
        replace: desktop13 ? ["mirrorMargins", "twoPerSheet"] : [],
        new: isWeb ? ["mirrorMargins", "embedFonts"] : ["mirrorMargins", "twoPerSheet", "embedFonts"],
      },
      printGuide: { editor: "word", web: isWeb, desktopPageSetup: desktop13 },
      pdfHelp: "Choose File, then Save As (Windows) or Export (Mac), and pick PDF. In Word on the web, choose File, then Export, then Download as PDF.",
    },

    loadState() {
      try {
        const raw = localStorage.getItem(STORE_KEY);
        return raw ? JSON.parse(raw) : null;
      } catch (e) {
        return null;
      }
    },

    saveState(state) {
      localStorage.setItem(STORE_KEY, JSON.stringify(state));
    },

    async createPublication(model, mode) {
      const images = prepareImages(model.blocks);
      const record = {
        templateId: model.templateId,
        theme: model.theme,
        samples: model.samples.slice(0, 80),
        page: model.page,
        print: model.print,
        createdAt: new Date().toISOString(),
      };
      // The file itself carries the template and theme as custom properties, so a document opened
      // from it can still be restyled and checked (getPublication reads them when settings are empty).
      const fileRecord = { templateId: model.templateId, theme: model.theme };

      if (mode === "new") {
        // A new document is built from the file itself, so it keeps book fold and other settings.
        const base64 = bytesToBase64(docxBytes(model, images, { bookFold: !!model.booklet, record: fileRecord }));
        await Word.run(async (context) => {
          context.application.createDocument(base64).open();
          await context.sync();
        });
        return { message: `${model.name} opened in a new document.` };
      }

      // Replacing this document: book fold only where Word lets Offizin turn it on afterward.
      const bookFold = !!model.booklet && desktop13;
      const base64 = bytesToBase64(docxBytes(model, images, { bookFold, record: fileRecord }));
      const before = await snapshotDocument();
      const data = { id: newId(), base64, model, images, setup: pageSetupFor(model, bookFold), record, before };
      await applyCreate(data);
      return {
        message: createMessage(model, bookFold),
        op: { kind: "create", label: `Create ${model.name.toLowerCase()}`, data },
      };
    },

    async insertPart(composed, width, partId, ctx = {}) {
      const name = ctx.label || "page part";
      const images = prepareImages(composed.blocks);
      const tag = TAG_PREFIX + newId();
      const pkg = flatPackage(composed.blocks, width, images, { tag, alias: `Offizin: ${name}` });
      const inner = flatPackage(composed.blocks, width, images);
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        if (api13) {
          const para = selection.paragraphs.getLast();
          para.getRange("Whole").insertOoxml(pkg, "After");
        } else {
          selection.insertOoxml(pkg, "After");
        }
        await context.sync();
        // The package asks for the hidden appearance; set it through the API as well (WordApi 1.1),
        // so no frame or title tab shows when the person clicks inside the part.
        try {
          const ccs = context.document.contentControls.getByTag(tag);
          ccs.load("items");
          await context.sync();
          for (const cc of ccs.items) cc.appearance = "Hidden";
          await context.sync();
        } catch (e) {
          console.warn("The part's marker could not be hidden", e);
        }
      });
      return {
        message: `${name.charAt(0).toUpperCase()}${name.slice(1)} inserted.`,
        op: { kind: "insert", label: `Insert ${name.toLowerCase()}`, data: { tag, inner, outer: pkg } },
      };
    },

    async contentWidth() {
      if (desktop13) {
        try {
          return await Word.run(async (context) => {
            const s = context.document.getSelection().sections.getFirst();
            s.pageSetup.load("pageWidth,pageHeight,leftMargin,rightMargin,bookFoldPrinting,twoPagesOnOne");
            await context.sync();
            const ps = s.pageSetup;
            // Book fold always splits the sheet into two side-by-side pages. Two pages per sheet
            // splits it side by side only on a landscape sheet; on a portrait sheet the pages are
            // full width and half height.
            const halves = ps.bookFoldPrinting || (ps.twoPagesOnOne && ps.pageWidth > ps.pageHeight);
            const page = halves ? ps.pageWidth / 2 : ps.pageWidth;
            return page - ps.leftMargin - ps.rightMargin;
          });
        } catch (e) {
          // Fall through to the saved publication.
        }
      }
      const pub = readSettings();
      if (pub && pub.page) return pub.page.width - pub.page.margins.left - pub.page.margins.right;
      return 468;
    },

    async applyPageSetup(setup) {
      if (!desktop13) throw new Error("This version of Word does not let add-ins change page size. Use Layout, then Size.");
      // The sections are remembered by index, so Undo restores each one's own values wherever
      // the cursor is by then.
      const indexes = await Word.run((context) => selectedSectionIndexes(context));
      const before = await Word.run((context) => readPageSetup(context, indexes));
      await Word.run((context) => writePageSetup(context, setup, indexes));
      return { message: "Page setup applied.", op: { kind: "pageSetup", label: "Change page setup", data: { before, after: setup, indexes } } };
    },

    async applyColumns({ count, line }) {
      if (!desktop13) throw new Error("This version of Word does not let add-ins set columns. Use Layout, then Columns.");
      const index = (await Word.run((context) => selectedSectionIndexes(context)))[0];
      const before = await Word.run((context) => readColumns(context, index));
      const after = { count, line: !!line, even: true };
      await Word.run((context) => writeColumns(context, index, after));
      return {
        message: `${count} column${count > 1 ? "s" : ""} applied.`,
        op: { kind: "columns", label: `Set ${count} column${count > 1 ? "s" : ""}`, data: { before, after, index } },
      };
    },

    /**
     * The publication record from the document settings, or, for a file made with Open in a new
     * document, the short record (template and theme) its custom properties carry (WordApi 1.3).
     * That record says mode "new": the document was built from the file itself, so it keeps the
     * file's own print settings. The panel composes the samples itself when a record has none.
     */
    async getPublication() {
      const saved = readSettings();
      if (saved) return saved;
      if (!api13) return null;
      try {
        return await Word.run(async (context) => {
          const props = context.document.properties.customProperties;
          props.load("items/key,items/value");
          await context.sync();
          const record = publicationFromProperties(props.items);
          return record ? { ...record, mode: "new" } : null;
        });
      } catch (e) {
        return null;
      }
    },

    async restyle(map, theme) {
      const current = (await host.getPublication()) || {};
      const previous = current.theme || null;
      await applyRestyle(map);
      await saveSettings({ ...current, theme });
      return { message: "Design applied. Pictures keep their colors.", op: { kind: "restyle", label: "Apply design", data: { map, theme, previous } } };
    },

    async undo(op) {
      const d = op.data;
      switch (op.kind) {
        case "create": {
          await Word.run(async (context) => {
            context.document.body.insertOoxml(d.before.ooxml, "Replace");
            await context.sync();
          });
          const notes = [];
          const restore = async (label, saved, fn, note) => {
            if (!saved) return null;
            try {
              return await Word.run((context) => fn(context, saved));
            } catch (e) {
              console.warn(label + " could not be restored", e);
              notes.push(note);
              return null;
            }
          };
          if (desktop13) await restore("Page setup", d.before.pageSetup, (c, s) => writePageSetup(c, { raw: s }), "The page setup could not be put back.");
          await restore("Headers and footers", d.before.headers, writeHeadersFooters, "The headers and footers could not be put back.");
          if (!d.before.headers) notes.push("The headers and footers could not be read before the publication, so the publication's footer stays.");
          const stylesNote = "so Normal, Title, and the headings keep the publication's look.";
          const added = api15 ? await restore("Styles", d.before.styles, writeStyles, `The style definitions could not be put back, ${stylesNote}`) : null;
          if (!d.before.styles) notes.push(api15 ? `The style definitions could not be read before the publication, ${stylesNote}` : `This version of Word does not let add-ins put back style definitions, ${stylesNote}`);
          if (added && added.length) notes.push(`The publication added the ${nameList(added)} style${added.length > 1 ? "s" : ""}, which keep${added.length > 1 ? "" : "s"} its look.`);
          await saveSettings(d.before.record);
          const message = desktop13
            ? "The document is back to how it was before the publication."
            : "The document's content is back. This version of Word keeps the publication's page size, so set the size again in Layout, then Size, if you need to.";
          return [message, ...notes].join(" ");
        }
        case "insert": {
          // Redo puts back what the part holds now, so the person's own typing in it survives.
          let inner = d.inner;
          let readFailed = false;
          await Word.run(async (context) => {
            const ccs = context.document.contentControls.getByTag(d.tag);
            ccs.load("items");
            await context.sync();
            if (!ccs.items.length) throw new Error("That page part is no longer in the document.");
            const cc = ccs.items[0];
            try {
              const current = api13 ? cc.getRange("Content").getOoxml() : cc.getOoxml();
              await context.sync();
              const unwrapped = unwrapControl(current.value, d.tag);
              if (unwrapped) inner = unwrapped;
              else readFailed = true;
            } catch (e) {
              console.warn("The part's current content could not be read; Redo will use what Offizin last had", e);
              readFailed = true;
            }
            cc.insertOoxml(ANCHOR_PACKAGE, "Replace");
            await context.sync();
          });
          anchors.add(d.tag);
          // The typing cannot be saved for Redo when Word would not let Offizin read the part.
          const lost = readFailed ? " What you typed in it could not be read, so Redo will put back the part as it was before." : "";
          return { message: `Removed: ${op.label.replace(/^Insert /, "")}.${lost}`, op: { ...op, data: { ...d, inner } } };
        }
        case "restyle":
          await applyRestyle(invertRestyleMap(d.map));
          if (d.previous) await saveSettings({ ...(readSettings() || {}), theme: d.previous });
          return "The previous design is back.";
        case "pageSetup":
          if (!Array.isArray(d.before)) throw new Error("The previous page setup is not known.");
          await Word.run((context) => writePageSetup(context, { raw: d.before }));
          return "Page setup restored.";
        case "columns":
          if (!d.before || typeof d.index !== "number") throw new Error("The previous columns are not known.");
          await Word.run((context) => writeColumns(context, d.index, d.before));
          return "Columns restored.";
        default:
          throw new Error("Offizin cannot undo that change.");
      }
    },

    async redo(op) {
      const d = op.data;
      switch (op.kind) {
        case "create":
          d.before = await snapshotDocument();
          await applyCreate(d);
          return { message: createMessage(d.model, d.setup.bookFoldPrinting), op };
        case "insert": {
          let atEnd = false;
          await Word.run(async (context) => {
            const ccs = context.document.contentControls.getByTag(d.tag);
            ccs.load("items");
            await context.sync();
            if (ccs.items.length) {
              ccs.items[0].insertOoxml(d.inner, "Replace");
            } else if (d.outer) {
              // Undo and Redo of the publication rebuilt the document, so the anchor is gone.
              context.document.body.insertOoxml(d.outer, "End");
              atEnd = true;
            } else {
              throw new Error("Offizin cannot find where that page part was.");
            }
            await context.sync();
          });
          anchors.delete(d.tag);
          const name = op.label.replace(/^Insert /, "");
          if (atEnd) return `Offizin could not find where the ${name} was, so it added it at the end of the document. Drag it where you want it.`;
          return `${op.label.replace(/^Insert /, "Inserted ")} again.`;
        }
        case "restyle":
          await applyRestyle(d.map);
          await saveSettings({ ...(readSettings() || {}), theme: d.theme });
          return "Design applied again.";
        case "pageSetup":
          await Word.run((context) => writePageSetup(context, d.after, d.indexes || [0]));
          return "Page setup applied again.";
        case "columns":
          await Word.run((context) => writeColumns(context, d.index || 0, d.after));
          return "Columns applied again.";
        default:
          throw new Error("Offizin cannot redo that change.");
      }
    },

    /** Ops leaving the history for good: remove the anchors left by undone page parts. */
    async discard(ops) {
      const tags = (ops || []).filter((op) => op && op.kind === "insert" && op.data && anchors.has(op.data.tag)).map((op) => op.data.tag);
      if (!tags.length) return;
      try {
        await Word.run(async (context) => {
          for (const tag of tags) {
            const ccs = context.document.contentControls.getByTag(tag);
            ccs.load("items");
            await context.sync();
            ccs.items.forEach((cc) => cc.delete(false));
          }
          await context.sync();
        });
      } catch (e) {
        console.warn("Anchors could not be removed", e);
      }
      tags.forEach((t) => anchors.delete(t));
    },

    /** Removes Offizin's invisible markers. Content stays; anchors of undone parts go. */
    async finish() {
      const count = await Word.run(async (context) => {
        const ccs = context.document.contentControls;
        ccs.load("items/tag");
        await context.sync();
        let n = 0;
        for (const cc of ccs.items) {
          if (!cc.tag || !cc.tag.startsWith(TAG_PREFIX)) continue;
          cc.delete(!anchors.has(cc.tag));
          n++;
        }
        await context.sync();
        return n;
      });
      anchors.clear();
      return count ? `Finished. Offizin removed ${count} invisible marker${count > 1 ? "s" : ""}; your content is unchanged.` : "Finished. There were no Offizin markers to remove.";
    },

    /** Searches the document for every sample it is given, in batches inside one Word.run. */
    async check({ samples = [], infoSamples = [] }) {
      return Word.run(async (context) => {
        const body = context.document.body;
        const findings = [];
        const unique = [...new Set(samples.filter((t) => typeof t === "string" && t.trim()))];
        const BATCH = 60;
        for (let i = 0; i < unique.length; i += BATCH) {
          const batch = unique.slice(i, i + BATCH).map((text) => ({ text, r: body.search(text, { matchCase: true }) }));
          batch.forEach((s) => s.r.load("items"));
          await context.sync();
          for (const s of batch) if (s.r.items.length) findings.push({ kind: "sample", text: s.text, search: s.text });
        }
        const infoSearches = infoSamples.map((f) => ({ f, r: body.search(f.sample, { matchCase: true }) }));
        infoSearches.forEach((s) => s.r.load("items"));
        const pics = body.inlinePictures;
        pics.load("items/altTextDescription");
        await context.sync();
        const placeholders = pics.items.filter((p) => p.altTextDescription === PLACEHOLDER_ALT).length;
        if (placeholders) findings.push({ kind: "picture", text: `${placeholders} picture placeholder${placeholders > 1 ? "s" : ""}. Right-click one and choose Change Picture.`, picture: true });
        for (const s of infoSearches) if (s.r.items.length) findings.push({ kind: "info", text: `${s.f.label} still shows the sample “${s.f.sample}”.`, search: s.f.sample });
        return findings;
      });
    },

    async select(f) {
      await Word.run(async (context) => {
        const body = context.document.body;
        if (f.picture) {
          const pics = body.inlinePictures;
          pics.load("items/altTextDescription");
          await context.sync();
          const hit = pics.items.find((p) => p.altTextDescription === PLACEHOLDER_ALT);
          if (!hit) throw new Error("No picture placeholders are left.");
          hit.select();
        } else {
          const r = body.search(f.search || f.text, { matchCase: true });
          r.load("items");
          await context.sync();
          if (!r.items.length) throw new Error("That text is no longer in the document.");
          r.items[0].select();
        }
        await context.sync();
      });
    },
  };

  return host;
}
