/**
 * The website demo host. It draws publications on a simulated page instead of in a document,
 * so people can try Offizin before installing it, and download the result as a Word file
 * (which Google Docs also opens).
 *
 * Host contract version 2 (every method may return a promise):
 *   kind, platform, label, capabilities
 *   loadState(), saveState(state)
 *   createPublication(model, mode, { options })  -> { message, op }
 *   insertPart(composed, width, partId)           -> { message, op }
 *   applyPageSetup(setup)                         -> { message, op }
 *   applyColumns({ count, line }), restyle(map, theme)
 *   undo(op), redo(op)                            -> { message, op? }
 *   discard(ops), finish()
 *   getPublication(), contentWidth(), check({ samples, infoSamples }), select(finding)
 *   themeChanged(theme)   optional, called whenever colors, fonts, or information change
 *
 * History: the demo keeps its own document state (template, options, parts, and page setup)
 * in memory. Each change Offizin makes stores a snapshot of that state before and after the
 * change, and the op only carries the snapshot's id. Undo and redo put a snapshot back and
 * redraw. Colors, fonts, and business information are not part of the history: the demo shows
 * them live as you change them in the panel.
 *
 * Demo helpers used by the studio page: setView, getView, geometry, redraw, currentModel,
 * setTitle, getTitle, downloadDocx, wordModel, snapshotCount.
 */
import { compose, composePart, BUSINESS_FIELDS } from "../core/compose.js";
import { findTemplate, findPart } from "../core/templates.js";
import { renderPreview, highlightText } from "../render/preview.js";
import { docxBytes } from "../render/ooxml.js";
import { prepareImages } from "../render/pictures.js";
import { PX_PER_PT, fitScale, matchPaper, unitFor, fileSlug, normalizeZoom } from "../studio/measure.js";

const STORE_KEY = "offizin.demo.state";

/** The demo's own document state, as stored in history snapshots. */
function emptyDoc() {
  return { current: null, parts: [], pageOverride: null };
}

function cloneDoc(doc) {
  return JSON.parse(JSON.stringify(doc));
}

/**
 * Creates the demo host.
 * canvas    the element the pages are drawn into.
 * viewport  the scrolling element around the pages (defaults to canvas), used to fit the zoom.
 * onRender  called after every draw with (model, result), or (null, null) when the page is empty.
 * render    the renderer (renderPreview by default; tests pass a stand in).
 * measure   optional function returning { width, height } of the space available for a page.
 * storage   optional localStorage stand in.
 */
export function createDemoHost({ canvas, viewport, onRender, render = renderPreview, measure, storage } = {}) {
  let theme = null;
  let doc = emptyDoc();
  let model = null;
  let overflow = [];
  let lastResult = null;
  let title = "";
  let view = { zoom: "fit", margins: true, folds: true, printPreview: false, spreads: false };
  const snapshots = new Map();
  let serial = 0;
  const store = storage || (typeof localStorage !== "undefined" ? localStorage : null);

  // -------------------------------------------------------------------------
  // Model
  // -------------------------------------------------------------------------

  /** Composes the current document for a platform, with page setup and inserted parts. */
  function composeDoc(platform) {
    if (!doc.current || !theme) return null;
    const t = findTemplate(doc.current.templateId);
    if (!t) return null;
    const m = compose(t, theme, platform, doc.current.options);
    if (doc.pageOverride) {
      const page = { ...m.page, ...doc.pageOverride };
      page.margins = { ...m.page.margins, ...(doc.pageOverride.margins || {}) };
      page.sizeName = matchPaper(page.width, page.height);
      m.page = page;
    }
    const width = m.page.width - m.page.margins.left - m.page.margins.right;
    for (const id of doc.parts) {
      const part = findPart(id);
      if (!part) continue;
      const composed = composePart(part, theme, platform, width);
      m.blocks.push(...composed.blocks);
      m.samples.push(...composed.samples);
      if (Array.isArray(m.pages) && m.pages.length) m.pages[m.pages.length - 1] = [...m.pages[m.pages.length - 1], ...composed.blocks];
    }
    return m;
  }

  function rebuild() {
    model = composeDoc("web");
  }

  function contentWidth() {
    if (!model) return 468;
    return model.page.width - model.page.margins.left - model.page.margins.right;
  }

  /** The size in points of one drawn sheet: a page, or a printer spread for booklets. */
  function sheetSize() {
    if (!model) return null;
    if (view.spreads && model.booklet && model.booklet.sheet) return { ...model.booklet.sheet };
    return { width: model.page.width, height: model.page.height };
  }

  function available() {
    if (measure) return measure();
    const box = viewport || canvas;
    return { width: Math.max(240, (box.clientWidth || 0) - 48), height: Math.max(240, (box.clientHeight || 0) - 48) };
  }

  function currentScale() {
    if (!model) return 1;
    if (view.zoom !== "fit") return Number(view.zoom) || 1;
    const sheet = sheetSize();
    const room = available();
    return fitScale({ availWidth: room.width, availHeight: room.height, sheetWidth: sheet.width, sheetHeight: sheet.height });
  }

  function draw() {
    if (!model) {
      canvas.innerHTML = "";
      overflow = [];
      lastResult = null;
      if (onRender) onRender(null, null);
      return;
    }
    const scale = currentScale();
    const showGuides = !view.printPreview && (view.margins || view.folds);
    const result = render(canvas, model, {
      scale,
      guides: showGuides,
      labels: true,
      gap: view.printPreview ? 64 : 28,
      checkOverflow: true,
      badges: !view.printPreview,
      spreads: !!(view.spreads && model.booklet),
      printPreview: !!view.printPreview,
    }) || { scale, overflow: [] };
    lastResult = { scale: result.scale ?? scale, overflow: result.overflow || [] };
    overflow = lastResult.overflow;
    if (onRender) onRender(model, lastResult);
  }

  // -------------------------------------------------------------------------
  // History snapshots
  // -------------------------------------------------------------------------

  function makeOp(kind, label, before) {
    const id = "demo-" + ++serial;
    snapshots.set(id, { before, after: cloneDoc(doc) });
    return { kind, label, data: { snapshot: id } };
  }

  function snapshotFor(op) {
    const id = op && op.data && op.data.snapshot;
    const snap = id ? snapshots.get(id) : null;
    if (!snap) throw new Error("This change is no longer in the demo's memory, so it cannot be undone or redone.");
    return snap;
  }

  function restore(state) {
    doc = cloneDoc(state);
    rebuild();
    draw();
  }

  // -------------------------------------------------------------------------
  // The host
  // -------------------------------------------------------------------------

  const host = {
    kind: "demo",
    platform: "web",
    label: "Live demo",
    capabilities: {
      createModes: [{ id: "show", label: "Show on the page" }],
      createLabel: "Show on the page",
      createNote: "In Word or Google Docs, this button builds the publication in your document. Here it appears on the page, and you can download it as a Word file.",
      pageSetup: true,
      columns: false,
      columnsHelp: "In Word for Windows and Mac, Offizin can also set text columns for the section you are working in.",
      restyle: false,
      insertHelp: "Click a part to add it to the end of the publication. In Word and Google Docs, parts go after the cursor.",
      pdfHelp: "In Word, choose File, then Save As (or Export), and pick PDF. In Google Docs, choose File, then Download, then PDF document. In this demo, download the Word file first.",
      history: true,
      historyNote: "In this demo, Offizin undo and redo put the page back the way it was before or after each Offizin change: creating a publication, adding a part, or changing the page setup. Colors, fonts, and business information update live, so they are not part of the history. The history lasts until you leave or reload the page.",
      finish: true,
      // The demo keeps every state, so undoing a publication loses nothing and needs no warning.
      safeHistory: true,
      printSettings: ["mirrorMargins", "twoPerSheet", "embedFonts"],
      printSettingsNote: "In this demo, these settings go into the Word file you download. In Word, Offizin saves them in the document.",
      printGuide: { editor: "web" },
    },

    loadState() {
      try {
        const raw = store ? store.getItem(STORE_KEY) : null;
        return raw ? JSON.parse(raw) : null;
      } catch (e) {
        return null;
      }
    },
    saveState(state) {
      try {
        if (store) store.setItem(STORE_KEY, JSON.stringify(state));
      } catch (e) {
        // Private browsing or full storage: the demo still works for this visit.
      }
    },

    themeChanged(next) {
      theme = next;
      rebuild();
      draw();
    },

    createPublication(m, mode, { options } = {}) {
      if (!findTemplate(m.templateId)) throw new Error("Offizin does not know this publication.");
      const before = cloneDoc(doc);
      doc = { current: { templateId: m.templateId, options: JSON.parse(JSON.stringify(options || {})) }, parts: [], pageOverride: null };
      if (!theme) theme = {};
      rebuild();
      draw();
      if (viewport) viewport.scrollTop = 0;
      else canvas.scrollTop = 0;
      return {
        message: `${m.name} is on the page. Change colors, fonts, or business information and watch it update.`,
        op: makeOp("create", `Create ${m.name.toLowerCase()}`, before),
      };
    },

    insertPart(composed, width, partId) {
      if (!doc.current) throw new Error("Create a publication first, then add parts to it.");
      const part = findPart(partId);
      if (!part) throw new Error("Offizin does not know this page part.");
      const before = cloneDoc(doc);
      doc.parts.push(partId);
      rebuild();
      draw();
      const sheets = canvas.querySelectorAll ? canvas.querySelectorAll(".ozp-sheet") : [];
      const last = sheets[sheets.length - 1];
      if (last && last.scrollIntoView) last.scrollIntoView({ block: "end", behavior: "smooth" });
      return { message: "Added to the end of the publication.", op: makeOp("insert", `Insert ${part.name.toLowerCase()}`, before) };
    },

    applyPageSetup(setup) {
      if (!doc.current) throw new Error("Create a publication first.");
      const before = cloneDoc(doc);
      doc.pageOverride = JSON.parse(JSON.stringify(setup || {}));
      rebuild();
      draw();
      return { message: "Page setup applied.", op: makeOp("pageSetup", "Change page setup", before) };
    },

    applyColumns() {
      throw new Error("Columns are available in Word for Windows and Mac.");
    },

    restyle() {
      return { message: "The demo page already follows the colors and fonts you choose.", op: null };
    },

    undo(op) {
      const snap = snapshotFor(op);
      restore(snap.before);
      return { message: `Undone: ${op.label || "Offizin change"}.` };
    },

    redo(op) {
      const snap = snapshotFor(op);
      restore(snap.after);
      return { message: `Redone: ${op.label || "Offizin change"}.`, op };
    },

    discard(ops) {
      for (const op of ops || []) {
        const id = op && op.data && op.data.snapshot;
        if (id) snapshots.delete(id);
      }
    },

    finish() {
      return "The demo page has no hidden Offizin markers to remove. In Word and Google Docs, Finish removes them so the document is plain and ready to share.";
    },

    getPublication() {
      return doc.current && model ? { templateId: doc.current.templateId, theme: model.theme, samples: model.samples } : null;
    },

    contentWidth,

    check() {
      if (!model) throw new Error("Create a publication first.");
      const findings = model.samples.map((text) => ({ kind: "sample", text }));
      let pictures = 0;
      const walk = (blocks) =>
        blocks.forEach((b) => {
          if (b.t === "image" && !b.src) pictures++;
          if (b.t === "grid") b.rows.forEach((r) => r.cells.forEach((c) => walk(c.blocks)));
          if (b.t === "columns") walk(b.blocks);
        });
      walk(model.blocks);
      if (pictures) findings.push({ kind: "picture", text: `${pictures} picture placeholder${pictures > 1 ? "s" : ""} to replace with your own pictures`, search: "picture" });
      for (const key of model.missingInfo) {
        const f = BUSINESS_FIELDS.find((x) => x.key === key);
        if (f) findings.push({ kind: "info", text: `${f.label} still shows the sample “${f.sample}”. Fill it in on the Info tab.`, search: f.sample });
      }
      overflow.forEach((over, i) => over && findings.push({ kind: "layout", text: `Text overflows side ${i + 1}. Shorten the text or make it smaller.`, selectable: false }));
      return findings;
    },

    select(f) {
      if (!highlightText(canvas, f.search || f.text)) throw new Error("That item is no longer on the page.");
    },

    // -----------------------------------------------------------------------
    // Demo helpers used by the studio page around the panel.
    // -----------------------------------------------------------------------

    /** Changes the view: { zoom: "fit" | number, margins, folds, guides, printPreview, spreads }. */
    setView(next = {}) {
      const v = { ...view, ...next };
      if ("guides" in next) {
        v.margins = !!next.guides;
        v.folds = !!next.guides;
        delete v.guides;
      }
      v.zoom = normalizeZoom(v.zoom);
      view = v;
      draw();
    },
    getView: () => ({ ...view }),

    /** What the rulers and status bar need to know about the drawn pages. */
    geometry() {
      if (!model) return null;
      const scale = lastResult ? lastResult.scale : currentScale();
      return {
        page: { ...model.page, margins: { ...model.page.margins } },
        sheet: sheetSize(),
        scale,
        pxPerPt: scale * PX_PER_PT,
        unit: unitFor(model.page, theme && theme.paper),
        overflow: overflow.slice(),
        booklet: !!model.booklet,
        spreads: !!(view.spreads && model.booklet),
        spreadCount: model.booklet && Array.isArray(model.booklet.spreads) ? model.booklet.spreads.length : 0,
        folds: !!model.folds,
        zoom: view.zoom,
      };
    },

    redraw: draw,
    currentModel: () => model,
    setTitle(next) {
      title = String(next || "").trim();
    },
    getTitle: () => title || (model ? model.name : ""),
    snapshotCount: () => snapshots.size,

    /** The publication composed for Word, as downloadDocx writes it. */
    wordModel() {
      if (!model) throw new Error("Create a publication first.");
      return composeDoc("word");
    },

    downloadDocx(name) {
      const wordModel = host.wordModel();
      const bytes = docxBytes(wordModel, prepareImages(wordModel.blocks));
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${fileSlug(name || title || wordModel.name)}.docx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        a.remove();
      }, 1500);
      return a.download;
    },
  };
  return host;
}
