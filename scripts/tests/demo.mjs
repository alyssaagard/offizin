/**
 * Tests for the demo host (host contract version 2 with in-memory history snapshots) and the
 * studio's DOM-free logic: zoom steps, fitting, ruler ticks and bands, printer's marks, labels,
 * keyboard shortcuts, and the page markup the studio expects.
 *
 * Run with: node scripts/check.mjs --only=demo
 */
import fs from "node:fs";
import path from "node:path";

export async function run({ ROOT, fail, ok, load }) {
  const expect = (cond, msg) => (cond ? ok(msg) : fail(msg));
  const near = (a, b, eps = 0.01) => Math.abs(a - b) <= eps;

  const measure = await load("docs/app/studio/measure.js");
  const keys = await load("docs/app/studio/keys.js");
  const { createDemoHost } = await load("docs/app/hosts/demo.js");
  const { compose } = await load("docs/app/core/compose.js");
  const { findTemplate, PARTS, TEMPLATES } = await load("docs/app/core/templates.js");
  const { createHistory, normalizeResult } = await load("docs/app/core/history.js");

  // -------------------------------------------------------------------------
  // Zoom
  // -------------------------------------------------------------------------
  {
    const { stepZoom, fitScale, normalizeZoom, zoomPercent, ZOOM_STEPS, ZOOM_CHOICES, PX_PER_PT } = measure;
    expect(JSON.stringify(ZOOM_STEPS) === "[0.5,0.75,1,1.25,1.5]" && ZOOM_CHOICES[0] === "fit", "zoom steps are fit, 50, 75, 100, 125, and 150 percent");
    expect(stepZoom(1, 1) === 1.25 && stepZoom(1, -1) === 0.75, "zoom in and out from 100 percent");
    expect(stepZoom(0.63, 1) === 0.75 && stepZoom(0.63, -1) === 0.5, "zoom steps from a fitted scale go to the nearest step");
    expect(stepZoom(1.5, 1) === 1.5 && stepZoom(0.5, -1) === 0.5, "zoom stops at 150 and 50 percent");
    expect(normalizeZoom("fit") === "fit" && normalizeZoom("0.75") === 0.75 && normalizeZoom(9) === 1.5 && normalizeZoom("x") === "fit", "zoom values are normalized");
    expect(zoomPercent(0.754) === "75%", "zoom percent text");
    const s = fitScale({ availWidth: 800, availHeight: 600, sheetWidth: 612, sheetHeight: 792 });
    expect(near(s, 600 / (792 * PX_PER_PT), 0.0001), `fit shows the whole Letter page on a desktop pasteboard (${s})`);
    const phone = fitScale({ availWidth: 340, availHeight: 200, sheetWidth: 612, sheetHeight: 792 });
    expect(near(phone, 340 / (612 * PX_PER_PT), 0.0001), "fit uses the width on a short, narrow screen");
    expect(fitScale({ availWidth: 5000, availHeight: 5000, sheetWidth: 612, sheetHeight: 792 }) === 1.5, "fit never goes above 150 percent");
  }

  // -------------------------------------------------------------------------
  // Units, paper, and rulers
  // -------------------------------------------------------------------------
  {
    const { unitFor, matchPaper, paperName, sizeText, orientationName, rulerScale, rulerTicks, rulerBands, rulerSvg, PX_PER_PT } = measure;
    expect(unitFor({ sizeName: "letter", width: 612, height: 792 }) === "in", "Letter rulers use inches");
    expect(unitFor({ sizeName: "a4", width: 595.28, height: 841.89 }) === "mm", "A4 rulers use millimeters");
    expect(unitFor({ sizeName: null, width: 300, height: 400 }, "a4") === "mm" && unitFor({ sizeName: null, width: 300, height: 400 }, "letter") === "in", "custom sizes follow the chosen paper");
    expect(matchPaper(841.89, 595.28) === "a4" && matchPaper(792, 612) === "letter" && matchPaper(100, 100) === null, "paper sizes are recognized in either orientation");
    expect(paperName({ sizeName: "letter", width: 612, height: 792 }) === "Letter", "paper name");
    expect(sizeText({ width: 612, height: 792 }, "in") === "8.5 x 11 in", `Letter size text (${sizeText({ width: 612, height: 792 }, "in")})`);
    expect(sizeText({ width: 595.28, height: 841.89 }, "mm") === "210 x 297 mm", "A4 size text");
    expect(orientationName({ width: 792, height: 612 }) === "Landscape", "orientation name");

    const ppp = 1 * PX_PER_PT;
    const sc = rulerScale("in", ppp);
    expect(sc.divisions === 8 && sc.labelEvery === 1, "at 100 percent an inch is divided into eighths");
    expect(rulerScale("in", 0.5 * PX_PER_PT).divisions === 4, "at 50 percent an inch is divided into quarters");
    const ticks = rulerTicks({ length: 300, origin: 30, pxPerPt: ppp, unit: "in" });
    const at = (pos) => ticks.find((t) => Math.abs(t.pos - pos) < 0.02);
    expect(at(30) && at(30).level === 0 && at(30).label === "0", "the ruler's zero sits at the page edge");
    expect(at(126) && at(126).label === "1", "one inch is 96 pixels at 100 percent");
    expect(at(78) && at(78).level === 1 && at(42) && at(42).level === 2, "half inch and eighth inch ticks");
    expect(ticks.every((t) => t.pos >= 0 && t.pos <= 300), "ticks stay on the ruler");
    const left = rulerTicks({ length: 300, origin: 200, pxPerPt: ppp, unit: "in" });
    expect(left.some((t) => Math.abs(t.pos - 104) < 0.02 && t.label === "1"), "labels count up on both sides of zero");
    const scrolled = rulerTicks({ length: 300, origin: -500, pxPerPt: ppp, unit: "in" });
    expect(scrolled.length > 0 && scrolled.some((t) => t.label === "6"), "ticks follow the page when it scrolls past the ruler");

    const mmPx = (72 / 25.4) * PX_PER_PT;
    const mmTicks = rulerTicks({ length: 400, origin: 0, pxPerPt: PX_PER_PT, unit: "mm" });
    expect(mmTicks.some((t) => t.label === "10" && Math.abs(t.pos - 10 * mmPx) < 0.02), "millimeter rulers number every 10 mm");
    expect(rulerScale("mm", 1.5 * PX_PER_PT).divisions === 10, "at 150 percent millimeter rulers show every millimeter");
    const zoomed = rulerTicks({ length: 400, origin: 0, pxPerPt: 1.5 * PX_PER_PT, unit: "in" });
    expect(zoomed.some((t) => t.label === "2" && Math.abs(t.pos - 288) < 0.02), "ticks follow zoom (2 in at 150 percent is 288 pixels)");

    const page = { width: 612, height: 792, margins: { top: 36, right: 54, bottom: 72, left: 36 } };
    const bx = rulerBands({ axis: "x", origin: 10, pxPerPt: ppp, page });
    expect(near(bx.sheet.from, 10) && near(bx.sheet.to, 10 + 816) && near(bx.bands[0].from, 10 + 48) && near(bx.bands[0].to, 10 + 558 * ppp), "horizontal ruler marks the left and right margins");
    const by = rulerBands({ axis: "y", origin: 0, pxPerPt: ppp, page });
    expect(near(by.bands[0].from, 48) && near(by.bands[0].to, 720 * ppp), "vertical ruler marks the top and bottom margins");
    const spread = rulerBands({ axis: "x", origin: 0, pxPerPt: ppp, page, sheetPt: 1224 });
    expect(spread.bands.length === 2 && near(spread.bands[1].from, (612 + 36) * ppp), "a printer spread shows margins for both pages");
    const svg = rulerSvg({ orientation: "h", length: 300, ticks, sheet: bx.sheet, bands: bx.bands });
    expect(svg.startsWith("<svg") && svg.includes('class="rl-tick"') && svg.includes(">1</text>") && svg.includes("rl-marker") && !/NaN|undefined/.test(svg), "horizontal ruler SVG");
    const vsvg = rulerSvg({ orientation: "v", length: 500, ticks: rulerTicks({ length: 500, origin: 0, pxPerPt: ppp, unit: "mm" }), sheet: by.sheet, bands: by.bands });
    expect(vsvg.includes("rotate(-90)") && !/NaN|undefined/.test(vsvg), "vertical ruler SVG");
  }

  // -------------------------------------------------------------------------
  // Labels, marks, and other helpers
  // -------------------------------------------------------------------------
  {
    const { overflowText, fileSlug, studioFontsHref, marksSvg, currentSheetIndex, sheetWord } = measure;
    expect(overflowText([false, false]) === "Text fits", "overflow text when everything fits");
    expect(overflowText([false, true], "Side") === "Text overflows on side 2", "overflow text for one side");
    expect(overflowText([true, true, true], "Page") === "Text overflows on pages 1, 2 and 3", "overflow text for several pages");
    expect(fileSlug("  Spring Café News!  ") === "spring-cafe-news" && fileSlug("") === "offizin-publication", "file names from titles");
    const { googleFontsHref } = await load("docs/app/core/schemes.js");
    const href = studioFontsHref(googleFontsHref());
    expect(href.includes("family=IBM+Plex+Sans") && href.includes("family=IBM+Plex+Mono") && href.includes("family=Lato") && href.endsWith("&display=swap") && href.split("display=swap").length === 2, "one Google Fonts request for IBM Plex and the template fonts");
    const marks = marksSvg(400, 300, 30);
    expect(marks.includes('width="460"') && marks.includes("mk-crop") && (marks.match(/class="mk-reg"/g) || []).length === 4 && !/NaN/.test(marks), "crop marks at the corners and four registration targets");
    expect(currentSheetIndex([{ top: -900, bottom: -100 }, { top: -50, bottom: 700 }, { top: 720, bottom: 1500 }], { top: 0, bottom: 800 }) === 1, "the current page is the one that fills most of the view");
    expect(currentSheetIndex([], { top: 0, bottom: 10 }) === -1, "no current page without pages");
    expect(sheetWord({ folds: { vertical: [0.5] }, page: { width: 612, height: 792 } }) === "Side" && sheetWord({ folds: null, page: { width: 612, height: 792 } }) === "Page" && sheetWord(null, true) === "Spread", "sheet words");

    // studio-4: the document title through Undo (empty page) and Redo (the same publication).
    const { titleAfterRender, UNTITLED } = measure;
    const news = { templateId: "newsletter", name: "Newsletter" };
    let t = titleAfterRender({ lastTemplate: null, titleEdited: false, title: UNTITLED }, news);
    expect(t.reset && t.lastTemplate === "newsletter" && !t.titleEdited && t.title === "Newsletter", "a new publication takes its template's name");
    t = titleAfterRender({ ...t, titleEdited: true, title: "Spring news" }, null);
    expect(!t.reset && t.title === "Spring news" && t.lastTemplate === "newsletter", "Undo to an empty page keeps the typed title and remembers the publication");
    t = titleAfterRender(t, news);
    expect(!t.reset && t.titleEdited && t.title === "Spring news", "Redo of the same publication keeps the typed title");
    t = titleAfterRender({ lastTemplate: "newsletter", titleEdited: false, title: "Newsletter" }, null);
    expect(t.title === UNTITLED, "an empty page with no typed title says Untitled publication");
    t = titleAfterRender(t, news);
    expect(!t.reset && t.title === "Newsletter", "Redo brings the template's name back when nothing was typed");
    t = titleAfterRender({ lastTemplate: "newsletter", titleEdited: true, title: "Spring news" }, { templateId: "flyer", name: "Flyer" });
    expect(t.reset && t.title === "Flyer" && !t.titleEdited, "a different publication gets its own title, like a new document");
  }

  // -------------------------------------------------------------------------
  // Keyboard
  // -------------------------------------------------------------------------
  {
    const { shortcutFor, isEditableTarget, inShortcutScope, escapeAction, historyView, SHORTCUT_SCOPE } = keys;
    const body = { tagName: "BODY", closest: () => null };
    // Elements inside the pasteboard or the page navigator (closest finds the scope) and outside it.
    const inScope = (tag, extra = {}) => {
      const t = { tagName: tag, ...extra };
      t.closest = (sel) => (sel === SHORTCUT_SCOPE ? t : null);
      return t;
    };
    const outside = (tag, extra = {}) => ({ tagName: tag, closest: () => null, ...extra });
    const board = inScope("DIV");
    const thumb = inScope("BUTTON");
    expect(shortcutFor({ key: "z", ctrlKey: true, target: body }) === "undo", "Ctrl+Z is Offizin undo");
    expect(shortcutFor({ key: "z", metaKey: true, target: body }) === "undo", "Command+Z is Offizin undo");
    expect(shortcutFor({ key: "Z", metaKey: true, shiftKey: true, target: body }) === "redo", "Command+Shift+Z is Offizin redo");
    expect(shortcutFor({ key: "y", ctrlKey: true, target: body }) === "redo", "Ctrl+Y is Offizin redo");
    expect(shortcutFor({ key: "z", ctrlKey: true, target: outside("BUTTON") }) === "undo", "Ctrl+Z works on panel buttons too (modifier shortcuts are page-wide)");
    expect(shortcutFor({ key: "z", ctrlKey: true, target: { tagName: "INPUT", type: "text" } }) === null, "Ctrl+Z in a text field undoes typing instead");
    expect(shortcutFor({ key: "+", target: board }) === "zoom-in" && shortcutFor({ key: "=", target: board }) === "zoom-in", "plus zooms in on the pasteboard");
    expect(shortcutFor({ key: "-", target: board }) === "zoom-out" && shortcutFor({ key: "0", target: board }) === "zoom-fit", "minus zooms out and 0 fits on the pasteboard");
    expect(shortcutFor({ key: "g", target: board }) === "guides" && shortcutFor({ key: "P", target: thumb }) === "print-preview", "G toggles guides and P print preview on the pasteboard and the navigator");
    expect(SHORTCUT_SCOPE.includes("#st-board") && SHORTCUT_SCOPE.includes("#st-nav-list") && inShortcutScope(board) && !inShortcutScope(body) && !inShortcutScope(null), "the single-key scope is the pasteboard and the page navigator");
    // studio-2: single keys never act on panel controls, buttons elsewhere, or the body (WCAG 2.1.4).
    const panelCheckbox = outside("INPUT", { type: "checkbox" });
    const panelRange = outside("INPUT", { type: "range" });
    expect(shortcutFor({ key: "g", target: panelCheckbox }) === null && shortcutFor({ key: "p", target: panelCheckbox }) === null, "G and P on a panel checkbox are left alone");
    expect(shortcutFor({ key: "-", target: panelRange }) === null && shortcutFor({ key: "+", target: outside("BUTTON") }) === null && shortcutFor({ key: "0", target: body }) === null, "zoom keys outside the pasteboard are left alone");
    expect(shortcutFor({ key: "g", target: { tagName: "TEXTAREA" } }) === null && shortcutFor({ key: "p", target: { isContentEditable: true } }) === null, "letters typed in text fields are left alone");
    expect(shortcutFor({ key: "+", ctrlKey: true, target: board }) === null && shortcutFor({ key: "0", metaKey: true, target: board }) === null, "browser zoom shortcuts are left alone");
    expect(shortcutFor({ key: "g", altKey: true, target: board }) === null, "Alt combinations are left alone");
    expect(shortcutFor({ key: "g", target: board, defaultPrevented: true }) === null, "a key another control already handled is left alone");
    expect(isEditableTarget({ tagName: "INPUT", type: "checkbox" }) === false && isEditableTarget({ tagName: "SELECT" }) === true, "editable targets");
    // studio-1: one Escape closes one thing. The panel's confirm and history list prevent the default.
    const esc = (extra = {}) => ({ key: "Escape", ...extra });
    expect(escapeAction(esc(), { sheetOpen: true, phone: true }) === "sheet", "Escape closes the panel sheet on a phone");
    expect(escapeAction(esc({ defaultPrevented: true }), { sheetOpen: true, phone: true }) === null, "Escape the panel already used (history list or confirm) leaves the sheet open");
    expect(escapeAction(esc(), { menuOpen: true, sheetOpen: true, phone: true }) === "menu" && escapeAction(esc({ defaultPrevented: true }), { menuOpen: true }) === null, "Escape closes the install menu first, unless it was already handled");
    expect(escapeAction(esc(), { sheetOpen: true, phone: false }) === null && escapeAction({ key: "g" }, { sheetOpen: true, phone: true }) === null, "Escape does nothing to a docked panel, and other keys are not Escape");
    expect(historyView({ undo() {}, redo() {} }).supported === false, "history buttons hide when the panel has no onHistoryChange");
    expect(historyView({}).supported === false && historyView(null).supported === false, "history buttons hide without the panel API");
    const fakePanel = { undo() {}, redo() {}, onHistoryChange() {} };
    const hv = historyView(fakePanel, { canUndo: true, canRedo: false, undoLabel: "Create tri-fold brochure", busy: false });
    expect(hv.supported && hv.canUndo && !hv.canRedo && hv.undoTitle === "Undo: Create tri-fold brochure" && hv.redoTitle === "Nothing to redo", "history button state from a snapshot");
    expect(historyView(fakePanel, { canUndo: true, canRedo: true, busy: true }).canUndo === false, "history buttons wait while the history is busy");
    const viaHistory = historyView({ ...fakePanel, history: { snapshot: () => ({ canUndo: false, canRedo: true, redoLabel: "Insert banner" }) } });
    expect(viaHistory.canRedo && viaHistory.redoTitle === "Redo: Insert banner", "history button state read from panel.history");
  }

  // -------------------------------------------------------------------------
  // Demo host: contract version 2 and history snapshots
  // -------------------------------------------------------------------------
  {
    const renders = [];
    let empties = 0;
    const canvas = { innerHTML: "", clientWidth: 900, clientHeight: 700, scrollTop: 0, querySelectorAll: () => [] };
    const fakeRender = (container, model, opts) => {
      renders.push({ model, opts });
      container.innerHTML = "drawn";
      const pages = model.blocks.filter((b) => b.t === "break").length + 1;
      return { scale: opts.scale, overflow: new Array(pages).fill(false) };
    };
    const memory = new Map();
    const storage = { getItem: (k) => (memory.has(k) ? memory.get(k) : null), setItem: (k, v) => memory.set(k, String(v)) };
    const host = createDemoHost({
      canvas,
      render: fakeRender,
      storage,
      measure: () => ({ width: 800, height: 600 }),
      onRender: (m) => {
        if (!m) empties++;
      },
    });
    const caps = host.capabilities;
    expect(caps.history === true && typeof caps.historyNote === "string" && caps.historyNote.length > 40, "capabilities: history with an honest note");
    expect(caps.finish === true && caps.printGuide && caps.printGuide.editor === "web", "capabilities: finish and the web print guide");
    expect(typeof host.finish() === "string" && host.finish().length > 20, "finish returns a message");
    host.saveState({ a: 1 });
    expect(host.loadState().a === 1, "settings are saved and loaded");

    const theme = { paper: "letter" };
    host.themeChanged(theme);
    expect(host.currentModel() === null && empties === 1, "an empty page before any publication");

    const trifold = findTemplate("trifold");
    const r1 = normalizeResult(host.createPublication(compose(trifold, theme, "web"), "show", { options: {} }));
    expect(r1.op && r1.op.kind === "create" && r1.op.label === "Create tri-fold brochure" && r1.message.includes("Tri-fold"), "createPublication returns { message, op }");
    expect(JSON.parse(JSON.stringify(r1.op)).data.snapshot === r1.op.data.snapshot, "ops are plain JSON");
    expect(host.currentModel().templateId === "trifold", "the tri-fold is on the page");
    const g1 = host.geometry();
    expect(g1 && g1.unit === "in" && g1.page.width === 792 && g1.folds === true && near(g1.pxPerPt, g1.scale * measure.PX_PER_PT, 1e-9), "geometry for the rulers (inches, landscape tri-fold)");

    const history = createHistory(host, { limit: 3 });
    history.push(r1.op);

    const second = TEMPLATES.find((t) => t.id === "newsletter") || TEMPLATES.find((t) => t.id !== "trifold");
    const r2 = normalizeResult(host.createPublication(compose(second, theme, "web"), "show", { options: {} }));
    history.push(r2.op);
    const blocksBefore = host.currentModel().blocks.length;
    const part = PARTS[0];
    const r3 = normalizeResult(host.insertPart(null, host.contentWidth(), part.id));
    history.push(r3.op);
    expect(r3.op.kind === "insert" && r3.op.label === `Insert ${part.name.toLowerCase()}`, "insertPart returns an insert op");
    const blocksAfter = host.currentModel().blocks.length;
    expect(blocksAfter > blocksBefore, `the part is added (${blocksBefore} to ${blocksAfter} blocks)`);

    const a4Landscape = { width: 841.89, height: 595.28, orientation: "landscape", margins: { top: 28.35, bottom: 28.35, left: 28.35, right: 28.35 } };
    const r4 = normalizeResult(host.applyPageSetup(a4Landscape));
    expect(r4.op && r4.op.kind === "pageSetup", "applyPageSetup returns a pageSetup op");
    const pageBefore = host.currentModel().page;
    expect(pageBefore.width === 841.89 && pageBefore.sizeName === "a4" && host.geometry().unit === "mm", "page setup applies and the rulers switch to millimeters");
    expect(host.wordModel().page.width === 841.89 && host.wordModel().blocks.length > 0, "the Word file uses the page setup and the inserted part");
    const snapsBefore = host.snapshotCount();
    history.push(r4.op);
    expect(history.size === 3 && host.snapshotCount() === snapsBefore - 1, "trimmed history entries free their snapshots (discard)");

    await history.undo();
    expect(host.currentModel().page.width !== 841.89 && host.currentModel().blocks.length === blocksAfter, "undo page setup restores the previous page");
    await history.undo();
    expect(host.currentModel().blocks.length === blocksBefore && host.currentModel().templateId === second.id, "undo insert removes exactly the part");
    await history.redo();
    expect(host.currentModel().blocks.length === blocksAfter, "redo insert puts the part back");
    await history.redo();
    expect(host.currentModel().page.width === 841.89, "redo page setup applies it again");
    await history.undo();
    await history.undo();
    await history.undo();
    expect(host.currentModel().templateId === "trifold", "undo create restores the publication before it");
    expect(!history.canUndo, "the oldest entry was trimmed, so there is nothing more to undo");

    // A fresh history from an empty page: undoing the first create empties the page.
    const h2 = createDemoHost({ canvas, render: fakeRender, storage, measure: () => ({ width: 800, height: 600 }), onRender: (m) => !m && empties++ });
    h2.themeChanged(theme);
    const before = empties;
    const op = normalizeResult(h2.createPublication(compose(trifold, theme, "web"), "show", {})).op;
    const undone = normalizeResult(await h2.undo(op));
    expect(h2.currentModel() === null && empties === before + 1 && undone.message.startsWith("Undone"), "undoing the first create leaves an empty page");
    const redone = normalizeResult(await h2.redo(op));
    expect(h2.currentModel().templateId === "trifold" && redone.op === op, "redo create brings it back");
    h2.discard([op]);
    let threw = false;
    try {
      await h2.undo(op);
    } catch (e) {
      threw = /no longer/.test(e.message);
    }
    expect(threw && h2.snapshotCount() === 0, "a discarded op can no longer be undone");
    let threwInsert = false;
    try {
      createDemoHost({ canvas, render: fakeRender }).insertPart(null, 468, part.id);
    } catch (e) {
      threwInsert = true;
    }
    expect(threwInsert, "inserting before creating a publication explains what to do");

    // View options reach the renderer.
    renders.length = 0;
    host.setView({ zoom: "0.75" });
    expect(renders.at(-1).opts.scale === 0.75 && host.geometry().scale === 0.75, "zoom reaches the renderer");
    host.setView({ zoom: "fit" });
    const fitted = renders.at(-1).opts.scale;
    const pg = host.currentModel().page;
    expect(near(fitted, measure.fitScale({ availWidth: 800, availHeight: 600, sheetWidth: pg.width, sheetHeight: pg.height }), 1e-6), "fit zoom uses the available space");
    host.setView({ margins: false, folds: false });
    expect(renders.at(-1).opts.guides === false, "guides off when both guide toggles are off");
    host.setView({ guides: true });
    expect(renders.at(-1).opts.guides === true && host.getView().margins && host.getView().folds, "the guides shortcut turns both guide kinds on");
    host.setView({ printPreview: true });
    const pp = renders.at(-1).opts;
    expect(pp.printPreview === true && pp.guides === false && pp.badges === false, "print preview hides guides and badges");
    host.setView({ printPreview: false, spreads: true });
    expect(renders.at(-1).opts.spreads === false, "spreads only apply to booklets");
    host.setView({ spreads: false });

    const findings = host.check({});
    expect(Array.isArray(findings) && findings.some((f) => f.kind === "sample"), "the design checker finds sample text");
    host.setTitle("Spring Open House");
    expect(host.getTitle() === "Spring Open House", "the document title is kept for the download");
  }

  // -------------------------------------------------------------------------
  // The page markup matches the studio code
  // -------------------------------------------------------------------------
  {
    const html = fs.readFileSync(path.join(ROOT, "docs/demo/index.html"), "utf8");
    const js = fs.readFileSync(path.join(ROOT, "docs/app/studio/studio.js"), "utf8");
    const ids = [...js.matchAll(/\$\("([a-z0-9-]+)"\)/g)].map((m) => m[1]);
    const missing = ids.filter((id) => !html.includes(`id="${id}"`));
    expect(ids.length > 20 && !missing.length, `every element the studio uses is on the page${missing.length ? ": missing " + missing.join(", ") : ""}`);
    expect(html.includes("../assets/brand/offizin-mark.svg") && html.includes('href="../word/"') && html.includes('href="../google-docs/"') && html.includes("../assets/studio.css"), "the page has the mark, install links, and studio styles");
    const css = fs.readFileSync(path.join(ROOT, "docs/assets/studio.css"), "utf8");
    expect(css.includes("IBM Plex Sans") && css.includes("IBM Plex Mono") && css.includes("radial-gradient") && css.includes("max-width: 760px"), "studio styles: Plex type, dot grid, and a phone layout");
    // a11y-2: the page has one visually hidden h1, and the hidden style keeps it in the accessibility tree.
    expect((html.match(/<h1\b/g) || []).length === 1 && /<h1 class="st-sr">Offizin live demo<\/h1>/.test(html) && /\.st-sr\s*\{[^}]*clip: rect\(0 0 0 0\)/.test(css) && !/\.st-sr\s*\{[^}]*display: none/.test(css), "the demo page has a visually hidden h1");
    // studio-3: a live region outside the panel sheet carries app bar messages on phones.
    const aside = html.slice(html.indexOf("<aside"), html.indexOf("</aside>"));
    expect(/<div class="st-toast" id="st-toast" role="status" aria-live="polite">/.test(html) && !aside.includes("st-toast") && js.includes("showToast(text, kind, sticky)"), "a status toast outside the sheet, which say() uses on phones");
    // studio-3: Undo and Redo results, which the panel writes on its own status line, reach the toast too.
    const runHistorySrc = js.slice(js.indexOf("async function runHistory("), js.indexOf('el.undo.addEventListener("click"'));
    expect(/await \(direction < 0 \? panel\.undo\(\) : panel\.redo\(\)\);[\s\S]*querySelector\("\.oz-status"\)[\s\S]*\btoast\(text, /.test(runHistorySrc), "runHistory repeats the panel's status line in the toast once an undo or redo settles");
    expect(/\.st-toast\s*\{[^}]*opacity: 0/.test(css) && !/\.st-toast\s*\{[^}]*(display: none|visibility: hidden)/.test(css) && css.includes(".st-toast.is-shown"), "the toast hides with opacity, so its live region stays in the accessibility tree");
    expect(/aria-label="Publication pages\.[^"]*plus and minus zoom[^"]*"/.test(html), "the pasteboard's label tells keyboard users which keys work while it has focus");
    // panel-6: the file buttons (a label around an invisible file input) show keyboard focus.
    const panelCss = fs.readFileSync(path.join(ROOT, "docs/app/ui/panel.css"), "utf8");
    expect(/\.oz-file:focus-within\s*\{[^}]*outline: 2px solid var\(--oz-focus\)/.test(panelCss), "Choose a picture, Open brand kit, and Open a .csv file show keyboard focus");
    const tryPage = fs.readFileSync(path.join(ROOT, "docs/try/index.html"), "utf8");
    expect(tryPage.includes('location.replace("../demo/" + location.hash)'), "the old /try/ address still redirects to the demo");
    const own = ["docs/demo/index.html", "docs/assets/studio.css", "docs/app/ui/panel.css", "docs/app/hosts/demo.js", "docs/app/studio/studio.js", "docs/app/studio/measure.js", "docs/app/studio/keys.js", "scripts/tests/demo.mjs"];
    const bad = own.filter((f) => new RegExp(String.fromCharCode(0x2014)).test(fs.readFileSync(path.join(ROOT, f), "utf8")));
    expect(!bad.length, `no em dashes in the studio files${bad.length ? ": " + bad.join(", ") : ""}`);
  }

  // -------------------------------------------------------------------------
  // The studio in a document: the demo page with studio.js bundled from source, driven in
  // jsdom at phone width. Skipped, with a note, when jsdom is not installed.
  // -------------------------------------------------------------------------
  await runStudioDomTests({ ROOT, ok, expect });
}

/** Finds jsdom: the repository's own node_modules, OFFIZIN_JSDOM, or a sibling folder that has it. */
async function findJsdom(ROOT) {
  const { createRequire } = await import("node:module");
  const candidates = [path.join(ROOT, "package.json"), process.env.OFFIZIN_JSDOM && path.join(process.env.OFFIZIN_JSDOM, "package.json"), path.join(ROOT, "..", "logotool", "package.json")].filter(Boolean);
  for (const c of candidates) {
    try {
      return createRequire(c)("jsdom");
    } catch (e) {
      // Try the next place.
    }
  }
  return null;
}

async function runStudioDomTests({ ROOT, ok, expect }) {
  const jsdom = await findJsdom(ROOT);
  if (!jsdom) {
    ok("studio in a document: skipped, jsdom not found (npm install jsdom, or set OFFIZIN_JSDOM to a folder that has it)");
    return;
  }
  const os = await import("node:os");
  const { pathToFileURL } = await import("node:url");
  const { bundle } = await import(pathToFileURL(path.join(ROOT, "scripts/build.mjs")).href);

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "offizin-studio-"));
  const entry = path.join(tmp, "entry.js");
  fs.writeFileSync(entry, `import { mountStudio } from "${path.join(ROOT, "docs/app/studio/studio.js")}";\nwindow.__studio = mountStudio(document);\n`);
  const js = bundle(entry);
  fs.rmSync(tmp, { recursive: true, force: true });

  const html = fs.readFileSync(path.join(ROOT, "docs/demo/index.html"), "utf8").replace(/<script type="module">[\s\S]*?<\/script>/, "");
  const { JSDOM, VirtualConsole } = jsdom;
  // The panel logs the two refused changes below on purpose; a quiet console keeps the run readable.
  const dom = new JSDOM(html, { url: "https://offizin.test/demo/#template=newsletter", runScripts: "outside-only", pretendToBeVisual: true, virtualConsole: new VirtualConsole() });
  const w = dom.window;
  w.HTMLCanvasElement.prototype.getContext = () => null;
  const media = { phone: true };
  w.matchMedia = (q) => ({ matches: media.phone && /max-width: 760px/.test(q), media: q, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} });
  w.eval(js);

  const $ = (s) => w.document.querySelector(s);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const until = async (test) => {
    for (let i = 0; i < 100 && !test(); i++) await sleep(20);
    return test();
  };
  const click = (el) => el.dispatchEvent(new w.MouseEvent("click", { bubbles: true, cancelable: true }));
  const key = (el, k, o = {}) => el.dispatchEvent(new w.KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...o }));
  const toast = $("#st-toast");
  const status = () => $("#st-panel .oz-status");
  /** Runs one undo or redo and waits for the panel's status line to show the result. */
  const settle = async (fn, word) => {
    click(toast); // Tapping the toast clears it, so an old message cannot pass for a new one.
    fn();
    await until(() => status().textContent.includes(word));
    await sleep(20);
    return { shown: toast.classList.contains("is-shown"), text: toast.textContent, kind: toast.dataset.kind, status: status().textContent };
  };
  /** Passes when the toast shows the same text and kind as the panel's status line; a failure prints what it saw. */
  const mirrored = (r, text, kind, msg) => {
    const pass = r.shown && r.status === text && r.text === text && r.kind === kind;
    expect(pass, pass ? msg : `${msg} (saw ${JSON.stringify(r)})`);
  };

  try {
    const ready = await until(() => !$("#st-undo").disabled);
    expect(ready && !$("#studio").classList.contains("st-sheet-open") && !!status().closest("#st-side"), "studio on a phone: the first publication is drawn, and the status line sits in the closed sheet");

    // studio-3: Undo and Redo from the app bar and from the keyboard show their result in the toast.
    let r = await settle(() => click($("#st-undo")), "Undone");
    mirrored(r, "Undone: Create newsletter.", "success", "studio on a phone: Undo shows its result in the toast");
    r = await settle(() => click($("#st-redo")), "Redone");
    mirrored(r, "Redone: Create newsletter.", "success", "studio on a phone: Redo shows its result in the toast");
    const board = $("#st-board");
    r = await settle(() => key(board, "z", { ctrlKey: true }), "Undone");
    mirrored(r, "Undone: Create newsletter.", "success", "studio on a phone: Ctrl+Z shows its result in the toast");
    r = await settle(() => key(board, "z", { ctrlKey: true, shiftKey: true }), "Redone");
    mirrored(r, "Redone: Create newsletter.", "success", "studio on a phone: Ctrl+Shift+Z shows its result in the toast");

    // A change the host refuses: the panel's error line is repeated too.
    const host = w.__studio.host;
    const realUndo = host.undo;
    host.undo = () => {
      throw new Error("The demo could not put the last page back.");
    };
    r = await settle(() => click($("#st-undo")), "could not put");
    mirrored(r, "The demo could not put the last page back.", "error", "studio on a phone: an undo that fails shows the reason in the toast");
    host.undo = realUndo;
    const realRedo = host.redo;
    r = await settle(() => click($("#st-undo")), "Undone");
    host.redo = () => Promise.reject(new Error("The demo could not build the page again."));
    r = await settle(() => click($("#st-redo")), "could not build");
    mirrored(r, "The demo could not build the page again.", "error", "studio on a phone: a redo that fails shows the reason in the toast");
    host.redo = realRedo;

    // With the sheet open, or on a wide screen, the status line is in view and the toast stays away.
    click($("#st-panel-toggle"));
    r = await settle(() => click($("#st-redo")), "Redone");
    expect($("#studio").classList.contains("st-sheet-open") && !r.shown && r.text === "" && r.status === "Redone: Create newsletter.", "studio on a phone: with the sheet open the status line carries the result, not the toast");
    click($("#st-panel-close"));
    media.phone = false;
    r = await settle(() => click($("#st-undo")), "Undone");
    expect(!r.shown && r.text === "" && r.status === "Undone: Create newsletter.", "studio on a wide screen: Undo never shows the toast");
  } finally {
    w.close();
  }
}
