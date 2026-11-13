/**
 * The Offizin studio: the live demo page dressed as a desktop publishing program.
 *
 * App bar (title, Offizin undo and redo, zoom, guides, print preview, download, install),
 * page navigator with live thumbnails, pasteboard with rulers, the Offizin panel docked on
 * the right (a sheet on phones), and a status bar. The page markup lives in docs/demo/index.html;
 * this module finds its elements by id and wires them to the demo host and the panel.
 *
 * The math (zoom, rulers, marks, labels) lives in measure.js and the shortcuts in keys.js, so
 * the tests can run them in Node.
 */
import { googleFontsHref } from "../core/schemes.js";
import { createDemoHost } from "../hosts/demo.js";
import { mountPanel } from "../ui/panel.js";
import {
  ZOOM_STEPS,
  stepZoom,
  zoomPercent,
  rulerTicks,
  rulerBands,
  rulerSvg,
  paperName,
  sizeText,
  orientationName,
  sheetWord,
  overflowText,
  currentSheetIndex,
  studioFontsHref,
  marksSvg,
  titleAfterRender,
  UNTITLED,
} from "./measure.js";
import { shortcutFor, escapeAction, historyView } from "./keys.js";

const PHONE = "(max-width: 760px)";
const MARK_BLEED = 30;
const RULER_SLACK = 240;

/**
 * Mounts the studio on a page that has the markup from docs/demo/index.html.
 * options.mountPanel replaces the panel's mount function (for test harnesses).
 */
export function mountStudio(doc = document, options = {}) {
  const win = doc.defaultView || window;
  const $ = (id) => doc.getElementById(id);
  const el = {
    app: $("studio"),
    fonts: $("st-fonts"),
    title: $("st-title"),
    undo: $("st-undo"),
    redo: $("st-redo"),
    history: $("st-history"),
    zoom: $("st-zoom"),
    zoomIn: $("st-zoom-in"),
    zoomOut: $("st-zoom-out"),
    margins: $("st-margins"),
    folds: $("st-folds"),
    spreads: $("st-spreads"),
    print: $("st-print"),
    download: $("st-download"),
    install: $("st-install"),
    installMenu: $("st-install-menu"),
    sheetToggle: $("st-panel-toggle"),
    sheetClose: $("st-panel-close"),
    side: $("st-side"),
    panel: $("st-panel"),
    navList: $("st-nav-list"),
    navCount: $("st-nav-count"),
    board: $("st-board"),
    pages: $("st-pages"),
    empty: $("st-empty"),
    rulerH: $("st-ruler-h"),
    rulerV: $("st-ruler-v"),
    corner: $("st-corner"),
    sPage: $("st-s-page"),
    sPaper: $("st-s-paper"),
    sOrient: $("st-s-orient"),
    sZoom: $("st-s-zoom"),
    sOverflow: $("st-s-overflow"),
    toast: $("st-toast"),
  };

  // Interface fonts (IBM Plex) and the template fonts in one Google Fonts request.
  if (el.fonts) el.fonts.href = studioFontsHref(googleFontsHref());

  let current = 0;
  let sheets = [];
  let model = null;
  let titleEdited = false;
  let lastTemplate = null;
  let frame = 0;
  let pinnedUntil = 0;
  const reduceMotion = () => !!(win.matchMedia && win.matchMedia("(prefers-reduced-motion: reduce)").matches);
  const isPhone = () => !!(win.matchMedia && win.matchMedia(PHONE).matches);

  // -------------------------------------------------------------------------
  // Host and panel
  // -------------------------------------------------------------------------

  const host = createDemoHost({
    canvas: el.pages,
    viewport: el.board,
    onRender: rendered,
    measure() {
      const w = el.board.clientWidth || 800;
      const h = el.board.clientHeight || 600;
      return { width: Math.max(200, w - 96), height: Math.max(200, h - 104) };
    },
  });
  const panel = (options.mountPanel || mountPanel)(el.panel, host, { siteUrl: ".." });

  function say(text, kind, sticky) {
    if (panel && panel.setStatus) panel.setStatus(text, kind, sticky);
    toast(text, kind, sticky);
  }

  /**
   * On a phone the panel's status line sits in the sheet, which is hidden while closed, so
   * the message also shows (and is announced) in a toast above the status bar.
   */
  function toast(text, kind, sticky) {
    if (isPhone() && !el.app.classList.contains("st-sheet-open")) showToast(text, kind, sticky);
    else hideToast();
  }

  let toastTimer = null;
  function showToast(text, kind = "info", sticky = false) {
    if (!el.toast) return;
    clearTimeout(toastTimer);
    el.toast.textContent = text;
    el.toast.dataset.kind = kind;
    el.toast.classList.add("is-shown");
    toastTimer = setTimeout(hideToast, sticky || kind === "error" ? 12000 : 6000);
  }
  function hideToast() {
    if (!el.toast) return;
    clearTimeout(toastTimer);
    el.toast.classList.remove("is-shown");
    el.toast.textContent = "";
  }
  if (el.toast) el.toast.addEventListener("click", hideToast);

  // -------------------------------------------------------------------------
  // After each draw
  // -------------------------------------------------------------------------

  function rendered(m, result) {
    model = m;
    sheets = m ? [...el.pages.querySelectorAll(".ozp-sheet")] : [];
    el.empty.hidden = !!m;
    el.app.classList.toggle("st-has-doc", !!m);
    // A different publication gets its own title, like a new document. The same one keeps
    // its title through Undo and Redo, which empty the page and bring it back.
    const t = titleAfterRender({ lastTemplate, titleEdited, title: el.title.value }, m);
    lastTemplate = t.lastTemplate;
    titleEdited = t.titleEdited;
    el.title.value = t.title;
    if (t.reset) {
      host.setTitle("");
      current = 0;
    }
    doc.title = `${el.title.value || UNTITLED} | Offizin live demo`;
    if (current >= sheets.length) current = Math.max(0, sheets.length - 1);
    decorate();
    buildNavigator(result);
    syncControls();
    updateCurrent(true);
  }

  /** Print preview: crop and registration marks around each sheet, unless the renderer drew them. */
  function decorate() {
    const view = host.getView();
    if (!view.printPreview) return;
    for (const sheet of sheets) {
      if (sheet.querySelector('[class*="crop"]') || (sheet.parentElement && sheet.parentElement.querySelector('[class*="crop"]'))) continue;
      const w = sheet.offsetWidth || parseFloat(sheet.style.width) || 0;
      const h = sheet.offsetHeight || parseFloat(sheet.style.height) || 0;
      const box = doc.createElement("div");
      box.className = "st-marks";
      box.setAttribute("aria-hidden", "true");
      box.style.left = box.style.top = `-${MARK_BLEED}px`;
      box.innerHTML = marksSvg(w, h, MARK_BLEED);
      sheet.appendChild(box);
    }
  }

  /** True when the renderer drew printer spreads (it may not support them yet). */
  function spreadsShown(g) {
    return !!(g && g.spreads && g.spreadCount && sheets.length === g.spreadCount);
  }

  function sheetLabel(sheet, i) {
    const prev = sheet.previousElementSibling;
    if (prev && prev.classList.contains("ozp-label") && prev.textContent.trim()) return prev.textContent.trim();
    const g = host.geometry();
    return `${sheetWord(model, spreadsShown(g))} ${i + 1}`;
  }

  function buildNavigator(result) {
    const list = el.navList;
    list.innerHTML = "";
    el.navCount.textContent = sheets.length ? String(sheets.length) : "";
    const over = (result && result.overflow) || [];
    sheets.forEach((sheet, i) => {
      const label = sheetLabel(sheet, i);
      const li = doc.createElement("li");
      const btn = doc.createElement("button");
      btn.type = "button";
      btn.className = "st-thumb";
      btn.dataset.index = String(i);
      btn.setAttribute("aria-label", `${label}${over[i] ? ", text overflows" : ""}`);
      const w = sheet.offsetWidth || parseFloat(sheet.style.width) || 100;
      const h = sheet.offsetHeight || parseFloat(sheet.style.height) || 130;
      const k = Math.min(132 / w, 150 / h);
      const art = doc.createElement("span");
      art.className = "st-thumb-art";
      art.style.width = `${Math.round(w * k)}px`;
      art.style.height = `${Math.round(h * k)}px`;
      const scaler = doc.createElement("span");
      scaler.className = "st-thumb-scale";
      scaler.style.width = `${w}px`;
      scaler.style.height = `${h}px`;
      scaler.style.transform = `scale(${k.toFixed(5)})`;
      const clone = sheet.cloneNode(true);
      clone.removeAttribute("data-page");
      clone.querySelectorAll(".ozp-overflow-badge, .st-marks, .ozp-hit").forEach((n) => (n.classList.contains("ozp-hit") ? n.classList.remove("ozp-hit") : n.remove()));
      clone.setAttribute("aria-hidden", "true");
      scaler.appendChild(clone);
      art.appendChild(scaler);
      const text = doc.createElement("span");
      text.className = "st-thumb-label";
      text.textContent = label;
      if (over[i]) {
        const flag = doc.createElement("span");
        flag.className = "st-thumb-flag";
        flag.title = "Text overflows";
        flag.setAttribute("aria-hidden", "true");
        text.appendChild(flag);
      }
      btn.append(art, text);
      li.appendChild(btn);
      list.appendChild(li);
    });
  }

  el.navList.addEventListener("click", (e) => {
    const btn = e.target.closest(".st-thumb");
    if (btn) scrollToSheet(Number(btn.dataset.index));
  });
  el.navList.addEventListener("keydown", (e) => {
    const btn = e.target.closest(".st-thumb");
    if (!btn) return;
    const all = [...el.navList.querySelectorAll(".st-thumb")];
    let i = all.indexOf(btn);
    if (e.key === "ArrowDown" || e.key === "ArrowRight") i = Math.min(all.length - 1, i + 1);
    else if (e.key === "ArrowUp" || e.key === "ArrowLeft") i = Math.max(0, i - 1);
    else if (e.key === "Home") i = 0;
    else if (e.key === "End") i = all.length - 1;
    else return;
    e.preventDefault();
    all[i].focus();
    scrollToSheet(i);
  });

  function scrollToSheet(i, instant) {
    const sheet = sheets[i];
    if (!sheet) return;
    const top = sheet.getBoundingClientRect().top - el.board.getBoundingClientRect().top + el.board.scrollTop - 40;
    const behavior = instant || reduceMotion() ? "auto" : "smooth";
    if (el.board.scrollTo) el.board.scrollTo({ top: Math.max(0, top), behavior });
    else el.board.scrollTop = Math.max(0, top);
    current = i;
    // Keep the chosen page current while the pasteboard scrolls to it.
    pinnedUntil = Date.now() + (behavior === "smooth" ? 900 : 120);
    updateCurrent(true);
  }

  // -------------------------------------------------------------------------
  // Current page, rulers, and status bar
  // -------------------------------------------------------------------------

  function updateCurrent(force) {
    if (sheets.length && Date.now() > pinnedUntil) {
      const b = el.board.getBoundingClientRect();
      const rects = sheets.map((s) => s.getBoundingClientRect());
      const next = currentSheetIndex(rects, { top: b.top, bottom: b.bottom });
      if (next >= 0 && (next !== current || force)) current = next;
    }
    el.navList.querySelectorAll(".st-thumb").forEach((t, i) => {
      if (i === current) t.setAttribute("aria-current", "page");
      else t.removeAttribute("aria-current");
    });
    updateStatus();
    updateRulers();
  }

  function schedule() {
    if (frame) return;
    const raf = win.requestAnimationFrame || ((fn) => setTimeout(fn, 16));
    frame = raf(() => {
      frame = 0;
      updateCurrent(false);
    });
  }
  el.board.addEventListener("scroll", schedule, { passive: true });

  function updateRulers() {
    const g = host.geometry();
    const hr = el.rulerH.getBoundingClientRect();
    const vr = el.rulerV.getBoundingClientRect();
    // A little extra length covers the rulers while a resize waits for its redraw.
    const hLen = (el.rulerH.clientWidth || hr.width) + RULER_SLACK;
    const vLen = (el.rulerV.clientHeight || vr.height) + RULER_SLACK;
    const hT = el.rulerH.clientHeight || 22;
    const vT = el.rulerV.clientWidth || 22;
    const sheet = sheets[current];
    if (!g || !sheet) {
      el.rulerH.innerHTML = rulerSvg({ orientation: "h", length: hLen, thickness: hT, ticks: [] });
      el.rulerV.innerHTML = rulerSvg({ orientation: "v", length: vLen, thickness: vT, ticks: [] });
      el.corner.textContent = "";
      return;
    }
    const sr = sheet.getBoundingClientRect();
    const pxPerPt = g.pxPerPt;
    const ox = sr.left - hr.left;
    const oy = sr.top - vr.top;
    const sheetW = sr.width / pxPerPt || g.sheet.width;
    const sheetH = sr.height / pxPerPt || g.sheet.height;
    const bx = rulerBands({ axis: "x", origin: ox, pxPerPt, page: g.page, sheetPt: sheetW });
    const by = rulerBands({ axis: "y", origin: oy, pxPerPt, page: g.page, sheetPt: sheetH });
    el.rulerH.innerHTML = rulerSvg({ orientation: "h", length: hLen, thickness: hT, ticks: rulerTicks({ length: hLen, origin: ox, pxPerPt, unit: g.unit }), sheet: bx.sheet, bands: bx.bands });
    el.rulerV.innerHTML = rulerSvg({ orientation: "v", length: vLen, thickness: vT, ticks: rulerTicks({ length: vLen, origin: oy, pxPerPt, unit: g.unit }), sheet: by.sheet, bands: by.bands });
    el.corner.textContent = g.unit;
    el.corner.title = g.unit === "mm" ? "Rulers show millimeters" : "Rulers show inches";
  }

  function updateStatus() {
    const g = host.geometry();
    if (!g || !model) {
      el.sPage.textContent = "No publication";
      el.sPaper.textContent = "";
      el.sOrient.textContent = "";
      el.sZoom.textContent = "";
      el.sOverflow.textContent = "";
      el.sOverflow.classList.remove("is-warn");
      return;
    }
    const word = sheetWord(model, spreadsShown(g));
    el.sPage.textContent = sheets.length ? `${word} ${current + 1} of ${sheets.length}` : word;
    el.sPaper.textContent = `${paperName(g.page, g.unit)} (${sizeText(g.page, g.unit)})`;
    el.sOrient.textContent = orientationName(g.page);
    el.sZoom.textContent = g.zoom === "fit" ? `${zoomPercent(g.scale)}, fit` : zoomPercent(g.scale);
    const text = overflowText(g.overflow, word);
    if (el.sOverflow.textContent !== text) el.sOverflow.textContent = text;
    el.sOverflow.classList.toggle("is-warn", g.overflow.some(Boolean));
  }

  // -------------------------------------------------------------------------
  // App bar controls
  // -------------------------------------------------------------------------

  function syncControls() {
    const v = host.getView();
    const g = host.geometry();
    const fitOption = el.zoom.querySelector('option[value="fit"]');
    if (fitOption) fitOption.textContent = v.zoom === "fit" && g ? `Fit (${zoomPercent(g.scale)})` : "Fit";
    el.zoom.value = v.zoom === "fit" ? "fit" : String(v.zoom);
    if (el.zoom.value !== (v.zoom === "fit" ? "fit" : String(v.zoom))) el.zoom.value = "fit";
    const scale = g ? g.scale : 1;
    el.zoomIn.disabled = !g || scale >= ZOOM_STEPS[ZOOM_STEPS.length - 1] - 0.001;
    el.zoomOut.disabled = !g || scale <= ZOOM_STEPS[0] + 0.001;
    el.margins.setAttribute("aria-pressed", String(!!v.margins));
    el.folds.setAttribute("aria-pressed", String(!!v.folds));
    const hasFolds = !!(g && g.folds);
    el.folds.disabled = !!g && !hasFolds;
    el.folds.title = g && !hasFolds ? "This publication has no folds" : "Show or hide fold guides (G)";
    el.print.setAttribute("aria-pressed", String(!!v.printPreview));
    el.spreads.hidden = !(g && g.booklet);
    el.spreads.setAttribute("aria-pressed", String(!!v.spreads));
    el.pages.classList.toggle("st-hide-margins", !v.margins);
    el.pages.classList.toggle("st-hide-folds", !v.folds);
    el.app.classList.toggle("st-print", !!v.printPreview);
    el.download.disabled = !g;
  }

  /** Changes the view and keeps the current page in place. */
  function setView(next) {
    const keep = current;
    host.setView(next);
    if (sheets[keep]) scrollToSheet(keep, true);
    else syncControls();
  }

  function zoomBy(dir) {
    const g = host.geometry();
    if (!g) return;
    setView({ zoom: stepZoom(g.scale, dir) });
  }

  function toggleGuides() {
    const v = host.getView();
    const on = !(v.margins || v.folds);
    setView({ margins: on, folds: on });
    say(on ? "Guides are showing. They never print." : "Guides are hidden.");
  }

  function togglePrint() {
    const on = !host.getView().printPreview;
    setView({ printPreview: on });
    say(on ? "Print preview: guides are hidden, and crop marks show where the paper is trimmed." : "Back to the editing view.");
  }

  el.zoom.addEventListener("change", () => setView({ zoom: el.zoom.value }));
  el.zoomIn.addEventListener("click", () => zoomBy(1));
  el.zoomOut.addEventListener("click", () => zoomBy(-1));
  el.margins.addEventListener("click", () => setView({ margins: el.margins.getAttribute("aria-pressed") !== "true" }));
  el.folds.addEventListener("click", () => setView({ folds: el.folds.getAttribute("aria-pressed") !== "true" }));
  el.spreads.addEventListener("click", () => {
    const on = el.spreads.getAttribute("aria-pressed") !== "true";
    current = 0;
    host.setView({ spreads: on });
    say(on ? "Printer spreads: pages are paired the way they print on each sheet." : "Reading order: pages in the order people read them.");
  });
  el.print.addEventListener("click", togglePrint);

  // Title
  el.title.addEventListener("input", () => {
    titleEdited = el.title.value.trim() !== "";
    host.setTitle(titleEdited ? el.title.value : "");
    doc.title = `${el.title.value || UNTITLED} | Offizin live demo`;
  });
  el.title.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === "Escape") {
      e.preventDefault();
      el.title.blur();
    }
  });
  el.title.addEventListener("blur", () => {
    if (!el.title.value.trim()) {
      titleEdited = false;
      host.setTitle("");
      el.title.value = model ? model.name : UNTITLED;
    }
  });

  // Download
  el.download.addEventListener("click", () => {
    try {
      const name = host.downloadDocx(el.title.value);
      say(`${name} downloaded. It also opens in Google Docs: upload it to Drive and open it with Google Docs.`, "success", true);
    } catch (e) {
      say(e.message, "error");
    }
  });

  // Install menu
  function setMenu(open, focusFirst) {
    el.installMenu.hidden = !open;
    el.install.setAttribute("aria-expanded", String(open));
    if (open && focusFirst) {
      const first = el.installMenu.querySelector("a");
      if (first) first.focus();
    }
  }
  el.install.addEventListener("click", () => setMenu(el.installMenu.hidden, false));
  el.install.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setMenu(true, true);
    }
  });
  el.installMenu.addEventListener("keydown", (e) => {
    const links = [...el.installMenu.querySelectorAll("a")];
    const i = links.indexOf(doc.activeElement);
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const n = (i + (e.key === "ArrowDown" ? 1 : -1) + links.length) % links.length;
      links[n].focus();
    } else if (e.key === "Escape") {
      e.preventDefault();
      setMenu(false);
      el.install.focus();
    } else if (e.key === "Tab") {
      setMenu(false);
    }
  });
  doc.addEventListener("click", (e) => {
    if (!el.installMenu.hidden && !e.target.closest("#st-install-wrap")) setMenu(false);
  });

  // Phone panel sheet
  function setSheet(open) {
    el.app.classList.toggle("st-sheet-open", open);
    el.sheetToggle.setAttribute("aria-expanded", String(open));
    if (open) (el.sheetClose || el.panel).focus();
    else if (isPhone()) el.sheetToggle.focus();
  }
  el.sheetToggle.addEventListener("click", () => setSheet(!el.app.classList.contains("st-sheet-open")));
  if (el.sheetClose) el.sheetClose.addEventListener("click", () => setSheet(false));

  // -------------------------------------------------------------------------
  // Offizin history buttons
  // -------------------------------------------------------------------------

  function showHistory(snap) {
    const hv = historyView(panel, snap);
    el.history.hidden = !hv.supported;
    el.undo.disabled = !hv.canUndo;
    el.redo.disabled = !hv.canRedo;
    el.undo.title = `${hv.undoTitle} (Ctrl+Z)`;
    el.redo.title = `${hv.redoTitle} (Ctrl+Shift+Z)`;
    el.undo.setAttribute("aria-label", hv.undoTitle === "Nothing to undo" ? "Undo" : hv.undoTitle);
    el.redo.setAttribute("aria-label", hv.redoTitle === "Nothing to redo" ? "Redo" : hv.redoTitle);
    return hv;
  }
  const history = showHistory();
  if (history.supported) {
    try {
      panel.onHistoryChange((snap) => showHistory(snap));
    } catch (e) {
      el.history.hidden = true;
    }
  }

  async function runHistory(direction) {
    const hv = historyView(panel);
    if (!hv.supported || hv.busy) return;
    if (direction < 0 && !hv.canUndo) return say("There is no Offizin change to undo.");
    if (direction > 0 && !hv.canRedo) return say("There is no Offizin change to redo.");
    try {
      await (direction < 0 ? panel.undo() : panel.redo());
      // The panel reports the result on its own status line: Undone, Redone, or why it failed.
      // The toast repeats that line, because a closed sheet hides it on a phone.
      const status = el.panel.querySelector(".oz-status");
      const text = status ? status.textContent : "";
      if (text) toast(text, status.dataset.kind || "info", text.length > 70);
    } catch (e) {
      say((e && e.message) || "That could not be undone.", "error");
    }
  }
  el.undo.addEventListener("click", () => runHistory(-1));
  el.redo.addEventListener("click", () => runHistory(1));

  // -------------------------------------------------------------------------
  // Keyboard
  // -------------------------------------------------------------------------

  doc.addEventListener("keydown", (e) => {
    // A key the panel (or another control) already handled is left alone.
    if (e.defaultPrevented) return;
    const esc = escapeAction(e, { menuOpen: !el.installMenu.hidden, sheetOpen: el.app.classList.contains("st-sheet-open"), phone: isPhone() });
    if (esc === "menu") {
      setMenu(false);
      el.install.focus();
      return;
    }
    if (esc === "sheet") {
      setSheet(false);
      return;
    }
    const action = shortcutFor(e);
    if (!action) return;
    if ((action === "undo" || action === "redo") && !historyView(panel).supported) return;
    e.preventDefault();
    if (action === "undo") runHistory(-1);
    else if (action === "redo") runHistory(1);
    else if (action === "zoom-in") zoomBy(1);
    else if (action === "zoom-out") zoomBy(-1);
    else if (action === "zoom-fit") setView({ zoom: "fit" });
    else if (action === "guides") toggleGuides();
    else if (action === "print-preview") togglePrint();
  });

  // -------------------------------------------------------------------------
  // Sizing, fonts, and links such as demo/#template=newsletter
  // -------------------------------------------------------------------------

  let resizeTimer = null;
  const onResize = () => {
    updateRulers();
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      if (host.getView().zoom === "fit") host.redraw();
      else updateCurrent(true);
    }, 120);
  };
  if (win.ResizeObserver) new win.ResizeObserver(onResize).observe(el.board);
  else win.addEventListener("resize", onResize);
  if (doc.fonts && doc.fonts.ready) doc.fonts.ready.then(() => host.redraw()).catch(() => {});

  function openFromHash() {
    const m = /template=([a-z0-9-]+)/.exec(win.location.hash || "");
    if (panel.openTemplate) panel.openTemplate(m ? m[1] : "trifold");
    const create = el.panel.querySelector('[data-action="create"]');
    if (create) create.click();
  }
  setTimeout(openFromHash, 0);
  win.addEventListener("hashchange", openFromHash);

  updateCurrent(true);
  syncControls();
  return { host, panel, setView, scrollToSheet };
}
