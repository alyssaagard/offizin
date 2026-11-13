/**
 * The Offizin panel: one user interface for the Word task pane, the Google Docs sidebar,
 * and the live demo on the website.
 *
 * The panel never touches a document directly. It talks to a "host" object that knows how
 * to create publications, insert parts, change page setup, restyle, and check a document in
 * its own editor (see app/hosts/). That keeps the three versions identical for users.
 *
 * Tabs:
 *   Create  pick a publication; customize its paper, margins, spacing, text, and print
 *           settings; and merge a pasted list into labels, badges, envelopes, and more
 *   Design  color and font schemes (or your own), default spacing and text, page setup,
 *           columns, and restyling a publication you already made
 *   Insert  page parts (building blocks) placed at the cursor
 *   Info    business information sets, logo, and a shareable brand kit file
 *   Check   the design checker, printing and folding help, and Finish
 *
 * Offizin's own history: every change a host reports as an op ({ kind, label, data }) is
 * recorded, and the Undo and Redo buttons in the panel's header reverse only those changes.
 * The editor's own undo list and the user's typing are left alone (see core/history.js).
 */
import { TEMPLATES, CATEGORIES, PARTS, FOLD_GUIDES, findTemplate, findPart } from "../core/templates.js";
import {
  compose,
  composePart,
  BUSINESS_FIELDS,
  restyleMap,
  customizeFor,
  resolvePage,
  clampLayout,
  twoUpSheet,
  DEFAULT_LAYOUT,
  LAYOUT_LIMITS,
  DEFAULT_PRINT,
} from "../core/compose.js";
import { COLOR_SCHEMES, FONT_SCHEMES, WORD_FONT_CHOICES, GOOGLE_FONT_CHOICES, FONT_EQUIVALENTS, findFontScheme, fontsFor } from "../core/schemes.js";
import { PAPER, MONTHS, escapeHtml, normalizeHex, round, inch, mm } from "../core/util.js";
import { parseRecords, recordsToText } from "../core/merge.js";
import { createHistory, normalizeResult, errorMessage } from "../core/history.js";
import { renderPreview, blocksHtml, ensurePreviewCss, splitPages } from "../render/preview.js";
import { readLogoFile } from "../render/pictures.js";
import { isEditableTarget } from "../studio/keys.js";
import { DIAGRAMS } from "./diagrams.js";
import { MARK_SVG } from "./mark.js";

const STATE_VERSION = 3;

/** How many information sets the saved settings keep, and the Info tab's New button allows. */
export const MAX_SETS = 20;

/** Returned by busy() when an action failed, so a successful action that returns nothing still counts. */
const FAILED = Symbol("failed");

/** A history request from outside the panel made while nothing on the page had the keyboard focus. */
const NO_FOCUS = "no-focus";

/** The large preview draws at most this many pages; long merges and booklets say how many there are. */
const PREVIEW_PAGES = 8;

/** The status while the saved settings are on their way (an Apps Script round trip in Google Docs). */
const LOADING_TEXT = "Loading your settings...";

const TABS = [
  { id: "create", name: "Create" },
  { id: "design", name: "Design" },
  { id: "insert", name: "Insert" },
  { id: "info", name: "Info" },
  { id: "check", name: "Check" },
];

/** Spacing and text controls, in the order the panel shows them. */
export const LAYOUT_CONTROLS = [
  { key: "textScale", label: "Text size", group: "text", help: "Every size at once, from headings to captions." },
  { key: "lineScale", label: "Line spacing", group: "spacing", help: "Space between the lines of a paragraph." },
  { key: "spaceScale", label: "Paragraph spacing", group: "spacing", help: "Space between paragraphs, headings, and blocks." },
  { key: "padScale", label: "Padding", group: "spacing", help: "Space inside panels, boxes, bands, and table cells." },
  { key: "gapScale", label: "Gaps", group: "spacing", help: "Space between columns, panels, and cards." },
];

/** Where Shrink to fit stops for each control, so text stays readable and the design keeps its shape. */
export const FIT_FLOORS = { textScale: 0.75, lineScale: 0.9, spaceScale: 0.5, padScale: 0.5, gapScale: 0.5 };

/**
 * One Shrink to fit step. Each allowed control goes down one slider step (5%), never below its
 * floor, and a control the user already set below its floor is left alone. Returns null when
 * nothing can shrink any further.
 */
export function shrinkStep(values, keys) {
  const next = { ...values };
  let moved = false;
  for (const k of keys) {
    if (!LAYOUT_LIMITS[k]) continue;
    const floor = Math.max(FIT_FLOORS[k] ?? LAYOUT_LIMITS[k][0], LAYOUT_LIMITS[k][0]);
    if (!(values[k] > floor + 0.001)) continue;
    next[k] = Math.max(floor, Math.round((values[k] - 0.05) * 100) / 100);
    moved = true;
  }
  return moved ? next : null;
}

/** Word's own margin presets, so margins mean the same thing in Offizin and in Layout, then Margins. */
export const MARGIN_PRESETS = [
  { id: "narrow", label: "Narrow", m: { top: 36, right: 36, bottom: 36, left: 36 } },
  { id: "moderate", label: "Moderate", m: { top: 72, right: 54, bottom: 72, left: 54 } },
  { id: "normal", label: "Normal", m: { top: 72, right: 72, bottom: 72, left: 72 } },
  { id: "wide", label: "Wide", m: { top: 72, right: 144, bottom: 72, left: 144 } },
];

/** Print settings an editor can keep in the document itself. Hosts list the ones they support. */
export const PRINT_TOGGLES = [
  {
    key: "mirrorMargins",
    label: "Mirror margins",
    help: "For pages printed on both sides and bound. The left margin becomes the inside margin, so facing pages mirror each other.",
  },
  {
    key: "twoPerSheet",
    label: "Two pages on each sheet",
    help: "Word prints two pages side by side on each sheet of {sheet} paper, which halves the paper for small pages.",
  },
  {
    key: "embedFonts",
    label: "Save fonts in the file",
    help: "Word saves the fonts inside the document when it saves, so the publication looks the same on computers without them. The file gets larger.",
  },
];

/** Example rows for "Fill in an example" when a template has no sample list of its own. */
const EXAMPLE_PEOPLE = [
  {
    name: "Avery Johnson",
    company: "Lakeside Library",
    address: "48 Lake Road",
    address2: "",
    city: "Springfield",
    state: "ST",
    postal: "12345",
    country: "",
    recipient: "Avery Johnson",
    reason: "Ten years of volunteer service",
    date: "June 12, 2026",
    seat: "Table 1",
    table: "Table 1",
    role: "Volunteer Coordinator",
    organization: "Lakeside Library",
  },
  {
    name: "Sam Okafor",
    company: "Northfield College",
    address: "1200 College Avenue",
    address2: "Room 214",
    city: "Northfield",
    state: "ST",
    postal: "12346",
    country: "",
    recipient: "Sam Okafor",
    reason: "Outstanding mentoring",
    date: "June 12, 2026",
    seat: "Table 2",
    table: "Table 2",
    role: "Guest Speaker",
    organization: "Northfield College",
  },
  {
    name: "Priya Raman",
    company: "Riverside Arts",
    address: "9 Harbor Street",
    address2: "",
    city: "Riverside",
    state: "ST",
    postal: "12347",
    country: "",
    recipient: "Priya Raman",
    reason: "Best new program of the year",
    date: "June 12, 2026",
    seat: "Table 2",
    table: "Table 2",
    role: "Board Member",
    organization: "Riverside Arts",
  },
];

/** What each mail merge record becomes, for the status line. */
const MERGE_ITEMS = { labels: "label", envelope: "envelope", "name-badges": "badge", tickets: "ticket", "place-cards": "place card", certificate: "certificate" };

const ICONS = {
  undo: '<svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true" focusable="false"><path d="M7.5 4.5 3.5 8.5l4 4" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/><path d="M4 8.5h7a4.5 4.5 0 0 1 0 9H8.5" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>',
  redo: '<svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true" focusable="false"><path d="M12.5 4.5l4 4-4 4" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/><path d="M16 8.5H9a4.5 4.5 0 0 0 0 9h2.5" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>',
  history: '<svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true" focusable="false"><path d="M3.6 10a6.4 6.4 0 1 0 1.9-4.6" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/><path d="M3.2 3.4v3.4h3.4" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/><path d="M10 6.6V10l2.4 1.6" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>',
};

// ---------------------------------------------------------------------------
// Saved settings
// ---------------------------------------------------------------------------

function guessPaper() {
  try {
    const lang = (navigator.language || "en-US").toUpperCase();
    return /-(US|CA|MX|PH|CL|CO|VE|GT|CR|SV|PR)$/.test(lang) ? "letter" : "a4";
  } catch (e) {
    return "letter";
  }
}

function defaultCustomFonts() {
  return { heading: { word: "Georgia", google: "Georgia" }, body: { word: "Calibri", google: "Open Sans" }, headingBold: true };
}

export function defaultState() {
  return {
    v: STATE_VERSION,
    theme: {
      colorScheme: "ink",
      customColors: { primary: "#1B3A5C", accent: "#C0392B", dark: "#1C2430" },
      fontScheme: "typewriter",
      customFonts: defaultCustomFonts(),
      paper: guessPaper(),
      layout: { ...DEFAULT_LAYOUT },
    },
    sets: { Main: {} },
    activeSet: "Main",
    logo: null,
  };
}

const own = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
const BAD_KEYS = new Set(["__proto__", "constructor", "prototype"]);
const pickFont = (list, v, fallback) => (typeof v === "string" && list.some((f) => f.name === v) ? v : fallback);

/** Accepts saved state or an imported brand kit, keeping only fields Offizin understands. */
export function sanitizeState(input) {
  const base = defaultState();
  if (!input || typeof input !== "object") return base;
  const t = input.theme && typeof input.theme === "object" ? input.theme : {};
  const out = { ...base, theme: { ...base.theme } };
  if (typeof t.colorScheme === "string" && (t.colorScheme === "custom" || COLOR_SCHEMES.some((s) => s.id === t.colorScheme))) out.theme.colorScheme = t.colorScheme;
  // Offizin 2.0 saved its untouched Harbor default with every change, so settings from before
  // version 2 move to Ink, the black and white default. Harbor can be chosen again on the Design tab.
  if (!(Number(input.v) >= 2) && out.theme.colorScheme === "harbor") out.theme.colorScheme = "ink";
  if (t.customColors && typeof t.customColors === "object") {
    out.theme.customColors = {
      primary: normalizeHex(t.customColors.primary || base.theme.customColors.primary),
      accent: normalizeHex(t.customColors.accent || base.theme.customColors.accent),
      dark: normalizeHex(t.customColors.dark || base.theme.customColors.dark),
    };
  }
  if (t.customFonts && typeof t.customFonts === "object") {
    const d = defaultCustomFonts();
    const part = (k) => {
      const src = t.customFonts[k] && typeof t.customFonts[k] === "object" ? t.customFonts[k] : {};
      return { word: pickFont(WORD_FONT_CHOICES, src.word, d[k].word), google: pickFont(GOOGLE_FONT_CHOICES, src.google, d[k].google) };
    };
    out.theme.customFonts = { heading: part("heading"), body: part("body"), headingBold: t.customFonts.headingBold !== false };
  }
  if (t.fontScheme === "custom" || FONT_SCHEMES.some((f) => f.id === t.fontScheme)) out.theme.fontScheme = t.fontScheme;
  // Settings saved before 2.1.1 with the old default fonts (Editorial) move to the new default.
  if (!(Number(input.v) >= 3) && out.theme.fontScheme === "editorial") out.theme.fontScheme = "typewriter";
  if (t.paper === "letter" || t.paper === "a4") out.theme.paper = t.paper;
  if (t.layout && typeof t.layout === "object") out.theme.layout = clampLayout(t.layout);
  if (input.sets && typeof input.sets === "object") {
    const sets = {};
    for (const [rawName, fields] of Object.entries(input.sets).slice(0, MAX_SETS)) {
      const name = String(rawName).slice(0, 40);
      if (!fields || typeof fields !== "object" || BAD_KEYS.has(name) || !name.trim()) continue;
      const clean = {};
      for (const f of BUSINESS_FIELDS) if (typeof fields[f.key] === "string") clean[f.key] = fields[f.key].slice(0, 300);
      sets[name] = clean;
    }
    if (Object.keys(sets).length) out.sets = sets;
  }
  out.activeSet = typeof input.activeSet === "string" && own(out.sets, input.activeSet) ? input.activeSet : Object.keys(out.sets)[0];
  if (input.logo && typeof input.logo.src === "string" && /^data:image\/(png|jpeg|gif|webp);base64,/.test(input.logo.src)) {
    // Only a real, positive size: a negative or infinite one would reach the Word file as a picture size.
    const dim = (v, d) => {
      const n = Number(v);
      return Number.isFinite(n) && n >= 1 ? Math.min(n, 4000) : d;
    };
    out.logo = { src: input.logo.src, width: dim(input.logo.width, 200), height: dim(input.logo.height, 100) };
  }
  return out;
}

// ---------------------------------------------------------------------------
// Plain helpers (no DOM), exported for tests
// ---------------------------------------------------------------------------

/** Rows for "Fill in an example": the template's own sample list, or a few made up people. */
export function exampleRecords(merge) {
  const fields = (merge && merge.fields) || [];
  const source = merge && Array.isArray(merge.sample) && merge.sample.length ? merge.sample : EXAMPLE_PEOPLE;
  return source.slice(0, 6).map((r, i) => {
    const out = {};
    for (const f of fields) out[f.key] = r[f.key] !== undefined ? String(r[f.key]) : `${f.label} ${i + 1}`;
    return out;
  });
}

/** True when two margin sets match to the nearest point. */
export function sameMargins(a, b) {
  if (!a || !b) return false;
  return ["top", "right", "bottom", "left"].every((k) => Math.abs((Number(a[k]) || 0) - (Number(b[k]) || 0)) < 0.5);
}

/** "1, 2, and 3" */
function listText(items) {
  const a = items.map(String);
  if (a.length < 2) return a.join("");
  if (a.length === 2) return `${a[0]} and ${a[1]}`;
  return `${a.slice(0, -1).join(", ")}, and ${a[a.length - 1]}`;
}

const pct = (v) => `${Math.round(Number(v) * 100)}%`;

function sidesText(duplex) {
  if (duplex === "long") return "Both sides, flip on the long edge";
  if (duplex === "short") return "Both sides, flip on the short edge";
  return "One side";
}

/**
 * Plain steps for printing a publication in the user's editor.
 * guide: the host's capabilities.printGuide ({ editor: "word" | "google" | "web", web, desktopPageSetup }).
 * info:  { booklet, mode } for the publication.
 */
export function printSteps(print = {}, guide = {}, info = {}) {
  const p = { ...DEFAULT_PRINT, ...print };
  const edge = p.duplex === "long" || p.duplex === "short" ? p.duplex : "";
  const steps = [];
  const editor = guide.editor || "web";
  if (editor === "word" && guide.web) {
    steps.push("Choose File, then Print, then Print. Word on the web makes a PDF and opens your browser's print window.");
    // Only a booklet made with Open in a new document carries book fold; Word on the web cannot
    // turn it on in the document it is editing, so that booklet has separate half-size pages.
    const pageByPage = !!info.booklet && info.mode !== "new";
    if (info.booklet && !pageByPage) {
      steps.push(
        "Book fold is saved in this document, so Word for Windows or Mac puts every page on the right sheet and side. To print the folded booklet, choose Editing, then Open in Desktop App, print from the Word app, and print on both sides with the flip on the short edge."
      );
    } else if (pageByPage) {
      steps.push(
        "Word on the web cannot turn on book fold, so this booklet has separate half-size pages that print one after another. To print a folded booklet, use your printer's booklet setting, which also chooses the flip edge. Or, before you edit the text, choose Editing, then Open in Desktop App, open Offizin there, and create the booklet again."
      );
    }
    if (edge && !info.booklet) steps.push(`In the print window, turn on two-sided printing and choose flip on ${edge} edge.`);
    steps.push("Set the scale to actual size (100 percent), not Fit to page, so folds and cut lines land in the right place.");
  } else if (editor === "word") {
    steps.push("Choose File, then Print.");
    if (info.booklet && (guide.desktopPageSetup || info.mode === "new")) {
      steps.push("Book fold is turned on for this booklet, so Word puts every page on the right sheet and side. Check the order in the print preview.");
    } else if (info.booklet) {
      steps.push("This version of Word could not turn on book fold, so the pages print one after another. Use your printer's booklet setting, or a newer Word app.");
    }
    if (p.twoPerSheet && !info.booklet) steps.push("Two pages per sheet is saved in the document, so the print preview already shows two pages on each sheet.");
    if (edge) {
      steps.push(
        `On Windows, change Print One Sided to Print on Both Sides, Flip pages on ${edge} edge. On a Mac, turn on Two-Sided and choose ${edge === "long" ? "Long-Edge" : "Short-Edge"} binding.`
      );
    }
    steps.push("Keep the scale at 100 percent: on Windows, leave 1 Page Per Sheet without scaling to a paper size; on a Mac, leave Scale at 100 percent.");
  } else if (editor === "google") {
    steps.push("Choose File, then Print. Google Docs opens your browser's print window.");
    if (info.booklet) steps.push("Offizin already placed the booklet pages in printing order, two to a sheet.");
    if (edge) steps.push(`In the print window, turn on two-sided printing (Print on both sides) and choose flip on ${edge} edge.`);
    steps.push("Set the scale to Default, Actual size, or 100 percent, not Fit to page, so folds and cut lines land in the right place.");
  } else {
    steps.push("Download the Word file and print it from Word, or make the publication with Offizin in Word or Google Docs.");
    if (edge) steps.push(`Print on both sides and flip on the ${edge} edge.`);
    steps.push("Print at actual size (100 percent), not Fit to page.");
  }
  if (p.paper) steps.push(`Paper: ${p.paper}`);
  if (p.finishing) steps.push(`After printing: ${p.finishing}`);
  return steps;
}

/** The history key for a key event (undo or redo), or null. Text fields keep their own undo. */
export function historyKey(ev) {
  if (!ev || ev.defaultPrevented || ev.isComposing || ev.altKey) return null;
  if (!(ev.ctrlKey || ev.metaKey)) return null;
  if (isEditableTarget(ev.target)) return null;
  const k = String(ev.key || "").toLowerCase();
  if (k === "z") return ev.shiftKey ? "redo" : "undo";
  if (k === "y" && ev.ctrlKey && !ev.metaKey && !ev.shiftKey) return "redo";
  return null;
}

// ---------------------------------------------------------------------------
// The panel
// ---------------------------------------------------------------------------

/**
 * Mounts the panel.
 * host: see app/hosts/demo.js for the full interface.
 * opts.siteUrl: link target for Help.
 *
 * Returns { state, composeTheme, setStatus, showTab, openTemplate, undo, redo, history,
 * onHistoryChange }. undo, redo, and onHistoryChange are present only when the host keeps
 * Offizin's history; onHistoryChange(fn) returns a function that stops the calls.
 */
export function mountPanel(root, host, opts = {}) {
  ensurePreviewCss();
  const caps = host.capabilities || {};
  const modes = Array.isArray(caps.createModes) && caps.createModes.length ? caps.createModes : [{ id: "replace", label: "Replace this document" }];
  const isMac = (() => {
    try {
      return /Mac|iPhone|iPad/i.test(navigator.platform || navigator.userAgent || "");
    } catch (e) {
      return false;
    }
  })();
  const MOD = isMac ? "Command" : "Ctrl";
  let state = defaultState();
  const ui = {
    tab: "create",
    category: "all",
    query: "",
    selected: null,
    custom: {}, // per template: { options, paperSize, orientation, margins, sameMargins, layout, print }
    merge: {}, // per template: { text, parsed }
    mode: modes[0].id,
    findings: null,
    foldType: null,
    lastPrint: null, // { templateId, print, booklet, mode } of the last publication made in this session
    historyOpen: false,
    confirm: null,
    open: { merge: true },
  };
  const history = caps.history && typeof host.undo === "function" && typeof host.redo === "function" ? createHistory(host) : null;
  let saveTimer = null;
  let thumbTimer = null;
  let statusTimer = null;
  let previewTimer = null;
  let mergeTimer = null;
  let themeTimer = null;
  let working = false;
  let busyText = null; // the status of the change that is running, for when a passing message expires
  let loaded = false; // the saved settings have arrived (or there were none)
  let loadError = null;
  let savePending = false;
  let changedEarly = false; // a setting was changed before the saved settings arrived
  let askedAgain = false; // the status line is asking for that change to be made again
  let loading = null;

  root.classList.add("oz");
  root.innerHTML = `
    <header class="oz-head">
      <span class="oz-mark">${MARK_SVG}</span>
      <span class="oz-wordmark">Offizin</span>
      ${
        history
          ? `<span class="oz-hist-btns" role="group" aria-label="Offizin undo and redo">
              <button type="button" class="oz-icon oz-undo" data-action="undo" disabled aria-label="Offizin undo" title="Nothing to undo">${ICONS.undo}</button>
              <button type="button" class="oz-icon oz-redo" data-action="redo" disabled aria-label="Offizin redo" title="Nothing to redo">${ICONS.redo}</button>
              <button type="button" class="oz-icon oz-histbtn" data-action="history" aria-expanded="false" aria-controls="oz-history" aria-label="Offizin history" title="Offizin history">${ICONS.history}<span class="oz-count" hidden></span></button>
            </span>`
          : ""
      }
    </header>
    <div class="oz-confirm" role="alertdialog" aria-labelledby="oz-confirm-title" aria-describedby="oz-confirm-text" hidden></div>
    <section class="oz-history" id="oz-history" aria-label="Offizin history" hidden></section>
    <nav class="oz-tabs" role="tablist" aria-label="Offizin">
      ${TABS.map((t) => `<button type="button" role="tab" id="oz-tab-${t.id}" aria-controls="oz-panel" data-tab="${t.id}" aria-selected="${t.id === ui.tab}">${t.name}</button>`).join("")}
    </nav>
    <main class="oz-body" id="oz-panel" role="tabpanel" tabindex="-1"></main>
    <footer class="oz-foot">
      <div class="oz-status" role="status" aria-live="polite">Ready</div>
      ${host.label ? `<span class="oz-host">${escapeHtml(host.label)}</span>` : ""}
      ${opts.siteUrl ? `<a class="oz-help" href="${escapeHtml(opts.siteUrl)}/support/" target="_blank" rel="noopener">Help</a>` : ""}
    </footer>`;
  const body = root.querySelector(".oz-body");
  const statusEl = root.querySelector(".oz-status");

  // -------------------------------------------------------------------------
  // State
  // -------------------------------------------------------------------------

  function composeTheme() {
    return {
      colorScheme: state.theme.colorScheme,
      customColors: state.theme.customColors,
      fontScheme: state.theme.fontScheme,
      customFonts: state.theme.customFonts,
      paper: state.theme.paper,
      layout: state.theme.layout,
      business: state.sets[state.activeSet] || {},
      logo: state.logo,
    };
  }

  function save() {
    if (!loaded) changedEarly = true;
    else if (askedAgain) {
      askedAgain = false;
      readyStatus();
    }
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      if (!loaded) {
        // Saving now would put the default state over the saved settings (and drop the logo),
        // so the save waits for them. After a failed load, the load is tried again first.
        savePending = true;
        if (loadError) loadSaved();
        return;
      }
      Promise.resolve()
        .then(() => host.saveState(state))
        .catch((e) => setStatus("Your settings could not be saved: " + errorMessage(e), "error"));
    }, 350);
  }

  /** Loads the saved settings once; a second call while one is running shares it. */
  function loadSaved() {
    if (loading) return loading;
    const retry = !!loadError;
    loading = Promise.resolve()
      .then(() => host.loadState())
      .then((saved) => {
        loaded = true;
        loadError = null;
        // The saved settings win over changes made while they were still loading, and the
        // status line says so when that drops a change.
        const dropped = !!saved && changedEarly;
        changedEarly = false;
        if (saved) {
          clearTimeout(saveTimer);
          state = sanitizeState(saved);
          savePending = false;
          render();
        }
        if (host.themeChanged) host.themeChanged(composeTheme());
        if (statusEl.textContent === LOADING_TEXT) readyStatus();
        if (dropped) {
          setStatus("Your saved settings have loaded. Make your last change again.");
          askedAgain = true;
        }
        else if (retry && saved) setStatus("Your saved settings have loaded.");
        if (savePending) {
          savePending = false;
          save();
        }
      })
      .catch((e) => {
        loadError = e;
        setStatus("Saved settings could not be loaded: " + errorMessage(e), "error");
      })
      .finally(() => {
        loading = null;
      });
    return loading;
  }

  /** Saves and tells the host. A delay (for sliders) groups quick changes into one redraw. */
  function themeChanged(delay = 0) {
    save();
    clearTimeout(themeTimer);
    if (!host.themeChanged) return;
    if (delay) themeTimer = setTimeout(() => host.themeChanged(composeTheme()), delay);
    else host.themeChanged(composeTheme());
  }

  function setStatus(text, kind = "info", sticky = false) {
    clearTimeout(statusTimer);
    statusEl.textContent = text;
    statusEl.dataset.kind = kind;
    if (!sticky && kind !== "error") statusTimer = setTimeout(readyStatus, 6000);
  }

  /** Back to Ready, or to the message of the change that is still running. */
  function readyStatus() {
    const running = working && busyText;
    statusEl.textContent = running ? busyText : "Ready";
    statusEl.dataset.kind = running ? "busy" : "info";
  }

  /** Runs one editor action at a time, with a status message while it works. */
  async function busy(text, fn) {
    if (working) {
      setStatus("Offizin is still working on the last change. Try again in a moment.");
      return FAILED;
    }
    working = true;
    busyText = text;
    setStatus(text, "busy", true);
    root.classList.add("oz-busy");
    try {
      return await fn();
    } catch (e) {
      console.error(e);
      setStatus(errorMessage(e), "error", true);
      return FAILED;
    } finally {
      working = false;
      busyText = null;
      root.classList.remove("oz-busy");
    }
  }

  /** Runs a host change, records the op it reports in Offizin's history, and returns { message, op }. */
  async function act(text, fn, fallback) {
    const result = await busy(text, fn);
    if (result === FAILED) return FAILED;
    const { message, op } = normalizeResult(result, fallback);
    if (op && history) {
      try {
        history.push(op, op.label);
      } catch (e) {
        console.warn("Offizin could not record this change", e);
      }
    }
    return { message, op };
  }

  // -------------------------------------------------------------------------
  // Offizin history: Undo, Redo, and the history list
  // -------------------------------------------------------------------------

  function syncHistory(snap) {
    if (!history) return;
    const s = snap || history.snapshot();
    const undo = root.querySelector(".oz-undo");
    const redo = root.querySelector(".oz-redo");
    const count = root.querySelector(".oz-count");
    undo.disabled = !s.canUndo;
    redo.disabled = !s.canRedo;
    undo.title = s.canUndo ? `Undo: ${s.undoLabel} (${MOD}+Z)` : "Nothing to undo";
    redo.title = s.canRedo ? `Redo: ${s.redoLabel} (${MOD}+Shift+Z)` : "Nothing to redo";
    undo.setAttribute("aria-label", s.canUndo ? `Offizin undo: ${s.undoLabel}` : "Offizin undo");
    redo.setAttribute("aria-label", s.canRedo ? `Offizin redo: ${s.redoLabel}` : "Offizin redo");
    count.hidden = !s.size;
    count.textContent = String(s.size);
    if (ui.historyOpen) renderHistory(s);
  }

  /** The control that had the keyboard focus inside the panel, so a redraw can give it back. */
  function focusedControl() {
    const doc = root.ownerDocument;
    const a = doc && doc.activeElement;
    return a && a !== doc.body && root.contains(a) ? a : null;
  }

  /** The element that has the keyboard focus outside the panel (the live demo's app bar or page), or null. */
  function focusedOutside() {
    const doc = root.ownerDocument;
    const a = doc && doc.activeElement;
    return a && a !== doc.body && !root.contains(a) ? a : null;
  }

  /**
   * Puts the keyboard focus on the first candidate that is still in the panel, enabled, and not
   * hidden: elements or selectors, in order. Returns the element that got the focus, or null.
   */
  function focusFirst(...candidates) {
    for (const c of candidates) {
      const el = typeof c === "string" ? root.querySelector(c) : c;
      if (el && el.isConnected && !el.disabled && !el.closest("[hidden]") && el.focus) {
        el.focus();
        return el;
      }
    }
    return null;
  }

  /**
   * After a history change redraws the controls, the focus goes back to the control that
   * started it, or to its nearest match: the current step in the history list, the Undo
   * button while it is enabled, then the history button, which is never disabled.
   */
  function focusHistoryControl(returnTo) {
    return focusFirst(returnTo, '.oz-history [aria-current="step"]', ".oz-undo", ".oz-histbtn", body);
  }

  async function doHistory(kind, target, returnTo, outside = null) {
    const text = kind === "undo" ? "Undoing..." : kind === "redo" ? "Redoing..." : "Moving through the history...";
    const run = kind === "undo" ? () => history.undo() : kind === "redo" ? () => history.redo() : () => history.jump(target);
    const r = await busy(text, run);
    // The history has settled and redrawn its controls by now. A request that came from outside
    // the panel (the live demo's own Undo and Redo) leaves the focus where it is; when a
    // question took the focus into the panel on the way, it goes back to where it was.
    if (!outside) focusHistoryControl(returnTo);
    else if (outside === NO_FOCUS) {
      // Nothing had the focus when the request came from outside. If a question took the focus
      // into the panel on the way, let it go again.
      const a = root.ownerDocument.activeElement;
      if (a && root.contains(a) && a.blur) a.blur();
    } else if (!focusedOutside()) focusFirst(outside);
    if (r === FAILED) return null;
    const msg = (r && r.message) || "Done.";
    setStatus(msg, "success", msg.length > 70);
    return r;
  }

  /** The history's position and entries, so a question asked about it can tell when it changed. */
  function historyStamp(s = history.snapshot()) {
    return `${s.position}:${s.entries.map((e) => e.id).join(",")}`;
  }

  /**
   * Undo ("undo"), redo ("redo"), or move to a place in the history ("jump", target position).
   * Undoing a new publication puts the old document back, which drops edits made since, so the
   * panel asks first (except in the demo, which keeps every state).
   */
  function requestHistory(kind, target, external = false) {
    if (!history) return Promise.reject(new Error("This editor cannot undo Offizin changes."));
    const s = history.snapshot();
    if (s.busy || working) {
      setStatus("Offizin is still working on the last change. Try again in a moment.");
      return Promise.resolve(null);
    }
    if (kind === "undo" && !s.canUndo) {
      setStatus("There is no Offizin change to undo.");
      return Promise.resolve(null);
    }
    if (kind === "redo" && !s.canRedo) {
      setStatus("There is no Offizin change to redo.");
      return Promise.resolve(null);
    }
    if (kind === "jump" && Number(target) === s.position) return Promise.resolve(null);
    const returnTo = focusedControl();
    const outside = focusedOutside() || (external && !returnTo ? NO_FOCUS : null);
    const undoing = kind === "undo" ? [s.entries[s.position - 1]] : kind === "jump" ? history.plan(target).undo : [];
    const creates = undoing.filter((e) => e && e.kind === "create");
    if (creates.length && !caps.safeHistory) {
      const stamp = historyStamp(s);
      return new Promise((resolve) => {
        askConfirm({
          title: `Undo “${creates[creates.length - 1].label}”?`,
          text:
            "The document goes back to how it was before Offizin made this publication. Anything you typed or changed in the publication since then is lost. Redo builds the publication again from its design, without those changes.",
          ok: "Undo anyway",
          onOk: () => {
            // The question named one change. If the history moved on while it was open, the
            // same request could undo something else, so it is asked again instead.
            if (historyStamp() !== stamp) {
              setStatus("Offizin's history changed while this question was open, so nothing was undone. Choose Undo again.", "error");
              resolve(null);
              return;
            }
            doHistory(kind, target, returnTo, outside).then(resolve);
          },
          onCancel: () => {
            setStatus("Nothing was undone.");
            resolve(null);
          },
          returnTo: returnTo || (outside === NO_FOCUS ? null : outside),
        });
      });
    }
    return doHistory(kind, target, returnTo, outside);
  }

  function askConfirm({ title, text, ok, onOk, onCancel, returnTo = focusedControl() }) {
    const box = root.querySelector(".oz-confirm");
    if (ui.confirm && ui.confirm.onCancel) ui.confirm.onCancel();
    ui.confirm = { onOk, onCancel, returnTo };
    box.innerHTML = `<p class="oz-confirm-title" id="oz-confirm-title">${escapeHtml(title)}</p>
      <p id="oz-confirm-text">${escapeHtml(text)}</p>
      <span class="oz-btns"><button type="button" class="oz-btn oz-small oz-danger" data-action="confirm-ok">${escapeHtml(ok)}</button>
      <button type="button" class="oz-btn oz-small" data-action="confirm-cancel">Cancel</button></span>`;
    box.hidden = false;
    box.querySelector('[data-action="confirm-cancel"]').focus();
  }

  function closeConfirm(accepted) {
    const c = ui.confirm;
    ui.confirm = null;
    const box = root.querySelector(".oz-confirm");
    box.hidden = true;
    box.innerHTML = "";
    if (!c) return;
    (accepted ? c.onOk : c.onCancel)();
    // The focus goes back to the control that opened the question. When the answer started a
    // history change, that control is disabled until the change is done, and doHistory moves
    // the focus once it has settled.
    if (!(history && history.busy)) focusHistoryControl(c.returnTo);
  }

  function timeText(ms) {
    try {
      return new Date(ms).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    } catch (e) {
      return "";
    }
  }

  function toggleHistory(open) {
    ui.historyOpen = open === undefined ? !ui.historyOpen : !!open;
    const btn = root.querySelector(".oz-histbtn");
    if (btn) btn.setAttribute("aria-expanded", String(ui.historyOpen));
    renderHistory();
  }

  function renderHistory(snap) {
    const box = root.querySelector(".oz-history");
    if (!history || !ui.historyOpen) {
      box.hidden = true;
      box.innerHTML = "";
      return;
    }
    const s = snap || history.snapshot();
    const rows = s.entries.map((e, i) => ({ e, pos: i + 1 })).reverse();
    const item = (pos, label, meta, cls, extra = "") =>
      `<li class="${cls}${pos === s.position ? " oz-here" : ""}"><button type="button" data-jump="${pos}"${pos === s.position ? ' aria-current="step"' : ""}${s.busy ? " disabled" : ""}>
        <span class="oz-hist-label">${escapeHtml(label)}</span><span class="oz-hist-meta">${escapeHtml(meta)}</span></button>${extra}</li>`;
    box.innerHTML = `
      <div class="oz-history-head"><strong>Offizin history</strong><span>${s.size ? `${s.position} of ${s.size} applied` : "No changes yet"}</span></div>
      ${
        s.size
          ? `<p class="oz-note">Choose a step to go back or forward to it. Only changes Offizin made are listed.</p>
             <ol class="oz-hist">
               ${rows
                 .map(({ e, pos }) =>
                   item(
                     pos,
                     e.label,
                     `${e.done ? "" : "Undone, "}${timeText(e.at)}`,
                     e.done ? "" : "oz-undone",
                     e.failed ? `<span class="oz-hist-fail">${escapeHtml(e.failed)} <button type="button" class="oz-link" data-forget="${e.id}">Remove from the history</button></span>` : ""
                   )
                 )
                 .join("")}
               ${item(0, "Before Offizin's first change", "Start", "oz-start")}
             </ol>`
          : `<p class="oz-note">Each publication, page part, design, and page setup change Offizin makes appears here, so you can undo it later without touching your own edits.</p>`
      }
      ${caps.historyNote ? `<p class="oz-note">${escapeHtml(caps.historyNote)}</p>` : ""}
      <span class="oz-btns">
        ${s.size ? `<button type="button" class="oz-btn oz-small" data-action="history-clear"${s.busy ? " disabled" : ""}>Clear history</button>` : ""}
        <button type="button" class="oz-btn oz-small" data-action="history-close">Close</button>
      </span>`;
    box.hidden = false;
  }

  // -------------------------------------------------------------------------
  // Rendering
  // -------------------------------------------------------------------------

  function render() {
    root.querySelectorAll(".oz-tabs [data-tab]").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.tab === ui.tab)));
    body.setAttribute("aria-labelledby", "oz-tab-" + ui.tab);
    const views = { create: renderCreate, design: renderDesign, insert: renderInsert, info: renderInfo, check: renderCheck };
    views[ui.tab]();
  }

  function innerWidth() {
    return Math.max(220, body.clientWidth - 28);
  }

  /** Remembers which collapsible sections are open, so a redraw keeps them that way. */
  function rememberOpen() {
    body.querySelectorAll("details[data-sec]").forEach((d) => (ui.open[d.dataset.sec] = d.open));
  }

  const isOpen = (sec) => !!ui.open[sec];

  const FOCUS_ATTRS = ["data-margin", "data-custom", "data-layout", "data-deflayout", "data-preset", "data-print", "data-orient", "data-paper", "data-option", "data-cfont", "data-scheme", "data-hex", "data-jump", "data-action"];

  /**
   * Redraws part of the panel and puts the keyboard focus back on the same control. When the
   * redraw disabled or hid that control, the focus goes to the fallback (a selector or element),
   * then to the summary of the section the control is in, then to the panel body.
   */
  function keepFocus(fn, fallback = null) {
    const a = focusedControl();
    let selector = null;
    if (a && a.hasAttribute) {
      for (const attr of FOCUS_ATTRS) {
        if (a.hasAttribute(attr)) {
          selector = `[${attr}="${a.getAttribute(attr)}"]`;
          break;
        }
      }
    }
    fn();
    if (!selector) return;
    const el = root.querySelector(selector);
    const section = el && el.closest("details");
    focusFirst(el, fallback, section && section.querySelector("summary"), body);
  }

  /** What a booklet looks like in each editor's preview: Google Docs gets printer spreads. */
  function drawable(model) {
    return host.platform === "google" && model.booklet ? { ...model, pages: null } : model;
  }

  // Create ------------------------------------------------------------------

  function galleryList() {
    const q = ui.query.trim().toLowerCase();
    let list = TEMPLATES.filter((t) => ui.category === "all" || t.category === ui.category);
    if (q) list = list.filter((t) => `${t.name} ${t.description} ${t.sizeLabel} ${t.plain ? "plain screening ats" : ""}`.toLowerCase().includes(q));
    // Plain resumes first: most applications go through screening software.
    if (ui.category === "resumes") list = [...list].sort((a, b) => (b.plain ? 1 : 0) - (a.plain ? 1 : 0));
    return list;
  }

  function cardsHtml() {
    const list = galleryList();
    if (!list.length) return `<p class="oz-muted oz-empty">No publications match “${escapeHtml(ui.query.trim())}”. Try another word, or choose All.</p>`;
    return list
      .map(
        (t) => `<button type="button" class="oz-card" data-template="${t.id}" aria-label="${escapeHtml(t.name)}: ${escapeHtml(t.sizeLabel)}${t.plain ? ". Plain format, easy for screening software to read" : ""}">
          <span class="oz-thumb" data-thumb="${t.id}"></span>
          <span class="oz-card-name">${escapeHtml(t.name)}</span>
          <span class="oz-card-meta">${escapeHtml(t.sizeLabel)}</span>
          ${t.plain ? `<span class="oz-tag">Plain format</span>` : ""}
        </button>`
      )
      .join("");
  }

  function renderCreate() {
    if (ui.selected) return renderDetail();
    body.innerHTML = `
      <div class="oz-toolbar">
        <input type="search" class="oz-search" data-action="search" placeholder="Search publications" aria-label="Search publications" value="${escapeHtml(ui.query)}" autocomplete="off">
        <label class="oz-field-inline"><span>Show</span>
          <select data-action="category">${CATEGORIES.map((c) => `<option value="${c.id}"${c.id === ui.category ? " selected" : ""}>${escapeHtml(c.name)}</option>`).join("")}</select>
        </label>
        <div class="oz-field-inline"><span id="oz-paper-label">Paper</span>
          <span class="oz-seg" role="group" aria-labelledby="oz-paper-label">
            <button type="button" data-paper="letter" aria-pressed="${state.theme.paper === "letter"}">Letter</button>
            <button type="button" data-paper="a4" aria-pressed="${state.theme.paper === "a4"}">A4</button>
          </span>
        </div>
      </div>
      ${ui.category === "resumes" ? `<p class="oz-tip">Most employers read resumes with screening software first. Plain formats, without tables, columns, or text in shapes, are the safest choice for online applications.</p>` : ""}
      <div class="oz-gallery" data-gallery>${cardsHtml()}</div>`;
    scheduleThumbs();
  }

  function scheduleThumbs() {
    clearTimeout(thumbTimer);
    const boxes = [...body.querySelectorAll("[data-thumb]")];
    const theme = composeTheme();
    let i = 0;
    const step = () => {
      const start = Date.now();
      while (i < boxes.length && Date.now() - start < 30) {
        const box = boxes[i++];
        if (!box.isConnected) return;
        const t = findTemplate(box.dataset.thumb);
        try {
          const model = drawable(compose(t, theme, host.platform));
          const bw = box.clientWidth || 120;
          const bh = box.clientHeight || 110;
          const pxW = model.page.width * (96 / 72);
          const pxH = model.page.height * (96 / 72);
          renderPreview(box, model, { scale: Math.min(bw / pxW, bh / pxH) * 0.94, pages: 1, guides: false, checkOverflow: false });
        } catch (e) {
          console.error(e);
        }
      }
      if (i < boxes.length) thumbTimer = setTimeout(step, 0);
    };
    thumbTimer = setTimeout(step, 0);
  }

  function customFor(t) {
    if (!ui.custom[t.id]) ui.custom[t.id] = { options: {}, paperSize: null, orientation: null, margins: null, sameMargins: false, layout: {}, print: {} };
    return ui.custom[t.id];
  }

  function mergeFor(t) {
    if (!ui.merge[t.id]) ui.merge[t.id] = { text: "", parsed: null };
    return ui.merge[t.id];
  }

  function optionValue(t, o) {
    const c = customFor(t);
    if (c.options[o.key] !== undefined) return c.options[o.key];
    return typeof o.default === "function" ? o.default() : o.default;
  }

  /** The print settings the host can keep in the document, for the chosen create mode. */
  function printKeys() {
    const ps = caps.printSettings;
    const list = Array.isArray(ps) ? ps : ps && typeof ps === "object" ? ps[ui.mode] || [] : [];
    return list.filter((k) => PRINT_TOGGLES.some((x) => x.key === k));
  }

  /** Print settings that make sense for this publication: two pages per sheet needs a small portrait page. */
  function printKeysFor(model) {
    return printKeys().filter((k) => (k === "twoPerSheet" ? !!model && !model.booklet && !!twoUpSheet(model.page) : true));
  }

  /** Everything compose() needs from the panel for a template: its options and Customize settings. */
  function composeOptions(t) {
    const c = customFor(t);
    const out = {};
    for (const o of t.options || []) out[o.key] = optionValue(t, o);
    if (c.paperSize) out.paperSize = c.paperSize;
    if (c.orientation) out.orientation = c.orientation;
    if (c.margins) out.margins = { ...c.margins };
    if (Object.keys(c.layout).length) out.layout = { ...c.layout };
    const print = {};
    for (const k of printKeys()) if (c.print[k] !== undefined) print[k] = !!c.print[k];
    if (Object.keys(print).length) out.print = print;
    const m = ui.merge[t.id];
    if (t.merge && m && m.parsed && m.parsed.records.length) out.records = m.parsed.records;
    return out;
  }

  function composeSelected() {
    const t = ui.selected;
    return compose(t, composeTheme(), host.platform, composeOptions(t));
  }

  function effectiveLayout(t) {
    const c = customFor(t);
    const base = { ...DEFAULT_LAYOUT, ...(state.theme.layout || {}) };
    const out = {};
    for (const k of Object.keys(DEFAULT_LAYOUT)) out[k] = c.layout[k] !== undefined ? c.layout[k] : base[k];
    return clampLayout(out);
  }

  function optionField(o, v) {
    const id = `oz-o-${o.key}`;
    const help = o.help ? `<span class="oz-help-text" id="${id}-help">${escapeHtml(o.help)}</span>` : "";
    const described = o.help ? ` aria-describedby="${id}-help"` : "";
    const wide = o.help || o.type === "text" || o.type === "boolean" ? " oz-span" : "";
    const label = `<span>${escapeHtml(o.label)}</span>`;
    if (o.type === "month") {
      return `<label class="oz-field${wide}">${label}<select data-option="${o.key}"${described}>${MONTHS.map((m, i) => `<option value="${i + 1}"${Number(v) === i + 1 ? " selected" : ""}>${m}</option>`).join("")}</select>${help}</label>`;
    }
    if (o.type === "select") {
      return `<label class="oz-field${wide}">${label}<select data-option="${o.key}"${described}>${(o.choices || [])
        .map((c) => `<option value="${escapeHtml(c.value)}"${String(c.value) === String(v) ? " selected" : ""}>${escapeHtml(c.label)}</option>`)
        .join("")}</select>${help}</label>`;
    }
    if (o.type === "boolean") {
      return `<div class="oz-field${wide}"><label class="oz-check"><input type="checkbox" data-option="${o.key}"${v === true || v === "true" ? " checked" : ""}${described}> ${escapeHtml(o.label)}</label>${help}</div>`;
    }
    if (o.type === "number") {
      const min = Number.isFinite(o.min) ? ` min="${o.min}"` : "";
      const max = Number.isFinite(o.max) ? ` max="${o.max}"` : "";
      return `<label class="oz-field${wide}">${label}<input type="number" data-option="${o.key}" value="${escapeHtml(v ?? "")}"${min}${max} step="${o.step || 1}" inputmode="numeric"${described}>${help}</label>`;
    }
    return `<label class="oz-field${wide}">${label}<input type="text" data-option="${o.key}" value="${escapeHtml(v ?? "")}" maxlength="${o.maxLength || 120}" autocomplete="off"${described}>${help}</label>`;
  }

  function optionsHtml(t) {
    if (!(t.options || []).length) return "";
    return `<h3 class="oz-h">Options</h3><div class="oz-options">${t.options.map((o) => optionField(o, optionValue(t, o))).join("")}</div>`;
  }

  // Mail merge

  function mergeCount(m) {
    const n = m && m.parsed ? m.parsed.records.length : 0;
    return n ? `${n} ${n === 1 ? "record" : "records"}` : "No list";
  }

  function mergeStatusHtml(t, m) {
    const fields = t.merge.fields;
    if (!m.parsed) return `<p class="oz-note">No list yet. ${escapeHtml(t.merge.note || "")}</p>`;
    const { records, matched, header, warnings } = m.parsed;
    const label = (k) => (fields.find((f) => f.key === k) || {}).label || k;
    const used = matched.filter(Boolean).map(label);
    const lines = [];
    lines.push(
      records.length
        ? `<p class="oz-ok-text">${records.length} ${records.length === 1 ? "record is" : "records are"} ready. Each one gets its own ${escapeHtml(MERGE_ITEMS[t.id] || t.merge.item || "copy")}.</p>`
        : `<p class="oz-warn-text">No records found. Put one person on each line.</p>`
    );
    lines.push(
      `<p class="oz-note">${header ? `Matched columns by their headings: ${escapeHtml(listText(used))}.` : `There is no heading row, so columns are read in this order: ${escapeHtml(listText(used))}.`}</p>`
    );
    for (const w of warnings) lines.push(`<p class="oz-warn-text">${escapeHtml(w)}</p>`);
    if (records.length) {
      const show = records.slice(0, 3);
      lines.push(
        `<ol class="oz-merge-rows">${show
          .map((r) => `<li>${escapeHtml(fields.map((f) => r[f.key]).filter(Boolean).join(", ").replace(/\n/g, " ") || "(empty)")}</li>`)
          .join("")}${records.length > 3 ? `<li class="oz-more">and ${records.length - 3} more</li>` : ""}</ol>`
      );
    }
    if (records.length > 300) lines.push(`<p class="oz-note">Long lists take longer to build in the editor. Offizin uses up to 1,000 rows.</p>`);
    return lines.join("");
  }

  function mergeHtml(t) {
    if (!t.merge) return "";
    const m = mergeFor(t);
    const fields = t.merge.fields;
    const example = `${fields.map((f) => f.label).join(", ")}`;
    return `<details class="oz-sec" data-sec="merge"${isOpen("merge") ? " open" : ""}>
      <summary><span class="oz-sec-title">Mail merge</span><span class="oz-sec-meta" data-merge-count>${escapeHtml(mergeCount(m))}</span></summary>
      <div class="oz-sec-body">
        <p class="oz-muted">${escapeHtml(t.merge.note || "Paste a list to print one item per person.")}</p>
        <p class="oz-note">Columns: ${escapeHtml(listText(fields.map((f) => f.label)))}. Copy the rows from Excel or Google Sheets with their heading row, paste a .csv file, or open one.</p>
        <textarea class="oz-merge-text" data-merge rows="6" spellcheck="false" aria-label="List for mail merge" placeholder="${escapeHtml(example)}">${escapeHtml(m.text)}</textarea>
        <div class="oz-merge-status" data-merge-status aria-live="polite">${mergeStatusHtml(t, m)}</div>
        <span class="oz-btns">
          <label class="oz-btn oz-small oz-file">Open a .csv file<input type="file" accept=".csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain" data-action="merge-file"></label>
          <button type="button" class="oz-btn oz-small" data-action="merge-example">Fill in an example</button>
          <button type="button" class="oz-btn oz-small" data-action="merge-clear"${m.text ? "" : " disabled"}>Clear the list</button>
        </span>
      </div>
    </details>`;
  }

  function setMergeText(t, text) {
    const m = mergeFor(t);
    m.text = String(text || "");
    m.parsed = m.text.trim() ? parseRecords(m.text, t.merge.fields) : null;
    const status = body.querySelector("[data-merge-status]");
    if (status) status.innerHTML = mergeStatusHtml(t, m);
    const count = body.querySelector("[data-merge-count]");
    if (count) count.textContent = mergeCount(m);
    const clear = body.querySelector('[data-action="merge-clear"]');
    if (clear) clear.disabled = !m.text;
    schedulePreview(0);
  }

  // Customize

  const inches = () => state.theme.paper === "letter";
  const unitLabel = () => (inches() ? "in" : "mm");
  const toUnit = (pt) => (inches() ? round(pt / 72, 2) : round((pt * 25.4) / 72, 1));
  function fromUnit(v) {
    const n = Number(v);
    if (v === "" || !Number.isFinite(n) || n < 0) return null;
    return round(inches() ? inch(n) : mm(n), 2);
  }

  function sizeName(page) {
    if (page.sizeName && PAPER[page.sizeName]) return PAPER[page.sizeName].label.replace(/ \(.*\)$/, "");
    return inches() ? `${round(page.width / 72, 2)} x ${round(page.height / 72, 2)} in` : `${Math.round((page.width * 25.4) / 72)} x ${Math.round((page.height * 25.4) / 72)} mm`;
  }

  function pageSectionHtml(t, allow, c, page, natural, model) {
    const parts = [];
    if (allow.paper && allow.paper.length) {
      const list = allow.paper.filter((k) => PAPER[k]);
      if (natural.sizeName && PAPER[natural.sizeName] && !list.includes(natural.sizeName)) list.unshift(natural.sizeName);
      parts.push(
        `<label class="oz-field"><span>Paper size</span><select data-custom="paperSize">${list
          .map((k) => `<option value="${k}"${k === page.sizeName ? " selected" : ""}>${escapeHtml(PAPER[k].label)}${k === natural.sizeName ? ", as designed" : ""}</option>`)
          .join("")}</select></label>`
      );
    } else {
      parts.push(
        `<div class="oz-field"><span id="oz-dpaper">Paper</span><span class="oz-seg" role="group" aria-labelledby="oz-dpaper">
          <button type="button" data-paper="letter" aria-pressed="${state.theme.paper === "letter"}">Letter</button>
          <button type="button" data-paper="a4" aria-pressed="${state.theme.paper === "a4"}">A4</button></span>
          <span class="oz-help-text">${t.category === "labels" || t.fold === "cards" ? "Sheets for cards and labels come in a Letter layout and an A4 layout." : "This publication follows your usual paper size."}</span></div>`
      );
    }
    if (allow.orientation) {
      parts.push(
        `<div class="oz-field"><span id="oz-dorient">Orientation</span><span class="oz-seg" role="group" aria-labelledby="oz-dorient">
          <button type="button" data-orient="portrait" aria-pressed="${page.orientation === "portrait"}">Portrait</button>
          <button type="button" data-orient="landscape" aria-pressed="${page.orientation === "landscape"}">Landscape</button></span></div>`
      );
    }
    if (allow.margins) {
      const mirror = !!(model && model.print && model.print.mirrorMargins);
      const names = { top: "Top", bottom: "Bottom", left: mirror ? "Inside" : "Left", right: mirror ? "Outside" : "Right" };
      const step = inches() ? 0.05 : 1;
      parts.push(
        `<div class="oz-field"><span>Margins (${unitLabel()})</span>
          <div class="oz-margins">${["top", "bottom", "left", "right"]
            .map(
              (s) =>
                `<label class="oz-mini"><span data-margin-name="${s}">${names[s]}</span><input type="number" data-margin="${s}" min="0" step="${step}" value="${toUnit(page.margins[s])}" inputmode="decimal"></label>`
            )
            .join("")}</div>
          <label class="oz-check"><input type="checkbox" data-custom="sameMargins"${c.sameMargins ? " checked" : ""}> Same on all sides</label>
          <span class="oz-chips" role="group" aria-label="Margin presets">
            <button type="button" class="oz-chip" data-preset="designed" aria-pressed="${!c.margins}">As designed</button>
            ${MARGIN_PRESETS.map((p) => `<button type="button" class="oz-chip" data-preset="${p.id}" aria-pressed="${!!c.margins && sameMargins(c.margins, p.m)}">${p.label}</button>`).join("")}
          </span>
          <span class="oz-help-text">Narrow, Moderate, Normal, and Wide match Word's own margin settings. Offizin keeps at least two inches (51 mm) of room inside the margins.${
            t.fold === "trifold"
              ? " Panels follow the folds, so wider side margins make the outer panels narrower. Offizin also keeps each panel at least one inch (25 mm) wide."
              : t.fold === "halffold"
              ? " Panels follow the folds, so wider side margins make the outer panels narrower."
              : ""
          }</span>
        </div>`
      );
    } else {
      parts.push(`<p class="oz-note">The margins are fixed so this layout lines up with ${t.category === "labels" ? "standard label sheets" : "its cut lines"}.</p>`);
    }
    return parts.join("");
  }

  function layoutControlsHtml(allow, values, scope) {
    return LAYOUT_CONTROLS.filter((x) => (x.group === "text" ? allow.text : allow.spacing))
      .map((x) => {
        const [lo, hi] = LAYOUT_LIMITS[x.key];
        const v = values[x.key];
        const id = `oz-${scope}-${x.key}`;
        const attr = scope === "custom" ? "data-layout" : "data-deflayout";
        return `<div class="oz-slider">
          <label for="${id}"><span>${x.label}</span><output data-out="${scope}-${x.key}">${pct(v)}</output></label>
          <input type="range" id="${id}" ${attr}="${x.key}" min="${lo}" max="${hi}" step="0.05" value="${v}" aria-describedby="${id}-help" aria-valuetext="${pct(v)}">
          <span class="oz-help-text" id="${id}-help">${x.help}</span>
        </div>`;
      })
      .join("");
  }

  function layoutMeta(values) {
    const changed = LAYOUT_CONTROLS.filter((x) => Math.abs(values[x.key] - 1) > 0.001);
    if (!changed.length) return "As designed";
    if (changed.length === 1) return `${changed[0].label} ${pct(values[changed[0].key])}`;
    return `${changed.length} changes`;
  }

  function printToggleHelp(x, model) {
    if (x.key !== "twoPerSheet") return x.help;
    const sheet = model ? twoUpSheet(model.page) : null;
    return x.help.replace("{sheet}", sheet && sheet.sizeName && PAPER[sheet.sizeName] ? PAPER[sheet.sizeName].label.replace(/ \(.*\)$/, "") : "larger");
  }

  function printSectionHtml(t, model) {
    const c = customFor(t);
    const print = model ? model.print : { ...DEFAULT_PRINT };
    const keys = printKeysFor(model);
    const rows = [
      ["Sides", sidesText(print.duplex)],
      ["Scale", "Actual size (100 percent)"],
    ];
    if (print.paper) rows.push(["Paper", print.paper]);
    if (print.finishing) rows.push(["Finishing", print.finishing]);
    const toggles = PRINT_TOGGLES.filter((x) => keys.includes(x.key))
      .map(
        (x) => `<div class="oz-toggle"><label class="oz-check"><input type="checkbox" data-print="${x.key}"${(c.print[x.key] ?? print[x.key]) ? " checked" : ""}> ${escapeHtml(x.label)}</label>
          <span class="oz-help-text">${escapeHtml(printToggleHelp(x, model))}</span></div>`
      )
      .join("");
    const where = keys.length
      ? caps.printSettingsNote || "Offizin saves these settings in the document, so Word uses them every time it prints."
      : caps.printSettingsNote || "This editor does not let add-ins save print settings, so choose them in the print window. The steps are below.";
    const steps = printSteps(print, caps.printGuide || {}, { booklet: !!(model && model.booklet), mode: ui.mode });
    return `
      <dl class="oz-print">${rows.map(([k, v]) => `<dt>${k}</dt><dd>${escapeHtml(v)}</dd>`).join("")}</dl>
      ${toggles}
      <p class="oz-note">${escapeHtml(where)}</p>
      <p class="oz-strong">How to print</p>
      <ol class="oz-steps">${steps.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ol>`;
  }

  function customizeHtml(t) {
    const allow = customizeFor(t);
    const c = customFor(t);
    const options = composeOptions(t);
    let model = null;
    try {
      model = compose(t, composeTheme(), host.platform, options);
    } catch (e) {
      model = null;
    }
    const page = resolvePage(t, state.theme.paper, options);
    const natural = resolvePage(t, state.theme.paper, { ...options, paperSize: null, orientation: null, margins: null });
    const layout = effectiveLayout(t);
    const showLayout = allow.text || allow.spacing;
    const printMeta = model ? sidesText(model.print.duplex).replace(/, flip on the (long|short) edge/, ", $1 edge") : "";
    return `
      <h3 class="oz-h">Customize</h3>
      <details class="oz-sec" data-sec="page"${isOpen("page") ? " open" : ""}>
        <summary><span class="oz-sec-title">Paper and margins</span><span class="oz-sec-meta">${escapeHtml(sizeName(page))}, ${page.orientation}</span></summary>
        <div class="oz-sec-body">${pageSectionHtml(t, allow, c, page, natural, model)}</div>
      </details>
      ${
        showLayout
          ? `<details class="oz-sec" data-sec="layout"${isOpen("layout") ? " open" : ""}>
              <summary><span class="oz-sec-title">Spacing and text</span><span class="oz-sec-meta" data-layout-meta>${escapeHtml(layoutMeta(layout))}</span></summary>
              <div class="oz-sec-body">
                <p class="oz-note">These change only this publication. To change every publication and page part, use Spacing and text on the Design tab.</p>
                ${layoutControlsHtml(allow, layout, "custom")}
                <button type="button" class="oz-btn oz-small" data-action="layout-reset"${Object.keys(c.layout).length ? "" : " disabled"}>Use the Design tab settings</button>
              </div>
            </details>`
          : ""
      }
      <details class="oz-sec" data-sec="print"${isOpen("print") ? " open" : ""}>
        <summary><span class="oz-sec-title">Print settings</span><span class="oz-sec-meta">${escapeHtml(printMeta)}</span></summary>
        <div class="oz-sec-body">${printSectionHtml(t, model)}</div>
      </details>
      <button type="button" class="oz-link oz-reset-all" data-action="custom-reset">Reset all customizations for this publication</button>`;
  }

  function refreshCustomize() {
    const box = body.querySelector("[data-customize]");
    if (!box || !ui.selected) return;
    rememberOpen();
    keepFocus(() => (box.innerHTML = customizeHtml(ui.selected)));
  }

  /** Puts the effective (clamped) margins into the inputs the user is not typing in. */
  function syncMarginInputs() {
    const t = ui.selected;
    if (!t) return;
    const page = resolvePage(t, state.theme.paper, composeOptions(t));
    const active = root.ownerDocument && root.ownerDocument.activeElement;
    body.querySelectorAll("[data-margin]").forEach((input) => {
      if (input !== active) input.value = toUnit(page.margins[input.dataset.margin]);
    });
    const c = customFor(t);
    body.querySelectorAll("[data-preset]").forEach((b) => {
      const p = MARGIN_PRESETS.find((x) => x.id === b.dataset.preset);
      b.setAttribute("aria-pressed", String(b.dataset.preset === "designed" ? !c.margins : !!c.margins && !!p && sameMargins(c.margins, p.m)));
    });
    const same = body.querySelector('[data-custom="sameMargins"]');
    if (same) same.checked = !!c.sameMargins;
  }

  function modeNote() {
    const m = modes.find((x) => x.id === ui.mode);
    return (m && m.note) || caps.createNote || "";
  }

  function renderDetail() {
    const t = ui.selected;
    rememberOpen();
    body.innerHTML = `
      <button type="button" class="oz-back" data-action="back">‹ All publications</button>
      <h2 class="oz-title">${escapeHtml(t.name)}</h2>
      ${t.plain ? `<p class="oz-tags"><span class="oz-tag">Plain format</span> <span class="oz-note">Easy for screening software to read.</span></p>` : ""}
      <p class="oz-muted">${escapeHtml(t.description)}</p>
      <p class="oz-meta">${escapeHtml(t.sizeLabel)}</p>
      <div class="oz-large" role="img" aria-label="Preview of ${escapeHtml(t.name)}"></div>
      <p class="oz-preview-note" data-preview-note aria-live="polite"></p>
      <div class="oz-fit" data-fit hidden>
        <button type="button" class="oz-btn oz-small" data-action="fit" aria-describedby="oz-fit-help">Shrink to fit</button>
        <span class="oz-help-text" id="oz-fit-help">Makes the text and spacing a little smaller, step by step, until every page fits.</span>
      </div>
      ${optionsHtml(t)}
      ${mergeHtml(t)}
      <div data-customize>${customizeHtml(t)}</div>
      ${
        modes.length > 1
          ? `<fieldset class="oz-modes"><legend>Where</legend>${modes
              .map((m) => `<label class="oz-radio"><input type="radio" name="oz-mode" value="${m.id}"${m.id === ui.mode ? " checked" : ""}><span>${escapeHtml(m.label)}</span></label>`)
              .join("")}</fieldset>`
          : ""
      }
      <button type="button" class="oz-btn oz-primary oz-wide" data-action="create">${escapeHtml(caps.createLabel || "Create publication")}</button>
      <p class="oz-note" data-mode-note>${escapeHtml(modeNote())}</p>`;
    renderLarge();
  }

  /** Draws the large preview and returns the numbers of the pages whose text runs past the bottom. */
  function renderLarge() {
    const box = body.querySelector(".oz-large");
    const note = body.querySelector("[data-preview-note]");
    const fit = body.querySelector("[data-fit]");
    if (fit) fit.hidden = true;
    if (!box || !ui.selected) return [];
    let model;
    try {
      model = drawable(composeSelected());
    } catch (e) {
      console.error(e);
      box.innerHTML = "";
      if (note) {
        note.textContent = "The preview could not be drawn: " + errorMessage(e);
        note.dataset.kind = "warn";
      }
      return [];
    }
    const result = renderPreview(box, model, { width: innerWidth(), labels: true, gap: 10, pages: PREVIEW_PAGES }) || {};
    // Running text (resumes and CVs) is split into pages when it is drawn, so the drawing knows the count.
    const total = result.total || (model.pages ? model.pages.length : splitPages(model.blocks).length);
    const over = (result.overflow || []).map((o, i) => (o ? i + 1 : 0)).filter(Boolean);
    const canShrink = over.length > 0 && !!shrinkStep(effectiveLayout(ui.selected), layoutKeys(ui.selected));
    if (fit) fit.hidden = !canShrink;
    if (!note) return over;
    const word = model.folds || model.page.width > model.page.height ? ["side", "sides"] : ["page", "pages"];
    const parts = [total > PREVIEW_PAGES ? `Showing the first ${PREVIEW_PAGES} of ${total} ${word[1]}.` : `${total} ${total === 1 ? word[0] : word[1]}.`];
    if (over.length) {
      parts.push(
        canShrink
          ? `Text runs past the bottom of ${word[0]} ${listText(over)}. Choose Shrink to fit, lower the text size or spacing under Customize, or shorten the text after you create it.`
          : `Text runs past the bottom of ${word[0]} ${listText(over)}. Shorten the text after you create it, or choose a larger paper size.`
      );
    }
    note.textContent = parts.join(" ");
    note.dataset.kind = over.length ? "warn" : "";
    return over;
  }

  /** The spacing and text controls this publication allows. */
  function layoutKeys(t) {
    const allow = customizeFor(t);
    return LAYOUT_CONTROLS.filter((x) => (x.group === "text" ? allow.text : allow.spacing)).map((x) => x.key);
  }

  /**
   * Shrink to fit: lowers the text size and spacing together, one step at a time, until the
   * preview shows no overflow or every control reaches its floor. It uses the same measurement
   * as the preview note, and only changes this publication's Customize settings.
   */
  function shrinkToFit(t) {
    const keys = layoutKeys(t);
    const c = customFor(t);
    let values = effectiveLayout(t);
    let over = renderLarge();
    let steps = 0;
    while (over.length && steps < 40) {
      const next = shrinkStep(values, keys);
      if (!next) break;
      values = next;
      for (const k of keys) c.layout[k] = values[k];
      steps++;
      over = renderLarge();
    }
    refreshCustomize();
    if (!steps) {
      setStatus(over.length ? "Text and spacing are already as small as they go." : "Everything already fits.");
      return;
    }
    // Show the new values, with the button that puts them back, under Spacing and text. The
    // Shrink to fit button is hidden once everything fits, so the focus moves to that section.
    const sec = body.querySelector('details[data-sec="layout"]');
    if (sec) {
      sec.open = true;
      ui.open.layout = true;
      const fit = body.querySelector('[data-action="fit"]');
      if (!fit || fit.closest("[hidden]")) focusFirst(sec.querySelector("summary"));
    }
    const text = pct(values.textScale);
    const space = pct(values.spaceScale);
    const sizes = !keys.includes("textScale") ? `spacing at ${space}` : keys.length === 1 ? `text at ${text}` : text === space ? `text and spacing at ${text}` : `text at ${text}, spacing at ${space}`;
    if (over.length) setStatus(`Still too long with ${sizes}. Shorten the text.`, "error");
    else setStatus(`Fits now, with ${sizes}.`, "success");
  }

  function schedulePreview(delay = 150) {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(renderLarge, delay);
  }

  async function createSelected() {
    const t = ui.selected;
    const options = composeOptions(t);
    let model;
    try {
      model = compose(t, composeTheme(), host.platform, options);
    } catch (e) {
      setStatus(errorMessage(e), "error", true);
      return;
    }
    const r = await act(`Creating ${t.name.toLowerCase()}...`, () => host.createPublication(model, ui.mode, { options, label: t.name }), `${t.name} created.`);
    if (r === FAILED) return;
    ui.foldType = t.fold;
    ui.lastPrint = { templateId: t.id, print: model.print, booklet: !!model.booklet, mode: ui.mode };
    setStatus(r.message, "success", r.message.length > 60);
  }

  // Design ------------------------------------------------------------------

  const fontSlot = () => (host.platform === "word" ? "word" : "google");
  const fontChoices = () => (host.platform === "word" ? WORD_FONT_CHOICES : GOOGLE_FONT_CHOICES);

  function fontListHtml() {
    const th = state.theme;
    const option = (scheme, id) => {
      const fn = fontsFor(scheme, host.platform);
      return `<button type="button" class="oz-fontopt" data-font="${id}" aria-pressed="${th.fontScheme === id}">
        <span class="oz-fontname" style="font-family:${escapeHtml(fn.heading.css)};font-weight:${scheme.headingBold ? 700 : 400}">${escapeHtml(scheme.name)}</span>
        <span class="oz-fontmeta" style="font-family:${escapeHtml(fn.body.css)}">${escapeHtml(fn.heading.name)} and ${escapeHtml(fn.body.name)}</span></button>`;
    };
    return FONT_SCHEMES.map((f) => option(f, f.id)).join("") + option(findFontScheme("custom", th.customFonts), "custom");
  }

  function customFontsHtml() {
    const cf = state.theme.customFonts;
    const slot = fontSlot();
    const select = (part) =>
      `<select data-cfont="${part}">${fontChoices()
        .map((f) => `<option value="${escapeHtml(f.name)}"${f.name === cf[part][slot] ? " selected" : ""}>${escapeHtml(f.name)}</option>`)
        .join("")}</select>`;
    return `<details class="oz-custom" data-sec="fonts"${state.theme.fontScheme === "custom" || isOpen("fonts") ? " open" : ""}>
      <summary>Choose your own fonts</summary>
      <label class="oz-field"><span>Headings</span>${select("heading")}</label>
      <label class="oz-field"><span>Body text</span>${select("body")}</label>
      <label class="oz-check"><input type="checkbox" data-cfont="bold"${cf.headingBold ? " checked" : ""}> Bold headings</label>
      <p class="oz-note">${
        host.platform === "word"
          ? "These fonts come with Microsoft 365. Offizin also picks a close Google Docs match, so a brand kit works in both editors."
          : "These are Google Fonts that Google Docs offers by name. Offizin also picks a close Word match, so a brand kit works in both editors."
      }</p>
    </details>`;
  }

  function renderDesign() {
    const t = state.theme;
    const unitIn = inches();
    const unit = unitLabel();
    const presetSize = t.paper === "letter" ? "letter" : "a4";
    const margin = 36;
    const layout = clampLayout({ ...DEFAULT_LAYOUT, ...(t.layout || {}) });
    rememberOpen();
    body.innerHTML = `
      <h3 class="oz-h">Color scheme</h3>
      <div class="oz-swatches">
        ${COLOR_SCHEMES.map(
          (s) => `<button type="button" class="oz-swatch" data-scheme="${s.id}" aria-pressed="${t.colorScheme === s.id}" title="${escapeHtml(s.name)}">
            <span class="oz-chipset"><i style="background:${s.primary}"></i><i style="background:${s.accent}"></i><i style="background:${s.dark}"></i></span>
            <span class="oz-swatch-name">${escapeHtml(s.name)}</span></button>`
        ).join("")}
        <button type="button" class="oz-swatch" data-scheme="custom" aria-pressed="${t.colorScheme === "custom"}">
          <span class="oz-chipset"><i style="background:${t.customColors.primary}"></i><i style="background:${t.customColors.accent}"></i><i style="background:${t.customColors.dark}"></i></span>
          <span class="oz-swatch-name">Your colors</span></button>
      </div>
      <details class="oz-custom"${t.colorScheme === "custom" ? " open" : ""}>
        <summary>Choose your own colors</summary>
        ${[
          ["primary", "Main color"],
          ["accent", "Accent color"],
          ["dark", "Text color"],
        ]
          .map(
            ([k, label]) =>
              `<label class="oz-color"><span>${label}</span><input type="color" data-color="${k}" value="${t.customColors[k]}"><input type="text" data-hex="${k}" value="${t.customColors[k]}" maxlength="7" spellcheck="false" aria-label="${label} hex code"></label>`
          )
          .join("")}
      </details>

      <h3 class="oz-h">Font scheme</h3>
      <div class="oz-fontlist" data-fontlist>${fontListHtml()}</div>
      ${customFontsHtml()}

      <h3 class="oz-h">Spacing and text</h3>
      <p class="oz-muted">Applies to every new publication and page part. Each publication can also have its own settings under Customize on the Create tab.</p>
      <div data-deflayout-box>${layoutControlsHtml({ text: true, spacing: true }, layout, "default")}</div>
      <button type="button" class="oz-btn oz-small" data-action="deflayout-reset">Reset to the designs' own spacing</button>

      <p class="oz-tip">Offizin builds publications with your editor's own styles: Title, Heading 1, Heading 2, and Normal. The plain resumes and CVs take their name and headings from those styles, so changing a style in ${
        host.platform === "word" ? "Word's Styles pane" : host.platform === "google" ? "Google Docs (Format, then Paragraph styles)" : "Word or Google Docs"
      } changes every matching line. In other designs, the fonts follow the styles, but some sizes and colors a design sets on purpose, such as a white title on a dark cover, stay as they are.</p>

      ${
        caps.restyle
          ? `<h3 class="oz-h">Update this publication</h3>
             <p class="oz-muted">Switch the colors and fonts of a publication you already made, and keep your text. Offizin's Undo switches them back.</p>
             <button type="button" class="oz-btn oz-wide" data-action="restyle">Apply this design to the publication</button>`
          : ""
      }

      <h3 class="oz-h">Page setup</h3>
      ${
        caps.pageSetup === true
          ? `<label class="oz-field"><span>Size</span><select data-ps="size">${Object.entries(PAPER)
              .map(([id, p]) => `<option value="${id}"${id === presetSize ? " selected" : ""}>${escapeHtml(p.label)}</option>`)
              .join("")}</select></label>
             <label class="oz-field"><span>Orientation</span><select data-ps="orientation"><option value="portrait">Portrait</option><option value="landscape">Landscape</option></select></label>
             <div class="oz-margins">
               ${["top", "bottom", "left", "right"]
                 .map(
                   (s) =>
                     `<label class="oz-mini"><span>${s[0].toUpperCase() + s.slice(1)} (${unit})</span><input type="number" step="${unitIn ? 0.05 : 1}" min="0" data-ps="${s}" value="${toUnit(margin)}"></label>`
                 )
                 .join("")}
             </div>
             <button type="button" class="oz-btn oz-wide" data-action="page-setup">Apply page setup</button>`
          : `<p class="oz-muted">${escapeHtml(caps.pageSetupHelp || "")}</p>`
      }
      ${
        caps.columns
          ? `<h3 class="oz-h">Columns</h3>
             <div class="oz-row2">
               <label class="oz-field"><span>Columns</span><select data-col="count"><option>1</option><option selected>2</option><option>3</option></select></label>
               <label class="oz-check"><input type="checkbox" data-col="line"> Line between</label>
             </div>
             <button type="button" class="oz-btn oz-wide" data-action="columns">Apply to this section</button>`
          : caps.columnsHelp
          ? `<h3 class="oz-h">Columns</h3><p class="oz-muted">${escapeHtml(caps.columnsHelp)}</p>`
          : ""
      }`;
  }

  function refreshFontList() {
    const list = body.querySelector("[data-fontlist]");
    if (list) list.innerHTML = fontListHtml();
  }

  function readPageSetup() {
    const q = (k) => body.querySelector(`[data-ps="${k}"]`);
    const toPt = (v) => (inches() ? inch(Number(v) || 0) : mm(Number(v) || 0));
    const size = PAPER[q("size").value] || PAPER.letter;
    const landscape = q("orientation").value === "landscape";
    return {
      width: round(landscape ? size.height : size.width),
      height: round(landscape ? size.width : size.height),
      orientation: landscape ? "landscape" : "portrait",
      margins: { top: round(toPt(q("top").value)), bottom: round(toPt(q("bottom").value)), left: round(toPt(q("left").value)), right: round(toPt(q("right").value)) },
    };
  }

  async function restyle() {
    const pub = await Promise.resolve()
      .then(() => (host.getPublication ? host.getPublication() : null))
      .catch(() => null);
    const from = pub && pub.theme ? pub.theme : null;
    if (!from) {
      setStatus("Offizin could not find the design this publication was made with. Restyling works on publications created with Offizin.", "error", true);
      return;
    }
    const to = composeTheme();
    const map = restyleMap(from, to, host.platform);
    if (!Object.keys(map.colors).length && !Object.keys(map.fonts).length) {
      setStatus("This publication already uses these colors and fonts.");
      return;
    }
    const theme = { colorScheme: to.colorScheme, customColors: to.customColors, fontScheme: to.fontScheme, customFonts: to.customFonts };
    const r = await act("Applying colors and fonts...", () => host.restyle(map, theme), "Design applied. Offizin's Undo brings the old look back.");
    if (r !== FAILED) setStatus(r.message, "success");
  }

  // Insert ------------------------------------------------------------------

  function renderInsert() {
    body.innerHTML = `
      <p class="oz-muted">${escapeHtml(caps.insertHelp || "Click a part to add it after the cursor. Parts use your current colors, fonts, spacing, and business information.")}</p>
      <div class="oz-parts">
        ${PARTS.map(
          (pt) => `<button type="button" class="oz-part" data-part="${pt.id}">
            <span class="oz-part-prev" data-prev="${pt.id}" aria-hidden="true"></span>
            <span class="oz-part-text"><strong>${escapeHtml(pt.name)}</strong><span>${escapeHtml(pt.description)}</span></span></button>`
        ).join("")}
      </div>`;
    const theme = composeTheme();
    body.querySelectorAll("[data-prev]").forEach((el) => {
      try {
        const { blocks } = composePart(findPart(el.dataset.prev), theme, host.platform, 400);
        const s = 84 / (400 * (96 / 72));
        el.innerHTML = `<span class="oz-part-scale" style="width:400pt;transform:scale(${s})">${blocksHtml(blocks)}</span>`;
      } catch (e) {
        console.error(e);
      }
    });
  }

  async function insertPart(id) {
    const part = findPart(id);
    const raw = await Promise.resolve()
      .then(() => (host.contentWidth ? host.contentWidth() : 468))
      .catch(() => 468);
    const width = Number(raw) > 72 ? Number(raw) : 468;
    const composed = composePart(part, composeTheme(), host.platform, width);
    const r = await act(`Inserting ${part.name.toLowerCase()}...`, () => host.insertPart(composed, width, part.id, { label: part.name }), `${part.name} inserted.`);
    if (r !== FAILED) setStatus(r.message, "success");
  }

  // Info --------------------------------------------------------------------

  function renderInfo() {
    const fields = state.sets[state.activeSet] || {};
    body.innerHTML = `
      <p class="oz-muted">Publications fill in this information for you.</p>
      <div class="oz-row2">
        <label class="oz-field"><span>Information set</span>
          <select data-action="set">${Object.keys(state.sets).map((n) => `<option value="${escapeHtml(n)}"${n === state.activeSet ? " selected" : ""}>${escapeHtml(n)}</option>`).join("")}</select>
        </label>
        <span class="oz-btns"><button type="button" class="oz-btn oz-small" data-action="new-set"${Object.keys(state.sets).length >= MAX_SETS ? " disabled" : ""}>New</button><button type="button" class="oz-btn oz-small" data-action="rename-set">Rename</button><button type="button" class="oz-btn oz-small" data-action="delete-set"${Object.keys(state.sets).length < 2 ? " disabled" : ""}>Delete</button></span>
      </div>
      <div class="oz-newset" hidden>
        <label class="oz-field"><span data-newset-label>Name for the new set</span><input type="text" data-newset maxlength="40"></label>
        <span class="oz-btns"><button type="button" class="oz-btn oz-small oz-primary" data-action="newset-ok">Save</button><button type="button" class="oz-btn oz-small" data-action="newset-cancel">Cancel</button></span>
      </div>
      ${BUSINESS_FIELDS.map(
        (f) => `<label class="oz-field"><span>${escapeHtml(f.label)}</span><input type="text" data-field="${f.key}" value="${escapeHtml(fields[f.key] || "")}" placeholder="${escapeHtml(f.sample)}" autocomplete="off"></label>`
      ).join("")}
      <div class="oz-field"><span>Logo</span>
        <div class="oz-logo">${state.logo ? `<img src="${escapeHtml(state.logo.src)}" alt="Your logo">` : `<span class="oz-muted">No logo yet. Templates use your organization name instead.</span>`}</div>
        <span class="oz-btns">
          <label class="oz-btn oz-small oz-file">Choose a picture<input type="file" accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml" data-action="logo"></label>
          ${state.logo ? `<button type="button" class="oz-btn oz-small" data-action="remove-logo">Remove</button>` : ""}
        </span>
      </div>
      <h3 class="oz-h">Brand kit</h3>
      <p class="oz-muted">Save your colors, fonts, spacing, information sets, and logo as one file. Share it with your team, or open it on another computer.</p>
      <span class="oz-btns"><button type="button" class="oz-btn" data-action="export">Save brand kit</button>
      <label class="oz-btn oz-file">Open brand kit<input type="file" accept=".json,application/json" data-action="import"></label></span>
      <textarea class="oz-export" hidden readonly aria-label="Brand kit file contents"></textarea>`;
  }

  function exportKit() {
    const kit = { offizinBrandKit: 1, savedAt: new Date().toISOString(), theme: state.theme, sets: state.sets, activeSet: state.activeSet, logo: state.logo };
    const text = JSON.stringify(kit, null, 2);
    const showText = (msg) => {
      const ta = body.querySelector(".oz-export");
      ta.hidden = false;
      ta.value = text;
      ta.select();
      setStatus(msg, "info", true);
    };
    if (caps.noDownloads) return showText("Copy this text into a new file named offizin-brand-kit.json.");
    try {
      const blob = new Blob([text], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "offizin-brand-kit.json";
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        a.remove();
      }, 1000);
      setStatus("Brand kit saved to your downloads.", "success");
    } catch (e) {
      // Some editors block downloads from panels. Show the file so it can be copied instead.
      showText("Copy this text and save it as offizin-brand-kit.json.");
    }
  }

  async function importKit(file) {
    const text = await file.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      throw new Error("That file is not an Offizin brand kit.");
    }
    if (!data || data.offizinBrandKit !== 1) throw new Error("That file is not an Offizin brand kit.");
    state = sanitizeState({ ...data, v: STATE_VERSION });
    themeChanged();
    render();
    setStatus("Brand kit opened.", "success");
  }

  // Check -------------------------------------------------------------------

  /**
   * The sample text the checker searches for: the publication record's own samples (or, for a
   * record without them, the samples of its template in its theme; or, with no record, those
   * of every template), plus the samples of every page part, which can be inserted anywhere.
   * The hosts search the whole list, so nothing here is cut short.
   */
  function checkSamples(pub) {
    const theme = composeTheme();
    const set = new Set();
    let recordSamples = pub && Array.isArray(pub.samples) && pub.samples.length ? pub.samples : null;
    if (!recordSamples && pub && pub.templateId && findTemplate(pub.templateId)) {
      try {
        recordSamples = compose(findTemplate(pub.templateId), { ...theme, ...(pub.theme || {}) }, host.platform).samples;
      } catch (e) {
        recordSamples = null;
      }
    }
    if (recordSamples) for (const s of recordSamples) set.add(s);
    else for (const t of TEMPLATES) for (const s of compose(t, theme, host.platform).samples) set.add(s);
    for (const pt of PARTS) for (const s of composePart(pt, theme, host.platform).samples) set.add(s);
    return [...set];
  }

  const FOLD_NAMES = {
    trifold: "Tri-fold",
    halffold: "Half-fold (bi-fold, cards, programs)",
    booklet: "Booklet (folded and stapled)",
    tent: "Tent card (place cards)",
    envelope: "Envelope",
    cards: "Cards, labels, badges, and tickets",
    none: "Flat (flyers, posters, letters, resumes)",
  };

  function renderCheck() {
    const type = ui.foldType && FOLD_GUIDES[ui.foldType] && DIAGRAMS[ui.foldType] ? ui.foldType : "trifold";
    const guide = FOLD_GUIDES[type];
    const findings = ui.findings;
    const last = ui.lastPrint;
    const print = last ? last.print : { ...DEFAULT_PRINT, ...((findTemplate("trifold") || {}).print || {}) };
    const steps = printSteps(print, caps.printGuide || {}, { booklet: !!(last && last.booklet), mode: last ? last.mode : ui.mode });
    body.innerHTML = `
      <h3 class="oz-h">Design checker</h3>
      <p class="oz-muted">Finds sample text and placeholder pictures that still need your content, and business information you have not filled in.</p>
      <button type="button" class="oz-btn oz-primary oz-wide" data-action="check">Check this publication</button>
      ${
        findings
          ? findings.length
            ? `<ul class="oz-findings">${findings
                .map(
                  (f, i) => `<li><span class="oz-kind oz-kind-${f.kind}">${escapeHtml(f.label)}</span><span class="oz-ftext">${escapeHtml(f.text)}</span>${
                    host.select && f.selectable !== false ? `<button type="button" class="oz-btn oz-small" data-finding="${i}">Show</button>` : ""
                  }</li>`
                )
                .join("")}</ul>`
            : `<p class="oz-ok">No problems found. Your publication is ready to print.</p>`
          : ""
      }
      <h3 class="oz-h">Print</h3>
      <p class="oz-muted">${last ? `For the ${escapeHtml((findTemplate(last.templateId) || { name: "publication" }).name.toLowerCase())} you made:` : "Make a publication on the Create tab to see its own print settings. Until then, these are the steps for a tri-fold brochure."}</p>
      <dl class="oz-print"><dt>Sides</dt><dd>${escapeHtml(sidesText(print.duplex))}</dd><dt>Scale</dt><dd>Actual size (100 percent)</dd></dl>
      <ol class="oz-steps">${steps.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ol>
      <h3 class="oz-h">Fold and finish</h3>
      <label class="oz-field"><span>Publication type</span>
        <select data-action="fold-type">
          ${Object.entries(FOLD_NAMES)
            .filter(([k]) => FOLD_GUIDES[k] && DIAGRAMS[k])
            .map(([k, v]) => `<option value="${k}"${k === type ? " selected" : ""}>${v}</option>`)
            .join("")}
        </select>
      </label>
      <div class="oz-diagram">${DIAGRAMS[type]}</div>
      <p class="oz-strong">${escapeHtml(guide.title)}</p>
      <ol class="oz-steps">${guide.steps.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ol>
      <p class="oz-tip">${escapeHtml(guide.tip)}</p>
      ${caps.pdfHelp ? `<h3 class="oz-h">Save as PDF</h3><p class="oz-muted">${escapeHtml(caps.pdfHelp)}</p>` : ""}
      ${
        caps.finish
          ? `<h3 class="oz-h">Finish publication</h3>
             <p class="oz-muted">When you are done, Finish removes the invisible markers Offizin keeps around inserted page parts. Your content stays exactly as it is. Afterward, Offizin's Undo can no longer remove those parts.</p>
             <button type="button" class="oz-btn oz-wide" data-action="finish">Finish publication</button>`
          : ""
      }`;
  }

  async function runCheck() {
    const pub = await Promise.resolve()
      .then(() => (host.getPublication ? host.getPublication() : null))
      .catch(() => null);
    if (pub && pub.templateId) {
      const t = findTemplate(pub.templateId);
      if (t) {
        ui.foldType = t.fold;
        if (!ui.lastPrint || ui.lastPrint.templateId !== t.id) {
          try {
            const m = compose(t, composeTheme(), host.platform);
            // A record read from a document made with Open in a new document says mode "new".
            ui.lastPrint = { templateId: t.id, print: pub.print || m.print, booklet: !!m.booklet, mode: pub.mode || "replace" };
          } catch (e) {
            // The checker still works without print details.
          }
        }
      }
    }
    const samples = checkSamples(pub);
    const infoSamples = BUSINESS_FIELDS.map((f) => ({ key: f.key, label: f.label, sample: f.sample }));
    const found = await busy("Checking...", () => host.check({ samples, infoSamples }));
    if (found === FAILED) return;
    const list = Array.isArray(found) ? found : [];
    ui.findings = list.map((f) => ({
      ...f,
      label: f.kind === "sample" ? "Sample text" : f.kind === "picture" ? "Placeholder" : f.kind === "info" ? "Business info" : "Layout",
    }));
    keepFocus(renderCheck);
    setStatus(list.length ? `${list.length} item${list.length === 1 ? "" : "s"} to review.` : "No problems found.", list.length ? "info" : "success");
  }

  async function finish() {
    const r = await act("Finishing...", () => host.finish(), "Finished.");
    if (r === FAILED) return;
    let forgot = 0;
    if (history) {
      // Inserted parts lose the markers their undo needs.
      for (const e of history.entries) {
        if (e.kind === "insert" && history.forget(e.id)) forgot++;
      }
    }
    setStatus(r.message + (forgot ? ` ${forgot} inserted part${forgot === 1 ? " was" : "s were"} removed from Offizin's history.` : ""), "success", true);
  }

  // -------------------------------------------------------------------------
  // Events
  // -------------------------------------------------------------------------

  function selectTemplate(t) {
    ui.selected = t;
    render();
    body.scrollTop = 0;
  }

  async function handleAction(action, el) {
    const t = ui.selected;
    switch (action) {
      case "undo":
      case "redo":
        await requestHistory(action);
        return;
      case "history":
        toggleHistory();
        return;
      case "history-close":
        toggleHistory(false);
        root.querySelector(".oz-histbtn")?.focus();
        return;
      case "history-clear":
        askConfirm({
          title: "Clear Offizin's history?",
          text: "Your document stays as it is, but Offizin can no longer undo or redo the changes in the list.",
          ok: "Clear history",
          onOk: () => {
            try {
              history.clear();
              setStatus("Offizin's history is empty.");
            } catch (e) {
              setStatus(errorMessage(e), "error");
            }
          },
          onCancel: () => {},
        });
        return;
      case "confirm-ok":
        closeConfirm(true);
        return;
      case "confirm-cancel":
        closeConfirm(false);
        return;
      case "back":
        rememberOpen();
        ui.selected = null;
        render();
        return;
      case "create":
        await createSelected();
        return;
      case "restyle":
        await restyle();
        return;
      case "page-setup": {
        const setup = readPageSetup();
        const r = await act("Applying page setup...", () => host.applyPageSetup(setup), "Page setup applied.");
        if (r !== FAILED) setStatus(r.message, "success");
        return;
      }
      case "columns": {
        const count = Number(body.querySelector('[data-col="count"]').value);
        const line = body.querySelector('[data-col="line"]').checked;
        const r = await act("Applying columns...", () => host.applyColumns({ count, line }), `${count} column${count > 1 ? "s" : ""} applied.`);
        if (r !== FAILED) setStatus(r.message, "success");
        return;
      }
      case "layout-reset":
        if (!t) return;
        customFor(t).layout = {};
        refreshCustomize();
        schedulePreview(0);
        return;
      case "fit":
        if (!t) return;
        clearTimeout(previewTimer);
        shrinkToFit(t);
        return;
      case "custom-reset":
        if (!t) return;
        delete ui.custom[t.id];
        refreshCustomize();
        body.querySelectorAll("[data-option]").forEach((input) => {
          const o = (t.options || []).find((x) => x.key === input.dataset.option);
          if (!o) return;
          const v = optionValue(t, o);
          if (input.type === "checkbox") input.checked = v === true || v === "true";
          else input.value = v;
        });
        schedulePreview(0);
        setStatus("This publication is back to its design.");
        return;
      case "deflayout-reset":
        state.theme.layout = { ...DEFAULT_LAYOUT };
        themeChanged();
        keepFocus(() => {
          const box = body.querySelector("[data-deflayout-box]");
          if (box) box.innerHTML = layoutControlsHtml({ text: true, spacing: true }, state.theme.layout, "default");
        });
        setStatus("Spacing and text are back to each design's own settings.");
        return;
      case "merge-example":
        if (!t || !t.merge) return;
        {
          const text = recordsToText(exampleRecords(t.merge), t.merge.fields);
          const ta = body.querySelector("[data-merge]");
          if (ta) ta.value = text;
          setMergeText(t, text);
        }
        return;
      case "merge-clear":
        if (!t || !t.merge) return;
        {
          const ta = body.querySelector("[data-merge]");
          if (ta) ta.value = "";
          setMergeText(t, "");
        }
        return;
      case "finish":
        await finish();
        return;
      case "check":
        await runCheck();
        return;
      case "export":
        exportKit();
        return;
      case "remove-logo":
        state.logo = null;
        themeChanged();
        keepFocus(renderInfo, '[data-action="logo"]');
        return;
      case "new-set":
      case "rename-set": {
        const box = body.querySelector(".oz-newset");
        box.hidden = false;
        box.dataset.mode = action;
        box.querySelector("[data-newset-label]").textContent = action === "new-set" ? "Name for the new set" : "New name for this set";
        const input = box.querySelector("[data-newset]");
        input.value = action === "new-set" ? "" : state.activeSet;
        input.focus();
        return;
      }
      case "newset-cancel":
        body.querySelector(".oz-newset").hidden = true;
        return;
      case "newset-ok": {
        const box = body.querySelector(".oz-newset");
        const adding = box.dataset.mode === "new-set";
        // Runs of spaces collapse, so the name reads the same in the list and in the saved settings.
        const name = box.querySelector("[data-newset]").value.trim().replace(/\s+/g, " ").slice(0, 40);
        if (!name) return setStatus("Type a name first.", "error");
        if (BAD_KEYS.has(name)) return setStatus("Choose a different name.", "error");
        // New refuses every name in use, including the active set's; Rename may keep its own name.
        if (own(state.sets, name) && (adding || name !== state.activeSet)) return setStatus("A set with that name already exists.", "error");
        if (adding && Object.keys(state.sets).length >= MAX_SETS) return setStatus(`Offizin keeps up to ${MAX_SETS} information sets. Delete one first.`, "error");
        if (adding) {
          state.sets[name] = {};
        } else {
          const fields = state.sets[state.activeSet];
          delete state.sets[state.activeSet];
          state.sets[name] = fields;
        }
        state.activeSet = name;
        themeChanged();
        keepFocus(renderInfo, '[data-action="set"]');
        return;
      }
      case "delete-set":
        if (Object.keys(state.sets).length < 2) return;
        delete state.sets[state.activeSet];
        state.activeSet = Object.keys(state.sets)[0];
        themeChanged();
        keepFocus(renderInfo, '[data-action="set"]');
        setStatus("Information set deleted.");
        return;
      default:
    }
  }

  root.addEventListener("click", async (ev) => {
    const tab = ev.target.closest("[data-tab]");
    if (tab) {
      if (ui.tab === "create") rememberOpen();
      ui.tab = tab.dataset.tab;
      render();
      return;
    }
    const el = ev.target.closest("button, [data-template], [data-part]");
    if (!el || el.disabled) return;
    const t = ui.selected;

    if (el.dataset.paper) {
      state.theme.paper = el.dataset.paper;
      themeChanged();
      if (ui.tab === "create" && t) {
        refreshCustomize();
        schedulePreview(0);
      } else {
        keepFocus(render);
      }
    } else if (el.dataset.template) {
      selectTemplate(findTemplate(el.dataset.template));
    } else if (el.dataset.jump !== undefined) {
      await requestHistory("jump", Number(el.dataset.jump));
    } else if (el.dataset.forget) {
      try {
        history.forget(el.dataset.forget);
      } catch (e) {
        setStatus(errorMessage(e), "error");
      }
      // The list was redrawn without this entry, and the Remove button with it.
      if (!focusedControl()) focusHistoryControl(null);
    } else if (el.dataset.orient && t) {
      const natural = resolvePage(t, state.theme.paper, { ...composeOptions(t), orientation: null });
      customFor(t).orientation = el.dataset.orient === natural.orientation ? null : el.dataset.orient;
      refreshCustomize();
      schedulePreview(0);
    } else if (el.dataset.preset && t) {
      const c = customFor(t);
      const p = MARGIN_PRESETS.find((x) => x.id === el.dataset.preset);
      c.margins = p ? { ...p.m } : null;
      // Same on all sides now says whether the sides the preset gives really are the same.
      const m = resolvePage(t, state.theme.paper, composeOptions(t)).margins;
      c.sameMargins = ["right", "bottom", "left"].every((s) => Math.abs(m[s] - m.top) < 0.5);
      syncMarginInputs();
      schedulePreview(0);
    } else if (el.dataset.scheme) {
      state.theme.colorScheme = el.dataset.scheme;
      themeChanged();
      keepFocus(renderDesign);
      setStatus(`Color scheme: ${el.dataset.scheme === "custom" ? "your colors" : COLOR_SCHEMES.find((s) => s.id === el.dataset.scheme).name}.`);
    } else if (el.dataset.font) {
      state.theme.fontScheme = el.dataset.font;
      themeChanged();
      refreshFontList();
      if (el.dataset.font === "custom") {
        const d = body.querySelector('details[data-sec="fonts"]');
        if (d) d.open = true;
      }
      body.querySelector(`[data-font="${el.dataset.font}"]`)?.focus();
      setStatus(`Font scheme: ${el.dataset.font === "custom" ? "your fonts" : FONT_SCHEMES.find((f) => f.id === el.dataset.font).name}.`);
    } else if (el.dataset.part) {
      await insertPart(el.dataset.part);
    } else if (el.dataset.finding !== undefined) {
      const f = ui.findings[Number(el.dataset.finding)];
      const r = await busy("Finding it...", () => host.select(f));
      if (r !== FAILED) setStatus(host.kind === "demo" ? "Highlighted on the page." : "Selected in your document.");
    } else if (el.dataset.action) {
      await handleAction(el.dataset.action, el);
    }
  });

  root.addEventListener("change", async (ev) => {
    const el = ev.target;
    const t = ui.selected;
    if (el.dataset.action === "category") {
      ui.category = el.value;
      keepFocus(renderCreate);
    } else if (el.dataset.option && t) {
      customFor(t).options[el.dataset.option] = el.type === "checkbox" ? el.checked : el.value;
      schedulePreview(0);
    } else if (el.dataset.custom === "paperSize" && t) {
      const natural = resolvePage(t, state.theme.paper, { ...composeOptions(t), paperSize: null });
      customFor(t).paperSize = el.value === natural.sizeName ? null : el.value;
      refreshCustomize();
      schedulePreview(0);
    } else if (el.dataset.custom === "sameMargins" && t) {
      const c = customFor(t);
      c.sameMargins = el.checked;
      if (el.checked) {
        const page = resolvePage(t, state.theme.paper, composeOptions(t));
        const v = page.margins.top;
        c.margins = { top: v, right: v, bottom: v, left: v };
        syncMarginInputs();
        schedulePreview(0);
      }
    } else if (el.dataset.margin && t) {
      syncMarginInputs();
    } else if (el.dataset.print && t) {
      customFor(t).print[el.dataset.print] = el.checked;
      refreshCustomize();
      schedulePreview(0);
    } else if (el.name === "oz-mode") {
      ui.mode = el.value;
      const note = body.querySelector("[data-mode-note]");
      if (note) note.textContent = modeNote();
      refreshCustomize();
      schedulePreview(0);
    } else if (el.dataset.cfont) {
      const cf = state.theme.customFonts;
      if (el.dataset.cfont === "bold") {
        cf.headingBold = el.checked;
      } else {
        const slot = fontSlot();
        const other = slot === "word" ? "google" : "word";
        cf[el.dataset.cfont][slot] = el.value;
        const match = FONT_EQUIVALENTS[slot] && FONT_EQUIVALENTS[slot][el.value];
        if (match) cf[el.dataset.cfont][other] = match;
      }
      state.theme.fontScheme = "custom";
      themeChanged();
      refreshFontList();
      setStatus("Font scheme: your fonts.");
    } else if (el.dataset.action === "set") {
      if (!own(state.sets, el.value)) return;
      state.activeSet = el.value;
      themeChanged();
      keepFocus(renderInfo);
    } else if (el.dataset.action === "logo" && el.files && el.files[0]) {
      try {
        state.logo = await readLogoFile(el.files[0]);
        themeChanged();
        keepFocus(renderInfo);
        setStatus("Logo added. New publications will use it.", "success");
      } catch (e) {
        setStatus(errorMessage(e), "error");
      }
    } else if (el.dataset.action === "import" && el.files && el.files[0]) {
      try {
        await importKit(el.files[0]);
      } catch (e) {
        setStatus(errorMessage(e), "error");
      }
    } else if (el.dataset.action === "merge-file" && el.files && el.files[0] && t && t.merge) {
      const file = el.files[0];
      el.value = "";
      if (file.size > 2 * 1024 * 1024) {
        setStatus("That file is larger than 2 MB. Save just the columns Offizin needs, or paste the rows instead.", "error");
        return;
      }
      try {
        const text = await file.text();
        const ta = body.querySelector("[data-merge]");
        if (ta) ta.value = text;
        setMergeText(t, text);
        setStatus(`${file.name} opened.`, "success");
      } catch (e) {
        setStatus("That file could not be read: " + errorMessage(e), "error");
      }
    } else if (el.dataset.action === "fold-type") {
      ui.foldType = el.value;
      keepFocus(renderCheck);
    } else if (el.dataset.hex) {
      const v = el.value.trim();
      if (/^#?[0-9a-f]{6}$/i.test(v)) {
        state.theme.customColors[el.dataset.hex] = normalizeHex(v);
        state.theme.colorScheme = "custom";
        themeChanged();
        keepFocus(renderDesign);
      } else {
        setStatus("Use a six digit color code, for example #1B3A5C.", "error");
      }
    }
  });

  root.addEventListener("input", (ev) => {
    const el = ev.target;
    const t = ui.selected;
    if (el.dataset.field) {
      const set = state.sets[state.activeSet] || (state.sets[state.activeSet] = {});
      set[el.dataset.field] = el.value;
      themeChanged(250);
    } else if (el.dataset.action === "search") {
      ui.query = el.value;
      const gallery = body.querySelector("[data-gallery]");
      if (gallery) {
        gallery.innerHTML = cardsHtml();
        scheduleThumbs();
      }
    } else if (el.dataset.option && t && el.type !== "checkbox" && el.tagName !== "SELECT") {
      customFor(t).options[el.dataset.option] = el.value;
      schedulePreview(300);
    } else if (el.dataset.margin && t) {
      const c = customFor(t);
      const v = fromUnit(el.value);
      if (v === null) return;
      const page = resolvePage(t, state.theme.paper, composeOptions(t));
      const next = c.margins ? { ...c.margins } : { ...page.margins };
      if (c.sameMargins) {
        next.top = next.right = next.bottom = next.left = v;
      } else {
        next[el.dataset.margin] = v;
      }
      c.margins = next;
      if (c.sameMargins) syncMarginInputs();
      schedulePreview(250);
    } else if (el.dataset.layout && t) {
      const v = Number(el.value);
      customFor(t).layout[el.dataset.layout] = v;
      const out = body.querySelector(`[data-out="custom-${el.dataset.layout}"]`);
      if (out) out.textContent = pct(v);
      el.setAttribute("aria-valuetext", pct(v));
      const meta = body.querySelector("[data-layout-meta]");
      if (meta) meta.textContent = layoutMeta(effectiveLayout(t));
      const reset = body.querySelector('[data-action="layout-reset"]');
      if (reset) reset.disabled = false;
      schedulePreview(120);
    } else if (el.dataset.deflayout) {
      const v = Number(el.value);
      state.theme.layout = clampLayout({ ...(state.theme.layout || {}), [el.dataset.deflayout]: v });
      const out = body.querySelector(`[data-out="default-${el.dataset.deflayout}"]`);
      if (out) out.textContent = pct(v);
      el.setAttribute("aria-valuetext", pct(v));
      themeChanged(200);
    } else if (el.dataset.merge !== undefined && t && t.merge) {
      clearTimeout(mergeTimer);
      const text = el.value;
      mergeTimer = setTimeout(() => setMergeText(t, text), 250);
    } else if (el.dataset.color) {
      state.theme.customColors[el.dataset.color] = normalizeHex(el.value);
      state.theme.colorScheme = "custom";
      const hex = body.querySelector(`[data-hex="${el.dataset.color}"]`);
      if (hex) hex.value = normalizeHex(el.value);
      body.querySelectorAll("[data-scheme]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.scheme === "custom")));
      const chips = body.querySelectorAll('[data-scheme="custom"] i');
      ["primary", "accent", "dark"].forEach((k, i) => chips[i] && (chips[i].style.background = state.theme.customColors[k]));
      themeChanged(120);
    }
  });

  root.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") {
      if (ui.confirm) {
        ev.preventDefault();
        closeConfirm(false);
        return;
      }
      if (ui.historyOpen) {
        ev.preventDefault();
        toggleHistory(false);
        root.querySelector(".oz-histbtn")?.focus();
        return;
      }
    }
    const hk = history ? historyKey(ev) : null;
    if (hk) {
      ev.preventDefault();
      requestHistory(hk);
      return;
    }
    const tab = ev.target.closest && ev.target.closest('[role="tab"]');
    if (!tab || (ev.key !== "ArrowRight" && ev.key !== "ArrowLeft")) return;
    const i = TABS.findIndex((x) => x.id === tab.dataset.tab);
    const next = TABS[(i + (ev.key === "ArrowRight" ? 1 : TABS.length - 1)) % TABS.length];
    if (ui.tab === "create") rememberOpen();
    ui.tab = next.id;
    render();
    root.querySelector(`[data-tab="${next.id}"]`).focus();
  });

  let resizeTimer = null;
  if (typeof ResizeObserver !== "undefined") {
    let lastWidth = 0;
    new ResizeObserver(() => {
      const w = body.clientWidth;
      if (Math.abs(w - lastWidth) < 8) return;
      lastWidth = w;
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => {
        if (ui.tab === "create") ui.selected ? renderLarge() : scheduleThumbs();
      }, 150);
    }).observe(body);
  }

  // -------------------------------------------------------------------------
  // Start
  // -------------------------------------------------------------------------

  if (history) history.onChange(syncHistory);
  render();
  // Settings changes made before the saved settings arrive wait for them (see save()).
  statusEl.textContent = LOADING_TEXT;
  loadSaved();

  const api = {
    get state() {
      return state;
    },
    composeTheme,
    setStatus,
    showTab(id) {
      if (!TABS.some((x) => x.id === id)) return;
      ui.tab = id;
      render();
    },
    openTemplate(id) {
      const t = findTemplate(id);
      if (!t) return;
      ui.tab = "create";
      ui.selected = t;
      render();
    },
    history,
  };
  if (history) {
    api.undo = () => requestHistory("undo", undefined, true);
    api.redo = () => requestHistory("redo", undefined, true);
    api.onHistoryChange = (fn) => history.onChange(fn);
  }
  return api;
}
