/**
 * The Offizin panel: one user interface for the Word task pane, the Google Docs sidebar,
 * and the live demo on the website.
 *
 * The panel never touches a document directly. It talks to a "host" object that knows how
 * to create publications, insert parts, change page setup, restyle, and check a document in
 * its own editor (see app/hosts/). That keeps the three versions identical for users.
 *
 * Tabs follow the way Publisher worked:
 *   Create  pick a publication from the gallery
 *   Design  color scheme, font scheme, page setup, and restyling
 *   Insert  page parts (building blocks) placed at the cursor
 *   Info    business information sets, logo, and a shareable brand kit file
 *   Check   the design checker plus printing and folding help
 */
import { TEMPLATES, CATEGORIES, PARTS, FOLD_GUIDES, findTemplate, findPart } from "../core/templates.js";
import { compose, composePart, BUSINESS_FIELDS, restyleMap } from "../core/compose.js";
import { COLOR_SCHEMES, FONT_SCHEMES, fontsFor } from "../core/schemes.js";
import { PAPER, MONTHS, escapeHtml, normalizeHex, round, inch, mm } from "../core/util.js";
import { renderPreview, blocksHtml, ensurePreviewCss } from "../render/preview.js";
import { readLogoFile } from "../render/pictures.js";
import { DIAGRAMS } from "./diagrams.js";
import { MARK_SVG } from "./mark.js";

const STATE_VERSION = 1;

/** Returned by busy() when an action failed, so a successful action that returns nothing still counts. */
const FAILED = Symbol("failed");
const TABS = [
  { id: "create", name: "Create" },
  { id: "design", name: "Design" },
  { id: "insert", name: "Insert" },
  { id: "info", name: "Info" },
  { id: "check", name: "Check" },
];

const PAGE_PRESETS = [
  ["letter", "Letter (8.5 x 11 in)"],
  ["legal", "Legal (8.5 x 14 in)"],
  ["tabloid", "Tabloid (11 x 17 in)"],
  ["halfLetter", "Half letter (5.5 x 8.5 in)"],
  ["card5x7", "Card (5 x 7 in)"],
  ["a3", "A3 (297 x 420 mm)"],
  ["a4", "A4 (210 x 297 mm)"],
  ["a5", "A5 (148 x 210 mm)"],
];

function guessPaper() {
  try {
    const lang = (navigator.language || "en-US").toUpperCase();
    return /-(US|CA|MX|PH|CL|CO|VE|GT|CR|SV|PR)$/.test(lang) ? "letter" : "a4";
  } catch (e) {
    return "letter";
  }
}

export function defaultState() {
  return {
    v: STATE_VERSION,
    theme: { colorScheme: "harbor", customColors: { primary: "#1B3A5C", accent: "#C0392B", dark: "#1C2430" }, fontScheme: "editorial", paper: guessPaper() },
    sets: { Main: {} },
    activeSet: "Main",
    logo: null,
  };
}

/** Accepts saved state or an imported brand kit, keeping only fields Offizin understands. */
export function sanitizeState(input) {
  const base = defaultState();
  if (!input || typeof input !== "object") return base;
  const t = input.theme || {};
  const out = { ...base, theme: { ...base.theme } };
  if (typeof t.colorScheme === "string" && (t.colorScheme === "custom" || COLOR_SCHEMES.some((s) => s.id === t.colorScheme))) out.theme.colorScheme = t.colorScheme;
  if (t.customColors && typeof t.customColors === "object") {
    out.theme.customColors = {
      primary: normalizeHex(t.customColors.primary || base.theme.customColors.primary),
      accent: normalizeHex(t.customColors.accent || base.theme.customColors.accent),
      dark: normalizeHex(t.customColors.dark || base.theme.customColors.dark),
    };
  }
  if (FONT_SCHEMES.some((f) => f.id === t.fontScheme)) out.theme.fontScheme = t.fontScheme;
  if (t.paper === "letter" || t.paper === "a4") out.theme.paper = t.paper;
  if (input.sets && typeof input.sets === "object") {
    const sets = {};
    for (const [name, fields] of Object.entries(input.sets).slice(0, 20)) {
      if (!fields || typeof fields !== "object") continue;
      const clean = {};
      for (const f of BUSINESS_FIELDS) if (typeof fields[f.key] === "string") clean[f.key] = fields[f.key].slice(0, 300);
      sets[String(name).slice(0, 40)] = clean;
    }
    if (Object.keys(sets).length) out.sets = sets;
  }
  out.activeSet = out.sets[input.activeSet] ? input.activeSet : Object.keys(out.sets)[0];
  if (input.logo && typeof input.logo.src === "string" && /^data:image\/(png|jpeg|gif|webp);base64,/.test(input.logo.src)) {
    out.logo = { src: input.logo.src, width: Number(input.logo.width) || 200, height: Number(input.logo.height) || 100 };
  }
  return out;
}

/**
 * Mounts the panel.
 * host: see app/hosts/demo.js for the full interface.
 * opts.siteUrl: link target for Help.
 */
export function mountPanel(root, host, opts = {}) {
  ensurePreviewCss();
  const caps = host.capabilities || {};
  let state = defaultState();
  const ui = { tab: "create", category: "all", selected: null, options: {}, mode: (caps.createModes || [{ id: "replace" }])[0].id, findings: null, foldType: null };
  let saveTimer = null;
  let thumbTimer = null;
  let statusTimer = null;

  root.classList.add("oz");
  root.innerHTML = `
    <header class="oz-head">
      <span class="oz-mark">${MARK_SVG}</span>
      <span class="oz-wordmark">Offizin</span>
      <span class="oz-host">${escapeHtml(host.label || "")}</span>
    </header>
    <nav class="oz-tabs" role="tablist" aria-label="Offizin">
      ${TABS.map((t) => `<button type="button" role="tab" id="oz-tab-${t.id}" aria-controls="oz-panel" data-tab="${t.id}" aria-selected="${t.id === ui.tab}">${t.name}</button>`).join("")}
    </nav>
    <main class="oz-body" id="oz-panel" role="tabpanel" tabindex="-1"></main>
    <footer class="oz-foot">
      <div class="oz-status" role="status" aria-live="polite">Ready</div>
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
      paper: state.theme.paper,
      business: state.sets[state.activeSet] || {},
      logo: state.logo,
    };
  }

  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      Promise.resolve(host.saveState(state)).catch((e) => setStatus("Your settings could not be saved: " + message(e), "error"));
    }, 350);
  }

  function themeChanged() {
    save();
    if (host.themeChanged) host.themeChanged(composeTheme());
  }

  function message(e) {
    return (e && (e.message || e.toString())) || "Something went wrong.";
  }

  function setStatus(text, kind = "info", sticky = false) {
    clearTimeout(statusTimer);
    statusEl.textContent = text;
    statusEl.dataset.kind = kind;
    if (!sticky && kind !== "error") statusTimer = setTimeout(() => {
      statusEl.textContent = "Ready";
      statusEl.dataset.kind = "info";
    }, 6000);
  }

  async function busy(text, fn) {
    setStatus(text, "busy", true);
    root.classList.add("oz-busy");
    try {
      const result = await fn();
      return result;
    } catch (e) {
      console.error(e);
      setStatus(message(e), "error", true);
      return FAILED;
    } finally {
      root.classList.remove("oz-busy");
    }
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

  // Create ------------------------------------------------------------------

  function renderCreate() {
    if (ui.selected) return renderDetail();
    const list = TEMPLATES.filter((t) => ui.category === "all" || t.category === ui.category);
    body.innerHTML = `
      <div class="oz-toolbar">
        <div class="oz-field-inline"><span id="oz-paper-label">Paper</span>
          <span class="oz-seg" role="group" aria-labelledby="oz-paper-label">
            <button type="button" data-paper="letter" aria-pressed="${state.theme.paper === "letter"}">Letter</button>
            <button type="button" data-paper="a4" aria-pressed="${state.theme.paper === "a4"}">A4</button>
          </span>
        </div>
        <label class="oz-field-inline"><span>Show</span>
          <select data-action="category">${CATEGORIES.map((c) => `<option value="${c.id}"${c.id === ui.category ? " selected" : ""}>${escapeHtml(c.name)}</option>`).join("")}</select>
        </label>
      </div>
      <div class="oz-gallery">
        ${list
          .map(
            (t) => `<button type="button" class="oz-card" data-template="${t.id}" aria-label="${escapeHtml(t.name)}: ${escapeHtml(t.sizeLabel)}">
              <span class="oz-thumb" data-thumb="${t.id}"></span>
              <span class="oz-card-name">${escapeHtml(t.name)}</span>
              <span class="oz-card-meta">${escapeHtml(t.sizeLabel)}</span>
            </button>`
          )
          .join("")}
      </div>`;
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
          const model = compose(t, theme, host.platform);
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

  function optionValue(t, o) {
    if (ui.options[o.key] !== undefined) return ui.options[o.key];
    return typeof o.default === "function" ? o.default() : o.default;
  }

  function renderDetail() {
    const t = ui.selected;
    const modes = caps.createModes || [{ id: "replace", label: "Replace this document" }];
    const opts = (t.options || [])
      .map((o) => {
        const v = optionValue(t, o);
        if (o.type === "month") {
          return `<label class="oz-field"><span>${escapeHtml(o.label)}</span><select data-option="${o.key}">${MONTHS.map((m, i) => `<option value="${i + 1}"${Number(v) === i + 1 ? " selected" : ""}>${m}</option>`).join("")}</select></label>`;
        }
        return `<label class="oz-field"><span>${escapeHtml(o.label)}</span><input type="number" data-option="${o.key}" value="${escapeHtml(v)}" min="1900" max="2200"></label>`;
      })
      .join("");
    body.innerHTML = `
      <button type="button" class="oz-back" data-action="back">‹ All publications</button>
      <h2 class="oz-title">${escapeHtml(t.name)}</h2>
      <p class="oz-muted">${escapeHtml(t.description)}</p>
      <p class="oz-meta">${escapeHtml(t.sizeLabel)}</p>
      <div class="oz-large" aria-label="Preview of ${escapeHtml(t.name)}"></div>
      ${opts ? `<div class="oz-options">${opts}</div>` : ""}
      ${
        modes.length > 1
          ? `<fieldset class="oz-modes"><legend>Where</legend>${modes
              .map((m) => `<label class="oz-radio"><input type="radio" name="oz-mode" value="${m.id}"${m.id === ui.mode ? " checked" : ""}><span>${escapeHtml(m.label)}</span></label>`)
              .join("")}</fieldset>`
          : ""
      }
      <button type="button" class="oz-btn oz-primary oz-wide" data-action="create">${escapeHtml(caps.createLabel || "Create publication")}</button>
      <p class="oz-note">${modes.find((m) => m.id === ui.mode)?.note || caps.createNote || ""}</p>`;
    renderLarge();
  }

  function renderLarge() {
    const box = body.querySelector(".oz-large");
    if (!box || !ui.selected) return;
    const model = compose(ui.selected, composeTheme(), host.platform, ui.options);
    renderPreview(box, model, { width: innerWidth(), labels: true, gap: 10 });
  }

  async function createSelected() {
    const t = ui.selected;
    const model = compose(t, composeTheme(), host.platform, ui.options);
    const result = await busy(`Creating ${t.name.toLowerCase()}...`, () => host.createPublication(model, ui.mode, { options: ui.options }));
    if (result !== FAILED) {
      ui.foldType = t.fold;
      setStatus(typeof result === "string" ? result : `${t.name} created.`, "success", typeof result === "string" && result.length > 60);
    }
  }

  // Design ------------------------------------------------------------------

  function renderDesign() {
    const th = state.theme;
    const fontsNow = (f) => fontsFor(f, host.platform);
    const unitIn = th.paper === "letter";
    const unit = unitIn ? "in" : "mm";
    const toUnit = (pt) => (unitIn ? round(pt / 72, 2) : round(pt / (72 / 25.4), 1));
    const presetSize = th.paper === "letter" ? "letter" : "a4";
    const margin = 36;
    body.innerHTML = `
      <h3 class="oz-h">Color scheme</h3>
      <div class="oz-swatches">
        ${COLOR_SCHEMES.map(
          (s) => `<button type="button" class="oz-swatch" data-scheme="${s.id}" aria-pressed="${th.colorScheme === s.id}" title="${escapeHtml(s.name)}">
            <span class="oz-chips"><i style="background:${s.primary}"></i><i style="background:${s.accent}"></i><i style="background:${s.dark}"></i></span>
            <span class="oz-swatch-name">${escapeHtml(s.name)}</span></button>`
        ).join("")}
        <button type="button" class="oz-swatch" data-scheme="custom" aria-pressed="${th.colorScheme === "custom"}">
          <span class="oz-chips"><i style="background:${th.customColors.primary}"></i><i style="background:${th.customColors.accent}"></i><i style="background:${th.customColors.dark}"></i></span>
          <span class="oz-swatch-name">Your colors</span></button>
      </div>
      <details class="oz-custom"${th.colorScheme === "custom" ? " open" : ""}>
        <summary>Choose your own colors</summary>
        ${[["primary", "Main color"], ["accent", "Accent color"], ["dark", "Text color"]]
          .map(
            ([k, label]) => `<label class="oz-color"><span>${label}</span><input type="color" data-color="${k}" value="${th.customColors[k]}"><input type="text" data-hex="${k}" value="${th.customColors[k]}" maxlength="7" spellcheck="false" aria-label="${label} hex code"></label>`
          )
          .join("")}
      </details>

      <h3 class="oz-h">Font scheme</h3>
      <div class="oz-fontlist">
        ${FONT_SCHEMES.map((f) => {
          const fn = fontsNow(f);
          return `<button type="button" class="oz-fontopt" data-font="${f.id}" aria-pressed="${th.fontScheme === f.id}">
            <span class="oz-fontname" style="font-family:${escapeHtml(fn.heading.css)};font-weight:${f.headingBold ? 700 : 400}">${escapeHtml(f.name)}</span>
            <span class="oz-fontmeta" style="font-family:${escapeHtml(fn.body.css)}">${escapeHtml(fn.heading.name)} and ${escapeHtml(fn.body.name)}</span></button>`;
        }).join("")}
      </div>

      ${
        caps.restyle
          ? `<h3 class="oz-h">Update this publication</h3>
             <p class="oz-muted">Switch the colors and fonts of a publication you already made, and keep your text.</p>
             <button type="button" class="oz-btn oz-wide" data-action="restyle">Apply this design to the publication</button>`
          : ""
      }

      <h3 class="oz-h">Page setup</h3>
      ${
        caps.pageSetup === true
          ? `<label class="oz-field"><span>Size</span><select data-ps="size">${PAGE_PRESETS.map(([id, label]) => `<option value="${id}"${id === presetSize ? " selected" : ""}>${label}</option>`).join("")}</select></label>
             <label class="oz-field"><span>Orientation</span><select data-ps="orientation"><option value="portrait">Portrait</option><option value="landscape">Landscape</option></select></label>
             <div class="oz-margins">
               ${["top", "bottom", "left", "right"].map((s) => `<label class="oz-field"><span>${s[0].toUpperCase() + s.slice(1)} (${unit})</span><input type="number" step="${unitIn ? 0.05 : 1}" min="0" data-ps="${s}" value="${toUnit(margin)}"></label>`).join("")}
             </div>
             <button type="button" class="oz-btn oz-wide" data-action="page-setup">Apply page setup</button>`
          : `<p class="oz-muted">${caps.pageSetupHelp || ""}</p>`
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
          ? `<h3 class="oz-h">Columns</h3><p class="oz-muted">${caps.columnsHelp}</p>`
          : ""
      }`;
  }

  function readPageSetup() {
    const q = (k) => body.querySelector(`[data-ps="${k}"]`);
    const unitIn = state.theme.paper === "letter";
    const toPt = (v) => (unitIn ? inch(Number(v) || 0) : mm(Number(v) || 0));
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
    const pub = await Promise.resolve(host.getPublication ? host.getPublication() : null).catch(() => null);
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
    const result = await busy("Applying colors and fonts...", () => host.restyle(map, { colorScheme: to.colorScheme, customColors: to.customColors, fontScheme: to.fontScheme }));
    if (result !== FAILED) setStatus(typeof result === "string" ? result : "Design applied. Use Undo if you want the old look back.", "success");
  }

  // Insert ------------------------------------------------------------------

  function renderInsert() {
    body.innerHTML = `
      <p class="oz-muted">${escapeHtml(caps.insertHelp || "Click a part to add it after the cursor. Parts use your current colors, fonts, and business information.")}</p>
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
    const width = await Promise.resolve(host.contentWidth ? host.contentWidth() : 468).catch(() => 468);
    const composed = composePart(part, composeTheme(), host.platform, width || 468);
    const result = await busy(`Inserting ${part.name.toLowerCase()}...`, () => host.insertPart(composed, width || 468, part.id));
    if (result !== FAILED) setStatus(typeof result === "string" ? result : `${part.name} inserted.`, "success");
  }

  // Info --------------------------------------------------------------------

  function renderInfo() {
    const fields = state.sets[state.activeSet] || {};
    body.innerHTML = `
      <p class="oz-muted">Publications fill in this information for you, the way business information sets worked in Publisher.</p>
      <div class="oz-row2">
        <label class="oz-field"><span>Information set</span>
          <select data-action="set">${Object.keys(state.sets).map((n) => `<option${n === state.activeSet ? " selected" : ""}>${escapeHtml(n)}</option>`).join("")}</select>
        </label>
        <span class="oz-btns"><button type="button" class="oz-btn oz-small" data-action="new-set">New</button><button type="button" class="oz-btn oz-small" data-action="rename-set">Rename</button><button type="button" class="oz-btn oz-small" data-action="delete-set"${Object.keys(state.sets).length < 2 ? " disabled" : ""}>Delete</button></span>
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
      <p class="oz-muted">Save your colors, fonts, information sets, and logo as one file. Share it with your team, or open it on another computer.</p>
      <span class="oz-btns"><button type="button" class="oz-btn" data-action="export">Save brand kit</button>
      <label class="oz-btn oz-file">Open brand kit<input type="file" accept=".json,application/json" data-action="import"></label></span>
      <textarea class="oz-export" hidden readonly aria-label="Brand kit file contents"></textarea>`;
  }

  function exportKit() {
    const kit = { offizinBrandKit: 1, savedAt: new Date().toISOString(), theme: state.theme, sets: state.sets, activeSet: state.activeSet, logo: state.logo };
    const text = JSON.stringify(kit, null, 2);
    if (caps.noDownloads) {
      const ta = body.querySelector(".oz-export");
      ta.hidden = false;
      ta.value = text;
      ta.select();
      setStatus("Copy this text into a new file named offizin-brand-kit.json.", "info", true);
      return;
    }
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
      const ta = body.querySelector(".oz-export");
      ta.hidden = false;
      ta.value = text;
      ta.select();
      setStatus("Copy this text and save it as offizin-brand-kit.json.", "info", true);
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

  function allSamples() {
    const theme = composeTheme();
    const set = new Set();
    for (const t of TEMPLATES) for (const s of compose(t, theme, host.platform).samples) set.add(s);
    for (const pt of PARTS) for (const s of composePart(pt, theme, host.platform).samples) set.add(s);
    return [...set];
  }

  function renderCheck() {
    const type = ui.foldType || "trifold";
    const guide = FOLD_GUIDES[type];
    const findings = ui.findings;
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
      <h3 class="oz-h">Print and fold</h3>
      <label class="oz-field"><span>Publication type</span>
        <select data-action="fold-type">
          ${Object.entries({ trifold: "Tri-fold", halffold: "Half-fold (bi-fold, cards, programs)", cards: "Cards and labels", none: "Flat (flyers, posters, letters)" })
            .map(([k, v]) => `<option value="${k}"${k === type ? " selected" : ""}>${v}</option>`)
            .join("")}
        </select>
      </label>
      <div class="oz-diagram">${DIAGRAMS[type]}</div>
      <p class="oz-strong">${escapeHtml(guide.title)}</p>
      <ol class="oz-steps">${guide.steps.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ol>
      <p class="oz-tip">${escapeHtml(guide.tip)}</p>
      ${caps.pdfHelp ? `<h3 class="oz-h">Save as PDF</h3><p class="oz-muted">${caps.pdfHelp}</p>` : ""}`;
  }

  async function runCheck() {
    const pub = await Promise.resolve(host.getPublication ? host.getPublication() : null).catch(() => null);
    if (pub && pub.templateId) {
      const t = findTemplate(pub.templateId);
      if (t) ui.foldType = t.fold;
    }
    const samples = pub && pub.samples && pub.samples.length ? pub.samples : allSamples();
    const infoSamples = BUSINESS_FIELDS.map((f) => ({ key: f.key, label: f.label, sample: f.sample }));
    const found = await busy("Checking...", () => host.check({ samples, infoSamples }));
    if (found === FAILED) return;
    ui.findings = found.map((f) => ({
      ...f,
      label: f.kind === "sample" ? "Sample text" : f.kind === "picture" ? "Placeholder" : f.kind === "info" ? "Business info" : "Layout",
    }));
    renderCheck();
    setStatus(found.length ? `${found.length} item${found.length === 1 ? "" : "s"} to review.` : "No problems found.", found.length ? "info" : "success");
  }

  // -------------------------------------------------------------------------
  // Events
  // -------------------------------------------------------------------------

  root.addEventListener("click", async (ev) => {
    const tab = ev.target.closest("[data-tab]");
    if (tab) {
      ui.tab = tab.dataset.tab;
      render();
      return;
    }
    const el = ev.target.closest("button, [data-template], [data-part]");
    if (!el || el.disabled) return;

    if (el.dataset.paper) {
      state.theme.paper = el.dataset.paper;
      themeChanged();
      render();
    } else if (el.dataset.template) {
      ui.selected = findTemplate(el.dataset.template);
      ui.options = {};
      render();
      body.scrollTop = 0;
    } else if (el.dataset.scheme) {
      state.theme.colorScheme = el.dataset.scheme;
      themeChanged();
      renderDesign();
      setStatus(`Color scheme: ${el.dataset.scheme === "custom" ? "your colors" : COLOR_SCHEMES.find((s) => s.id === el.dataset.scheme).name}.`);
    } else if (el.dataset.font) {
      state.theme.fontScheme = el.dataset.font;
      themeChanged();
      renderDesign();
      setStatus(`Font scheme: ${FONT_SCHEMES.find((f) => f.id === el.dataset.font).name}.`);
    } else if (el.dataset.part) {
      await insertPart(el.dataset.part);
    } else if (el.dataset.finding !== undefined) {
      const f = ui.findings[Number(el.dataset.finding)];
      const r = await busy("Finding it...", () => host.select(f));
      if (r !== FAILED) setStatus(host.kind === "demo" ? "Highlighted on the page." : "Selected in your document.");
    } else {
      const action = el.dataset.action;
      if (action === "back") {
        ui.selected = null;
        render();
      } else if (action === "create") {
        await createSelected();
      } else if (action === "restyle") {
        await restyle();
      } else if (action === "page-setup") {
        const setup = readPageSetup();
        const r = await busy("Applying page setup...", () => host.applyPageSetup(setup));
        if (r !== FAILED) setStatus("Page setup applied.", "success");
      } else if (action === "columns") {
        const count = Number(body.querySelector('[data-col="count"]').value);
        const line = body.querySelector('[data-col="line"]').checked;
        const r = await busy("Applying columns...", () => host.applyColumns({ count, line }));
        if (r !== FAILED) setStatus(`${count} column${count > 1 ? "s" : ""} applied.`, "success");
      } else if (action === "new-set" || action === "rename-set") {
        const box = body.querySelector(".oz-newset");
        box.hidden = false;
        box.dataset.mode = action;
        box.querySelector("[data-newset-label]").textContent = action === "new-set" ? "Name for the new set" : "New name for this set";
        const input = box.querySelector("[data-newset]");
        input.value = action === "new-set" ? "" : state.activeSet;
        input.focus();
      } else if (action === "newset-cancel") {
        body.querySelector(".oz-newset").hidden = true;
      } else if (action === "newset-ok") {
        const box = body.querySelector(".oz-newset");
        const name = box.querySelector("[data-newset]").value.trim().slice(0, 40);
        if (!name) return setStatus("Type a name first.", "error");
        if (state.sets[name] && name !== state.activeSet) return setStatus("A set with that name already exists.", "error");
        if (box.dataset.mode === "new-set") {
          state.sets[name] = {};
        } else {
          const fields = state.sets[state.activeSet];
          delete state.sets[state.activeSet];
          state.sets[name] = fields;
        }
        state.activeSet = name;
        themeChanged();
        renderInfo();
      } else if (action === "delete-set") {
        if (Object.keys(state.sets).length < 2) return;
        delete state.sets[state.activeSet];
        state.activeSet = Object.keys(state.sets)[0];
        themeChanged();
        renderInfo();
        setStatus("Information set deleted.");
      } else if (action === "remove-logo") {
        state.logo = null;
        themeChanged();
        renderInfo();
      } else if (action === "export") {
        exportKit();
      } else if (action === "check") {
        await runCheck();
      }
    }
  });

  root.addEventListener("change", async (ev) => {
    const el = ev.target;
    if (el.dataset.action === "category") {
      ui.category = el.value;
      renderCreate();
    } else if (el.dataset.option) {
      ui.options[el.dataset.option] = el.value;
      renderLarge();
    } else if (el.name === "oz-mode") {
      ui.mode = el.value;
      const note = body.querySelector(".oz-note");
      const modes = caps.createModes || [];
      if (note) note.innerHTML = modes.find((m) => m.id === ui.mode)?.note || caps.createNote || "";
    } else if (el.dataset.action === "set") {
      state.activeSet = el.value;
      themeChanged();
      renderInfo();
    } else if (el.dataset.action === "logo" && el.files && el.files[0]) {
      try {
        state.logo = await readLogoFile(el.files[0]);
        themeChanged();
        renderInfo();
        setStatus("Logo added. New publications will use it.", "success");
      } catch (e) {
        setStatus(message(e), "error");
      }
    } else if (el.dataset.action === "import" && el.files && el.files[0]) {
      try {
        await importKit(el.files[0]);
      } catch (e) {
        setStatus(message(e), "error");
      }
    } else if (el.dataset.action === "fold-type") {
      ui.foldType = el.value;
      renderCheck();
    } else if (el.dataset.hex) {
      const v = el.value.trim();
      if (/^#?[0-9a-f]{6}$/i.test(v)) {
        state.theme.customColors[el.dataset.hex] = normalizeHex(v);
        state.theme.colorScheme = "custom";
        themeChanged();
        renderDesign();
      } else {
        setStatus("Use a six digit color code, for example #1B3A5C.", "error");
      }
    }
  });

  root.addEventListener("input", (ev) => {
    const el = ev.target;
    if (el.dataset.field) {
      const set = state.sets[state.activeSet] || (state.sets[state.activeSet] = {});
      set[el.dataset.field] = el.value;
      themeChanged();
    } else if (el.dataset.color) {
      state.theme.customColors[el.dataset.color] = normalizeHex(el.value);
      state.theme.colorScheme = "custom";
      const hex = body.querySelector(`[data-hex="${el.dataset.color}"]`);
      if (hex) hex.value = normalizeHex(el.value);
      body.querySelectorAll('[data-scheme]').forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.scheme === "custom")));
      const chips = body.querySelectorAll('[data-scheme="custom"] i');
      ["primary", "accent", "dark"].forEach((k, i) => chips[i] && (chips[i].style.background = state.theme.customColors[k]));
      themeChanged();
    }
  });

  root.addEventListener("keydown", (ev) => {
    const tab = ev.target.closest && ev.target.closest('[role="tab"]');
    if (!tab || (ev.key !== "ArrowRight" && ev.key !== "ArrowLeft")) return;
    const i = TABS.findIndex((t) => t.id === tab.dataset.tab);
    const next = TABS[(i + (ev.key === "ArrowRight" ? 1 : TABS.length - 1)) % TABS.length];
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

  render();
  Promise.resolve(host.loadState())
    .then((saved) => {
      if (saved) {
        state = sanitizeState(saved);
        render();
      }
      if (host.themeChanged) host.themeChanged(composeTheme());
    })
    .catch((e) => setStatus("Saved settings could not be loaded: " + message(e), "error"));

  return {
    get state() {
      return state;
    },
    composeTheme,
    setStatus,
    showTab(id) {
      ui.tab = id;
      render();
    },
    openTemplate(id) {
      const t = findTemplate(id);
      if (!t) return;
      ui.tab = "create";
      ui.selected = t;
      ui.options = {};
      render();
    },
  };
}
