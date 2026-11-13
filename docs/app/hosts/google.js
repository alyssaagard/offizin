/**
 * The Google Docs host, running inside the Offizin sidebar.
 *
 * The sidebar composes publications and draws picture placeholders in the browser, then
 * hands the finished model to the Apps Script functions in Code.gs (google.script.run),
 * which build it with DocumentApp. Nothing leaves Google: there is no Offizin server.
 *
 * Settings are saved with PropertiesService (per user). Logos are too large for that, so the
 * logo is kept in the sidebar's browser storage when the browser allows it.
 *
 * Offizin's own history (contract version 2):
 *   create     Code.gs keeps a copy of the previous document in the script cache for six hours;
 *              Undo restores it, and Redo builds the publication again from the same design.
 *              When building fails, Code.gs puts the copy back itself; when Apps Script stops
 *              the script, recoverCreate below does it from the pending record Code.gs left
 *   insert     the part is a named range; Undo removes it (keeping what it held, so Redo puts
 *              back exactly that) and leaves a tiny anchor for Redo (or, when Undo and Redo of
 *              the publication removed the anchor, Redo uses the part's position)
 *   restyle    Undo applies the reverse color and font map
 *   pageSetup  Undo restores the page size and margins read before the change
 */
/* global google */
import { walkBlocks, invertRestyleMap } from "../core/compose.js";
import { prepareImages } from "../render/pictures.js";
import { mountPanel } from "../ui/panel.js";

const LOGO_KEY = "offizin.google.logo";

function call(fn, ...args) {
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler(resolve)
      .withFailureHandler((e) => reject(e instanceof Error ? e : new Error((e && e.message) || String(e))))
      [fn](...args);
  });
}

/** Adds picture data to image blocks and drops the logo data URL, which the server does not need twice. */
function withImages(blocks) {
  const copy = JSON.parse(JSON.stringify(blocks));
  const lookup = prepareImages(copy);
  walkBlocks(copy, (b) => {
    if (b.t !== "image") return;
    const img = lookup(b);
    if (img) b.image = img;
    delete b.src;
  });
  return copy;
}

/** The model the server needs: without merge records (they can be large) or booklet page copies. */
function serverModel(model) {
  const { pages, options, ...rest } = model;
  return { ...rest, blocks: withImages(model.blocks) };
}

let logoInMemory = null;
let opCounter = 0;
const newId = () => `${Date.now().toString(36)}${(opCounter++).toString(36)}`;

/**
 * After ozCreate failed: Code.gs puts the document back itself when building throws, but when
 * Apps Script stops the script (its time limit, for example) only its pending record remains.
 * Restores from that record and returns the error to show.
 */
async function recoverCreate(error, run) {
  const reason = ((error && error.message) || String(error)).replace(/\.$/, "");
  let res;
  try {
    res = await call("ozRecoverCreate", run);
  } catch (e) {
    return new Error(`Offizin could not finish the publication (${reason}). ${(e && e.message) || "Use File, then Version history, to go back."}`);
  }
  if (res && res.restored) return new Error(`Offizin could not finish the publication, so your document is back as it was. (${reason})`);
  return error instanceof Error ? error : new Error(reason);
}

/** Plain words for what a snapshot could not keep. */
function unsupportedNote(list, where = "the old document") {
  const kinds = [...new Set(list || [])];
  if (!kinds.length) return "";
  return ` Offizin's Undo cannot bring back ${kinds.join(", ")} from ${where}; use File, then Version history, for those.`;
}

export function createGoogleHost() {
  const host = {
    kind: "google",
    platform: "google",
    label: "for Google Docs",
    capabilities: {
      createModes: [
        {
          id: "replace",
          label: "Replace this document",
          note: "Replaces everything in this document. Offizin's Undo brings it back for up to six hours. If the document is pageless, choose File, then Page setup, then Pages first.",
        },
      ],
      pageSetup: true,
      columns: false,
      columnsHelp: "To flow text in columns, use Format, then Columns, in Google Docs.",
      restyle: true,
      noDownloads: true,
      history: true,
      historyNote:
        "Offizin's Undo reverses only what Offizin did, and leaves Google Docs' Undo alone. Undoing an inserted page part removes it along with anything you typed in it, and Redo puts that back exactly. Undoing a new publication brings back the document as it was before, for up to six hours, so changes made to the publication since then are lost. The history lasts while this sidebar is open.",
      finish: true,
      printSettings: [],
      printGuide: { editor: "google" },
      pdfHelp: "Choose File, then Download, then PDF document (.pdf).",
    },

    async loadState() {
      const state = await call("ozGetState");
      let logo = logoInMemory;
      try {
        const raw = localStorage.getItem(LOGO_KEY);
        if (raw) logo = JSON.parse(raw);
      } catch (e) {
        // Browser storage is not available in this sidebar; the logo lasts for this session.
      }
      return state ? { ...state, logo } : logo ? { logo } : null;
    },

    async saveState(state) {
      const { logo, ...rest } = state;
      logoInMemory = logo;
      try {
        if (logo) localStorage.setItem(LOGO_KEY, JSON.stringify(logo));
        else localStorage.removeItem(LOGO_KEY);
      } catch (e) {
        // See loadState.
      }
      await call("ozSaveState", JSON.stringify(rest));
    },

    async createPublication(model) {
      const payload = serverModel(model);
      // Each attempt has its own id, so a failed attempt only ever restores its own copy.
      const run = newId();
      let res;
      try {
        res = await call("ozCreate", { ...payload, run });
      } catch (e) {
        throw await recoverCreate(e, run);
      }
      const snapshot = res && res.snapshot;
      return {
        message: (res && res.message ? res.message : `${model.name} created.`) + unsupportedNote(snapshot && snapshot.unsupported),
        op: { kind: "create", label: `Create ${model.name.toLowerCase()}`, data: { snapshot, payload } },
      };
    },

    async insertPart(composed, width, partId, ctx = {}) {
      const name = ctx.label || "page part";
      const id = newId();
      const blocks = withImages(composed.blocks);
      const res = await call("ozInsert", blocks, width, id);
      return {
        message: `${name.charAt(0).toUpperCase()}${name.slice(1)} inserted.`,
        op: { kind: "insert", label: `Insert ${name.toLowerCase()}`, data: { id, blocks, width, index: res && typeof res.index === "number" ? res.index : null } },
      };
    },

    contentWidth() {
      return call("ozContentWidth");
    },

    async applyPageSetup(setup) {
      const res = await call("ozPageSetup", setup);
      return { message: res.message, op: { kind: "pageSetup", label: "Change page setup", data: { before: res.before, after: setup } } };
    },

    applyColumns() {
      throw new Error("Use Format, then Columns, in Google Docs.");
    },

    getPublication() {
      return call("ozGetPublication");
    },

    async restyle(map, theme) {
      const res = await call("ozRestyle", map, theme);
      return { message: res.message, op: { kind: "restyle", label: "Apply design", data: { map, theme, previous: res.previous } } };
    },

    async undo(op) {
      const d = op.data;
      switch (op.kind) {
        case "create":
          return call("ozUndoCreate", d.snapshot);
        case "insert": {
          const res = await call("ozUndoInsert", d.id);
          const index = res && typeof res.index === "number" ? res.index : d.index;
          // What was removed, with the person's own typing, so Redo puts back exactly that.
          const content = res && Array.isArray(res.content) && res.content.length ? res.content : null;
          const lost = unsupportedNote(res && res.unsupported, "that page part").replace(/Undo cannot/, "Redo cannot");
          return { message: `Removed: ${op.label.replace(/^Insert /, "")}.${lost}`, op: { ...op, data: { ...d, index, content } } };
        }
        case "restyle": {
          const res = await call("ozRestyle", invertRestyleMap(d.map), d.previous);
          return res && res.message ? "The previous design is back." : "The previous design is back.";
        }
        case "pageSetup":
          await call("ozPageSetup", d.before);
          return "Page setup restored.";
        default:
          throw new Error("Offizin cannot undo that change.");
      }
    },

    async redo(op) {
      const d = op.data;
      switch (op.kind) {
        case "create": {
          // Redo builds the publication again, so a stopped script is recovered as in createPublication.
          const run = newId();
          let res;
          try {
            res = await call("ozCreate", { ...d.payload, run });
          } catch (e) {
            throw await recoverCreate(e, run);
          }
          return { message: res.message, op: { ...op, data: { ...d, snapshot: res.snapshot } } };
        }
        case "insert": {
          // Without its anchor (the publication was rebuilt), the part returns to its earlier position.
          const res = await call("ozRedoInsert", d.id, d.blocks, d.width, typeof d.index === "number" ? d.index : null, d.content || null);
          const index = res && typeof res.index === "number" ? res.index : d.index;
          const where = res && /earlier position/.test(res.message || "") ? " at its earlier position" : "";
          return { message: `${op.label.replace(/^Insert /, "Inserted ")} again${where}.`, op: { ...op, data: { ...d, index } } };
        }
        case "restyle":
          await call("ozRestyle", d.map, d.theme);
          return "Design applied again.";
        case "pageSetup":
          await call("ozPageSetup", d.after);
          return "Page setup applied again.";
        default:
          throw new Error("Offizin cannot redo that change.");
      }
    },

    /** Ops leaving the history for good: remove anchors of undone parts that can no longer be redone. */
    async discard(ops) {
      const ids = (ops || []).filter((op) => op && op.kind === "insert" && op.data).map((op) => op.data.id);
      if (ids.length) await call("ozDiscardAnchors", ids).catch(() => {});
    },

    finish() {
      return call("ozFinish");
    },

    /** Every sample the panel sends is searched; Code.gs does it in one call. */
    check({ samples, infoSamples }) {
      return call("ozCheck", samples || [], infoSamples || []);
    },

    select(finding) {
      return call("ozSelect", finding);
    },
  };
  return host;
}

function start() {
  const root = document.getElementById("app");
  mountPanel(root, createGoogleHost(), { siteUrl: "https://offizin.alyssaagard.dev" });
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start);
else start();
