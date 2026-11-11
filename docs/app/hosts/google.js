/**
 * The Google Docs host, running inside the Offizin sidebar.
 *
 * The sidebar composes publications and draws picture placeholders in the browser, then
 * hands the finished model to the Apps Script functions in Code.gs (google.script.run),
 * which build it with DocumentApp. Nothing leaves Google: there is no Offizin server.
 *
 * Settings are saved with PropertiesService (per user). Logos are too large for that, so the
 * logo is kept in the sidebar's browser storage when the browser allows it.
 */
/* global google */
import { walkBlocks } from "../core/compose.js";
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

let logoInMemory = null;

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
          note: "Replaces everything in this document. Undo (Ctrl+Z or Command+Z) brings it back. If the document is pageless, choose File, then Page setup, then Pages first.",
        },
      ],
      pageSetup: true,
      columns: false,
      columnsHelp: "To flow text in columns, use Format, then Columns, in Google Docs.",
      restyle: true,
      noDownloads: true,
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

    createPublication(model) {
      const payload = { ...model, blocks: withImages(model.blocks) };
      return call("ozCreate", payload);
    },

    insertPart(composed, width) {
      return call("ozInsert", withImages(composed.blocks), width);
    },

    contentWidth() {
      return call("ozContentWidth");
    },

    applyPageSetup(setup) {
      return call("ozPageSetup", setup);
    },

    applyColumns() {
      throw new Error("Use Format, then Columns, in Google Docs.");
    },

    getPublication() {
      return call("ozGetPublication");
    },

    restyle(map, theme) {
      return call("ozRestyle", map, theme);
    },

    check({ samples, infoSamples }) {
      return call("ozCheck", samples.slice(0, 120), infoSamples);
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
