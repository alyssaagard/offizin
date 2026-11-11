/**
 * The Microsoft Word host (Office.js). Loaded by word/taskpane.html.
 *
 * How Offizin builds a publication in Word:
 *   1. The panel composes the publication and this file writes it as a complete .docx in the
 *      browser (exact page size, margins, tables, pictures). Nothing is sent to a server.
 *   2. The .docx replaces the open document (Document.insertFileFromBase64, WordApi 1.5,
 *      which also copies section properties such as page size), or opens as a new document
 *      (Application.createDocument, WordApi 1.3).
 *   3. Older versions fall back to Body.insertFileFromBase64 (WordApi 1.1).
 *   4. Where Word allows add-ins to set page size (WordApiDesktop 1.3: Word for Windows and
 *      Mac), the page setup is applied again to be certain.
 *
 * Requirement sets are checked at run time, because an XML manifest cannot require
 * desktop-only sets. Sources: learn.microsoft.com/javascript/api/requirement-sets/word/.
 */
/* global Office, Word */
import { round } from "../core/util.js";
import { docxBytes, flatPackage } from "../render/ooxml.js";
import { prepareImages } from "../render/pictures.js";
import { bytesToBase64 } from "../render/zip.js";

const STORE_KEY = "offizin.word.state";
const SETTINGS_KEY = "offizin";
const PLACEHOLDER_ALT = "Offizin picture placeholder";

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
      Office.context.document.settings.set(SETTINGS_KEY, data);
      Office.context.document.settings.saveAsync(() => resolve());
    } catch (e) {
      resolve();
    }
  });
}

/** Applies page size and margins with WordApiDesktop 1.3, where Word allows it. */
async function applyPageSetupDesktop(context, page, onlySelectionSection = false) {
  const sections = onlySelectionSection ? context.document.getSelection().sections : context.document.sections;
  sections.load("items");
  await context.sync();
  for (const s of sections.items) {
    const ps = s.pageSetup;
    const landscape = page.width > page.height;
    const PO = Word.PageOrientation;
    ps.orientation = PO ? (landscape ? PO.landscape : PO.portrait) : landscape ? "Landscape" : "Portrait";
    ps.pageWidth = page.width;
    ps.pageHeight = page.height;
    ps.topMargin = page.margins.top;
    ps.bottomMargin = page.margins.bottom;
    ps.leftMargin = page.margins.left;
    ps.rightMargin = page.margins.right;
  }
  await context.sync();
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function createWordHost() {
  const api13 = supports("WordApi", "1.3");
  const api15 = supports("WordApi", "1.5");
  const desktop13 = supports("WordApiDesktop", "1.3");
  const isWeb = Office.context.platform === Office.PlatformType.OfficeOnline;

  const createModes = [
    {
      id: "replace",
      label: "Replace this document",
      note: "Replaces everything in this document. Undo (Ctrl+Z or Command+Z) brings it back.",
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
      const base64 = bytesToBase64(docxBytes(model, images));
      const record = { templateId: model.templateId, theme: model.theme, samples: model.samples.slice(0, 80), page: model.page, createdAt: new Date().toISOString() };

      if (mode === "new") {
        await Word.run(async (context) => {
          context.application.createDocument(base64).open();
          await context.sync();
        });
        return `${model.name} opened in a new document.`;
      }

      // Each attempt runs in its own batch, so a failure cannot affect the next attempt.
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
      let how = "";
      if (api15 && (await attempt("Document.insertFileFromBase64", (c) => c.document.insertFileFromBase64(base64, "Replace", { importTheme: false, importStyles: true, importParagraphSpacing: true, importPageColor: false })))) {
        how = "document";
      } else if (await attempt("Body.insertFileFromBase64", (c) => c.document.body.insertFileFromBase64(base64, "Replace"))) {
        how = "body";
      } else {
        // Last resort: content only. Page size must then be set by hand.
        const width = model.page.width - model.page.margins.left - model.page.margins.right;
        await Word.run(async (context) => {
          context.document.body.insertOoxml(flatPackage(model.blocks, width, images), "Replace");
          await context.sync();
        });
        how = "ooxml";
      }
      if (desktop13) {
        try {
          await Word.run((context) => applyPageSetupDesktop(context, model.page));
        } catch (e) {
          console.warn("Page setup could not be applied", e);
        }
      }
      console.info(`Offizin built ${model.templateId} using ${how}`);
      await saveSettings(record);
      // Only Word for Windows and Mac let add-ins confirm the page size, so elsewhere say what to check.
      if (desktop13) return `${model.name} created.`;
      return `${model.name} created. The page should be ${sizeText(model.page)}. If it is not, set it in Layout, then Size.`;
    },

    async insertPart(composed, width) {
      const pkg = flatPackage(composed.blocks, width, prepareImages(composed.blocks));
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        if (api13) {
          const para = selection.paragraphs.getLast();
          para.getRange("Whole").insertOoxml(pkg, "After");
        } else {
          selection.insertOoxml(pkg, "After");
        }
        await context.sync();
      });
    },

    async contentWidth() {
      if (desktop13) {
        try {
          return await Word.run(async (context) => {
            const s = context.document.getSelection().sections.getFirst();
            s.pageSetup.load("pageWidth,leftMargin,rightMargin");
            await context.sync();
            return s.pageSetup.pageWidth - s.pageSetup.leftMargin - s.pageSetup.rightMargin;
          });
        } catch (e) {
          // Fall through to the saved publication.
        }
      }
      const pub = host.getPublication();
      if (pub && pub.page) return pub.page.width - pub.page.margins.left - pub.page.margins.right;
      return 468;
    },

    async applyPageSetup(setup) {
      if (!desktop13) throw new Error("This version of Word does not let add-ins change page size. Use Layout, then Size.");
      await Word.run((context) => applyPageSetupDesktop(context, setup, true));
    },

    async applyColumns({ count, line }) {
      if (!desktop13) throw new Error("This version of Word does not let add-ins set columns. Use Layout, then Columns.");
      await Word.run(async (context) => {
        const section = context.document.getSelection().sections.getFirst();
        const cols = section.pageSetup.textColumns;
        cols.setCount(count);
        cols.setIsEvenlySpaced(true);
        cols.setHasLineBetween(!!line);
        await context.sync();
      });
    },

    getPublication() {
      try {
        return Office.context.document.settings.get(SETTINGS_KEY) || null;
      } catch (e) {
        return null;
      }
    },

    async restyle(map, theme) {
      await Word.run(async (context) => {
        const body = context.document.body;
        const ooxml = body.getOoxml();
        await context.sync();
        let xml = ooxml.value;
        for (const [from, to] of Object.entries(map.colors)) {
          xml = xml.replace(new RegExp(`"${from.replace("#", "")}"`, "gi"), `"${to.replace("#", "")}"`);
        }
        for (const [from, to] of Object.entries(map.fonts)) {
          xml = xml.replace(new RegExp(`(w:(?:ascii|hAnsi|eastAsia|cs)=")${escapeRegExp(from)}(")`, "g"), `$1${to}$2`);
        }
        body.insertOoxml(xml, "Replace");
        await context.sync();
      });
      const pub = host.getPublication() || {};
      await saveSettings({ ...pub, theme });
      return "Design applied. Pictures keep their colors. Use Undo if you want the old look back.";
    },

    async check({ samples, infoSamples }) {
      return Word.run(async (context) => {
        const body = context.document.body;
        const sampleSearches = samples.slice(0, 80).map((text) => ({ text, r: body.search(text, { matchCase: true }) }));
        const infoSearches = infoSamples.map((f) => ({ f, r: body.search(f.sample, { matchCase: true }) }));
        sampleSearches.forEach((s) => s.r.load("items"));
        infoSearches.forEach((s) => s.r.load("items"));
        const pics = body.inlinePictures;
        pics.load("items/altTextDescription");
        await context.sync();
        const findings = [];
        for (const s of sampleSearches) if (s.r.items.length) findings.push({ kind: "sample", text: s.text, search: s.text });
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
