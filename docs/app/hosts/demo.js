/**
 * The website demo host. It draws publications on a simulated page instead of in a document,
 * so people can try Offizin before installing it, and download the result as a Word file
 * (which Google Docs also opens).
 *
 * Host interface used by the panel (every method may return a promise):
 *   kind, platform, label, capabilities
 *   loadState(), saveState(state)
 *   createPublication(model, mode, { options })
 *   insertPart(composed, width, partId)
 *   applyPageSetup(setup), applyColumns({ count, line })
 *   restyle(map, theme), getPublication(), contentWidth()
 *   check({ samples, infoSamples }) -> findings, select(finding)
 *   themeChanged(theme)   optional, called whenever colors, fonts, or information change
 */
import { compose, composePart, BUSINESS_FIELDS } from "../core/compose.js";
import { findTemplate, findPart } from "../core/templates.js";
import { renderPreview, highlightText } from "../render/preview.js";
import { docxBytes } from "../render/ooxml.js";
import { prepareImages } from "../render/pictures.js";

const STORE_KEY = "offizin.demo.state";

export function createDemoHost({ canvas, onRender }) {
  let theme = null;
  let current = null; // { templateId, options }
  let parts = []; // part ids appended at the end
  let pageOverride = null;
  let model = null;
  let overflow = [];
  let view = { zoom: "fit", guides: true };

  function rebuild() {
    if (!current || !theme) return;
    const t = findTemplate(current.templateId);
    model = compose(t, theme, "web", current.options);
    if (pageOverride) model.page = { ...model.page, ...pageOverride };
    for (const id of parts) {
      const composed = composePart(findPart(id), theme, "web", contentWidth());
      model.blocks.push(...composed.blocks);
      model.samples.push(...composed.samples);
    }
  }

  function contentWidth() {
    if (!model) return 468;
    return model.page.width - model.page.margins.left - model.page.margins.right;
  }

  function draw() {
    if (!model) return;
    const available = Math.max(240, canvas.clientWidth - 48);
    const pagePx = model.page.width * (96 / 72);
    const zoom = view.zoom === "fit" ? Math.min(1.25, available / pagePx) : Number(view.zoom);
    const result = renderPreview(canvas, model, { scale: zoom, guides: view.guides, labels: true });
    overflow = result.overflow;
    if (onRender) onRender(model, result);
  }

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
      try {
        localStorage.setItem(STORE_KEY, JSON.stringify(state));
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
      current = { templateId: m.templateId, options: { ...(options || {}) } };
      parts = [];
      pageOverride = null;
      rebuild();
      draw();
      canvas.scrollTop = 0;
      return `${m.name} is on the page. Change colors, fonts, or business information and watch it update.`;
    },

    insertPart(composed, width, partId) {
      if (!current) throw new Error("Create a publication first, then add parts to it.");
      parts.push(partId);
      rebuild();
      draw();
      const sheets = canvas.querySelectorAll(".ozp-sheet");
      const last = sheets[sheets.length - 1];
      if (last && last.scrollIntoView) last.scrollIntoView({ block: "end", behavior: "smooth" });
      return "Added to the end of the publication.";
    },

    applyPageSetup(setup) {
      if (!current) throw new Error("Create a publication first.");
      pageOverride = setup;
      rebuild();
      draw();
    },

    applyColumns() {
      throw new Error("Columns are available in Word for Windows and Mac.");
    },

    getPublication() {
      return current && model ? { templateId: current.templateId, theme: model.theme, samples: model.samples } : null;
    },

    contentWidth,

    check() {
      if (!model) throw new Error("Create a publication first.");
      const findings = model.samples.map((text) => ({ kind: "sample", text }));
      let pictures = 0;
      const walk = (blocks) => blocks.forEach((b) => {
        if (b.t === "image" && !b.src) pictures++;
        if (b.t === "grid") b.rows.forEach((r) => r.cells.forEach((c) => walk(c.blocks)));
      });
      walk(model.blocks);
      if (pictures) findings.push({ kind: "picture", text: `${pictures} picture placeholder${pictures > 1 ? "s" : ""} to replace with your own pictures`, search: "picture" });
      for (const key of model.missingInfo) {
        const f = BUSINESS_FIELDS.find((x) => x.key === key);
        findings.push({ kind: "info", text: `${f.label} still shows the sample “${f.sample}”. Fill it in on the Info tab.`, search: f.sample });
      }
      overflow.forEach((over, i) => over && findings.push({ kind: "layout", text: `Text overflows side ${i + 1}. Shorten the text or make it smaller.`, selectable: false }));
      return findings;
    },

    select(f) {
      if (!highlightText(canvas, f.search || f.text)) throw new Error("That item is no longer on the page.");
    },

    // Demo-only helpers used by the page around the panel.
    setView(next) {
      view = { ...view, ...next };
      draw();
    },
    redraw: draw,
    currentModel: () => model,
    downloadDocx() {
      if (!model) throw new Error("Create a publication first.");
      const wordModel = compose(findTemplate(current.templateId), theme, "word", current.options);
      if (pageOverride) wordModel.page = { ...wordModel.page, ...pageOverride };
      for (const id of parts) wordModel.blocks.push(...composePart(findPart(id), theme, "word", contentWidth()).blocks);
      const bytes = docxBytes(wordModel, prepareImages(wordModel.blocks));
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${wordModel.name.toLowerCase().replace(/[^a-z0-9]+/g, "-")}.docx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        a.remove();
      }, 1500);
    },
  };
  return host;
}
