/**
 * The composer turns a template plus the user's design choices into a resolved layout model.
 *
 * Templates describe publications with roles ("title", "body"), color tokens ("primary"),
 * and business information fields ("{org}"). The composer resolves all of that into plain
 * values: font names, point sizes, hex colors, and column widths. Each renderer (the web
 * preview, Word, and Google Docs) then draws the same model, so the three always match.
 *
 * Resolved block types: p, rule, spacer, break, image, grid, columns.
 * Lists, bands, data tables, and merge scopes in templates are converted into those.
 *
 * Model additions in version 2:
 *   model.styles    the default look of each role, for native styles (Title, Heading 1, Normal)
 *   model.print     print settings: duplex, book fold, mirror margins, gutter, and more
 *   model.footer    optional footer text; {page} and {pages} are left for the renderer
 *   model.pages     booklets only: each page's blocks in reading order
 *   model.booklet   booklets only: page count, sheet size, and printer spread order
 *   model.options   the options the publication was composed with
 */
import { PAPER, pageSize, round, mix, clamp, contrastRatio } from "./util.js";
import { palette, fontsFor, findColorScheme, findFontScheme } from "./schemes.js";

export const MODEL_VERSION = 2;

/** Typography roles. Sizes are in points; line is a multiple of single spacing. */
export const ROLES = {
  display: { font: "heading", size: 40, bold: true, line: 1.0, after: 6, color: "primary" },
  title: { font: "heading", size: 28, bold: true, line: 1.05, after: 6, color: "primary" },
  headline: { font: "heading", size: 20, bold: true, line: 1.1, after: 6, color: "primary" },
  h1: { font: "heading", size: 15, bold: true, line: 1.15, after: 5, color: "primary" },
  h2: { font: "heading", size: 12, bold: true, line: 1.2, after: 4, color: "dark" },
  kicker: { font: "body", size: 8, bold: true, caps: true, line: 1.2, after: 4, color: "accent" },
  lead: { font: "body", size: 12, line: 1.35, after: 8, color: "dark" },
  body: { font: "body", size: 10, line: 1.3, after: 6, color: "dark" },
  small: { font: "body", size: 8.5, line: 1.3, after: 3, color: "dark" },
  caption: { font: "body", size: 8, italic: true, line: 1.25, after: 3, color: "muted" },
  quote: { font: "heading", size: 15, italic: true, line: 1.25, after: 6, color: "primary" },
  label: { font: "body", size: 7.5, bold: true, caps: true, line: 1.2, after: 2, color: "muted" },
  button: { font: "body", size: 11, bold: true, caps: true, line: 1.2, after: 0, color: "onAccent" },
};

/**
 * Native styles for each role. Word uses the style ID (built-in IDs for Title, Heading 1,
 * Heading 2, Normal, Quote, and Caption, so the Styles pane, Navigation pane, and screen
 * readers understand the document). Google Docs uses the named paragraph style.
 */
export const ROLE_STYLES = {
  display: { word: "OffizinDisplay", name: "Offizin Display", google: "TITLE" },
  title: { word: "Title", name: "Title", google: "TITLE" },
  headline: { word: "Heading1", name: "heading 1", google: "HEADING1" },
  h1: { word: "Heading1", name: "heading 1", google: "HEADING1" },
  h2: { word: "Heading2", name: "heading 2", google: "HEADING2" },
  kicker: { word: "OffizinKicker", name: "Offizin Kicker", google: "NORMAL" },
  lead: { word: "OffizinLead", name: "Offizin Lead", google: "SUBTITLE" },
  body: { word: "Normal", name: "Normal", google: "NORMAL" },
  small: { word: "OffizinSmall", name: "Offizin Small", google: "NORMAL" },
  caption: { word: "Caption", name: "caption", google: "NORMAL" },
  quote: { word: "Quote", name: "Quote", google: "NORMAL" },
  label: { word: "OffizinLabel", name: "Offizin Label", google: "NORMAL" },
  button: { word: "OffizinButton", name: "Offizin Button", google: "NORMAL" },
};

/** Business information, in the spirit of Publisher's business information sets. */
export const BUSINESS_FIELDS = [
  { key: "name", label: "Your name", sample: "Jordan Rivera" },
  { key: "title", label: "Job title", sample: "Program Director" },
  { key: "org", label: "Organization", sample: "Riverside Arts Collective" },
  { key: "tagline", label: "Tagline or motto", sample: "Creating together since 1998" },
  { key: "address", label: "Street address", sample: "123 Main Street" },
  { key: "city", label: "City, state, and postal code", sample: "Springfield, ST 12345" },
  { key: "phone", label: "Phone", sample: "(555) 123-4567" },
  { key: "email", label: "Email", sample: "hello@example.org" },
  { key: "web", label: "Website", sample: "www.example.org" },
  { key: "social", label: "Social media", sample: "@riversidearts" },
  { key: "profile", label: "LinkedIn or portfolio link", sample: "linkedin.com/in/jordanrivera" },
];

/**
 * Spacing and text controls. Each is a multiplier (1 = the template's design).
 *   textScale   font sizes          lineScale   line spacing
 *   spaceScale  space between paragraphs and blocks
 *   padScale    padding inside panels, boxes, and cells
 *   gapScale    gaps between columns
 */
export const DEFAULT_LAYOUT = { textScale: 1, lineScale: 1, spaceScale: 1, padScale: 1, gapScale: 1 };
export const LAYOUT_LIMITS = {
  textScale: [0.7, 1.5],
  lineScale: [0.8, 1.8],
  spaceScale: [0, 3],
  padScale: [0, 3],
  gapScale: [0, 3],
};

/** Print settings carried by every model. Renderers apply what their editor supports. */
export const DEFAULT_PRINT = {
  duplex: "none", // "none", "long" (flip on long edge), or "short" (flip on short edge)
  scale: 100, // always print at actual size
  bookFold: false, // Word prints the pages as a folded booklet (imposition done by Word)
  mirrorMargins: false, // inside and outside margins for facing pages
  gutter: 0, // extra binding margin in points
  twoPerSheet: false, // Word prints two pages on each sheet
  embedFonts: false, // Word for Windows embeds fonts when the file is saved
  paper: "", // a plain-language paper suggestion
  finishing: "", // cutting, folding, or binding notes
};

export const DEFAULT_THEME = {
  colorScheme: "ink",
  customColors: null, // { primary, accent, dark } when the user picks their own
  fontScheme: "typewriter",
  customFonts: null, // { heading: { word, google }, body: { word, google }, headingBold } with fontScheme "custom"
  paper: "letter", // "letter" or "a4"
  business: {},
  logo: null, // { src: dataURL, width, height } in pixels
  layout: DEFAULT_LAYOUT,
};

/** What the Customize panel may change for a template. */
export function customizeFor(template) {
  const c = template.customize || {};
  return {
    paper: c.paper === false ? null : c.paper || null, // a list of PAPER keys, or null for Letter/A4 only
    orientation: !!c.orientation,
    margins: c.margins !== false,
    spacing: c.spacing !== false,
    text: c.text !== false,
  };
}

/** The three chosen colors for a theme, from a named scheme or custom colors. */
export function themeColors(theme) {
  if (theme.colorScheme === "custom" && theme.customColors) return theme.customColors;
  return findColorScheme(theme.colorScheme);
}

/** Field values with samples filled in, plus the list of fields that used samples. */
export function businessValues(business = {}) {
  const values = {};
  const missing = [];
  for (const f of BUSINESS_FIELDS) {
    const v = String(business[f.key] || "").trim();
    values[f.key] = v || f.sample;
    if (!v) missing.push(f.key);
  }
  return { values, missing };
}

function normMargins(m) {
  if (typeof m === "number") return { top: m, right: m, bottom: m, left: m };
  return { top: Number(m.top) || 0, right: Number(m.right) || 0, bottom: Number(m.bottom) || 0, left: Number(m.left) || 0 };
}

export function clampLayout(layout = {}) {
  const out = {};
  for (const [k, [lo, hi]] of Object.entries(LAYOUT_LIMITS)) {
    const v = Number(layout[k]);
    out[k] = Number.isFinite(v) ? clamp(v, lo, hi) : DEFAULT_LAYOUT[k];
  }
  return out;
}

/** Resolves a template's page for the chosen paper and any Customize overrides. */
export function resolvePage(template, paper, options = {}) {
  const spec = typeof template.page === "function" ? template.page(paper, options) : template.page;
  const allow = customizeFor(template);
  let sizeName = spec.size === "paper" || !spec.size ? paper : spec.size;
  if (options.paperSize && allow.paper && allow.paper.includes(options.paperSize)) sizeName = options.paperSize;
  let orientation = spec.orientation || "portrait";
  if (allow.orientation && (options.orientation === "portrait" || options.orientation === "landscape")) orientation = options.orientation;
  const size = spec.width ? { width: spec.width, height: spec.height } : pageSize(sizeName, orientation);
  let margins = normMargins(spec.margins ?? 36);
  if (options.margins && allow.margins) {
    const m = normMargins(options.margins);
    // Keep at least two inches of usable space in each direction.
    let maxSide = Math.max(0, (size.width - 144) / 2);
    // A tri-fold also keeps each outer panel at least one inch wide.
    if (template.fold === "trifold") maxSide = Math.max(0, Math.min(maxSide, size.width / 3 - 72));
    const maxTop = Math.max(0, (size.height - 144) / 2);
    margins = {
      top: round(clamp(m.top, 0, maxTop)),
      bottom: round(clamp(m.bottom, 0, maxTop)),
      left: round(clamp(m.left, 0, maxSide)),
      right: round(clamp(m.right, 0, maxSide)),
    };
  }
  return {
    width: round(size.width),
    height: round(size.height),
    orientation: size.width > size.height ? "landscape" : "portrait",
    margins,
    sizeName: spec.width ? null : sizeName,
  };
}

/** Cleans merge records: plain string values, at most 1000 records. */
export function normalizeRecords(records) {
  if (!Array.isArray(records)) return [];
  return records.slice(0, 1000).map((r) => {
    const out = {};
    for (const [k, v] of Object.entries(r || {})) out[String(k)] = v == null ? "" : String(v).trim();
    return out;
  });
}

/** The standard paper size a sheet matches in either orientation (within about 1.5 mm), or null. */
export function matchPaperSize(width, height) {
  const near = (a, b) => Math.abs(a - b) < 4.5;
  for (const [name, p] of Object.entries(PAPER)) {
    if (near(p.width, width) && near(p.height, height)) return { name, width: p.width, height: p.height };
    if (near(p.height, width) && near(p.width, height)) return { name, width: p.height, height: p.width };
  }
  return null;
}

/**
 * The sheet that holds two portrait pages side by side (booklets, and two pages per sheet),
 * snapped to standard paper: two A5 pages are 296 mm wide, and the sheet is A4 at 297 mm.
 * Returns { width, height, sizeName } or null for a landscape page.
 */
export function twoUpSheet(page) {
  if (!page || page.width > page.height) return null;
  const width = round(page.width * 2);
  const match = matchPaperSize(width, page.height);
  if (!match) return null;
  return { width: round(match.width), height: round(match.height), sizeName: match.name };
}

/** The page size Word stores for a model: the whole sheet for book fold and for two pages per sheet. */
export function wordSheet(model, bookFold) {
  if (bookFold && model.booklet) return { width: model.booklet.sheet.width, height: model.booklet.sheet.height };
  if (model.print && model.print.twoPerSheet && !model.booklet) {
    const sheet = twoUpSheet(model.page);
    if (sheet) return { width: sheet.width, height: sheet.height };
  }
  return { width: model.page.width, height: model.page.height };
}

/** Printer spread order for a booklet of n pages (n is a multiple of 4), as [left, right] page indexes. */
export function bookletSpreads(n) {
  const spreads = [];
  for (let s = 0; s < n / 4; s++) {
    spreads.push([n - 1 - 2 * s, 2 * s]); // front of sheet s
    spreads.push([2 * s + 1, n - 2 - 2 * s]); // back of sheet s
  }
  return spreads;
}

// ---------------------------------------------------------------------------
// Composition
// ---------------------------------------------------------------------------

/** Fill tokens that are always light, so accent text on them uses its darker text tone. */
const LIGHT_FILLS = new Set(["white", "tint", "accentTint", "line"]);

class Composer {
  constructor(theme, platform, layout) {
    this.theme = { ...DEFAULT_THEME, ...theme };
    this.platform = platform;
    this.layout = clampLayout(layout || this.theme.layout);
    this.colors = palette(themeColors(this.theme));
    this.fontScheme = findFontScheme(this.theme.fontScheme, this.theme.customFonts);
    this.fonts = fontsFor(this.fontScheme, platform);
    const { values, missing } = businessValues(this.theme.business);
    this.info = values;
    this.missingInfo = missing;
    this.samples = new Set();
    this.vars = {};
    // The fill tokens of the cells being composed, innermost last, so text can pick a readable tone.
    this.backgrounds = [];
    // A template's own look for some roles (see template.styles in compose()).
    this.styleOverrides = {};
  }

  /**
   * The color for text. Text in the accent color uses the accent's darker text tone on paper and
   * light fills, and the accent itself on other fills. The choice goes by the fill's token, not
   * its value, so a publication restyled to another scheme matches one made in that scheme.
   */
  textColor(token, fallback) {
    if ((token || fallback) !== "accent") return this.color(token, fallback);
    const fill = this.backgrounds[this.backgrounds.length - 1];
    const light = fill === undefined || (typeof fill === "string" && !fill.startsWith("#") ? LIGHT_FILLS.has(fill) : contrastRatio(this.color(fill), "#FFFFFF") < 2);
    return this.color(light ? "accentText" : "accent");
  }

  color(token, fallback = "dark") {
    if (!token) token = fallback;
    if (typeof token === "string" && token.startsWith("#")) return token.toUpperCase();
    if (typeof token === "object" && token.mix) {
      return mix(this.color(token.mix[0]), this.color(token.mix[1]), token.mix[2]);
    }
    return this.colors[token] || this.colors[fallback];
  }

  size(pt) {
    return Math.max(4, Math.round(pt * this.layout.textScale * 2) / 2);
  }

  space(pt) {
    return round((pt || 0) * this.layout.spaceScale, 1);
  }

  /** Replaces {field} tokens with merge values, template variables, or business information. */
  fill(text) {
    return String(text).replace(/\{(\w+)\}/g, (m, key) => {
      if (key in this.vars) return String(this.vars[key]);
      if (key in this.info) return this.info[key];
      return m;
    });
  }

  /**
   * Remembers sample text, so the design checker can find what is still unreplaced. Short text
   * is skipped unless it is known to be a sample: marked { sample: true }, marked { fixed: false }
   * (the sample records in mail merge templates), or a placeholder such as "Job title".
   */
  recordSample(raw, block, caps) {
    if (block.fixed) return;
    if (/\{\w+\}/.test(raw)) return;
    // Search one line at a time and stop at a tab, since neither is a space in the editors.
    const firstLine = String(raw).split("\n")[0].split("\t")[0];
    // Keep double spaces: the editors keep them, so the search has to as well.
    let clean = firstLine.replace(/\*\*/g, "").trim();
    if (!/[a-z]/i.test(clean)) return;
    const known = block.sample === true || block.fixed === false || /\S\s+(name|title)$/i.test(clean);
    if (clean.length < (known ? 3 : 12)) return;
    // Text styled in capitals is written in capitals, and the editors search with matching case.
    if (caps) clean = clean.toUpperCase();
    // The checker searches for the start of each sample, which keeps searches short and reliable.
    this.samples.add(clean.slice(0, 48).trim());
  }

  /** "**bold** text" becomes runs. */
  runs(text, caps) {
    const out = [];
    const parts = String(text).split(/(\*\*[^*]+\*\*)/g);
    for (const part of parts) {
      if (!part) continue;
      const bold = part.startsWith("**") && part.endsWith("**");
      let t = bold ? part.slice(2, -2) : part;
      if (caps) t = t.toUpperCase();
      out.push(bold ? { text: t, bold: true } : { text: t });
    }
    return out.length ? out : [{ text: "" }];
  }

  /** The default look of one role, after the theme and the text controls. */
  roleStyle(roleName) {
    const role = ROLES[roleName] || ROLES.body;
    const font = this.fonts[role.font];
    const isHeading = role.font === "heading";
    const caps = !!role.caps;
    return {
      role: roleName,
      word: ROLE_STYLES[roleName].word,
      wordName: ROLE_STYLES[roleName].name,
      google: ROLE_STYLES[roleName].google,
      fontKey: role.font,
      font: font.name,
      css: font.css,
      size: this.size(role.size),
      color: this.textColor(role.color),
      bold: role.bold ? (isHeading ? this.fonts.headingBold : true) : false,
      italic: !!role.italic,
      caps,
      tracking: caps ? 0.08 : 0,
      line: round(role.line * this.layout.lineScale, 3),
      before: 0,
      after: this.space(role.after),
      ...this.overrideFor(roleName),
    };
  }

  /**
   * A template's own size, color, and spacing for a role, scaled the same way as paragraphs, so
   * paragraphs that use it match their style exactly and carry no direct formatting in Word.
   */
  overrideFor(roleName) {
    const o = this.styleOverrides[roleName];
    if (!o) return {};
    const out = {};
    if (o.size !== undefined) out.size = this.size(o.size);
    if (o.color !== undefined) out.color = this.textColor(o.color);
    if (o.before !== undefined) out.before = this.space(o.before);
    if (o.after !== undefined) out.after = this.space(o.after);
    if (o.line !== undefined) out.line = round(Math.max(0.8, o.line * this.layout.lineScale), 3);
    return out;
  }

  styles() {
    const out = {};
    for (const r of Object.keys(ROLES)) out[r] = this.roleStyle(r);
    return out;
  }

  paragraph(b, extra = {}) {
    const roleName = ROLES[b.role] ? b.role : "body";
    const role = ROLES[roleName];
    const fontKey = b.font || role.font;
    const font = this.fonts[fontKey];
    const isHeading = fontKey === "heading";
    const bold = b.bold ?? (role.bold ? (isHeading ? this.fonts.headingBold : true) : false);
    const raw = b.text ?? "";
    const caps = b.caps ?? role.caps ?? false;
    this.recordSample(raw, b, caps);
    const p = {
      t: "p",
      style: roleName,
      runs: this.runs(this.fill(raw), caps),
      font: font.name,
      css: font.css,
      size: this.size(b.size ?? role.size),
      color: this.textColor(b.color, role.color),
      bold,
      italic: b.italic ?? role.italic ?? false,
      align: b.align || "left",
      before: this.space(b.before ?? 0),
      after: this.space(b.after ?? role.after),
      line: round(Math.max(0.8, (b.line ?? role.line) * this.layout.lineScale), 3),
      tracking: b.tracking ?? (caps ? 0.08 : 0),
      caps: !!caps,
      ...extra,
    };
    if (b.indent) p.indent = { left: Number(b.indent.left) || 0, hanging: Number(b.indent.hanging) || 0 };
    if (b.keepNext) p.keepNext = true;
    return p;
  }

  /** Resolves a list of template blocks for a container `width` points wide. */
  blocks(list, width) {
    const out = [];
    for (const b of list.flat(Infinity)) {
      if (!b) continue;
      switch (b.t) {
        case "p":
          out.push(this.paragraph(b));
          break;
        case "list": {
          const indent = b.indent ?? 12;
          b.items.forEach((item, i) => {
            const last = i === b.items.length - 1;
            out.push(
              this.paragraph(
                { ...b, t: "p", text: item, indent: undefined, after: last ? (b.after ?? 6) : (b.gap ?? 3) },
                { bullet: { char: b.bullet || "•", color: this.color(b.bulletColor || "accent"), indent } }
              )
            );
          });
          break;
        }
        case "rule":
          out.push({
            t: "rule",
            color: this.color(b.color || "line"),
            weight: b.weight ?? 1,
            before: this.space(b.before ?? 4),
            after: this.space(b.after ?? 8),
            width: round(b.width ? Math.min(b.width, width) : width),
            align: b.align || "left",
          });
          break;
        case "spacer":
          out.push({ t: "spacer", height: this.space(b.height ?? 12) });
          break;
        case "break":
          out.push({ t: "break" });
          break;
        case "image":
          out.push(this.image(b, width));
          break;
        case "band":
          out.push(this.band(b, width));
          break;
        case "table":
          out.push(this.table(b, width));
          break;
        case "grid":
          out.push(this.grid(b, width));
          break;
        case "columns":
          out.push(this.columns(b, width));
          break;
        case "scope": {
          const saved = this.vars;
          this.vars = { ...saved, ...b.vars };
          out.push(...this.blocks(b.blocks || [], width));
          this.vars = saved;
          break;
        }
        default:
          throw new Error("Unknown block type: " + b.t);
      }
    }
    return out;
  }

  image(b, width) {
    let w = b.width === "full" ? width : Math.min(b.width ?? width, width);
    let h = b.height ?? w * 0.66;
    const logo = b.kind === "logo" && this.theme.logo && this.theme.logo.src ? this.theme.logo : null;
    if (logo) {
      // Fit the logo inside the frame and keep its proportions.
      const ratio = logo.width / logo.height;
      if (w / h > ratio) w = h * ratio;
      else h = w / ratio;
    }
    return {
      t: "image",
      kind: b.kind || "photo",
      // Never a zero or negative size, which Word cannot open and Google Docs refuses.
      width: Math.max(1, round(w)),
      height: Math.max(1, round(h)),
      align: b.align || "left",
      label: logo ? "" : this.fill(b.label || (b.kind === "logo" ? "Your logo" : "Your picture")),
      fill: this.color(b.fill || "tint"),
      ink: this.color(b.ink || "primary"),
      src: logo ? logo.src : null,
      before: this.space(b.before ?? 0),
      after: this.space(b.after ?? 6),
      alt: logo ? "Logo" : "Offizin picture placeholder",
    };
  }

  band(b, width) {
    const w = b.width ? Math.min(b.width, width) : width;
    return this.grid(
      {
        t: "grid",
        widths: [w],
        abs: true,
        align: b.align || "center",
        rows: [{ height: b.height, exact: b.exact, cells: [{ blocks: b.blocks, fill: b.fill, pad: b.pad ?? 12, valign: b.valign || "middle", border: b.border }] }],
      },
      width
    );
  }

  table(b, width) {
    const head = {
      fill: b.headFill || "primary",
      pad: [5, 6, 5, 6],
      border: null,
    };
    const rows = [];
    if (b.header) {
      rows.push({
        cells: b.header.map((text) => ({ ...head, blocks: [{ t: "p", text, role: "label", color: b.headColor || "onPrimary", after: 0, fixed: true }] })),
      });
    }
    b.rows.forEach((r, i) => {
      rows.push({
        cells: r.map((text, c) => ({
          fill: b.zebra !== false && i % 2 === 1 ? "tint" : null,
          pad: [5, 6, 5, 6],
          border: { bottom: { color: "line", weight: 0.75 } },
          blocks: [{ t: "p", text, role: "small", after: 0, bold: c === 0 && b.firstBold ? true : undefined, color: c === 0 && b.firstColor ? b.firstColor : undefined }],
        })),
      });
    });
    return this.grid({ t: "grid", widths: b.widths || b.rows[0].map(() => 1), rows, align: b.align }, width);
  }

  cellBorder(border) {
    if (!border) return null;
    const side = (s) => (s ? { color: this.color(s.color || "line"), weight: s.weight ?? 0.75, style: s.style || "solid" } : null);
    if (border.all) {
      const s = side(border.all);
      return { top: s, right: s, bottom: s, left: s };
    }
    return { top: side(border.top), right: side(border.right), bottom: side(border.bottom), left: side(border.left) };
  }

  grid(b, width) {
    const total = b.width ? Math.min(b.width, width) : width;
    // Sheet layouts (cards, labels) keep exact gaps so they line up with printed stock.
    const fixedGap = b.fixedGap || b.abs;
    const gap = (b.gap ?? 0) * (fixedGap ? 1 : this.layout.gapScale);
    const n = b.widths.length;
    let cols;
    if (b.abs) {
      cols = b.widths.slice();
    } else {
      const sum = b.widths.reduce((a, c) => a + c, 0);
      const avail = total - gap * (n - 1);
      cols = b.widths.map((f) => (f / sum) * avail);
    }
    cols = cols.map((c) => round(c));
    // Stop with a plain message, before anything reaches the document, if margins leave no room.
    if (cols.some((c) => !(c > 0))) throw new Error("This layout does not fit the page with these margins. Make the margins smaller under Customize, then try again.");
    const rows = b.rows.map((r) => {
      const row = Array.isArray(r) ? { cells: r } : r;
      const cells = [];
      row.cells.forEach((c, i) => {
        const raw = typeof (c.pad ?? 0) === "number" ? Array(4).fill(c.pad ?? 0) : c.pad;
        const pad = raw.map((v) => round(v * this.layout.padScale, 1));
        const inner = Math.max(12, cols[i] - pad[1] - pad[3]);
        if (c.fill) this.backgrounds.push(c.fill);
        const blocks = this.blocks(c.blocks || [], inner);
        if (c.fill) this.backgrounds.pop();
        cells.push({ fill: c.fill ? this.color(c.fill) : null, pad, valign: c.valign || "top", border: this.cellBorder(c.border ?? b.border), blocks });
        if (gap && i < row.cells.length - 1) cells.push({ fill: null, pad: [0, 0, 0, 0], valign: "top", border: null, blocks: [], spacer: true });
      });
      return { height: row.height ? round(row.height) : null, exact: !!row.exact, cells };
    });
    let widths = cols;
    if (gap) {
      widths = [];
      cols.forEach((c, i) => {
        widths.push(c);
        if (i < cols.length - 1) widths.push(round(gap));
      });
    }
    const used = round(widths.reduce((a, c) => a + c, 0));
    return { t: "grid", width: used, cols: widths, rows, align: b.align || (used < width - 1 ? "center" : "left") };
  }

  /** Flowing columns. Word uses native section columns; other editors use the estimated split. */
  columns(b, width) {
    const count = clamp(Math.round(b.count || 2), 1, 4);
    const gap = round((b.gap ?? 18) * this.layout.gapScale, 1);
    const colWidth = round((width - gap * (count - 1)) / count);
    const blocks = this.blocks(b.blocks || [], colWidth);
    return {
      t: "columns",
      count,
      gap,
      width: round(width),
      colWidth,
      rule: b.rule ? { color: this.color((b.rule && b.rule.color) || "line"), weight: (b.rule && b.rule.weight) || 0.75 } : null,
      blocks,
      split: splitForColumns(blocks, count, colWidth),
    };
  }
}

/** Rough height of a resolved block in points, used to balance columns. */
export function estimateHeight(b, width) {
  switch (b.t) {
    case "p": {
      const text = b.runs.map((r) => r.text).join("");
      const lines = text.split("\n").reduce((n, line) => n + Math.max(1, Math.ceil((line.length * b.size * 0.5) / Math.max(24, width - (b.bullet ? b.bullet.indent : 0)))), 0);
      return lines * b.size * b.line * 1.17 + b.before + b.after;
    }
    case "rule":
      return b.before + b.after + b.weight;
    case "spacer":
      return b.height;
    case "image":
      return b.height + b.before + b.after;
    case "grid":
      return b.rows.reduce((sum, r) => {
        const content = Math.max(0, ...r.cells.map((c, i) => c.pad[0] + c.pad[2] + c.blocks.reduce((s, x) => s + estimateHeight(x, Math.max(12, b.cols[i] - c.pad[1] - c.pad[3])), 0)));
        return sum + Math.max(r.height || 0, content);
      }, 0);
    case "columns":
      return b.blocks.reduce((s, x) => s + estimateHeight(x, b.colWidth), 0) / b.count;
    default:
      return 0;
  }
}

/** Indexes where columns 2..n begin, balancing estimated heights without splitting a block. */
export function splitForColumns(blocks, count, width) {
  if (count < 2 || blocks.length < 2) return [];
  const heights = blocks.map((b) => estimateHeight(b, width));
  const total = heights.reduce((a, c) => a + c, 0);
  const starts = [];
  let acc = 0;
  let k = 1;
  for (let i = 0; i < blocks.length && k < count; i++) {
    if (acc >= (total * k) / count && i > (starts[starts.length - 1] ?? 0)) {
      starts.push(i);
      k++;
    }
    acc += heights[i];
  }
  return starts;
}

/** Helpers handed to template build functions. */
function makeContext(composer, page, options, template, records) {
  const W = round(page.width - page.margins.left - page.margins.right);
  const H = round(page.height - page.margins.top - page.margins.bottom);
  return {
    paper: composer.theme.paper,
    page,
    W,
    H,
    info: composer.info,
    opt: options,
    layout: composer.layout,
    hasLogo: !!(composer.theme.logo && composer.theme.logo.src),
    /** Merge records (from a pasted list), or the template's sample records. */
    records,
    hasRecords: !!(options.records && options.records.length),
    /**
     * Column widths for n equal panels that line up with the folds of the full sheet.
     * The outer panels give up the page margin; the inner ones are a full third (or half).
     */
    panels(n) {
      const each = page.width / n;
      return Array.from({ length: n }, (_, i) => {
        let w = each;
        if (i === 0) w -= page.margins.left;
        if (i === n - 1) w -= page.margins.right;
        return round(w);
      });
    },
    setVar(key, value) {
      composer.vars[key] = value;
    },
    template,
  };
}

function resolveOptions(template, options) {
  const opts = {};
  for (const o of template.options || []) {
    const fallback = typeof o.default === "function" ? o.default() : o.default;
    let v = options[o.key] ?? fallback;
    if (o.type === "number") {
      v = Number.isFinite(Number(v)) ? Number(v) : fallback;
      if (Number.isFinite(o.min)) v = Math.max(o.min, v);
      if (Number.isFinite(o.max)) v = Math.min(o.max, v);
    }
    if (o.type === "boolean") v = v === true || v === "true";
    if (o.type === "select" && o.choices && !o.choices.some((c) => c.value === v)) v = fallback;
    opts[o.key] = v;
  }
  return opts;
}

/** Builds the footer (if the template has one), leaving {page} and {pages} for the renderer. */
function resolveFooter(composer, template, ctx) {
  const spec = typeof template.footer === "function" ? template.footer(ctx) : template.footer;
  if (!spec) return null;
  const p = composer.paragraph({ t: "p", text: spec.text, role: spec.role || "small", align: spec.align || "center", color: spec.color || "muted", after: 0, fixed: true });
  return { ...p, align: spec.align || "center" };
}

/**
 * Composes a template into a resolved model.
 * @param template  an entry from TEMPLATES
 * @param theme     { colorScheme, customColors, fontScheme, customFonts, paper, business, logo, layout }
 * @param platform  "word", "google", or "web"
 * @param options   the template's own options plus Customize settings:
 *                  { paperSize, orientation, margins, layout, records, print }
 */
export function compose(template, theme, platform, options = {}) {
  const layout = { ...DEFAULT_LAYOUT, ...((theme && theme.layout) || {}), ...(options.layout || {}) };
  const composer = new Composer(theme, platform, layout);
  const page = resolvePage(template, composer.theme.paper, options);
  const opts = resolveOptions(template, options);
  const records = normalizeRecords(options.records && options.records.length ? options.records : (template.merge && template.merge.sample) || []);
  const ctx = makeContext(composer, page, { ...opts, records: options.records || null }, template, records);
  // A template can give some roles its own look, so its paragraphs follow Word and Google Docs styles.
  if (template.styles) composer.styleOverrides = (typeof template.styles === "function" ? template.styles(ctx) : template.styles) || {};
  const raw = template.build(ctx);

  const printSpec = typeof template.print === "function" ? template.print(ctx) : template.print;
  const print = { ...DEFAULT_PRINT, ...(printSpec || {}), ...(options.print || {}) };

  let modelPage = page;
  let blocks;
  let pages = null;
  let booklet = null;
  if (raw && !Array.isArray(raw) && raw.pages) {
    // Booklets: pages in reading order, padded to a multiple of four.
    pages = raw.pages.map((pg) => composer.blocks(pg, ctx.W));
    while (pages.length % 4) pages.splice(pages.length - 1, 0, []);
    const sheet = twoUpSheet(page) || { width: round(page.width * 2), height: page.height };
    booklet = {
      pages: pages.length,
      sheet: { width: sheet.width, height: sheet.height },
      spreads: bookletSpreads(pages.length),
    };
    if (platform === "google") {
      // Google Docs has no booklet printing, so Offizin lays out the printer spreads itself.
      const m = page.margins;
      modelPage = { width: sheet.width, height: sheet.height, orientation: "landscape", margins: { ...m }, sizeName: (twoUpSheet(page) || {}).sizeName || null };
      const rowH = round(page.height - m.top - m.bottom - 10);
      blocks = [];
      booklet.spreads.forEach(([l, r], i) => {
        if (i) blocks.push({ t: "break" });
        blocks.push({
          t: "grid",
          width: round(page.width * 2 - m.left - m.right),
          cols: [round(page.width - m.left), round(page.width - m.right)],
          rows: [
            {
              height: rowH,
              exact: false,
              cells: [
                { fill: null, pad: [0, m.right, 0, 0], valign: "top", border: null, blocks: pages[l] },
                { fill: null, pad: [0, 0, 0, m.left], valign: "top", border: null, blocks: pages[r] },
              ],
            },
          ],
          align: "left",
        });
      });
      print.bookFold = false;
      print.duplex = "short";
    } else {
      blocks = [];
      pages.forEach((pg, i) => {
        if (i) blocks.push({ t: "break" });
        blocks.push(...pg);
      });
      if (platform === "word") print.bookFold = true;
    }
  } else {
    blocks = composer.blocks(raw, ctx.W);
  }
  // Two pages per sheet needs a small portrait page that fits twice on standard paper.
  if (print.twoPerSheet && (booklet || !twoUpSheet(modelPage))) print.twoPerSheet = false;

  return {
    version: MODEL_VERSION,
    templateId: template.id,
    name: template.name,
    page: modelPage,
    folds: booklet && platform === "google" ? { vertical: [0.5], horizontal: [] } : template.folds ? template.folds(ctx) : null,
    colors: composer.colors,
    fonts: composer.fonts,
    styles: composer.styles(),
    blocks,
    pages,
    // Running text (resumes and CVs) continues onto more pages, as it does in Word and Google Docs.
    flow: !!template.flow && !pages,
    booklet,
    footer: resolveFooter(composer, template, ctx),
    print,
    samples: [...composer.samples],
    missingInfo: composer.missingInfo,
    options: { ...opts, paperSize: options.paperSize || null, orientation: options.orientation || null, margins: options.margins || null, layout: composer.layout, records: options.records || null },
    theme: {
      colorScheme: composer.theme.colorScheme,
      customColors: composer.theme.customColors,
      fontScheme: composer.theme.fontScheme,
      customFonts: composer.theme.customFonts,
    },
  };
}

/** Composes a building block (a page part) for insertion at the cursor. */
export function composePart(part, theme, platform, width = 468) {
  const composer = new Composer(theme, platform, theme && theme.layout);
  const ctx = {
    W: width,
    info: composer.info,
    layout: composer.layout,
    hasLogo: !!(composer.theme.logo && composer.theme.logo.src),
    opt: {},
    setVar(key, value) {
      composer.vars[key] = value;
    },
  };
  const blocks = composer.blocks(part.build(ctx), width);
  return { blocks, samples: [...composer.samples], colors: composer.colors, fonts: composer.fonts, styles: composer.styles() };
}

/**
 * Old-to-new color and font replacements for restyling an existing publication,
 * Publisher style: switch the scheme and the publication follows.
 */
export function restyleMap(fromTheme, toTheme, platform) {
  const a = palette(themeColors({ ...DEFAULT_THEME, ...fromTheme }));
  const b = palette(themeColors({ ...DEFAULT_THEME, ...toTheme }));
  const colors = {};
  for (const key of Object.keys(a)) {
    if (a[key] !== b[key] && !colors[a[key]]) colors[a[key]] = b[key];
  }
  const fa = fontsFor(findFontScheme(fromTheme.fontScheme, fromTheme.customFonts), platform);
  const fb = fontsFor(findFontScheme(toTheme.fontScheme, toTheme.customFonts), platform);
  const fonts = {};
  if (fa.heading.name !== fb.heading.name) fonts[fa.heading.name] = fb.heading.name;
  if (fa.body.name !== fb.body.name && !fonts[fa.body.name]) fonts[fa.body.name] = fb.body.name;
  return { colors, fonts };
}

/** The reverse of a restyle map, for undo. */
export function invertRestyleMap(map) {
  const flip = (o) => Object.fromEntries(Object.entries(o || {}).map(([k, v]) => [v, k]));
  return { colors: flip(map.colors), fonts: flip(map.fonts) };
}

/** Walks every block, including those nested in grid cells and columns. */
export function walkBlocks(blocks, visit) {
  for (const b of blocks) {
    visit(b);
    if (b.t === "grid") for (const r of b.rows) for (const c of r.cells) walkBlocks(c.blocks, visit);
    if (b.t === "columns") walkBlocks(b.blocks, visit);
  }
}

