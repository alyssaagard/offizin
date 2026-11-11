/**
 * The composer turns a template plus the user's design choices into a resolved layout model.
 *
 * Templates describe publications with roles ("title", "body"), color tokens ("primary"),
 * and business information fields ("{org}"). The composer resolves all of that into plain
 * values: font names, point sizes, hex colors, and column widths. Each renderer (the web
 * preview, Word, and Google Docs) then draws the same model, so the three always match.
 *
 * Resolved block types: p, rule, spacer, break, image, grid. Lists, bands, and data tables
 * in templates are converted into those.
 */
import { pageSize, round, mix } from "./util.js";
import { palette, fontsFor, findColorScheme, findFontScheme } from "./schemes.js";

export const MODEL_VERSION = 1;

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
];

export const DEFAULT_THEME = {
  colorScheme: "harbor",
  customColors: null, // { primary, accent, dark } when the user picks their own
  fontScheme: "editorial",
  paper: "letter", // "letter" or "a4"
  business: {},
  logo: null, // { src: dataURL, width, height } in pixels
};

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
  return { top: m.top, right: m.right, bottom: m.bottom, left: m.left };
}

/** Resolves a template's page description for the chosen paper. */
export function resolvePage(template, paper) {
  const spec = typeof template.page === "function" ? template.page(paper) : template.page;
  const size = spec.width ? { width: spec.width, height: spec.height } : pageSize(spec.size === "paper" || !spec.size ? paper : spec.size, spec.orientation);
  const margins = normMargins(spec.margins ?? 36);
  return {
    width: round(size.width),
    height: round(size.height),
    orientation: size.width > size.height ? "landscape" : "portrait",
    margins,
  };
}

// ---------------------------------------------------------------------------
// Composition
// ---------------------------------------------------------------------------

class Composer {
  constructor(theme, platform) {
    this.theme = { ...DEFAULT_THEME, ...theme };
    this.platform = platform;
    this.colors = palette(themeColors(this.theme));
    this.fontScheme = findFontScheme(this.theme.fontScheme);
    this.fonts = fontsFor(this.fontScheme, platform);
    const { values, missing } = businessValues(this.theme.business);
    this.info = values;
    this.missingInfo = missing;
    this.samples = new Set();
    this.vars = {};
  }

  color(token, fallback = "dark") {
    if (!token) token = fallback;
    if (typeof token === "string" && token.startsWith("#")) return token.toUpperCase();
    if (typeof token === "object" && token.mix) {
      return mix(this.color(token.mix[0]), this.color(token.mix[1]), token.mix[2]);
    }
    return this.colors[token] || this.colors[fallback];
  }

  /** Replaces {field} tokens with business information or template variables. */
  fill(text) {
    return String(text).replace(/\{(\w+)\}/g, (m, key) => {
      if (key in this.vars) return String(this.vars[key]);
      if (key in this.info) return this.info[key];
      return m;
    });
  }

  recordSample(raw, block) {
    if (block.fixed) return;
    if (/\{\w+\}/.test(raw)) return;
    const clean = raw.replace(/\*\*/g, "").replace(/\s+/g, " ").trim();
    if (clean.length < 12 || !/[a-z]/i.test(clean)) return;
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

  paragraph(b, extra = {}) {
    const role = ROLES[b.role || "body"] || ROLES.body;
    const fontKey = b.font || role.font;
    const font = this.fonts[fontKey];
    const isHeading = fontKey === "heading";
    const bold = b.bold ?? (role.bold ? (isHeading ? this.fonts.headingBold : true) : false);
    const raw = b.text ?? "";
    this.recordSample(raw, b);
    const caps = b.caps ?? role.caps ?? false;
    const p = {
      t: "p",
      runs: this.runs(this.fill(raw), caps),
      font: font.name,
      css: font.css,
      size: b.size ?? role.size,
      color: this.color(b.color, role.color),
      bold,
      italic: b.italic ?? role.italic ?? false,
      align: b.align || "left",
      before: b.before ?? 0,
      after: b.after ?? role.after,
      line: b.line ?? role.line,
      tracking: b.tracking ?? (caps ? 0.08 : 0),
      ...extra,
    };
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
                { ...b, t: "p", text: item, after: last ? (b.after ?? 6) : (b.gap ?? 3) },
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
            before: b.before ?? 4,
            after: b.after ?? 8,
            width: round(b.width ? Math.min(b.width, width) : width),
            align: b.align || "left",
          });
          break;
        case "spacer":
          out.push({ t: "spacer", height: b.height ?? 12 });
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
      width: round(w),
      height: round(h),
      align: b.align || "left",
      label: logo ? "" : this.fill(b.label || (b.kind === "logo" ? "Your logo" : "Your picture")),
      fill: this.color(b.fill || "tint"),
      ink: this.color(b.ink || "primary"),
      src: logo ? logo.src : null,
      before: b.before ?? 0,
      after: b.after ?? 6,
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
    const gap = b.gap ?? 0;
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
    const rows = b.rows.map((r) => {
      const row = Array.isArray(r) ? { cells: r } : r;
      const cells = [];
      row.cells.forEach((c, i) => {
        const pad = typeof (c.pad ?? 0) === "number" ? Array(4).fill(c.pad ?? 0) : c.pad;
        const inner = Math.max(12, cols[i] - pad[1] - pad[3]);
        cells.push({
          fill: c.fill ? this.color(c.fill) : null,
          pad,
          valign: c.valign || "top",
          border: this.cellBorder(c.border ?? b.border),
          blocks: this.blocks(c.blocks || [], inner),
        });
        if (gap && i < row.cells.length - 1) cells.push({ fill: null, pad: [0, 0, 0, 0], valign: "top", border: null, blocks: [], spacer: true });
      });
      return { height: row.height ? round(row.height) : null, exact: !!row.exact, cells };
    });
    let widths = cols;
    if (gap) {
      widths = [];
      cols.forEach((c, i) => {
        widths.push(c);
        if (i < cols.length - 1) widths.push(gap);
      });
    }
    const used = round(widths.reduce((a, c) => a + c, 0));
    return { t: "grid", width: used, cols: widths, rows, align: b.align || (used < width - 1 ? "center" : "left") };
  }
}

/** Helpers handed to template build functions. */
function makeContext(composer, page, options, template) {
  const W = round(page.width - page.margins.left - page.margins.right);
  const H = round(page.height - page.margins.top - page.margins.bottom);
  return {
    paper: composer.theme.paper,
    page,
    W,
    H,
    info: composer.info,
    opt: options,
    hasLogo: !!(composer.theme.logo && composer.theme.logo.src),
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

/**
 * Composes a template into a resolved model.
 * @param template  an entry from TEMPLATES
 * @param theme     { colorScheme, customColors, fontScheme, paper, business, logo }
 * @param platform  "word", "google", or "web"
 * @param options   values for the template's own options (for example, the calendar month)
 */
export function compose(template, theme, platform, options = {}) {
  const composer = new Composer(theme, platform);
  const page = resolvePage(template, composer.theme.paper);
  const opts = {};
  for (const o of template.options || []) opts[o.key] = options[o.key] ?? (typeof o.default === "function" ? o.default() : o.default);
  const ctx = makeContext(composer, page, opts, template);
  const raw = template.build(ctx);
  const blocks = composer.blocks(raw, ctx.W);
  return {
    version: MODEL_VERSION,
    templateId: template.id,
    name: template.name,
    page,
    folds: template.folds ? template.folds(ctx) : null,
    colors: composer.colors,
    fonts: composer.fonts,
    blocks,
    samples: [...composer.samples],
    missingInfo: composer.missingInfo,
    theme: {
      colorScheme: composer.theme.colorScheme,
      customColors: composer.theme.customColors,
      fontScheme: composer.theme.fontScheme,
    },
  };
}

/** Composes a building block (a page part) for insertion at the cursor. */
export function composePart(part, theme, platform, width = 468) {
  const composer = new Composer(theme, platform);
  const ctx = {
    W: width,
    info: composer.info,
    hasLogo: !!(composer.theme.logo && composer.theme.logo.src),
    opt: {},
    setVar(key, value) {
      composer.vars[key] = value;
    },
  };
  const blocks = composer.blocks(part.build(ctx), width);
  return { blocks, samples: [...composer.samples], colors: composer.colors, fonts: composer.fonts };
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
  const fa = fontsFor(findFontScheme(fromTheme.fontScheme), platform);
  const fb = fontsFor(findFontScheme(toTheme.fontScheme), platform);
  const fonts = {};
  if (fa.heading.name !== fb.heading.name) fonts[fa.heading.name] = fb.heading.name;
  if (fa.body.name !== fb.body.name && !fonts[fa.body.name]) fonts[fa.body.name] = fb.body.name;
  return { colors, fonts };
}

/** Walks every block, including those nested in grid cells. */
export function walkBlocks(blocks, visit) {
  for (const b of blocks) {
    visit(b);
    if (b.t === "grid") for (const r of b.rows) for (const c of r.cells) walkBlocks(c.blocks, visit);
  }
}
