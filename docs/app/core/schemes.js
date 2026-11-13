/**
 * Color schemes and font schemes.
 *
 * A color scheme has three chosen colors. Everything else (tints, muted text, text on
 * colored backgrounds) is derived, so any custom brand colors work in every template.
 *
 * A font scheme names a heading font and a body font for each platform. Word and Google Docs
 * ship different fonts, so each scheme pairs fonts that are installed on that platform:
 *   word:   fonts included with Microsoft 365 (Windows, Mac, and Word on the web)
 *   google: Google Fonts families, which Google Docs offers by name
 *   css:    the stack used by the on-screen preview
 */
import { mix, readableOn, normalizeHex, contrastRatio, hexToRgb, rgbToHex } from "./util.js";

/**
 * The first two schemes are black and white, for publications that print well on any printer.
 * "Ink" is the default. Each scheme keeps its main color and its text color distinct (even when
 * both are near black), so restyling can tell colored panels from text.
 */
export const COLOR_SCHEMES = [
  { id: "ink", name: "Ink", primary: "#111111", accent: "#5C5C5C", dark: "#1A1A1A" },
  { id: "graphite", name: "Graphite", primary: "#404040", accent: "#111111", dark: "#1A1A1A" },
  { id: "harbor", name: "Harbor", primary: "#1B3A5C", accent: "#C0392B", dark: "#1C2430" },
  { id: "ink-gold", name: "Ink and Gold", primary: "#121212", accent: "#A88B5C", dark: "#1A1A1A" },
  { id: "orchard", name: "Orchard", primary: "#2F5D3A", accent: "#D98E04", dark: "#1E2A21" },
  { id: "terracotta", name: "Terracotta", primary: "#9C4A2F", accent: "#2E6E73", dark: "#2B1D18" },
  { id: "lagoon", name: "Lagoon", primary: "#0E6E72", accent: "#F28C38", dark: "#13292A" },
  { id: "plum", name: "Plum", primary: "#5B2A5E", accent: "#C9962B", dark: "#251628" },
  { id: "slate", name: "Slate", primary: "#36454F", accent: "#3B9EA1", dark: "#1F262B" },
  { id: "rosewood", name: "Rosewood", primary: "#7A2E3A", accent: "#B88A4A", dark: "#2A1519" },
  { id: "citrus", name: "Citrus", primary: "#1F6F8B", accent: "#E5A400", dark: "#14262E" },
  { id: "cobalt", name: "Cobalt", primary: "#2848A8", accent: "#E4572E", dark: "#151B33" },
  { id: "sage", name: "Sage", primary: "#4F6B5B", accent: "#B5677D", dark: "#1F2924" },
  { id: "charcoal", name: "Charcoal", primary: "#2B2D42", accent: "#E63946", dark: "#16171F" },
];

export const FONT_SCHEMES = [
  {
    // The default: headings in a typewriter-like code face, the same one the live demo uses.
    // Word has no IBM Plex Mono, so it uses Consolas, the closest face that every copy of Word has.
    id: "typewriter",
    name: "Typewriter",
    headingBold: true,
    heading: { word: "Consolas", google: "IBM Plex Mono", css: "'IBM Plex Mono', Consolas, 'SF Mono', Menlo, monospace" },
    body: { word: "Arial", google: "Arial", css: "Arial, 'Helvetica Neue', Helvetica, sans-serif" },
  },
  {
    id: "editorial",
    name: "Editorial",
    headingBold: true,
    heading: { word: "Georgia", google: "Georgia", css: "Georgia, 'Times New Roman', serif" },
    body: { word: "Arial", google: "Arial", css: "Arial, 'Helvetica Neue', Helvetica, sans-serif" },
  },
  {
    id: "classic",
    name: "Classic",
    headingBold: false,
    heading: { word: "Garamond", google: "EB Garamond", css: "Garamond, 'EB Garamond', serif" },
    body: { word: "Calibri", google: "Lato", css: "Calibri, Lato, 'Segoe UI', sans-serif" },
  },
  {
    id: "modern",
    name: "Modern",
    headingBold: true,
    heading: { word: "Century Gothic", google: "Montserrat", css: "'Century Gothic', Montserrat, sans-serif" },
    body: { word: "Calibri", google: "Open Sans", css: "Calibri, 'Open Sans', sans-serif" },
  },
  {
    id: "bold",
    name: "Bold",
    headingBold: true,
    heading: { word: "Franklin Gothic Medium", google: "Oswald", css: "'Franklin Gothic Medium', Oswald, 'Arial Narrow', sans-serif" },
    body: { word: "Georgia", google: "Lora", css: "Georgia, Lora, serif" },
  },
  {
    id: "elegant",
    name: "Elegant",
    headingBold: false,
    heading: { word: "Palatino Linotype", google: "Cormorant Garamond", css: "'Palatino Linotype', 'Cormorant Garamond', Palatino, serif" },
    body: { word: "Gill Sans MT", google: "Raleway", css: "'Gill Sans MT', 'Gill Sans', Raleway, sans-serif" },
  },
  {
    id: "friendly",
    name: "Friendly",
    headingBold: true,
    heading: { word: "Rockwell", google: "Roboto Slab", css: "Rockwell, 'Roboto Slab', serif" },
    body: { word: "Corbel", google: "Nunito", css: "Corbel, Nunito, sans-serif" },
  },
  {
    id: "clean",
    name: "Clean",
    headingBold: true,
    heading: { word: "Aptos Display", google: "Inter", css: "'Aptos Display', Inter, 'Segoe UI', sans-serif" },
    body: { word: "Aptos", google: "Inter", css: "Aptos, Inter, 'Segoe UI', sans-serif" },
  },
  {
    id: "traditional",
    name: "Traditional",
    headingBold: true,
    heading: { word: "Cambria", google: "Libre Baskerville", css: "Cambria, 'Libre Baskerville', serif" },
    body: { word: "Calibri", google: "Source Sans 3", css: "Calibri, 'Source Sans 3', sans-serif" },
  },
];

/** Google Fonts families used by the font schemes, for preview pages that load them. */
export const GOOGLE_FONT_FAMILIES = [
  "EB Garamond:ital,wght@0,400;0,700;1,400",
  "Merriweather:ital,wght@0,400;0,700;1,400",
  "Playfair Display:ital,wght@0,400;0,700;1,400",
  "Roboto:ital,wght@0,400;0,700;1,400",
  "Lato:ital,wght@0,400;0,700;1,400",
  "Montserrat:ital,wght@0,400;0,700;1,400",
  "Open Sans:ital,wght@0,400;0,700;1,400",
  "Oswald:wght@400;600",
  "Lora:ital,wght@0,400;0,700;1,400",
  "Cormorant Garamond:ital,wght@0,500;0,700;1,500",
  "Raleway:ital,wght@0,400;0,700;1,400",
  "Roboto Slab:wght@400;700",
  "Nunito:ital,wght@0,400;0,700;1,400",
  "Inter:ital,wght@0,400;0,700;1,400",
  "Libre Baskerville:ital,wght@0,400;0,700;1,400",
  "Source Sans 3:ital,wght@0,400;0,700;1,400",
  "IBM Plex Mono:ital,wght@0,400;0,500;0,700;1,400",
];

export function googleFontsHref() {
  const q = GOOGLE_FONT_FAMILIES.map((f) => "family=" + f.replace(/ /g, "+")).join("&");
  return "https://fonts.googleapis.com/css2?" + q + "&display=swap";
}

export function findColorScheme(id) {
  return COLOR_SCHEMES.find((s) => s.id === id) || COLOR_SCHEMES[0];
}

/**
 * Fonts people can pick for their own font scheme. Word fonts come with Microsoft 365
 * (installed on Windows and Mac, and available as cloud fonts in Word on the web).
 * Google fonts are Google Fonts families that Google Docs offers by name.
 * `css` is the preview fallback stack.
 */
export const WORD_FONT_CHOICES = [
  { name: "Aptos", css: "Aptos, 'Segoe UI', sans-serif" },
  { name: "Aptos Display", css: "'Aptos Display', Aptos, 'Segoe UI', sans-serif" },
  { name: "Arial", css: "Arial, Helvetica, sans-serif" },
  { name: "Bookman Old Style", css: "'Bookman Old Style', Georgia, serif" },
  { name: "Calibri", css: "Calibri, Carlito, sans-serif" },
  { name: "Cambria", css: "Cambria, Caladea, serif" },
  { name: "Century Gothic", css: "'Century Gothic', Montserrat, sans-serif" },
  { name: "Consolas", css: "Consolas, 'IBM Plex Mono', Menlo, monospace" },
  { name: "Constantia", css: "Constantia, Georgia, serif" },
  { name: "Corbel", css: "Corbel, 'Segoe UI', sans-serif" },
  { name: "Franklin Gothic Medium", css: "'Franklin Gothic Medium', 'Arial Narrow', sans-serif" },
  { name: "Garamond", css: "Garamond, 'EB Garamond', serif" },
  { name: "Georgia", css: "Georgia, serif" },
  { name: "Gill Sans MT", css: "'Gill Sans MT', 'Gill Sans', sans-serif" },
  { name: "Palatino Linotype", css: "'Palatino Linotype', Palatino, serif" },
  { name: "Rockwell", css: "Rockwell, 'Roboto Slab', serif" },
  { name: "Segoe UI", css: "'Segoe UI', system-ui, sans-serif" },
  { name: "Times New Roman", css: "'Times New Roman', Times, serif" },
  { name: "Trebuchet MS", css: "'Trebuchet MS', sans-serif" },
  { name: "Verdana", css: "Verdana, sans-serif" },
];

export const GOOGLE_FONT_CHOICES = [
  { name: "Arial", css: "Arial, Helvetica, sans-serif" },
  { name: "Cormorant Garamond", css: "'Cormorant Garamond', Garamond, serif" },
  { name: "EB Garamond", css: "'EB Garamond', Garamond, serif" },
  { name: "Georgia", css: "Georgia, serif" },
  { name: "IBM Plex Mono", css: "'IBM Plex Mono', Consolas, monospace" },
  { name: "Inter", css: "Inter, sans-serif" },
  { name: "Lato", css: "Lato, sans-serif" },
  { name: "Libre Baskerville", css: "'Libre Baskerville', serif" },
  { name: "Lora", css: "Lora, serif" },
  { name: "Merriweather", css: "Merriweather, serif" },
  { name: "Montserrat", css: "Montserrat, sans-serif" },
  { name: "Nunito", css: "Nunito, sans-serif" },
  { name: "Open Sans", css: "'Open Sans', sans-serif" },
  { name: "Oswald", css: "Oswald, sans-serif" },
  { name: "Playfair Display", css: "'Playfair Display', serif" },
  { name: "Raleway", css: "Raleway, sans-serif" },
  { name: "Roboto", css: "Roboto, sans-serif" },
  { name: "Roboto Slab", css: "'Roboto Slab', serif" },
  { name: "Source Sans 3", css: "'Source Sans 3', sans-serif" },
  { name: "Times New Roman", css: "'Times New Roman', Times, serif" },
  { name: "Verdana", css: "Verdana, sans-serif" },
];

/**
 * The closest font on the other editor, so a font scheme of your own works in Word and in
 * Google Docs. FONT_EQUIVALENTS.word maps Word fonts to Google fonts, and .google the reverse.
 */
export const FONT_EQUIVALENTS = {
  word: {
    Aptos: "Inter",
    "Aptos Display": "Inter",
    Arial: "Arial",
    "Bookman Old Style": "Libre Baskerville",
    Calibri: "Open Sans",
    Cambria: "Merriweather",
    "Century Gothic": "Montserrat",
    Consolas: "IBM Plex Mono",
    Constantia: "Lora",
    Corbel: "Source Sans 3",
    "Franklin Gothic Medium": "Oswald",
    Garamond: "EB Garamond",
    Georgia: "Georgia",
    "Gill Sans MT": "Lato",
    "Palatino Linotype": "Libre Baskerville",
    Rockwell: "Roboto Slab",
    "Segoe UI": "Open Sans",
    "Times New Roman": "Times New Roman",
    "Trebuchet MS": "Nunito",
    Verdana: "Verdana",
  },
  google: {
    Arial: "Arial",
    "Cormorant Garamond": "Garamond",
    "EB Garamond": "Garamond",
    Georgia: "Georgia",
    "IBM Plex Mono": "Consolas",
    Inter: "Aptos",
    Lato: "Gill Sans MT",
    "Libre Baskerville": "Bookman Old Style",
    Lora: "Constantia",
    Merriweather: "Cambria",
    Montserrat: "Century Gothic",
    Nunito: "Trebuchet MS",
    "Open Sans": "Calibri",
    Oswald: "Franklin Gothic Medium",
    "Playfair Display": "Georgia",
    Raleway: "Century Gothic",
    Roboto: "Arial",
    "Roboto Slab": "Rockwell",
    "Source Sans 3": "Corbel",
    "Times New Roman": "Times New Roman",
    Verdana: "Verdana",
  },
};

const cssFor = (list, name, generic) => (list.find((f) => f.name === name) || { css: `'${name}', ${generic}` }).css;

/**
 * A font scheme by ID. "custom" builds one from the user's own picks:
 * customFonts = { heading: { word, google }, body: { word, google }, headingBold }.
 */
export function findFontScheme(id, customFonts) {
  if (id === "custom" && customFonts && customFonts.heading && customFonts.body) {
    const h = customFonts.heading;
    const b = customFonts.body;
    return {
      id: "custom",
      name: "Your fonts",
      headingBold: customFonts.headingBold !== false,
      heading: { word: h.word || "Georgia", google: h.google || "Georgia", css: cssFor(WORD_FONT_CHOICES, h.word || "Georgia", "serif") },
      body: { word: b.word || "Arial", google: b.google || "Arial", css: cssFor(WORD_FONT_CHOICES, b.word || "Arial", "sans-serif") },
    };
  }
  return FONT_SCHEMES.find((s) => s.id === id) || FONT_SCHEMES[0];
}

/**
 * The full palette a template can use, from three chosen colors.
 * Token names are what templates write: "primary", "accent", "tint", "onPrimary", and so on.
 */
export function palette(base) {
  const primary = normalizeHex(base.primary);
  const accent = normalizeHex(base.accent);
  const dark = normalizeHex(base.dark || "#1A1A1A");
  const tint = mix(primary, "#FFFFFF", 0.9);
  const accentTint = mix(accent, "#FFFFFF", 0.86);
  // The order matters: when two tokens share a color, the earlier one keeps it exactly.
  return distinct({
    primary,
    accent,
    dark,
    white: "#FFFFFF",
    muted: mix(dark, "#FFFFFF", 0.3),
    line: mix(dark, "#FFFFFF", 0.78),
    tint,
    accentTint,
    deep: mix(primary, "#000000", 0.25),
    onPrimary: readableOn(primary, dark),
    onAccent: readableOn(accent, dark),
    onDark: readableOn(dark, dark),
    // Text set in the accent color: the accent, darkened just enough to read at 4.5:1 on the
    // light backgrounds it sits on. Rules, fills, and bullets keep the accent itself.
    accentText: readableTone(accent, dark, ["#FFFFFF", tint, accentTint]),
  });
}

/** Mixes a color toward dark in 5% steps until it reaches 4.5:1 against every background. */
function readableTone(color, dark, backgrounds) {
  for (let i = 0; i <= 20; i++) {
    const c = mix(color, dark, i / 20);
    if (backgrounds.every((bg) => contrastRatio(c, bg) >= 4.5)) return c;
  }
  return dark;
}

/**
 * Gives every token its own color value. Apply design and Offizin's Undo replace one color value
 * with another, so two tokens that share a value (white text on the main color and white text on
 * the accent, for example) could not be told apart and would both go the same way. A repeated
 * value moves by one step in its blue channel, which no one can see in print or on screen.
 */
function distinct(tokens) {
  const used = new Set();
  const out = {};
  for (const [key, value] of Object.entries(tokens)) {
    let c = value;
    for (let step = 1; used.has(c); step++) {
      const [r, g, b] = hexToRgb(value);
      c = rgbToHex([r, g, b >= 128 ? b - step : b + step]);
    }
    used.add(c);
    out[key] = c;
  }
  return out;
}

/** Fonts for one platform: "word", "google", or "web" (the demo, which uses Google Fonts). */
export function fontsFor(scheme, platform) {
  const pick = (f) => ({
    name: platform === "word" ? f.word : f.google,
    css: platform === "word" ? f.css : `'${f.google}', ${f.css}`,
  });
  return { heading: pick(scheme.heading), body: pick(scheme.body), headingBold: scheme.headingBold };
}
