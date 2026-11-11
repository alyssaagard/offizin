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
import { mix, readableOn, normalizeHex } from "./util.js";

export const COLOR_SCHEMES = [
  { id: "harbor", name: "Harbor", primary: "#1B3A5C", accent: "#C0392B", dark: "#1C2430" },
  { id: "ink-gold", name: "Ink and Gold", primary: "#1A1A1A", accent: "#A88B5C", dark: "#1A1A1A" },
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
];

export function googleFontsHref() {
  const q = GOOGLE_FONT_FAMILIES.map((f) => "family=" + f.replace(/ /g, "+")).join("&");
  return "https://fonts.googleapis.com/css2?" + q + "&display=swap";
}

export function findColorScheme(id) {
  return COLOR_SCHEMES.find((s) => s.id === id) || COLOR_SCHEMES[0];
}

export function findFontScheme(id) {
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
  return {
    primary,
    accent,
    dark,
    white: "#FFFFFF",
    muted: mix(dark, "#FFFFFF", 0.38),
    line: mix(dark, "#FFFFFF", 0.78),
    tint: mix(primary, "#FFFFFF", 0.9),
    accentTint: mix(accent, "#FFFFFF", 0.86),
    deep: mix(primary, "#000000", 0.25),
    onPrimary: readableOn(primary, dark),
    onAccent: readableOn(accent, dark),
    onDark: readableOn(dark, dark),
  };
}

/** Fonts for one platform: "word", "google", or "web" (the demo, which uses Google Fonts). */
export function fontsFor(scheme, platform) {
  const pick = (f) => ({
    name: platform === "word" ? f.word : f.google,
    css: platform === "word" ? f.css : `'${f.google}', ${f.css}`,
  });
  return { heading: pick(scheme.heading), body: pick(scheme.body), headingBold: scheme.headingBold };
}
