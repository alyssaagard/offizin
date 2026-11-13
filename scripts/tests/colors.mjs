/**
 * Color tests: Apply design and Offizin's Undo for every pair of color schemes, and readable
 * text in every scheme.
 *
 *  1. Every scheme's palette gives each token its own value, so a restyle can tell them apart.
 *  2. Restyling a publication from scheme A to scheme B gives exactly the colors of the same
 *     publication made in scheme B, and undoing the restyle gives back exactly scheme A.
 *  3. Accent and gray text reach 4.5:1 against the paper or the light fill behind them.
 *
 * Run with: node scripts/check.mjs --only=colors
 */

const TEMPLATE_IDS = ["trifold", "bifold", "newsletter", "event-flyer", "sale-flyer", "poster", "postcards", "certificate", "resume", "calendar", "program", "menu"];

/** Every color in a model, in a fixed order, leaving out picture placeholders (their fills are blends). */
function colorsOf(value, out = [], path = "") {
  if (typeof value === "string") {
    if (/^#[0-9A-F]{6}$/i.test(value)) out.push([path, value.toUpperCase()]);
  } else if (Array.isArray(value)) {
    value.forEach((v, i) => colorsOf(v, out, `${path}[${i}]`));
  } else if (value && typeof value === "object") {
    if (value.t === "image") return out;
    for (const k of Object.keys(value)) colorsOf(value[k], out, `${path}.${k}`);
  }
  return out;
}

/** Paragraphs with the fill behind them (the innermost cell fill, or white paper). */
function paragraphsOnFills(blocks, bg = "#FFFFFF", out = []) {
  for (const b of blocks) {
    if (b.t === "p") out.push({ p: b, bg });
    if (b.t === "grid") for (const r of b.rows) for (const c of r.cells) paragraphsOnFills(c.blocks, c.fill || bg, out);
    if (b.t === "columns") paragraphsOnFills(b.blocks, bg, out);
  }
  return out;
}

export async function run({ fail, ok, load }) {
  const { compose, restyleMap, invertRestyleMap } = await load("docs/app/core/compose.js");
  const { COLOR_SCHEMES, palette } = await load("docs/app/core/schemes.js");
  const { TEMPLATES } = await load("docs/app/core/templates.js");
  const { contrastRatio } = await load("docs/app/core/util.js");
  const templates = TEMPLATE_IDS.map((id) => TEMPLATES.find((t) => t.id === id)).filter(Boolean);
  if (templates.length !== TEMPLATE_IDS.length) fail("colors: a template in TEMPLATE_IDS no longer exists");

  // 1. Distinct palettes ------------------------------------------------------------
  {
    const bad = COLOR_SCHEMES.filter((s) => {
      const values = Object.values(palette(s));
      return new Set(values).size !== values.length;
    });
    bad.length ? bad.forEach((s) => fail(`colors: two tokens share a value in ${s.name}`)) : ok(`colors: all ${COLOR_SCHEMES.length} palettes give every token its own value`);
  }

  // 2. Restyle and undo for every pair of schemes -------------------------------------
  {
    let problems = 0;
    let pairs = 0;
    const report = (msg) => {
      problems++;
      if (problems <= 12) fail(msg);
    };
    const models = new Map();
    const colorsFor = (scheme, t) => {
      const key = `${scheme}/${t.id}`;
      if (!models.has(key)) models.set(key, colorsOf(compose(t, { colorScheme: scheme, paper: "letter" }, "word")));
      return models.get(key);
    };
    for (const a of COLOR_SCHEMES) {
      for (const b of COLOR_SCHEMES) {
        if (a === b) continue;
        pairs++;
        const map = restyleMap({ colorScheme: a.id }, { colorScheme: b.id }, "word").colors;
        const back = invertRestyleMap({ colors: map, fonts: {} }).colors;
        for (const t of templates) {
          const before = colorsFor(a.id, t);
          const want = colorsFor(b.id, t);
          const after = before.map(([p, c]) => [p, map[c] || c]);
          const wrong = after.findIndex(([p, c], i) => !want[i] || want[i][0] !== p || want[i][1] !== c);
          if (wrong >= 0) {
            report(`colors: ${a.name} to ${b.name} on ${t.id}: ${after[wrong][0]} became ${after[wrong][1]}, but a ${b.name} publication has ${want[wrong] ? want[wrong][1] : "nothing"}`);
            continue;
          }
          const undone = after.map(([p, c]) => [p, back[c] || c]);
          const lost = undone.findIndex(([p, c], i) => before[i][1] !== c);
          if (lost >= 0) report(`colors: undo of ${a.name} to ${b.name} on ${t.id}: ${undone[lost][0]} came back as ${undone[lost][1]}, not ${before[lost][1]}`);
        }
      }
    }
    if (!problems) ok(`colors: Apply design matches a fresh publication and Undo restores every color, for all ${pairs} scheme pairs on ${templates.length} templates`);
  }

  // 3. Readable accent and gray text --------------------------------------------------
  {
    let problems = 0;
    let checked = 0;
    const report = (msg) => {
      problems++;
      if (problems <= 12) fail(msg);
    };
    for (const s of COLOR_SCHEMES) {
      const pal = palette(s);
      for (const t of TEMPLATES) {
        const m = compose(t, { colorScheme: s.id, paper: "letter" }, "word");
        for (const { p, bg } of paragraphsOnFills(m.blocks)) {
          const role = p.color === pal.accentText ? "accent" : p.color === pal.muted ? "gray" : p.color === pal.accent ? "accent" : null;
          if (!role) continue;
          // Accent text on a dark fill uses the accent itself; only light backgrounds are held to 4.5:1 here.
          if (contrastRatio(bg, "#FFFFFF") >= 2) continue;
          checked++;
          const ratio = contrastRatio(p.color, bg);
          const text = p.runs.map((r) => r.text).join("").slice(0, 30);
          if (ratio < 4.5) report(`colors: ${s.name} ${t.id}: ${role} text "${text}" is ${ratio.toFixed(2)}:1 on ${bg}`);
        }
      }
    }
    if (!problems) ok(`colors: ${checked} accent and gray lines read at 4.5:1 or better on their backgrounds in every scheme`);
  }
}
