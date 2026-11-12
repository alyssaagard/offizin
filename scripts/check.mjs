/**
 * Offizin's automated checks. Run with: npm run check   (after npm run build)
 *
 *  1. House style: no em dashes in any text file.
 *  2. Every template composes for Word, Google Docs, and the web, on Letter and A4,
 *     with no missing values and nothing wider than the page.
 *  3. Every template becomes a .docx whose XML is well formed (uses xmllint when available).
 *  4. The Word manifest is well formed, and the Google sidebar bundle parses.
 *  5. The Google Docs code (Code.gs) builds every template and page part against a
 *     simulated DocumentApp that enforces the same rules Docs does: valid indexes, and every
 *     table cell and the document ending in a paragraph.
 */
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { makeGoogle, invariants } from "./lib/google-mock.mjs";
import { execFileSync } from "node:child_process";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const ONLY = (process.argv.find((a) => a.startsWith("--only=")) || "").slice(7) || null;
const load = (p) => import(pathToFileURL(path.join(ROOT, p)).href);
let failures = 0;
const fail = (msg) => {
  failures++;
  console.log("  FAIL " + msg);
};
const ok = (msg) => console.log("  ok   " + msg);

const { TEMPLATES, PARTS } = await load("docs/app/core/templates.js");
const { compose, composePart, walkBlocks } = await load("docs/app/core/compose.js");
const { docxBytes, documentXml } = await load("docs/app/render/ooxml.js");

const STUB_PNG = { mime: "image/png", base64: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==" };

/** Each file in scripts/tests/ exports run({ ROOT, fail, ok, load, STUB_PNG }). */
async function runAreaTests() {
  const testDir = path.join(ROOT, "scripts/tests");
  const files = fs.existsSync(testDir) ? fs.readdirSync(testDir).filter((f) => f.endsWith(".mjs")).sort() : [];
  for (const f of files) {
    if (ONLY && !f.startsWith(ONLY)) continue;
    console.log("Tests: " + f.replace(/\.mjs$/, ""));
    try {
      const mod = await import(pathToFileURL(path.join(testDir, f)).href);
      await mod.run({ ROOT, fail, ok, load, STUB_PNG });
    } catch (e) {
      fail(`${f}: ${(e && e.stack) || e}`);
    }
  }
}

// npm run check -- --only=word   runs just scripts/tests/word.mjs
if (ONLY) {
  await runAreaTests();
  console.log(failures ? `\n${failures} problem${failures > 1 ? "s" : ""} found.` : "\nAll checks passed.");
  process.exit(failures ? 1 : 0);
}

// 1. House style ------------------------------------------------------------
console.log("House style");
{
  const skip = new Set([".git", "node_modules"]);
  const textExt = /\.(js|mjs|gs|html|css|json|md|xml|txt|svg|yml)$/i;
  const bad = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      if (skip.has(e.name)) continue;
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else if (textExt.test(e.name) && fs.readFileSync(p, "utf8").includes("\u2014")) bad.push(path.relative(ROOT, p));
    }
  };
  walk(ROOT);
  bad.length ? bad.forEach((f) => fail("em dash in " + f)) : ok("no em dashes");
}

// 2. Composition ------------------------------------------------------------
console.log("Templates");
for (const t of TEMPLATES) {
  for (const platform of ["word", "google", "web"]) {
    for (const paper of ["letter", "a4"]) {
      const m = compose(t, { paper, colorScheme: "plum", fontScheme: "elegant", business: { org: "Test & Co. <Ltd>" } }, platform);
      const W = m.page.width - m.page.margins.left - m.page.margins.right;
      const H = m.page.height - m.page.margins.top - m.page.margins.bottom;
      walkBlocks(m.blocks, (b) => {
        const s = JSON.stringify(b);
        if (/NaN|undefined|Infinity/.test(s)) fail(`${t.id}/${platform}/${paper}: bad value in ${b.t}`);
        if (b.t === "grid") {
          b.rows.forEach((r) => r.cells.length !== b.cols.length && fail(`${t.id}: row has ${r.cells.length} cells for ${b.cols.length} columns`));
          b.rows.forEach((r) => r.height && r.height > H && fail(`${t.id}/${paper}: row taller than the page (${r.height} > ${H})`));
        }
      });
      for (const b of m.blocks) if (b.t === "grid" && b.width > W + 0.6) fail(`${t.id}/${paper}: grid ${b.width}pt wider than ${W}pt`);
    }
  }
}
for (const p of PARTS) for (const platform of ["word", "google", "web"]) composePart(p, {}, platform, 468);
ok(`${TEMPLATES.length} templates and ${PARTS.length} parts compose`);

// 3. Word documents ---------------------------------------------------------
console.log("Word documents");
let xmllint = true;
try {
  execFileSync("xmllint", ["--version"], { stdio: "ignore" });
} catch (e) {
  xmllint = false;
  console.log("  note xmllint not found; checking structure only");
}
const wellFormed = (xml, label) => {
  if (!xmllint) return;
  try {
    execFileSync("xmllint", ["--noout", "-"], { input: xml, stdio: ["pipe", "ignore", "pipe"] });
  } catch (e) {
    fail(`${label}: ${String(e.stderr || e.message).split("\n")[0]}`);
  }
};
for (const t of TEMPLATES) {
  const m = compose(t, { paper: "a4", business: { org: "Ampersand & <Angle>" } }, "word");
  const bytes = docxBytes(m, () => STUB_PNG);
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) fail(`${t.id}: .docx is not a zip`);
  wellFormed(documentXml(m, { add: () => ({ id: "rIdTest", name: "media/test.png" }), nextDrawingId: 1 }), `${t.id} document.xml`);
}
ok(`${TEMPLATES.length} .docx files built`);

// 4. Generated files --------------------------------------------------------
console.log("Generated files");
{
  const manifest = path.join(ROOT, "docs/word/manifest.xml");
  const sidebar = path.join(ROOT, "docs/google-docs/Sidebar.html");
  if (!fs.existsSync(manifest) || !fs.existsSync(sidebar)) {
    fail("run npm run build first");
  } else {
    wellFormed(fs.readFileSync(manifest, "utf8"), "manifest.xml");
    const html = fs.readFileSync(sidebar, "utf8");
    const js = html.slice(html.indexOf("<script>") + 8, html.lastIndexOf("</script>"));
    try {
      new vm.Script(js);
      ok("manifest and sidebar bundle");
    } catch (e) {
      fail("sidebar bundle does not parse: " + e.message);
    }
  }
}

// 5. Google Docs, simulated -------------------------------------------------
console.log("Google Docs (simulated)");
const CODE_GS = path.join(ROOT, "google-docs/Code.gs");
const checkBody = (body, label) => invariants(body).forEach((m) => fail(`${label}: ${m}`));

function withImages(blocks) {
  const copy = JSON.parse(JSON.stringify(blocks));
  walkBlocks(copy, (b) => b.t === "image" && (b.image = STUB_PNG));
  return copy;
}

for (const t of TEMPLATES) {
  const { context, body } = makeGoogle(CODE_GS);
  const m = compose(t, { paper: "letter", colorScheme: "lagoon", fontScheme: "modern" }, "google");
  try {
    context.ozCreate({ ...m, blocks: withImages(m.blocks) });
    checkBody(body, t.id);
    const findings = context.ozCheck(m.samples, [{ key: "phone", label: "Phone", sample: "(555) 123-4567" }]);
    if (m.samples.length && !findings.some((f) => f.kind === "sample")) fail(`${t.id}: checker found no sample text`);
    for (const f of findings.slice(0, 3)) context.ozSelect(f);
    const pub = context.ozGetPublication();
    if (!pub || pub.templateId !== t.id) fail(`${t.id}: publication record missing`);
    context.ozRestyle({ colors: { [m.colors.primary]: "#123456" }, fonts: { [m.fonts.heading.name]: "Lora" } }, { colorScheme: "harbor", fontScheme: "bold" });
  } catch (e) {
    fail(`${t.id}: ${e.message}`);
  }
}
{
  const { context, body } = makeGoogle(CODE_GS);
  for (const p of PARTS) {
    try {
      context.ozInsert(withImages(composePart(p, {}, "google", 468).blocks), 468);
    } catch (e) {
      fail(`part ${p.id}: ${e.message}`);
    }
  }
  checkBody(body, "parts");
}
ok(`Code.gs built ${TEMPLATES.length} templates and ${PARTS.length} parts`);

// 6. Area tests ----------------------------------------------------------------
await runAreaTests();

console.log(failures ? `\n${failures} problem${failures > 1 ? "s" : ""} found.` : "\nAll checks passed.");
process.exit(failures ? 1 : 0);
