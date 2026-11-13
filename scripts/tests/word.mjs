/**
 * Word output tests: every template becomes a valid .docx that uses Word's own features.
 */
import { execFileSync } from "node:child_process";

/** Reads the stored (uncompressed) zip files Offizin writes. */
function unzip(bytes) {
  const files = {};
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let pos = 0;
  while (pos + 30 <= bytes.length && dv.getUint32(pos, true) === 0x04034b50) {
    const size = dv.getUint32(pos + 18, true);
    const nameLen = dv.getUint16(pos + 26, true);
    const extra = dv.getUint16(pos + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(pos + 30, pos + 30 + nameLen));
    const start = pos + 30 + nameLen + extra;
    files[name] = bytes.subarray(start, start + size);
    pos = start + size;
  }
  return files;
}

/** The CT_Settings element order from ECMA-376 Part 1, for the elements Offizin writes. */
const SETTINGS_ORDER = ["view", "zoom", "embedTrueTypeFonts", "saveSubsetFonts", "mirrorMargins", "defaultTabStop", "bookFoldPrinting", "characterSpacingControl", "printTwoOnOne", "compat"];

export async function run({ fail, ok, load, STUB_PNG }) {
  const { TEMPLATES, PARTS } = await load("docs/app/core/templates.js");
  const { compose, composePart, walkBlocks } = await load("docs/app/core/compose.js");
  const { docxBytes, flatPackage, settingsXml, publicationProperties, publicationFromProperties, PROPERTY_LIMIT } = await load("docs/app/render/ooxml.js");
  const { escapeXml } = await load("docs/app/core/util.js");

  let xmllint = true;
  try {
    execFileSync("xmllint", ["--version"], { stdio: "ignore" });
  } catch (e) {
    xmllint = false;
  }
  const wellFormed = (xml, label) => {
    if (!xmllint) return;
    try {
      execFileSync("xmllint", ["--noout", "-"], { input: xml, stdio: ["pipe", "ignore", "pipe"] });
    } catch (e) {
      fail(`${label}: ${String(e.stderr || e.message).split("\n")[0]}`);
    }
  };
  const images = () => STUB_PNG;
  const decode = (b) => new TextDecoder().decode(b);

  // 1. Every template on every paper: valid parts and no lost text.
  let docs = 0;
  for (const t of TEMPLATES) {
    for (const paper of ["letter", "a4"]) {
      const m = compose(t, { paper, colorScheme: "cobalt", fontScheme: "classic", business: { org: "Ampersand & <Angle>" } }, "word");
      const files = unzip(docxBytes(m, images));
      docs++;
      for (const [name, data] of Object.entries(files)) if (name.endsWith(".xml") || name.endsWith(".rels")) wellFormed(decode(data), `${t.id}/${paper} ${name}`);
      const doc = decode(files["word/document.xml"] || new Uint8Array());
      const text = (doc.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join("\u0000");
      const missing = [];
      const blocks = m.pages ? m.pages.flat() : m.blocks;
      walkBlocks(blocks, (b) => {
        if (b.t !== "p") return;
        for (const r of b.runs) for (const line of r.text.split("\n")) for (const seg of line.split("\t")) if (seg && !text.includes(escapeXml(seg))) missing.push(seg);
      });
      if (missing.length) fail(`${t.id}/${paper}: ${missing.length} text pieces missing from document.xml, for example "${missing[0]}"`);
      const usesBullets = (() => {
        let found = false;
        walkBlocks(blocks, (b) => b.t === "p" && b.bullet && (found = true));
        return found;
      })();
      if (usesBullets && !files["word/numbering.xml"]) fail(`${t.id}: bullets without a numbering part`);
      if (usesBullets && !/<w:numPr>/.test(doc)) fail(`${t.id}: bullets do not use native numbering`);
      if (m.footer && !files["word/footer1.xml"]) fail(`${t.id}: footer missing`);
      if (m.footer && /\{page\}/.test(m.footer.runs.map((r) => r.text).join("")) && !/w:instr=" PAGE "/.test(decode(files["word/footer1.xml"]))) fail(`${t.id}: page number field missing`);
      const settings = decode(files["word/settings.xml"]);
      const order = [...settings.matchAll(/<w:(\w+)[\s/>]/g)].map((x) => x[1]).filter((n) => SETTINGS_ORDER.includes(n));
      const sorted = [...order].sort((a, b) => SETTINGS_ORDER.indexOf(a) - SETTINGS_ORDER.indexOf(b));
      if (order.join() !== sorted.join()) fail(`${t.id}: settings elements out of schema order: ${order.join(", ")}`);
      if (m.blocks.some((b) => b.t === "columns") && !/<w:cols w:num="\d"/.test(doc)) fail(`${t.id}: flowing columns are not native section columns`);
      if (/<w:pStyle w:val="Heading1"\/>/.test(doc) === false && blocks.some((b) => b.t === "p" && (b.style === "h1" || b.style === "headline"))) fail(`${t.id}: headings do not use Heading 1`);
    }
  }
  ok(`${docs} Word documents valid, complete, and using native styles, bullets, columns, and footers`);

  // 2. Styles: built-in names so Word recognizes them.
  {
    const m = compose(TEMPLATES[0], {}, "word");
    const styles = decode(unzip(docxBytes(m, images))["word/styles.xml"]);
    for (const [id, name] of [["Normal", "Normal"], ["Title", "Title"], ["Heading1", "heading 1"], ["Heading2", "heading 2"], ["Quote", "Quote"], ["Caption", "caption"]]) {
      if (!styles.includes(`w:styleId="${id}"><w:name w:val="${name}"/>`)) fail(`styles.xml lacks the built-in ${name} style`);
    }
    if (!/w:styleId="Heading1">[\s\S]*?<w:outlineLvl w:val="0"\/>/.test(styles)) fail("Heading 1 has no outline level");
    ok("native styles use Word's built-in names and outline levels");
  }

  // 3. Booklets: book fold uses the sheet; page by page uses the page.
  const booklet = TEMPLATES.find((t) => t.pagesOption || (t.id && t.id.includes("booklet")));
  if (booklet) {
    const m = compose(booklet, { paper: "letter" }, "word");
    if (!m.booklet) fail(`${booklet.id}: no booklet data in the Word model`);
    else {
      const fold = unzip(docxBytes(m, images, { bookFold: true }));
      const flat = unzip(docxBytes(m, images, { bookFold: false }));
      const w = (f) => Number((decode(f["word/document.xml"]).match(/<w:pgSz w:w="(\d+)"/) || [])[1]);
      if (!/<w:bookFoldPrinting\/>/.test(decode(fold["word/settings.xml"]))) fail("book fold setting missing");
      if (/<w:bookFoldPrinting\/>/.test(decode(flat["word/settings.xml"]))) fail("page by page booklet should not use book fold");
      if (w(fold) !== Math.round(m.booklet.sheet.width * 20)) fail("book fold page size should be the whole sheet");
      if (w(flat) !== Math.round(m.page.width * 20)) fail("page by page booklet should use the page size");
      ok("booklets: book fold on the whole sheet, or page by page");
    }
  } else {
    ok("no booklet template yet; booklet checks skipped");
  }

  // 4. Print settings.
  {
    const s = settingsXml({ mirrorMargins: true, twoPerSheet: true, embedFonts: true }, false);
    for (const tag of ["<w:mirrorMargins/>", "<w:printTwoOnOne/>", "<w:embedTrueTypeFonts/>", '<w:view w:val="print"/>']) if (!s.includes(tag)) fail(`settings missing ${tag}`);
    wellFormed(s, "settings.xml");
    ok("print settings are written in schema order");
  }

  // 5. Parts: full formatting, a hidden tagged content control, and no numbering part.
  for (const part of PARTS) {
    const c = composePart(part, {}, "word", 468);
    const pkg = flatPackage(c.blocks, 468, images, { tag: "offizin:op:test", alias: "Offizin: " + part.name });
    wellFormed(pkg, `part ${part.id}`);
    if (!pkg.includes('<w:tag w:val="offizin:op:test"/>')) fail(`part ${part.id}: no content control tag`);
    if (!/<w:sdtPr>[^]*?<w15:appearance w15:val="hidden"\/><\/w:sdtPr>/.test(pkg)) fail(`part ${part.id}: the content control is not hidden`);
    if (!pkg.includes('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')) fail(`part ${part.id}: the w15 namespace is not declared`);
    if (pkg.includes("<w:numPr>")) fail(`part ${part.id}: parts should keep tab bullets`);
  }
  ok(`${PARTS.length} page parts wrap in hidden tagged content controls`);

  // 6. The publication record travels in the file as custom document properties, so a document
  // opened from the file can still be restyled and checked.
  {
    const theme = { colorScheme: "custom", customColors: { accent: "#AA0000", ink: "#101010", paper: "#FFFFF0", light: "#EEEEEE" }, fontScheme: "custom", customFonts: { heading: "Georgia & Friends", body: "Arial" } };
    const m = compose(TEMPLATES[0], { paper: "letter", ...theme }, "word");
    const record = { templateId: m.templateId, theme };
    const files = unzip(docxBytes(m, images, { record }));
    const custom = decode(files["docProps/custom.xml"] || new Uint8Array());
    if (!custom) fail("custom.xml missing");
    wellFormed(custom, "custom.xml");
    const types = decode(files["[Content_Types].xml"]);
    if (!types.includes('PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"')) fail("custom.xml has no content type override");
    if (!/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/custom-properties" Target="docProps\/custom\.xml"/.test(decode(files["_rels/.rels"]))) fail("custom.xml has no package relationship");
    const props = publicationProperties(record);
    if (props.some((p) => p.value.length > PROPERTY_LIMIT)) fail("a custom property is longer than Word keeps");
    const pids = [...custom.matchAll(/pid="(\d+)"/g)].map((x) => Number(x[1]));
    if (pids[0] !== 2 || pids.some((p, i) => p !== i + 2)) fail(`custom property pids must start at 2 and run in order (${pids.join(",")})`);
    const back = publicationFromProperties(props.map((p) => ({ key: p.name.toUpperCase(), value: p.value })));
    if (!back || back.templateId !== m.templateId || JSON.stringify(back.theme) !== JSON.stringify(theme)) fail(`the record did not survive the round trip: ${JSON.stringify(back)}`);
    // A long theme is split across several properties and still comes back whole.
    const long = { ...theme, customFonts: { heading: "H".repeat(300), body: "B".repeat(300) } };
    const longProps = publicationProperties({ templateId: "x", theme: long });
    if (longProps.filter((p) => /^OffizinTheme\d+$/.test(p.name)).length < 3) fail("a long theme should be split into several properties");
    if (JSON.stringify(publicationFromProperties(longProps).theme) !== JSON.stringify(long)) fail("a split theme did not come back whole");
    if (publicationFromProperties([]) !== null || publicationFromProperties([{ key: "Other", value: "x" }]) !== null) fail("documents without the record should give null");
    if (unzip(docxBytes(m, images))["docProps/custom.xml"]) fail("a .docx without a record should have no custom.xml");
    ok("the publication record travels in docProps/custom.xml, in properties of at most 255 characters");
  }
}
