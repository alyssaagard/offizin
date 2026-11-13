/**
 * Native styles and page limits:
 *  1. The plain resumes and CVs put their name and heading look into the Title, Heading 1, and
 *     Heading 2 styles, so those paragraphs carry no direct formatting in Word, and changing a
 *     style in Word or Google Docs changes every matching line.
 *  2. Customize margins can never make a tri-fold panel narrower than one inch.
 *  3. A layout that cannot fit stops with a plain message before anything reaches the document,
 *     and no picture is ever zero or negative in size.
 *
 * Run with: node scripts/check.mjs --only=styles
 */

export async function run({ fail, ok, load }) {
  const expect = (cond, msg) => (cond ? ok(msg) : fail(msg));
  const { compose, resolvePage } = await load("docs/app/core/compose.js");
  const { TEMPLATES } = await load("docs/app/core/templates.js");
  const { documentXml } = await load("docs/app/render/ooxml.js");
  const media = { add: () => ({ id: "rIdTest", name: "media/test.png" }), nextDrawingId: 1 };

  // 1. Resume headings follow their styles ---------------------------------------------
  {
    const plain = TEMPLATES.filter((t) => t.plain);
    const problems = [];
    for (const t of plain) {
      for (const headingColor of ["dark", "primary", "accent"]) {
        for (const layout of [undefined, { textScale: 0.9, spaceScale: 1.4, lineScale: 1.2 }]) {
          const m = compose(t, { paper: "letter", colorScheme: "harbor" }, "word", { headingColor, layout });
          const paras = documentXml(m, media).match(/<w:p>.*?<\/w:p>/g) || [];
          for (const id of ["Title", "Heading1", "Heading2"]) {
            const mine = paras.filter((x) => x.includes(`<w:pStyle w:val="${id}"/>`));
            const direct = mine.filter((x) => /<w:sz |<w:color |<w:spacing w:before/.test(x));
            if (direct.length) problems.push(`${t.id} (${headingColor}${layout ? ", custom spacing" : ""}): ${direct.length} ${id} paragraphs carry their own size, color, or spacing`);
          }
          if (!paras.some((x) => x.includes('<w:pStyle w:val="Heading1"/>'))) problems.push(`${t.id}: no Heading 1 paragraphs`);
        }
      }
    }
    problems.slice(0, 8).forEach((x) => fail(`styles: ${x}`));
    if (!problems.length) ok(`styles: the ${plain.length} plain resumes and CVs take their name and heading look from Word's own styles`);
  }

  // 2. Tri-fold margins ---------------------------------------------------------------------
  {
    const t = TEMPLATES.find((x) => x.id === "trifold");
    const problems = [];
    for (const paper of ["letter", "a4"]) {
      for (const side of [72, 200, 300, 400, 10000]) {
        const page = resolvePage(t, paper, { margins: { top: 36, right: side, bottom: 36, left: side } });
        const outer = page.width / 3 - Math.max(page.margins.left, page.margins.right);
        if (outer < 71.5) problems.push(`${paper}, ${side} pt sides: an outer panel is ${Math.round(outer)} pt wide`);
        try {
          compose(t, { paper }, "word", { margins: { top: 36, right: side, bottom: 36, left: side } });
        } catch (e) {
          problems.push(`${paper}, ${side} pt sides: ${e.message}`);
        }
      }
    }
    problems.forEach((x) => fail(`styles: tri-fold ${x}`));
    if (!problems.length) ok("styles: Customize margins keep every tri-fold panel at least one inch wide");
  }

  // 3. Safe sizes -----------------------------------------------------------------------------
  {
    const booklet = TEMPLATES.find((x) => x.id === "booklet");
    let smallest = Infinity;
    let threw = null;
    try {
      for (const pages of [4, 8, 16]) {
        const m = compose(booklet, { paper: "letter" }, "word", { pages, margins: { top: 10000, right: 36, bottom: 10000, left: 36 } });
        const walk = (blocks) =>
          blocks.forEach((b) => {
            if (b.t === "image") smallest = Math.min(smallest, b.width, b.height);
            if (b.t === "grid") b.rows.forEach((r) => r.cells.forEach((c) => walk(c.blocks)));
            if (b.t === "columns") walk(b.blocks);
          });
        walk(m.blocks);
        (m.pages || []).forEach((pg) => walk(pg.blocks || []));
      }
    } catch (e) {
      threw = e.message;
    }
    expect(threw === null ? smallest >= 1 : /does not fit the page/.test(threw), "styles: deep margins never give a picture a zero or negative size");
  }
}
