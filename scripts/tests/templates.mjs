/**
 * Template tests for Offizin 2.1. Run with: node scripts/check.mjs --only=templates
 *
 *  1. Every template composes for Word, Google Docs, and the web, on Letter and A4, with and
 *     without a logo, with no NaN, undefined, or unfilled {fields}.
 *  2. No grid is wider than the page or the cell that holds it, and sheets of cards fit the page.
 *  3. Plain resumes and CVs use only paragraphs, keep spacing small, and keep contact details
 *     out of the footer.
 *  4. Mail merge templates print the right number of sheets for 1, 6, 7, and 31 records.
 *  5. Booklets have a multiple of four pages and printer spreads that use every page once.
 *  6. Fold guides, print settings, Customize settings, and options are complete.
 *  7. The design checker can find every sample.
 *  8. Printed text and placeholders read right: no name twice, sponsor slots stay placeholders,
 *     invoice terms, detail samples, page guides, deep margins, badge notes, and booklet filler.
 */
import fs from "node:fs";
import path from "node:path";

const PLATFORMS = ["word", "google", "web"];
const PAPERS = ["letter", "a4"];
const LOGO = { src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==", width: 300, height: 100 };

/** Sheets (or pages) per group of records, for each mail merge template. */
const PER_SHEET = {
  labels: (paper) => (paper === "a4" ? 21 : 30),
  certificate: () => 1,
  "name-badges": () => 6,
  tickets: () => 10,
  "place-cards": () => 4,
  envelope: () => 1,
};

/** Finds NaN, Infinity, undefined values, and leftover "undefined" or "NaN" text anywhere in a model. */
function badValues(value, where = "model", out = []) {
  if (out.length > 5) return out;
  if (value === undefined) out.push(`${where} is undefined`);
  else if (typeof value === "number" && !Number.isFinite(value)) out.push(`${where} is ${value}`);
  else if (typeof value === "string" && /\bundefined\b|\bNaN\b/.test(value)) out.push(`${where} contains "${value.slice(0, 40)}"`);
  else if (Array.isArray(value)) value.forEach((v, i) => badValues(v, `${where}[${i}]`, out));
  else if (value && typeof value === "object") for (const [k, v] of Object.entries(value)) badValues(v, `${where}.${k}`, out);
  return out;
}

/** Checks grid widths against their container, all the way down. Returns problems. */
function widthProblems(blocks, width, where, out = []) {
  for (const b of blocks) {
    if (b.t === "grid") {
      if (b.width > width + 0.6) out.push(`${where}: grid ${b.width}pt wider than its ${round(width)}pt container`);
      const sum = b.cols.reduce((a, c) => a + c, 0);
      if (Math.abs(sum - b.width) > 0.6) out.push(`${where}: grid columns add up to ${round(sum)}pt, not ${b.width}pt`);
      for (const r of b.rows) {
        if (r.cells.length !== b.cols.length) out.push(`${where}: a row has ${r.cells.length} cells for ${b.cols.length} columns`);
        r.cells.forEach((c, i) => widthProblems(c.blocks, Math.max(12, b.cols[i] - c.pad[1] - c.pad[3]), where, out));
      }
    } else if (b.t === "image" || b.t === "rule") {
      if (b.width > width + 0.6) out.push(`${where}: ${b.t} ${b.width}pt wider than its ${round(width)}pt container`);
    } else if (b.t === "columns") {
      if (b.width > width + 0.6) out.push(`${where}: columns ${b.width}pt wider than ${round(width)}pt`);
      widthProblems(b.blocks, b.colWidth, where, out);
    }
  }
  return out;
}

const round = (n) => Math.round(n * 100) / 100;

/** Every paragraph's text, including paragraphs inside grids. */
function allText(blocks, walkBlocks) {
  const out = [];
  walkBlocks(blocks, (b) => b.t === "p" && out.push(b.runs.map((r) => r.text).join("")));
  return out;
}

/** Splits resolved blocks into printed pages at each page break. */
function splitPages(blocks) {
  const pages = [[]];
  for (const b of blocks) b.t === "break" ? pages.push([]) : pages[pages.length - 1].push(b);
  return pages;
}

function recordsFor(template, n) {
  return Array.from({ length: n }, (_, i) => {
    const r = {};
    for (const f of template.merge.fields) r[f.key] = `${f.label} ${i + 1}`;
    return r;
  });
}

export async function run({ ROOT, fail, ok, load }) {
  const { TEMPLATES, CATEGORIES, FOLD_GUIDES } = await load("docs/app/core/templates.js");
  const { MORE_TEMPLATES, MORE_FOLD_GUIDES } = await load("docs/app/core/templates/more.js");
  const { compose, walkBlocks, estimateHeight, customizeFor, resolvePage } = await load("docs/app/core/compose.js");
  const { PAPER } = await load("docs/app/core/util.js");
  const { parseRecords } = await load("docs/app/core/merge.js");

  // 1 and 2. Composition, values, and widths ----------------------------------------
  {
    let count = 0;
    let problems = 0;
    const report = (msg) => {
      problems++;
      if (problems <= 40) fail(msg);
    };
    for (const t of TEMPLATES) {
      for (const platform of PLATFORMS) {
        for (const paper of PAPERS) {
          for (const logo of [null, LOGO]) {
            const theme = { paper, colorScheme: "plum", fontScheme: "elegant", business: { org: "Test & Co. <Ltd>" }, logo };
            const where = `${t.id}/${platform}/${paper}${logo ? "/logo" : ""}`;
            let m;
            try {
              m = compose(t, theme, platform);
            } catch (e) {
              report(`${where}: compose threw ${e.message}`);
              continue;
            }
            count++;
            for (const msg of badValues({ page: m.page, blocks: m.blocks, pages: m.pages, booklet: m.booklet, footer: m.footer, print: m.print, folds: m.folds })) report(`${where}: ${msg}`);
            const W = m.page.width - m.page.margins.left - m.page.margins.right;
            const H = m.page.height - m.page.margins.top - m.page.margins.bottom;
            if (!(W > 72 && H > 72)) report(`${where}: usable area ${W} x ${H} is too small`);
            for (const msg of widthProblems(m.blocks, W, where)) report(msg);
            if (m.pages) {
              const pageW = resolvePage(t, paper, {}).width - m.page.margins.left - m.page.margins.right;
              m.pages.forEach((pg, i) => widthProblems(pg, pageW, `${where} page ${i + 1}`).forEach(report));
            }
            // Fields left unfilled, such as "{nmae}". The footer may keep {page} and {pages}.
            for (const text of allText(m.blocks, walkBlocks)) {
              const left = text.match(/\{\w+\}/);
              if (left) report(`${where}: unfilled field ${left[0]}`);
            }
            if (m.footer) {
              const left = m.footer.runs.map((r) => r.text).join("").replace(/\{pages?\}/g, "").match(/\{\w+\}/);
              if (left) report(`${where}: unfilled field ${left[0]} in the footer`);
            }
            // A sheet of fixed-height rows (cards, badges, labels) must fit on one page.
            for (const b of m.blocks) {
              if (b.t !== "grid" || !b.rows.length || !b.rows.every((r) => r.exact && r.height)) continue;
              const total = b.rows.reduce((s, r) => s + r.height, 0);
              if (total > H + 0.6) report(`${where}: a sheet of fixed rows is ${round(total)}pt tall, more than the page's ${round(H)}pt`);
            }
            for (const b of m.blocks) if (b.t === "grid") for (const r of b.rows) if (r.height && r.height > H + 0.6) report(`${where}: row taller than the page (${r.height} > ${round(H)})`);
          }
        }
      }
    }
    if (problems > 40) fail(`${problems - 40} more composition problems not shown`);
    if (!problems) ok(`${TEMPLATES.length} templates compose ${count} ways with no bad values, and no grid is wider than its page or cell`);
  }

  // Content fits: the new publications should not run past the bottom of a page, and text in
  // fixed-height pieces (badges, tickets, cards) should not be cut off, even with long names.
  // Heights are estimates (estimateHeight in compose.js), so this catches clear overflows only.
  {
    let problems = 0;
    const LONG = {
      name: "Maria Fernanda Gonzalez Rodriguez",
      role: "Director of Community Partnerships",
      organization: "Northfield Regional Arts Council",
      company: "Northfield Regional Arts Council",
      address: "1200 Lakeshore Boulevard",
      address2: "Suite 400",
      city: "Springfield",
      state: "ST",
      postal: "12345",
      country: "United States",
      table: "Table 12",
      seat: "Row C, seat 14",
    };
    // The new publications, plus the core mail merge templates (labels, certificate), which also take long names.
    const fitted = [...MORE_TEMPLATES, ...TEMPLATES.filter((x) => x.merge && !MORE_TEMPLATES.includes(x))];
    for (const t of fitted) {
      const variants = [{}];
      const pages = (t.options || []).find((o) => o.key === "pages");
      if (pages) pages.choices.forEach((c) => variants.push({ pages: c.value }));
      if (t.merge) variants.push({ records: [LONG, LONG, LONG] });
      // Every paper size and orientation the Customize panel offers.
      const allow = customizeFor(t);
      for (const paperSize of allow.paper || []) for (const orientation of allow.orientation ? ["portrait", "landscape"] : [null]) variants.push({ paperSize, orientation });
      for (const paper of PAPERS) {
        for (const logo of [null, LOGO]) {
          for (const options of variants) {
            const m = compose(t, { paper, logo }, "web", options);
            const where = `${t.id}/${paper}${logo ? "/logo" : ""}${options.pages ? "/" + options.pages : ""}${options.records ? "/long names" : ""}${options.paperSize ? "/" + options.paperSize : ""}${options.orientation ? "/" + options.orientation : ""}`;
            const W = m.page.width - m.page.margins.left - m.page.margins.right;
            const H = m.page.height - m.page.margins.top - m.page.margins.bottom;
            (m.pages || splitPages(m.blocks)).forEach((pg, i) => {
              const h = pg.reduce((s, b) => s + estimateHeight(b, W), 0);
              if (h > H + 1) {
                problems++;
                fail(`${where} page ${i + 1}: content is about ${Math.round(h)}pt tall, more than the ${Math.round(H)}pt page`);
              }
            });
            walkBlocks(m.blocks, (b) => {
              if (b.t !== "grid") return;
              for (const r of b.rows) {
                if (!r.exact || !r.height) continue;
                r.cells.forEach((c, i) => {
                  const h = c.pad[0] + c.pad[2] + c.blocks.reduce((s, x) => s + estimateHeight(x, Math.max(12, b.cols[i] - c.pad[1] - c.pad[3])), 0);
                  if (h > r.height + 1) {
                    problems++;
                    fail(`${where}: a fixed-height cell needs about ${Math.round(h)}pt but has ${r.height}pt, so text would be cut off`);
                  }
                });
              }
            });
          }
        }
      }
    }
    if (!problems) ok(`${fitted.length} publications fit their pages and pieces on Letter and A4, with every booklet length, Customize paper size, and long names`);
  }

  // 3. Plain resumes and CVs ---------------------------------------------------------
  {
    const plain = TEMPLATES.filter((t) => t.plain === true);
    if (plain.length < 6) fail(`expected at least 6 plain resumes and CVs, found ${plain.length}`);
    const contactKeys = ["phone", "email", "address", "city", "web", "profile", "social"];
    let problems = 0;
    for (const t of plain) {
      if (t.category !== "resumes") fail(`${t.id}: plain templates belong in the resumes category`);
      const spec = typeof t.footer === "function" ? null : t.footer;
      if (spec && contactKeys.some((k) => spec.text.includes(`{${k}}`))) {
        problems++;
        fail(`${t.id}: the footer contains contact details`);
      }
      for (const platform of PLATFORMS) {
        for (const paper of PAPERS) {
          for (const headingColor of ["dark", "accent"]) {
            const business = { phone: "(555) 010-9999", email: "pat@example.com", city: "Fairview, ST 54321", web: "pat.example.com", profile: "linkedin.com/in/pat", address: "9 Elm Street" };
            const m = compose(t, { paper, business }, platform, { headingColor });
            const where = `${t.id}/${platform}/${paper}`;
            walkBlocks(m.blocks, (b) => {
              if (b.t === "p") return;
              if (b.t === "spacer" && b.height <= 24) return;
              if (b.t === "break") return;
              problems++;
              fail(`${where}: plain templates may use only paragraphs, but found ${b.t}${b.t === "spacer" ? ` of ${b.height}pt` : ""}`);
            });
            if (m.pages) {
              problems++;
              fail(`${where}: a plain template must not be a booklet`);
            }
            if (m.footer) {
              const text = m.footer.runs.map((r) => r.text).join("");
              const leaked = Object.values(business).filter((v) => text.includes(v));
              if (leaked.length) {
                problems++;
                fail(`${where}: the footer shows contact details (${leaked.join(", ")})`);
              }
            }
            const body = allText(m.blocks, walkBlocks).join("\n");
            if (!body.includes(business.email) && !body.includes(business.phone)) {
              problems++;
              fail(`${where}: contact details should appear in the body`);
            }
            if (/\d{4}\s*[\u2013\u2014-]\s*(\d{4}|present)/i.test(body)) {
              problems++;
              fail(`${where}: dates and ranges should be written without dashes`);
            }
          }
        }
      }
    }
    if (!problems) ok(`${plain.length} plain resumes and CVs use only paragraphs and keep contact details in the body`);
  }

  // 4. Mail merge ------------------------------------------------------------------------
  {
    let problems = 0;
    const merges = TEMPLATES.filter((t) => t.merge);
    for (const t of merges) {
      if (!PER_SHEET[t.id]) {
        problems++;
        fail(`${t.id}: add the number of records per sheet to PER_SHEET in this test`);
        continue;
      }
      if (!Array.isArray(t.merge.fields) || !t.merge.fields.length || !t.merge.note) {
        problems++;
        fail(`${t.id}: merge needs fields and a note`);
      }
      for (const paper of PAPERS) {
        const per = PER_SHEET[t.id](paper);
        for (const n of [1, 6, 7, 31]) {
          const records = recordsFor(t, n);
          for (const platform of PLATFORMS) {
            const m = compose(t, { paper }, platform, { records });
            const sheets = splitPages(m.blocks).length;
            const want = Math.ceil(n / per);
            if (sheets !== want) {
              problems++;
              fail(`${t.id}/${platform}/${paper}: ${n} records made ${sheets} sheets, expected ${want}`);
            }
            const text = allText(m.blocks, walkBlocks).join("\n");
            const firstKey = t.merge.fields[0];
            for (const i of [1, n]) {
              if (!text.includes(`${firstKey.label} ${i}`)) {
                problems++;
                fail(`${t.id}/${platform}/${paper}: record ${i} of ${n} is missing from the publication`);
              }
            }
          }
        }
        // With no list, the preview still shows something useful.
        const empty = compose(t, { paper }, "web");
        if (!allText(empty.blocks, walkBlocks).some((s) => s.trim())) {
          problems++;
          fail(`${t.id}/${paper}: the preview is empty without a list`);
        }
      }
      // The panel's parser matches each field by its key, label, and aliases.
      const head = t.merge.fields.map((f) => f.label).join("\t");
      const row = t.merge.fields.map((f) => `x ${f.key}`).join("\t");
      const parsed = parseRecords(`${head}\n${row}`, t.merge.fields);
      if (parsed.records.length !== 1 || t.merge.fields.some((f) => parsed.records[0][f.key] !== `x ${f.key}`)) {
        problems++;
        fail(`${t.id}: a pasted list with the field labels as headings is not read correctly`);
      }
      for (const f of t.merge.fields) {
        for (const alias of f.aliases || []) {
          const other = t.merge.fields.find((g) => g !== f && [g.key, g.label].some((s) => s.toLowerCase().replace(/[^a-z0-9]/g, "") === alias.toLowerCase().replace(/[^a-z0-9]/g, "")));
          if (other) {
            problems++;
            fail(`${t.id}: the alias "${alias}" of ${f.key} is also the name of ${other.key}`);
          }
        }
      }
    }
    // Specific fields the badges need.
    const badges = TEMPLATES.find((t) => t.id === "name-badges");
    if (badges) {
      const parsed = parseRecords("Name\tJob title\tCompany\nAda Lovelace\tAnalyst\tEngines Ltd", badges.merge.fields);
      const r = parsed.records[0] || {};
      if (r.name !== "Ada Lovelace" || r.role !== "Analyst" || r.organization !== "Engines Ltd") {
        problems++;
        fail("name-badges: the headings Name, Job title, and Company are not matched to name, role, and organization");
      }
      if (!badges.merge.sample || badges.merge.sample.length < 1) fail("name-badges: needs sample records for the preview");
    }
    if (!problems) ok(`${merges.length} mail merge templates make the right number of sheets for 1, 6, 7, and 31 records`);
  }

  // Ticket and certificate numbering and option limits.
  {
    let problems = 0;
    const tickets = TEMPLATES.find((t) => t.id === "tickets");
    if (tickets) {
      const m = compose(tickets, { paper: "letter" }, "web", { sheets: 3, startNumber: 41 });
      const text = allText(m.blocks, walkBlocks).join("\n");
      if (splitPages(m.blocks).length !== 3) (problems++, fail("tickets: 3 sheets did not make 3 pages"));
      if (!text.includes("No. 0041") || !text.includes("No. 0070") || text.includes("No. 0071")) (problems++, fail("tickets: numbers should run from No. 0041 to No. 0070"));
      const capped = compose(tickets, { paper: "letter" }, "web", { sheets: 500 });
      if (splitPages(capped.blocks).length !== 20) (problems++, fail("tickets: the number of sheets should be limited to 20"));
    }
    const gifts = TEMPLATES.find((t) => t.id === "gift-certificates");
    if (gifts) {
      const m = compose(gifts, { paper: "a4" }, "word", { amount: "\u20AC75", sheets: 2 });
      const text = allText(m.blocks, walkBlocks).join("\n");
      if (splitPages(m.blocks).length !== 2) (problems++, fail("gift-certificates: 2 sheets did not make 2 pages"));
      if (!text.includes("\u20AC75") || !text.toLowerCase().includes("no. 0006")) (problems++, fail("gift-certificates: the amount or the numbers are missing"));
      const blankAmount = compose(gifts, { paper: "letter" }, "web", { amount: "  " });
      if (!allText(blankAmount.blocks, walkBlocks).join("\n").includes("$50")) (problems++, fail("gift-certificates: a blank amount should fall back to $50"));
    }
    if (!problems) ok("tickets and gift certificates are numbered in order and respect their limits");
  }

  // 5. Booklets ---------------------------------------------------------------------------
  {
    let problems = 0;
    const booklets = TEMPLATES.filter((t) => t.fold === "booklet");
    if (!booklets.length) (problems++, fail("no booklet template found"));
    for (const t of booklets) {
      const choice = (t.options || []).find((o) => o.key === "pages");
      const counts = choice ? choice.choices.map((c) => c.value) : [undefined];
      for (const count of [...counts, "5"]) {
        for (const paper of PAPERS) {
          for (const platform of PLATFORMS) {
            const where = `${t.id}/${count || "default"}/${platform}/${paper}`;
            const m = compose(t, { paper }, platform, count ? { pages: count } : {});
            if (!m.pages || !m.booklet) {
              problems++;
              fail(`${where}: a booklet must build pages`);
              continue;
            }
            const n = m.pages.length;
            const want = count === "5" ? 8 : Number(count) || n;
            if (n % 4 !== 0) (problems++, fail(`${where}: ${n} pages is not a multiple of four`));
            if (n !== want) (problems++, fail(`${where}: made ${n} pages, expected ${want}`));
            if (m.booklet.pages !== n) (problems++, fail(`${where}: booklet.pages is ${m.booklet.pages}, not ${n}`));
            if (m.pages.some((pg) => !pg.length)) (problems++, fail(`${where}: a page is empty`));
            const seen = new Array(n).fill(0);
            for (const s of m.booklet.spreads) {
              if (s.length !== 2) (problems++, fail(`${where}: a spread does not have two pages`));
              for (const i of s) if (Number.isInteger(i) && i >= 0 && i < n) seen[i]++;
              if (s[0] + s[1] !== n - 1) (problems++, fail(`${where}: spread ${s.join(" and ")} would not face correctly when folded`));
              // Odd page numbers (even indexes) are right-hand pages.
              if (s[0] % 2 !== 1 || s[1] % 2 !== 0) (problems++, fail(`${where}: spread ${s.join(" and ")} puts a left-hand page on the right`));
            }
            if (seen.some((c) => c !== 1)) (problems++, fail(`${where}: the spreads do not use every page exactly once`));
            if (m.booklet.spreads.length !== n / 2) (problems++, fail(`${where}: expected ${n / 2} spreads`));
            const [first, second] = m.booklet.spreads;
            if (!first || first[0] !== n - 1 || first[1] !== 0 || !second || second[0] !== 1 || second[1] !== n - 2) (problems++, fail(`${where}: the first sheet should carry the back and front covers, then pages 2 and ${n - 1}`));
            const sheet = m.booklet.sheet;
            const single = resolvePage(t, paper, {});
            // The sheet is two pages wide, snapped to real paper: two A5 pages (296 mm) print on A4 (297 mm).
            const sheetPaper = paper === "a4" ? PAPER.a4 : PAPER.letter;
            if (Math.abs(sheet.width - single.width * 2) > 4.5 || Math.abs(sheet.height - single.height) > 0.6) (problems++, fail(`${where}: the sheet is not two pages wide`));
            if (Math.abs(sheet.width - sheetPaper.height) > 0.6 || Math.abs(sheet.height - sheetPaper.width) > 0.6) (problems++, fail(`${where}: the sheet should be ${paper === "a4" ? "A4" : "Letter"} landscape`));
            const expectSize = paper === "a4" ? PAPER.a5 : PAPER.halfLetter;
            if (Math.abs(single.width - expectSize.width) > 0.6 || Math.abs(single.height - expectSize.height) > 0.6) (problems++, fail(`${where}: pages should be ${paper === "a4" ? "A5" : "half letter"}`));
            if (platform === "google") {
              const grids = m.blocks.filter((b) => b.t === "grid").length;
              if (grids !== n / 2) (problems++, fail(`${where}: Google Docs should get ${n / 2} printer spreads, got ${grids}`));
              if (m.print.duplex !== "short") (problems++, fail(`${where}: Google spreads print on both sides, flipping on the short edge`));
            } else if (splitPages(m.blocks).length !== n) {
              problems++;
              fail(`${where}: expected ${n} pages in reading order`);
            }
            if (platform === "word" && !m.print.bookFold) (problems++, fail(`${where}: Word should print with book fold`));
            if (m.print.duplex !== "short") (problems++, fail(`${where}: booklets print on both sides, flipping on the short edge`));
          }
        }
      }
      // The contents page lists real page numbers.
      const m = compose(t, { paper: "letter" }, "web", { pages: "16" });
      const contentsIndex = m.pages.findIndex((pg) => allText(pg, walkBlocks).some((s) => s === "Contents"));
      if (contentsIndex < 0) (problems++, fail(`${t.id}: the 16 page booklet has no contents page`));
      else {
        const text = allText(m.pages[contentsIndex], walkBlocks);
        const nums = text.filter((s) => /^\d+$/.test(s)).map(Number);
        if (!nums.length || nums.some((x) => x < 1 || x > 16) || nums.some((x, i) => i && x <= nums[i - 1])) (problems++, fail(`${t.id}: contents page numbers are missing or out of order (${nums.join(", ")})`));
      }
    }
    if (!problems) ok(`${booklets.length} booklet template makes 4, 8, 12, and 16 pages with spreads that use every page once`);
  }

  // 6. Fold guides, settings, and options --------------------------------------------------
  {
    let problems = 0;
    const sentence = (s) => typeof s === "string" && /^[A-Z\u201C"(]/.test(s) && /[.)]$/.test(s.trim());
    for (const key of ["booklet", "tent", "envelope"]) {
      const g = MORE_FOLD_GUIDES[key];
      if (!g) {
        problems++;
        fail(`MORE_FOLD_GUIDES is missing ${key}`);
        continue;
      }
      if (typeof g.title !== "string" || !g.title) (problems++, fail(`${key} guide needs a title`));
      if (!Array.isArray(g.steps) || g.steps.length < 2 || !g.steps.every(sentence)) (problems++, fail(`${key} guide needs at least two steps written as complete sentences`));
      if (!sentence(g.tip)) (problems++, fail(`${key} guide needs a tip written as a complete sentence`));
      if (FOLD_GUIDES[key] !== g) (problems++, fail(`${key} guide is not in FOLD_GUIDES`));
    }
    const ids = new Set();
    const categories = new Set(CATEGORIES.map((c) => c.id));
    for (const t of TEMPLATES) {
      if (ids.has(t.id)) (problems++, fail(`duplicate template id ${t.id}`));
      ids.add(t.id);
      if (!categories.has(t.category)) (problems++, fail(`${t.id}: unknown category ${t.category}`));
      if (!FOLD_GUIDES[t.fold]) (problems++, fail(`${t.id}: no fold guide for "${t.fold}"`));
      for (const o of t.options || []) {
        if (!o.key || !o.label || !["month", "number", "select", "boolean", "text"].includes(o.type)) (problems++, fail(`${t.id}: option ${o.key} is incomplete`));
        const d = typeof o.default === "function" ? o.default() : o.default;
        if (o.type === "select" && !(o.choices || []).some((c) => c.value === d)) (problems++, fail(`${t.id}: the default of ${o.key} is not one of its choices`));
        if (o.type === "number" && (!Number.isFinite(Number(d)) || (o.min != null && d < o.min) || (o.max != null && d > o.max))) (problems++, fail(`${t.id}: the default of ${o.key} is out of range`));
      }
    }
    const wanted = { booklet: "programs", "name-badges": "labels", tickets: "labels", "gift-certificates": "calendars", bookmarks: "cards", "place-cards": "cards", envelope: "stationery", invoice: "stationery", "photo-page": "flyers" };
    const folds = { booklet: "booklet", "name-badges": "cards", tickets: "cards", "gift-certificates": "cards", bookmarks: "cards", "place-cards": "tent", envelope: "envelope" };
    for (const [id, category] of Object.entries(wanted)) {
      const t = MORE_TEMPLATES.find((x) => x.id === id);
      if (!t) {
        problems++;
        fail(`more.js is missing ${id}`);
        continue;
      }
      if (t.category !== category) (problems++, fail(`${id}: category should be ${category}`));
      if (folds[id] && t.fold !== folds[id]) (problems++, fail(`${id}: fold should be ${folds[id]}`));
      if (!t.print || !t.customize) (problems++, fail(`${id}: needs print and customize settings`));
      if (!t.name || !t.description || !t.sizeLabel) (problems++, fail(`${id}: needs a name, description, and size label`));
    }
    for (const id of ["name-badges", "envelope"]) {
      const t = MORE_TEMPLATES.find((x) => x.id === id);
      const c = t && customizeFor(t);
      if (c && (c.paper || c.margins)) (problems++, fail(`${id}: paper size and margins must stay fixed`));
    }
    const booklet = MORE_TEMPLATES.find((t) => t.id === "booklet");
    if (booklet && (booklet.print.duplex !== "short" || booklet.print.finishing !== "Fold the stack in half and staple twice on the fold.")) (problems++, fail("booklet: print settings do not match the specification"));
    // Envelope sizes: #10 on Letter and DL on A4, landscape.
    const env = MORE_TEMPLATES.find((t) => t.id === "envelope");
    if (env) {
      const us = resolvePage(env, "letter", {});
      const eu = resolvePage(env, "a4", {});
      if (us.width !== 684 || us.height !== 297 || us.orientation !== "landscape") (problems++, fail("envelope: Letter should give a #10 envelope, 9.5 x 4.125 in"));
      if (Math.abs(eu.width - 623.62) > 0.6 || Math.abs(eu.height - 311.81) > 0.6 || eu.orientation !== "landscape") (problems++, fail("envelope: A4 should give a DL envelope, 220 x 110 mm"));
    }
    // Badges are 4 x 3 inches, two across and three down, on both papers.
    const badges = MORE_TEMPLATES.find((t) => t.id === "name-badges");
    if (badges) {
      for (const paper of PAPERS) {
        const m = compose(badges, { paper }, "web");
        const g = m.blocks[0];
        const heights = [];
        for (let i = 0; i < g.rows.length; i += 2) heights.push(g.rows[i].height + g.rows[i + 1].height);
        if (g.cols.join() !== "288,288" || heights.length !== 3 || heights.some((h) => h !== 216)) (problems++, fail(`name-badges/${paper}: badges should be 4 x 3 in, 2 across and 3 down`));
      }
    }
    if (!problems) ok("fold guides, categories, print and Customize settings, and options are complete");
  }

  // 7. The design checker can find every sample, and finds short placeholders --------------
  {
    let problems = 0;
    const report = (msg) => {
      problems++;
      if (problems <= 20) fail(msg);
    };
    const { ALT_DEUTSCH } = await load("docs/app/core/filler.js");
    let count = 0;
    for (const t of TEMPLATES) {
      for (const platform of PLATFORMS) {
        const m = compose(t, { paper: "letter" }, platform);
        // The editors search the document with matching case, so each sample must appear exactly,
        // capitals and double spaces included, inside one paragraph.
        const texts = allText(m.blocks, walkBlocks);
        for (const s of m.samples) {
          count++;
          if (!texts.some((x) => x.includes(s))) report(`${t.id}/${platform}: the checker would never find the sample "${s}"`);
        }
      }
    }
    const samplesOf = (id, options) => compose(TEMPLATES.find((t) => t.id === id), { paper: "letter" }, "word", options).samples;
    const expect = [
      ["resume-plain", "Job title"],
      ["resume-plain", "Degree"],
      ["resume", "Degree"],
      ["menu", "Dish name"],
      ["event-flyer", "Event name"],
      ["booklet", "Event name"],
      ["sale-flyer", "Spring sale"],
      ["name-badges", "Priya Raman"],
      ["place-cards", "TABLE 1"],
      ["newsletter", "MEMBER SPOTLIGHT"],
      ["resume-plain", "Organization name, City, ST  |  March 2021"],
    ];
    for (const [id, s] of expect) if (!samplesOf(id).some((x) => x === s || x.startsWith(s + " "))) report(`${id}: the checker does not list the sample "${s}"`);
    // Text from a pasted list is the user's own, so it is never listed as a sample.
    const badges = samplesOf("name-badges", { records: [{ name: "Ada Lovelace", role: "Analyst", organization: "Engines Ltd" }] });
    if (badges.some((s) => /Ada Lovelace|Analyst|Engines Ltd/.test(s))) report("name-badges: names from a pasted list are listed as samples");
    // A certificate list without a reason column leaves that line blank instead of printing filler.
    const cert = TEMPLATES.find((t) => t.id === "certificate");
    const merged = allText(compose(cert, { paper: "letter" }, "word", { records: [{ recipient: "Ada Lovelace" }, { recipient: "Ben Ode" }] }).blocks, walkBlocks).join("\n");
    if (ALT_DEUTSCH.some((s) => merged.includes(s))) report("certificate: a list without reasons prints the sample reason");
    if (!/Ada Lovelace/.test(merged) || !/Ben Ode/.test(merged)) report("certificate: merged names are missing");
    const withReason = allText(compose(cert, { paper: "letter" }, "word", { records: [{ recipient: "Ada Lovelace", reason: "Ten years of service" }] }).blocks, walkBlocks).join("\n");
    if (!withReason.includes("Ten years of service")) report("certificate: a reason from the list is missing");
    if (!problems) ok(`the design checker can find all ${count} samples, including short placeholders, capitals, and double spaces`);
  }

  // 8. Printed text and placeholders that the 2.1 review found wrong ---------------------
  {
    let problems = 0;
    const report = (msg) => {
      problems++;
      fail(msg);
    };
    const { ALT_DEUTSCH, ALT_SHORT, ALT_QUOTES } = await load("docs/app/core/filler.js");
    const byId = (id) => TEMPLATES.find((t) => t.id === id);
    const images = (blocks) => {
      const out = [];
      walkBlocks(blocks, (b) => b.t === "image" && out.push(b));
      return out;
    };

    // The menu names the organization once, as the title, and the greeting card's back panel only in
    // "Made with care by", with or without a logo. Without a logo there is no wordmark above either.
    for (const [id, want] of [
      ["menu", 1],
      ["greeting-card", 0],
    ]) {
      for (const logo of [null, LOGO]) {
        const m = compose(byId(id), { paper: "letter", logo }, "word");
        const org = allText(m.blocks, walkBlocks).filter((s) => s === "Riverside Arts Collective").length;
        const logos = images(m.blocks).filter((b) => b.kind === "logo").length;
        if (org !== want) report(`${id}${logo ? "/logo" : ""}: the organization name is printed ${org} times as its own line, expected ${want}`);
        if (logos !== (logo ? 1 : 0)) report(`${id}${logo ? "/logo" : ""}: expected ${logo ? "one logo" : "no logo box"}, found ${logos}`);
      }
    }

    // Sponsor slots are picture placeholders, so the organization's own logo never fills them.
    for (const logo of [null, LOGO]) {
      const m = compose(byId("booklet"), { paper: "letter", logo }, "web", { pages: "16" });
      const slots = m.pages.flatMap((pg) => images(pg)).filter((b) => b.label === "Sponsor logo");
      if (slots.length !== 6) report(`booklet/16${logo ? "/logo" : ""}: expected six sponsor slots, found ${slots.length}`);
      if (slots.some((b) => b.kind !== "photo" || b.src)) report(`booklet/16${logo ? "/logo" : ""}: a sponsor slot carries the organization's own logo`);
    }

    // Invoice terms read correctly for 0, 1, and more days.
    const terms = (due) => allText(compose(byId("invoice"), { paper: "letter" }, "web", { due }).blocks, walkBlocks).find((s) => s.startsWith("Payment is due")) || "";
    if (!terms(0).startsWith("Payment is due on receipt.")) report(`invoice: with 0 days the terms read "${terms(0)}"`);
    if (!terms(1).startsWith("Payment is due within 1 day of the invoice date.")) report(`invoice: with 1 day the terms read "${terms(1)}"`);
    if (!terms(30).startsWith("Payment is due within 30 days of the invoice date.")) report(`invoice: with 30 days the terms read "${terms(30)}"`);

    // Short detail values such as a time are listed as samples, like the date and place beside them.
    const samplesOf = (id, options) => compose(byId(id), { paper: "letter" }, "word", options).samples;
    for (const [id, s] of [
      ["event-flyer", "6 to 9 PM"],
      ["event-flyer", "Saturday, June 14"],
      ["poster", "9 AM to noon"],
    ]) if (!samplesOf(id).includes(s)) report(`${id}: the checker does not list the detail "${s}"`);

    // The 4-page booklet's page guide is generated text, not a sample to replace.
    const nav = samplesOf("booklet", { pages: "4" }).filter((s) => /^(Welcome|Schedule|Contact details), (page \d|back cover)/.test(s));
    if (nav.length) report(`booklet/4: the page guide is listed as sample text: ${nav.join(", ")}`);

    // Deep top and bottom margins never give a picture a zero or negative height.
    for (const paper of PAPERS) {
      const m = compose(byId("booklet"), { paper }, "web", { pages: "16", margins: { top: 9999, bottom: 9999, left: 36, right: 36 } });
      const bad = m.pages.flatMap((pg) => images(pg)).filter((b) => !(b.height >= 1) || !(b.width >= 1));
      if (bad.length) report(`booklet/16/${paper}: with deep top and bottom margins, pictures come out ${bad.map((b) => b.width + "x" + b.height).join(", ")}`);
      // The photo pages size their pictures from what is left of the page, so they keep a floor of 36 pt.
      const photoPages = m.pages.filter((pg) => allText(pg, walkBlocks).some((s) => s === "Highlights from last year" || s === "Moments from the weekend"));
      if (photoPages.length !== 2) report(`booklet/16/${paper}: expected two photo pages, found ${photoPages.length}`);
      const flat = photoPages.flatMap((pg) => images(pg)).filter((b) => b.height < 36);
      if (flat.length) report(`booklet/16/${paper}: photo page pictures come out ${flat.map((b) => b.width + "x" + b.height).join(", ")} with deep margins`);
    }

    // Name badges stay 4 x 3 inches on A4 and the print note says how close to the edge they come.
    const badgeNote = (paper) => compose(byId("name-badges"), { paper }, "web").print.finishing || "";
    if (!/3\.4 mm/.test(badgeNote("a4")) || !/\.$/.test(badgeNote("a4").trim())) report(`name-badges/a4: the print note should say the badges come within 3.4 mm of the side edges, found "${badgeNote("a4")}"`);
    if (/3\.4 mm/.test(badgeNote("letter"))) report("name-badges/letter: the A4 edge note should not appear on Letter");

    // No filler line repeats on one printed side of a booklet.
    const norm = (s) => s.replace(/[“”".!?:,]/g, "").trim().toLowerCase();
    const FILLER = new Set([...ALT_DEUTSCH, ...ALT_SHORT, ...ALT_QUOTES].map(norm));
    const fillerLines = (blocks) =>
      allText(blocks, walkBlocks)
        .flatMap((s) => s.split(/(?<=[.!?])\s+/))
        .map(norm)
        .filter((s) => FILLER.has(s));
    for (const pages of ["4", "8", "12", "16"]) {
      const m = compose(byId("booklet"), { paper: "letter" }, "word", { pages });
      m.booklet.spreads.forEach(([l, r]) => {
        const lines = [...fillerLines(m.pages[l]), ...fillerLines(m.pages[r])];
        const twice = lines.filter((s, i) => lines.indexOf(s) !== i);
        if (twice.length) report(`booklet/${pages}: pages ${l + 1} and ${r + 1} print on one side and both carry "${twice[0]}"`);
      });
    }
    if (!problems) ok("names, sponsor slots, invoice terms, detail samples, page guides, deep margins, badge notes, and booklet filler read right");
  }

  // House style in the template files.
  {
    const files = ["docs/app/core/templates/more.js", "docs/app/core/templates/resumes.js", "scripts/tests/templates.mjs"];
    const bad = files.filter((f) => fs.readFileSync(path.join(ROOT, f), "utf8").includes(String.fromCharCode(0x2014)));
    bad.length ? bad.forEach((f) => fail(`em dash in ${f}`)) : ok("template files follow the house style");
  }
}
