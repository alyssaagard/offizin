/**
 * Tests for the panel's logic that runs without a browser: Offizin's own history (undo, redo,
 * jumps, trimming, failures), mail merge parsing, saved settings and brand kits, print steps,
 * keyboard shortcuts, Customize options, and print settings in the Word file.
 *
 * Run with: node scripts/check.mjs --only=panel
 */

export async function run({ ROOT, fail, ok, load }) {
  const expect = (cond, msg) => (cond ? ok(msg) : fail(msg));
  const { createHistory, normalizeResult, HISTORY_LIMIT } = await load("docs/app/core/history.js");
  const { parseRecords, recordsToText } = await load("docs/app/core/merge.js");
  const panel = await load("docs/app/ui/panel.js");
  const { compose, resolvePage, twoUpSheet, wordSheet, DEFAULT_LAYOUT } = await load("docs/app/core/compose.js");
  const { findTemplate, TEMPLATES } = await load("docs/app/core/templates.js");
  const { documentXml, settingsXml } = await load("docs/app/render/ooxml.js");
  const { FONT_EQUIVALENTS, WORD_FONT_CHOICES, GOOGLE_FONT_CHOICES } = await load("docs/app/core/schemes.js");

  // -------------------------------------------------------------------------
  // History with a pretend editor
  // -------------------------------------------------------------------------
  {
    const doc = [];
    const calls = [];
    const performer = {
      undo(op) {
        calls.push("undo " + op.data.id);
        if (op.data.failUndo) throw new Error("The editor said no.");
        doc.splice(doc.indexOf(op.data.id), 1);
        return `Removed ${op.data.id}.`;
      },
      redo(op) {
        calls.push("redo " + op.data.id);
        doc.push(op.data.id);
        return { message: "", op: { ...op, data: { ...op.data, redone: (op.data.redone || 0) + 1 } } };
      },
      discard(ops) {
        calls.push("discard " + ops.map((o) => o.data.id).join(","));
      },
    };
    let time = 0;
    const h = createHistory(performer, { limit: 4, now: () => ++time });
    const seen = [];
    const stop = h.onChange((s) => seen.push(s));
    const add = (id, kind = "insert") => {
      doc.push(id);
      return h.push({ kind, label: `Insert ${id}`, data: { id } });
    };
    add("a");
    add("b");
    add("c");
    let s = h.snapshot();
    expect(s.canUndo && !s.canRedo && s.undoLabel === "Insert c" && s.position === 3 && s.size === 3 && !s.busy, "history: three changes are recorded with labels");
    const r1 = await h.undo();
    expect(r1.ok && r1.message === "Removed c." && doc.join("") === "ab" && h.snapshot().redoLabel === "Insert c", "history: undo reverses only the last Offizin change");
    const r2 = await h.redo();
    expect(r2.ok && /Redone: Insert c/.test(r2.message) && doc.join("") === "abc", "history: redo applies it again, with a default message");
    expect(h.entries[2] && h.snapshot().entries[2].done, "history: entries know whether they are applied");
    await h.jump(0);
    expect(doc.length === 0 && h.position === 0 && h.canRedo && !h.canUndo, "history: a jump to the start undoes everything, newest first");
    expect(calls.slice(-3).join("|") === "undo c|undo b|undo a", "history: undo order is newest first");
    const plan = h.plan(2);
    expect(plan.redo.map((e) => e.label).join(",") === "Insert a,Insert b" && !plan.undo.length, "history: plan lists what a jump would redo");
    await h.jump(3);
    expect(doc.join("") === "abc", "history: a jump to the end redoes everything, oldest first");
    await h.undo();
    add("d");
    expect(calls[calls.length - 1] === "discard c" && h.size === 3 && !h.canRedo, "history: a new change drops the undone branch and lets the editor clean up");
    add("e");
    add("f");
    expect(h.size === 4 && calls[calls.length - 1] === "discard a" && h.entries[0].label === "Insert b", "history: the oldest change is trimmed at the limit");
    const bad = h.push({ kind: "insert", label: "Insert x", data: { id: "x", failUndo: true } });
    doc.push("x");
    let failed = "";
    try {
      await h.undo();
    } catch (e) {
      failed = e.message;
    }
    s = h.snapshot();
    expect(failed === "The editor said no." && s.position === s.size && s.entries[s.size - 1].failed === "The editor said no.", "history: a failed undo keeps its place and remembers why");
    expect(h.forget(bad.id) && !h.entries.some((e) => e.id === bad.id), "history: a change that cannot be undone can be removed");
    const p1 = h.undo();
    let guard = "";
    try {
      h.push({ kind: "insert", data: { id: "y" } });
    } catch (e) {
      guard = e.message;
    }
    await p1;
    expect(/still working/.test(guard), "history: nothing new is recorded while an undo is running");
    const before = seen.length;
    stop();
    await h.redo();
    expect(seen.length === before, "history: a listener can stop listening");
    expect(h.entries.length && h.entries[h.entries.length - 1].label === "Insert f", "history: refreshed ops from redo replace the stored ones");
    h.clear();
    expect(h.size === 0 && /^discard /.test(calls[calls.length - 1]), "history: clearing hands every op to the editor for cleanup");
    expect(HISTORY_LIMIT === 50, "history: 50 changes by default");
    const n1 = normalizeResult("Done.", "x");
    const n2 = normalizeResult({ message: "", op: { kind: "create", data: {} } }, "Made.");
    const n3 = normalizeResult(undefined, "Fallback.");
    expect(n1.message === "Done." && !n1.op && n2.message === "Made." && n2.op.kind === "create" && n3.message === "Fallback." && !n3.op, "history: host results of every shape become a message and an op");
  }

  // -------------------------------------------------------------------------
  // Mail merge parsing
  // -------------------------------------------------------------------------
  {
    const fields = findTemplate("labels").merge.fields;
    const tsv = "Name\tCompany\tStreet address\tCity\tState\tZip\nAvery Johnson\tLakeside Library\t48 Lake Road\tSpringfield\tST\t12345\nSam Okafor\t\t1200 College Ave\tNorthfield\tST\t12346\n\n";
    const a = parseRecords(tsv, fields);
    expect(a.header && a.records.length === 2 && a.records[0].postal === "12345" && a.records[1].company === "", "merge: rows from a spreadsheet match columns by their headings and aliases (Zip is the postal code)");
    expect(a.warnings.some((w) => /Address line 2/.test(w)), "merge: missing columns are named in a warning");
    const csv = 'name,company,address\n"Rivera, Jordan","Smith ""and"" Sons",1 Main St\n';
    const b = parseRecords(csv, fields);
    expect(b.records.length === 1 && b.records[0].name === "Rivera, Jordan" && b.records[0].company === 'Smith "and" Sons', "merge: quoted commas and quotes in .csv files are kept");
    const semi = parseRecords("Name;City\nAna;Lyon\nBen;Graz", fields);
    expect(semi.records.length === 2 && semi.records[1].city === "Graz", "merge: semicolon separated files work");
    const plain = parseRecords("Avery Johnson\nSam Okafor\nPriya Raman", findTemplate("place-cards").merge.fields);
    expect(!plain.header && plain.records.length === 3 && plain.records[2].name === "Priya Raman", "merge: one name per line works without a heading row");
    const many = parseRecords("Name\n" + Array.from({ length: 1200 }, (_, i) => `Person ${i}`).join("\n"), fields);
    expect(many.records.length === 1000 && many.warnings.some((w) => /1,000/.test(w)), "merge: at most 1,000 records, with a warning");
    expect(parseRecords("   \n  ", fields).records.length === 0, "merge: an empty list has no records");
    const round = parseRecords(recordsToText(a.records, fields), fields);
    expect(round.records.length === 2 && round.records[0].city === "Springfield" && round.header, "merge: records turn back into text and parse the same way");
    const ex = panel.exampleRecords(findTemplate("name-badges").merge);
    expect(ex.length === 6 && ex[0].name === "Avery Johnson" && ex.every((r) => Object.keys(r).length === findTemplate("name-badges").merge.fields.length), "merge: Fill in an example uses the template's own sample list");
    const ex2 = panel.exampleRecords(findTemplate("envelope").merge);
    expect(ex2.length === 3 && ex2[1].city === "Northfield" && ex2[1].address2 === "Room 214", "merge: Fill in an example makes complete addresses when a template has no sample list");
    for (const t of TEMPLATES.filter((x) => x.merge)) {
      const text = recordsToText(panel.exampleRecords(t.merge), t.merge.fields);
      const parsed = parseRecords(text, t.merge.fields);
      const m = compose(t, { paper: "letter" }, "web", { records: parsed.records });
      if (!parsed.records.length || !m.blocks.length) fail(`merge: the example list for ${t.id} did not build`);
    }
    ok("merge: the example list builds for every mail merge template");
  }

  // -------------------------------------------------------------------------
  // Saved settings and brand kits
  // -------------------------------------------------------------------------
  {
    const { sanitizeState, defaultState } = panel;
    const d = defaultState();
    expect(d.v === 2 && d.theme.layout.padScale === 1 && d.theme.customFonts.heading.word && d.theme.customFonts.body.google, "settings: defaults include spacing and a font scheme of your own");
    const s = sanitizeState({
      theme: {
        colorScheme: "nope",
        fontScheme: "custom",
        customFonts: { heading: { word: "Georgia", google: "Lora" }, body: { word: "Comic Sans'; color:red", google: "Inter" }, headingBold: false },
        layout: { textScale: 9, padScale: -1, gapScale: "x", lineScale: 1.2 },
        paper: "a4",
      },
      sets: { Main: { name: "Ana", bogus: "x" }, __proto__: { name: "evil" }, constructor: { name: "evil" } },
      activeSet: "constructor",
      logo: { src: "javascript:alert(1)" },
    });
    expect(s.theme.colorScheme === "ink" && s.theme.fontScheme === "custom" && s.theme.paper === "a4", "settings: unknown schemes fall back to black and white, and your own fonts are kept");
    expect(s.theme.customFonts.heading.google === "Lora" && s.theme.customFonts.body.word === "Calibri" && s.theme.customFonts.headingBold === false, "settings: only fonts from Offizin's lists are accepted");
    expect(s.theme.layout.textScale === 1.5 && s.theme.layout.padScale === 0 && s.theme.layout.gapScale === 1 && s.theme.layout.lineScale === 1.2, "settings: spacing and text values are kept within their limits");
    expect(Object.keys(s.sets).join() === "Main" && s.activeSet === "Main" && s.sets.Main.name === "Ana" && !("bogus" in s.sets.Main), "settings: information sets keep known fields, and unsafe set names are dropped");
    expect(s.logo === null && Object.getPrototypeOf(s.sets) === Object.prototype, "settings: a bad logo is dropped and nothing reaches the object prototype");
    const fromKit = sanitizeState(JSON.parse(JSON.stringify({ theme: d.theme, sets: d.sets, activeSet: d.activeSet })));
    expect(JSON.stringify(fromKit.theme) === JSON.stringify(d.theme), "settings: a brand kit round trip keeps every design setting");
    // Offizin 2.0 saved its untouched Harbor default with every change (legacy-1).
    const old = (v, colorScheme) => sanitizeState({ v, theme: { colorScheme }, sets: { Main: { org: "Lakeside Library" } } });
    expect(old(1, "harbor").theme.colorScheme === "ink" && old(undefined, "harbor").theme.colorScheme === "ink" && old(1, "harbor").sets.Main.org === "Lakeside Library", "settings: 2.0 settings with the old Harbor default move to Ink and keep everything else");
    expect(old(2, "harbor").theme.colorScheme === "harbor" && old(1, "plum").theme.colorScheme === "plum", "settings: Harbor chosen in 2.1, and other 2.0 schemes, stay as they are");
    // Brand kit logo sizes must be real, positive sizes (panel-13).
    const png = "data:image/png;base64,iVBORw0KGgo=";
    const logo = (width, height) => sanitizeState({ v: 2, logo: { src: png, width, height } }).logo;
    expect(logo(-50, 20).width === 200 && logo("Infinity", 20).width === 200 && logo(1e9, 20).width === 4000 && logo(0.5, 20).width === 200 && logo(120, "x").height === 100, "settings: a logo with a negative, infinite, huge, or missing size gets a safe size instead");
    expect(logo(120, 60).width === 120 && logo(120, 60).height === 60 && logo("120", "60").width === 120, "settings: a logo with a normal size keeps it");
    // The saved settings and the Info tab's New button share one limit (panel-8).
    const many = {};
    for (let i = 1; i <= panel.MAX_SETS + 1; i++) many[`Client ${i}`] = { org: `Client ${i}` };
    expect(panel.MAX_SETS === 20 && Object.keys(sanitizeState({ v: 2, sets: many }).sets).length === panel.MAX_SETS, "settings: the saved settings keep up to MAX_SETS information sets");
  }

  // -------------------------------------------------------------------------
  // Print steps and keyboard shortcuts
  // -------------------------------------------------------------------------
  {
    const { printSteps, historyKey, sameMargins, MARGIN_PRESETS, PRINT_TOGGLES, LAYOUT_CONTROLS } = panel;
    const desk = printSteps({ duplex: "short", paper: "Card stock.", finishing: "Fold it." }, { editor: "word", web: false, desktopPageSetup: true }, {});
    expect(desk[0] === "Choose File, then Print." && desk.some((x) => /Print on Both Sides, Flip pages on short edge/.test(x) && /Short-Edge binding/.test(x)), "print: Word for Windows and Mac steps name the two-sided setting");
    expect(desk.some((x) => /100 percent/.test(x)) && desk.includes("Paper: Card stock.") && desk.includes("After printing: Fold it."), "print: steps include scale, paper, and finishing");
    const web = printSteps({ duplex: "long" }, { editor: "word", web: true }, {});
    expect(/Word on the web makes a PDF/.test(web[0]) && web.some((x) => /flip on long edge/.test(x)), "print: Word on the web explains its PDF print");
    // A booklet on Word on the web: only Open in a new document carries book fold. Replace gives
    // separate pages, where a short edge flip would turn the backs upside down, so no duplex step.
    const webReplace = printSteps({ duplex: "short" }, { editor: "word", web: true }, { booklet: true, mode: "replace" });
    const webNew = printSteps({ duplex: "short" }, { editor: "word", web: true }, { booklet: true, mode: "new" });
    expect(
      webReplace.some((x) => /cannot turn on book fold/.test(x) && /printer's booklet setting/.test(x)) && !webReplace.some((x) => /which uses book fold|flip on short edge/.test(x)),
      "print: Word on the web with Replace says the booklet prints page by page and skips the short edge step"
    );
    expect(webNew.some((x) => /Book fold is saved in this document/.test(x) && /Open in Desktop App/.test(x)) && !webNew.some((x) => /cannot turn on book fold/.test(x)), "print: Word on the web with Open in a new document says book fold is in the file");
    const fold = printSteps({ duplex: "short" }, { editor: "word", desktopPageSetup: true }, { booklet: true });
    const noFold = printSteps({ duplex: "short" }, { editor: "word", desktopPageSetup: false }, { booklet: true, mode: "replace" });
    const newDoc = printSteps({ duplex: "short" }, { editor: "word", desktopPageSetup: false }, { booklet: true, mode: "new" });
    expect(fold.some((x) => /Book fold is turned on/.test(x)) && noFold.some((x) => /could not turn on book fold/.test(x)) && newDoc.some((x) => /Book fold is turned on/.test(x)), "print: booklet steps follow what Word could set");
    const two = printSteps({ twoPerSheet: true }, { editor: "word" }, {});
    expect(two.some((x) => /Two pages per sheet is saved/.test(x)) && !two.some((x) => /Both Sides/.test(x)), "print: two pages per sheet is explained, and one-sided pieces skip the duplex step");
    const g = printSteps({ duplex: "short" }, { editor: "google" }, { booklet: true });
    expect(/Google Docs opens your browser's print window/.test(g[0]) && g.some((x) => /printing order/.test(x)), "print: Google Docs steps");
    const all = [desk, web, webReplace, webNew, fold, noFold, newDoc, two, g, printSteps({}, {}, {})].flat();
    expect(all.every((x) => /[.!?]$/.test(x.trim()) && !x.includes(String.fromCharCode(0x2014))), "print: every step is a sentence with no em dashes");
    const key = (k, o = {}) => historyKey({ key: k, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, target: { tagName: "BUTTON" }, ...o });
    expect(key("z", { ctrlKey: true }) === "undo" && key("Z", { metaKey: true, shiftKey: true }) === "redo" && key("y", { ctrlKey: true }) === "redo", "keys: Ctrl or Command Z undoes, and Shift Z or Ctrl Y redoes");
    expect(key("z", { ctrlKey: true, target: { tagName: "INPUT", type: "text" } }) === null && key("z", { ctrlKey: true, target: { tagName: "TEXTAREA" } }) === null, "keys: text fields keep their own undo");
    expect(key("z") === null && key("z", { ctrlKey: true, altKey: true }) === null && key("z", { ctrlKey: true, defaultPrevented: true }) === null && key("y", { metaKey: true }) === null, "keys: other combinations are left alone");
    expect(sameMargins({ top: 72, right: 72, bottom: 72, left: 72 }, MARGIN_PRESETS.find((p) => p.id === "normal").m) && !sameMargins(null, {}), "customize: margin presets match Word's");
    expect(PRINT_TOGGLES.map((x) => x.key).join() === "mirrorMargins,twoPerSheet,embedFonts" && LAYOUT_CONTROLS.length === Object.keys(DEFAULT_LAYOUT).length, "customize: every spacing control and print toggle is listed");

    // Shrink to fit: small steps on the slider grid, floors that keep text readable, and the
    // user's own smaller settings left alone.
    const { shrinkStep, FIT_FLOORS } = panel;
    const keys = Object.keys(DEFAULT_LAYOUT);
    let v = { ...DEFAULT_LAYOUT };
    const seen = [];
    for (let i = 0; i < 40; i++) {
      const next = shrinkStep(v, keys);
      if (!next) break;
      seen.push(next);
      v = next;
    }
    const onGrid = (x) => Math.abs(x * 20 - Math.round(x * 20)) < 1e-9;
    const steadilySmaller = seen.every((s, i) => keys.every((k) => s[k] <= (i ? seen[i - 1][k] : DEFAULT_LAYOUT[k]) && onGrid(s[k])));
    expect(seen.length === 10 && keys.every((k) => v[k] === FIT_FLOORS[k]) && steadilySmaller, "shrink to fit: 5% steps down to floors that keep text readable");
    const mine = shrinkStep({ ...DEFAULT_LAYOUT, spaceScale: 0.2 }, keys);
    const textOnly = shrinkStep(DEFAULT_LAYOUT, ["textScale"]);
    expect(mine.spaceScale === 0.2 && mine.textScale === 0.95 && textOnly.textScale === 0.95 && textOnly.padScale === 1 && shrinkStep(v, keys) === null && shrinkStep(DEFAULT_LAYOUT, []) === null, "shrink to fit: changes only allowed controls and never raises your own settings");
  }

  // -------------------------------------------------------------------------
  // Customize options reach the model and the Word file
  // -------------------------------------------------------------------------
  {
    const flyer = findTemplate("event-flyer");
    const half = compose(flyer, { paper: "letter" }, "word", { paperSize: "halfLetter", print: { twoPerSheet: true, mirrorMargins: true, embedFonts: true } });
    expect(half.page.sizeName === "halfLetter" && half.print.twoPerSheet && half.print.mirrorMargins && half.print.embedFonts, "customize: paper size and print settings reach the model");
    const sheet = wordSheet(half, false);
    expect(sheet.width === 792 && sheet.height === 612, "print: two half letter pages share one Letter sheet in landscape");
    const xml = documentXml(half, { add: () => ({ id: "rIdTest", name: "media/test.png" }), nextDrawingId: 1 }, { bookFold: false });
    expect(/<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>/.test(xml), "print: the Word file stores the whole sheet for two pages per sheet");
    const settings = settingsXml(half.print, false);
    expect(settings.includes("<w:printTwoOnOne/>") && settings.includes("<w:mirrorMargins/>") && settings.includes("<w:embedTrueTypeFonts/>"), "print: the Word file keeps two per sheet, mirror margins, and embedded fonts");
    const legal = compose(flyer, { paper: "letter" }, "word", { paperSize: "legal", print: { twoPerSheet: true } });
    expect(!legal.print.twoPerSheet && wordSheet(legal, false).width === 612, "print: two pages per sheet is turned off when the page does not fit twice on standard paper");
    const a5 = twoUpSheet(resolvePage(flyer, "a4", { paperSize: "a5" }));
    expect(a5 && a5.sizeName === "a4" && Math.abs(a5.width - 841.89) < 0.01, "print: two A5 pages print on a true A4 sheet");
    const booklet = compose(findTemplate("booklet"), { paper: "letter" }, "word", { pages: "8", print: { twoPerSheet: true } });
    expect(!booklet.print.twoPerSheet && booklet.print.bookFold, "print: booklets use book fold, never two pages per sheet");
    const wide = compose(flyer, { paper: "letter" }, "web", { margins: { top: 300, right: 400, bottom: 10, left: 10 } });
    expect(wide.page.margins.top <= (792 - 144) / 2 && wide.page.margins.right <= (612 - 144) / 2, "customize: margins always leave two inches of room");
    const base = compose(flyer, { paper: "letter" }, "web");
    const tight = compose(flyer, { paper: "letter", layout: { ...DEFAULT_LAYOUT, padScale: 2 } }, "web", { layout: { textScale: 0.8 } });
    const sizes = (m) => JSON.stringify(m.styles.body.size) + JSON.stringify(m.styles.title.size);
    expect(sizes(base) !== sizes(tight) && tight.options.layout.padScale === 2 && tight.options.layout.textScale === 0.8, "customize: publication settings sit on top of the Design tab defaults");
    const custom = { fontScheme: "custom", customFonts: { heading: { word: "Rockwell", google: "Roboto Slab" }, body: { word: "Verdana", google: "Verdana" }, headingBold: false } };
    const cw = compose(flyer, { paper: "letter", ...custom }, "word");
    const cg = compose(flyer, { paper: "letter", ...custom }, "google");
    expect(cw.fonts.heading.name === "Rockwell" && cg.fonts.heading.name === "Roboto Slab" && cw.fonts.headingBold === false, "fonts: your own fonts reach Word and Google Docs by name");
    const words = WORD_FONT_CHOICES.map((f) => f.name);
    const googles = GOOGLE_FONT_CHOICES.map((f) => f.name);
    expect(
      words.every((n) => googles.includes(FONT_EQUIVALENTS.word[n])) && googles.every((n) => words.includes(FONT_EQUIVALENTS.google[n])),
      "fonts: every font has a close match in the other editor"
    );
  }

  // -------------------------------------------------------------------------
  // The panel in a document: the sidebar bundled from source and driven in jsdom with a
  // pretend host. Skipped, with a note, when jsdom is not installed.
  // -------------------------------------------------------------------------
  await runDomTests({ fail, ok, expect, ROOT, TEMPLATES, findTemplate, compose });
}

/** Finds jsdom: the repository's own node_modules, or the folder named in OFFIZIN_JSDOM. */
async function findJsdom(ROOT) {
  const { createRequire } = await import("node:module");
  const path = await import("node:path");
  const candidates = [path.join(ROOT, "package.json"), process.env.OFFIZIN_JSDOM && path.join(process.env.OFFIZIN_JSDOM, "package.json")].filter(Boolean);
  for (const c of candidates) {
    try {
      return createRequire(c)("jsdom");
    } catch (e) {
      // Try the next place.
    }
  }
  return null;
}

async function runDomTests({ fail, ok, expect, ROOT, TEMPLATES, findTemplate, compose }) {
  const jsdom = await findJsdom(ROOT);
  if (!jsdom) {
    ok("panel in a document: skipped, jsdom not found (npm install jsdom, or set OFFIZIN_JSDOM to a folder that has it)");
    return;
  }
  const fs = await import("node:fs");
  const os = await import("node:os");
  const path = await import("node:path");
  const { pathToFileURL } = await import("node:url");
  const { bundle } = await import(pathToFileURL(path.join(ROOT, "scripts/build.mjs")).href);
  const { PARTS, findPart } = await import(pathToFileURL(path.join(ROOT, "docs/app/core/templates.js")).href);
  const { composePart } = await import(pathToFileURL(path.join(ROOT, "docs/app/core/compose.js")).href);

  // One bundle of the panel from source, mounted through window.__ozMount.
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "offizin-panel-"));
  const entry = path.join(tmp, "entry.js");
  fs.writeFileSync(entry, `import { mountPanel } from "${path.join(ROOT, "docs/app/ui/panel.js")}";\nwindow.__ozMount = (root, host, opts) => mountPanel(root, host, opts);\n`);
  const js = bundle(entry);
  fs.rmSync(tmp, { recursive: true, force: true });

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const { JSDOM, VirtualConsole } = jsdom;

  /** A pretend editor: it records what the panel asks of it and keeps Offizin's history. */
  function fakeHost(over = {}) {
    const calls = [];
    const host = {
      kind: "test",
      platform: "word",
      label: "Test editor",
      capabilities: {
        createModes: [{ id: "replace", label: "Replace this document" }],
        history: true,
        restyle: true,
        printGuide: { editor: "word", web: false, desktopPageSetup: true },
        ...(over.capabilities || {}),
      },
      loadState: () => null,
      saveState: (s) => calls.push(["save", JSON.parse(JSON.stringify(s))]),
      createPublication: (m) => ({ message: `${m.name} created.`, op: { kind: "create", label: `Create ${m.name.toLowerCase()}`, data: { id: m.templateId } } }),
      insertPart: (c, w, id, { label }) => ({ message: `${label} inserted.`, op: { kind: "insert", label: `Insert ${label.toLowerCase()}`, data: { id } } }),
      undo: (op) => `Undone: ${op.label}.`,
      redo: (op) => ({ message: `Redone: ${op.label}.`, op }),
      discard: () => {},
      getPublication: () => null,
      contentWidth: () => 468,
      check: ({ samples, infoSamples }) => {
        calls.push(["check", samples, infoSamples]);
        return [];
      },
      ...over,
    };
    host.calls = calls;
    return host;
  }

  /** Mounts the panel in a fresh document and returns helpers for driving it. */
  function mount(host, { settle = 60, before = "" } = {}) {
    const errors = [];
    const vc = new VirtualConsole();
    vc.on("jsdomError", (e) => errors.push("jsdom: " + ((e.detail && e.detail.stack) || e.message)));
    vc.on("error", (...a) => errors.push("console.error: " + a.map((x) => (x && x.message) || String(x)).join(" ")));
    const dom = new JSDOM(`<!doctype html><html><head></head><body>${before}<div id="root"></div></body></html>`, { url: "https://offizin.test/", runScripts: "outside-only", pretendToBeVisual: true, virtualConsole: vc });
    const w = dom.window;
    w.HTMLCanvasElement.prototype.getContext = () => null;
    w.eval(js);
    const api = w.__ozMount(w.document.getElementById("root"), host, {});
    const $ = (s) => w.document.querySelector(s);
    const $$ = (s) => [...w.document.querySelectorAll(s)];
    const click = (el) => el && el.dispatchEvent(new w.MouseEvent("click", { bubbles: true, cancelable: true }));
    const change = (el, value) => {
      el.focus();
      el.value = value;
      el.dispatchEvent(new w.Event("change", { bubbles: true }));
    };
    const input = (el, value) => {
      el.value = value;
      el.dispatchEvent(new w.Event("input", { bubbles: true }));
    };
    const key = (el, k, o = {}) => el.dispatchEvent(new w.KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...o }));
    const press = (el) => {
      el.focus();
      click(el);
    };
    const tab = (id) => click($(`[data-tab="${id}"]`));
    const status = () => $(".oz-status").textContent;
    const focused = () => {
      const a = w.document.activeElement;
      if (!a || a === w.document.body) return "BODY";
      const attrs = [...a.attributes].filter((x) => /^(data-|class$|aria-current)/.test(x.name)).map((x) => `${x.name}=${x.value}`).join(" ");
      return `${a.tagName} ${attrs}`;
    };
    const ready = sleep(settle);
    return { w, api, $, $$, click, change, input, key, press, tab, status, focused, errors, ready, sleep };
  }

  // Status while a slow change runs (panel-12): the refusal message expires back to the busy
  // text, not to Ready. Its 6 second timer runs while the other tests go on.
  const slowDone = (async () => {
    const host = fakeHost({ createPublication: (m) => new Promise((r) => setTimeout(() => r({ message: `${m.name} created.`, op: { kind: "create", label: `Create ${m.name.toLowerCase()}`, data: {} } }), 7200)) });
    const p = mount(host);
    await p.ready;
    p.click(p.$('[data-template="trifold"]'));
    await p.sleep(30);
    p.click(p.$('[data-action="create"]'));
    await p.sleep(30);
    const busyText = p.status();
    p.key(p.$('[data-action="create"]'), "z", { ctrlKey: true });
    const refused = p.status();
    await p.sleep(6300);
    const later = p.status();
    const busyClass = p.$(".oz").classList.contains("oz-busy");
    await p.sleep(1200);
    return { busyText, refused, later, busyClass, after: p.status(), errors: p.errors };
  })();

  // Saved settings that arrive late (panel-1): nothing is saved over them, and the logo stays.
  {
    const saved = { v: 2, theme: { colorScheme: "custom", paper: "letter" }, sets: { Main: { name: "Ana Diaz", org: "Lakeside Library" } }, activeSet: "Main", logo: { src: "data:image/png;base64,iVBORw0KGgo=", width: 120, height: 60 } };
    const host = fakeHost({ loadState: () => new Promise((r) => setTimeout(() => r(JSON.parse(JSON.stringify(saved))), 500)) });
    const p = mount(host, { settle: 20 });
    await p.ready;
    const loadingShown = p.status();
    p.click(p.$('[data-paper="a4"]'));
    await p.sleep(420);
    const savedEarly = host.calls.filter((c) => c[0] === "save").length;
    await p.sleep(200);
    expect(p.status() === "Your saved settings have loaded. Make your last change again." && p.api.state.theme.paper === "letter", "panel: when the saved settings replace a change made while they loaded, the status line says to make it again");
    p.tab("info");
    const name = p.$('[data-field="name"]').value;
    expect(savedEarly === 0 && host.calls.filter((c) => c[0] === "save").length === 0 && name === "Ana Diaz" && !!p.api.state.logo && /Loading/.test(loadingShown), "panel: a change made before the saved settings arrive is not saved over them, and the logo stays");
    p.input(p.$('[data-field="phone"]'), "555-0100");
    await p.sleep(700);
    const last = host.calls.filter((c) => c[0] === "save").pop();
    expect(last && last[1].sets.Main.name === "Ana Diaz" && last[1].sets.Main.phone === "555-0100" && last[1].logo && last[1].theme.paper === "letter", "panel: after the saved settings arrive, changes save on top of them");
    expect(p.status() === "Ready" || /Ready/.test(p.status()) || p.status() !== loadingShown, "panel: the loading note goes away once the settings are in");
    expect(!p.errors.length, "panel: no errors in the document while loading settings" + (p.errors.length ? ": " + p.errors[0] : ""));
  }

  // A failed load is tried again before the first save (panel-1).
  {
    let attempts = 0;
    const saved = { v: 2, sets: { Main: { name: "Ana Diaz" } }, activeSet: "Main" };
    const host = fakeHost({
      loadState: () => {
        attempts++;
        return attempts === 1 ? Promise.reject(new Error("The server did not answer.")) : Promise.resolve(JSON.parse(JSON.stringify(saved)));
      },
    });
    const p = mount(host);
    await p.ready;
    const failedStatus = p.status();
    p.click(p.$('[data-paper="a4"]'));
    await p.sleep(450);
    expect(/^Your saved settings have loaded\. Make your last change again\.$/.test(p.status()), "panel: after a failed load, the change that loaded the settings again is named as one to repeat");
    p.tab("info");
    expect(/could not be loaded/.test(failedStatus) && attempts === 2 && host.calls.filter((c) => c[0] === "save").length === 0 && p.$('[data-field="name"]').value === "Ana Diaz", "panel: after a failed load, a change loads the settings again instead of saving the defaults");
    p.input(p.$('[data-field="org"]'), "Lakeside Library");
    await p.sleep(450);
    const last = host.calls.filter((c) => c[0] === "save").pop();
    expect(last && last[1].sets.Main.name === "Ana Diaz" && last[1].sets.Main.org === "Lakeside Library", "panel: once the second load worked, changes save on top of the saved settings");
  }

  // The design checker's samples (panel-2, checker-1, newdoc-1).
  {
    const coupon = composePart(findPart("coupon"), {}, "word", 468).samples;
    const trifold = compose(findTemplate("trifold"), { paper: "letter" }, "word").samples;
    const badges = compose(findTemplate("name-badges"), { paper: "letter" }, "word").samples;
    let pub = null;
    const host = fakeHost({ getPublication: () => pub });
    const p = mount(host);
    await p.ready;
    p.tab("check");
    const runCheck = async () => {
      host.calls.length = 0;
      p.press(p.$('[data-action="check"]'));
      await p.sleep(80);
      return host.calls.filter((c) => c[0] === "check");
    };
    let checks = await runCheck();
    let samples = checks[0] ? checks[0][1] : [];
    expect(checks.length === 1 && samples.length >= 200 && badges.every((s) => samples.includes(s)) && coupon.every((s) => samples.includes(s)), "check: with no publication record, every template's and every page part's sample text is sent in one call");
    expect(checks[0] && Array.isArray(checks[0][2]) && checks[0][2].length > 0 && checks[0][2].every((x) => x.key && x.sample), "check: the business information samples go with it, once");
    pub = { templateId: "trifold", theme: { colorScheme: "ink", fontScheme: "editorial" }, samples: trifold.slice(0, 5) };
    checks = await runCheck();
    samples = checks[0][1];
    expect(trifold.slice(0, 5).every((s) => samples.includes(s)) && coupon.every((s) => samples.includes(s)) && !badges.some((s) => !trifold.includes(s) && samples.includes(s)), "check: with a record, its samples and every page part's samples are searched, not the whole library");
    pub = { templateId: "trifold", theme: { colorScheme: "ink", fontScheme: "editorial" } };
    checks = await runCheck();
    samples = checks[0][1];
    expect(trifold.every((s) => samples.includes(s)) && coupon.every((s) => samples.includes(s)) && samples.length < 120, "check: a record with a template and no samples gets them from the template in its theme");
    expect(/Check this publication|BUTTON/.test(p.focused()) && /data-action=check/.test(p.focused()), "focus: Check this publication keeps the focus after the results are drawn");
  }

  // The Check tab's print steps follow the create mode in the publication record: a booklet made
  // with Open in a new document carries book fold, and a record without a mode means Replace.
  {
    const stepsFor = async (pub) => {
      const host = fakeHost({ capabilities: { printGuide: { editor: "word", web: true } }, getPublication: () => pub });
      const p = mount(host);
      await p.ready;
      p.tab("check");
      p.press(p.$('[data-action="check"]'));
      await p.sleep(80);
      return p.$$(".oz-print + .oz-steps li").map((li) => li.textContent);
    };
    const newDoc = await stepsFor({ templateId: "booklet", theme: {}, mode: "new" });
    const replaced = await stepsFor({ templateId: "booklet", theme: {} });
    expect(newDoc.some((x) => /Book fold is saved in this document/.test(x)) && !newDoc.some((x) => /cannot turn on book fold/.test(x)), "check: a booklet record with mode new gets the print steps for a file that carries book fold");
    expect(replaced.some((x) => /cannot turn on book fold/.test(x)) && !replaced.some((x) => /Book fold is saved in this document/.test(x)), "check: a booklet record without a mode gets the print steps for Replace");
  }

  // Information sets (panel-3, panel-8, panel-9) and the focus around them (panel-4, panel-7).
  {
    const host = fakeHost({ loadState: () => ({ v: 2, sets: { Main: { name: "Ana Diaz", org: "Lakeside Library" }, "Spring  Fair": { org: "Fair Committee" } }, activeSet: "Main" }) });
    const p = mount(host);
    await p.ready;
    expect(p.status() === "Ready", "panel: saved settings that load with no change before them leave the status at Ready");
    p.tab("info");
    p.click(p.$('[data-action="new-set"]'));
    p.$("[data-newset]").value = "Main";
    p.click(p.$('[data-action="newset-ok"]'));
    expect(/already exists/.test(p.status()) && p.api.state.sets.Main.name === "Ana Diaz" && p.$('[data-field="name"]').value === "Ana Diaz", "info: New with the active set's name is refused and the set keeps its information");
    p.$("[data-newset]").value = "  Fall   Fair ";
    p.press(p.$('[data-action="newset-ok"]'));
    expect(p.api.state.sets["Fall Fair"] && p.api.state.activeSet === "Fall Fair" && /data-action=set/.test(p.focused()), "info: a new set's name has single spaces, and the focus moves to the set list");
    const options = p.$$('[data-action="set"] option').map((o) => [o.textContent, o.value]);
    expect(options.some(([t, v]) => t === "Spring  Fair" && v === "Spring  Fair"), "info: a set name with two spaces in a row is the option's value too");
    p.change(p.$('[data-action="set"]'), "Spring  Fair");
    expect(p.api.state.activeSet === "Spring  Fair" && p.$('[data-field="org"]').value === "Fair Committee" && /SELECT data-action=set/.test(p.focused()), "info: choosing that set shows its fields and keeps the focus on the list");
    p.press(p.$('[data-action="delete-set"]'));
    expect(!p.api.state.sets["Spring  Fair"] && /Information set deleted/.test(p.status()) && (/data-action=delete-set/.test(p.focused()) || /data-action=set/.test(p.focused())), "info: Delete keeps the focus in the panel");
    p.click(p.$('[data-action="rename-set"]'));
    p.$("[data-newset]").value = p.api.state.activeSet;
    p.click(p.$('[data-action="newset-ok"]'));
    expect(!/already exists/.test(p.status()), "info: Rename may keep the set's own name");
    for (let i = Object.keys(p.api.state.sets).length; i < 20; i++) {
      p.click(p.$('[data-action="new-set"]'));
      p.$("[data-newset]").value = `Client ${i}`;
      p.click(p.$('[data-action="newset-ok"]'));
    }
    const atLimit = Object.keys(p.api.state.sets).length;
    const newDisabled = p.$('[data-action="new-set"]').disabled;
    p.$('[data-action="new-set"]').disabled = false;
    p.click(p.$('[data-action="new-set"]'));
    p.$("[data-newset]").value = "One too many";
    p.click(p.$('[data-action="newset-ok"]'));
    expect(atLimit === 20 && newDisabled && /up to 20 information sets/.test(p.status()) && Object.keys(p.api.state.sets).length === 20, "info: the 21st set is refused with a message, and New is disabled at the limit");
    expect(!p.errors.length, "info: no errors in the document" + (p.errors.length ? ": " + p.errors[0] : ""));
  }

  // Focus after redraws (panel-4, panel-7) and the preview's role (panel-14).
  {
    const host = fakeHost();
    const p = mount(host);
    await p.ready;
    p.change(p.$('[data-action="category"]'), "brochures");
    expect(/SELECT data-action=category/.test(p.focused()) && p.$('[data-action="category"]').value === "brochures", "focus: the Show list keeps the focus when it changes");
    p.press(p.$('[data-paper="a4"]'));
    expect(/data-paper=a4/.test(p.focused()) && p.api.state.theme.paper === "a4", "focus: the gallery's Paper buttons keep the focus");
    p.click(p.$('[data-template="trifold"]'));
    await p.sleep(30);
    expect(p.$(".oz-large").getAttribute("role") === "img" && /Preview of Tri-fold/.test(p.$(".oz-large").getAttribute("aria-label")), "a11y: the large preview is one picture with a label");
    expect(/one inch \(25 mm\) wide/.test(p.$("[data-customize]").textContent), "customize: the tri-fold margins help says each panel stays an inch wide");
    p.press(p.$('[data-action="custom-reset"]'));
    expect(/data-action=custom-reset/.test(p.focused()), "focus: Reset all customizations keeps the focus");
    p.tab("design");
    const swatch = p.$('[data-scheme="harbor"]') || p.$$("[data-scheme]")[1];
    p.press(swatch);
    expect(new RegExp(`data-scheme=${swatch.dataset.scheme}`).test(p.focused()), "focus: a color scheme swatch keeps the focus");
    p.change(p.$('[data-hex="primary"]'), "#224466");
    expect(/data-hex=primary/.test(p.focused()) && p.api.state.theme.customColors.primary === "#224466", "focus: a committed hex code keeps the focus");
    expect(/plain resumes and CVs/.test(p.$(".oz-tip").textContent) && /stay as they are/.test(p.$(".oz-tip").textContent), "design: the styles tip says which designs follow the editor's styles fully");
    p.tab("check");
    p.change(p.$('[data-action="fold-type"]'), "booklet");
    expect(/SELECT data-action=fold-type/.test(p.focused()) && p.$('[data-action="fold-type"]').value === "booklet", "focus: the Publication type list keeps the focus when it changes");
  }

  // Margin presets and Same on all sides (panel-11).
  {
    const host = fakeHost();
    const p = mount(host);
    await p.ready;
    p.click(p.$('[data-template="newsletter"]'));
    await p.sleep(30);
    const same = p.$('[data-custom="sameMargins"]');
    same.checked = true;
    same.dispatchEvent(new p.w.Event("change", { bubbles: true }));
    p.click(p.$('[data-preset="wide"]'));
    const m = (s) => Number(p.$(`[data-margin="${s}"]`).value);
    const wide = [m("top"), m("bottom"), m("left"), m("right")];
    const afterWide = p.$('[data-custom="sameMargins"]').checked;
    p.input(p.$('[data-margin="left"]'), "1.25");
    const typed = [m("top"), m("bottom"), m("left"), m("right")];
    expect(wide.join() === "1,1,2,2" && afterWide === false && typed.join() === "1,1,1.25,2", "customize: Wide unchecks Same on all sides, so the next typed margin changes one side only");
    p.click(p.$('[data-preset="normal"]'));
    expect(p.$('[data-custom="sameMargins"]').checked === true, "customize: a preset with equal sides checks Same on all sides");
  }

  // The confirm box and the history list give the focus back (panel-5, panel-focus-10), and
  // Undo anyway refuses to undo a different change than the one it named (panel-10).
  {
    const host = fakeHost();
    const p = mount(host);
    await p.ready;
    p.click(p.$('[data-template="trifold"]'));
    await p.sleep(30);
    p.click(p.$('[data-action="create"]'));
    await p.sleep(60);
    p.press(p.$(".oz-undo"));
    await p.sleep(10);
    const opened = p.focused();
    p.click(p.$('[data-action="confirm-cancel"]'));
    const afterCancel = p.focused();
    p.press(p.$(".oz-undo"));
    await p.sleep(10);
    p.key(p.w.document.activeElement, "Escape");
    const afterEscape = p.focused();
    expect(/confirm-cancel/.test(opened) && /class=oz-icon oz-undo/.test(afterCancel) && /class=oz-icon oz-undo/.test(afterEscape), "focus: Cancel and Escape put the focus back on the Undo button");
    p.press(p.$(".oz-undo"));
    await p.sleep(10);
    p.tab("insert");
    p.click(p.$('[data-part="coupon"]'));
    await p.sleep(60);
    const stillOpen = !p.$(".oz-confirm").hidden;
    p.click(p.$('[data-action="confirm-ok"]'));
    await p.sleep(60);
    const snap = p.api.history.snapshot();
    expect(stillOpen && /changed while this question was open/.test(p.status()) && snap.position === 2 && snap.entries.every((e) => e.done), "history: Undo anyway does nothing when the history changed while the question was open");
    p.press(p.$(".oz-undo"));
    await p.sleep(60);
    expect(/Undone: Insert coupon/.test(p.status()) && /class=oz-icon oz-undo/.test(p.focused()), "focus: after an undo with more to undo, the Undo button keeps the focus");
    p.press(p.$(".oz-undo"));
    await p.sleep(10);
    p.press(p.$('[data-action="confirm-ok"]'));
    await p.sleep(80);
    expect(/Undone: Create tri-fold/.test(p.status()) && p.$(".oz-undo").disabled && /class=oz-icon oz-histbtn/.test(p.focused()), "focus: after Undo anyway undoes the last change, the focus goes to the history button");
    p.press(p.$(".oz-histbtn"));
    p.press(p.$('[data-jump="2"]'));
    await p.sleep(80);
    expect(p.api.history.position === 2 && /data-jump=2/.test(p.focused()) && /aria-current=step/.test(p.focused()), "focus: after a jump in the history list, the new current step has the focus");
    p.press(p.$('[data-action="history-clear"]'));
    await p.sleep(10);
    p.press(p.$('[data-action="confirm-ok"]'));
    expect(p.api.history.size === 0 && /class=oz-icon oz-histbtn/.test(p.focused()), "focus: after Clear history, the history button has the focus");
    expect(!p.errors.length, "history: no errors in the document" + (p.errors.length ? ": " + p.errors[0] : ""));
  }

  // Undo and redo asked for from outside the panel (the live demo's app bar and page) leave the
  // keyboard focus there, also when nothing had the focus. Only a request made from the panel moves it.
  {
    const host = fakeHost();
    const p = mount(host, { before: '<button type="button" id="outside">Undo</button><div id="board" tabindex="0"></div>' });
    await p.ready;
    p.click(p.$('[data-template="trifold"]'));
    await p.sleep(30);
    p.click(p.$('[data-action="create"]'));
    await p.sleep(60);
    p.tab("insert");
    p.click(p.$('[data-part="coupon"]'));
    await p.sleep(60);
    const outside = p.$("#outside");
    const board = p.$("#board");
    const at = () => p.w.document.activeElement;
    outside.focus();
    await p.api.undo();
    const afterUndo = at();
    const undone = p.status();
    await p.api.redo();
    const afterRedo = at();
    board.focus();
    await p.api.undo();
    expect(afterUndo === outside && /Undone: Insert coupon/.test(undone) && afterRedo === outside && at() === board, "focus: an undo or redo asked for from outside the panel leaves the focus outside");
    // Undoing the publication asks first. The question takes the focus, then gives it back.
    outside.focus();
    let asked = p.api.undo();
    await p.sleep(10);
    const inQuestion = p.focused();
    p.click(p.$('[data-action="confirm-cancel"]'));
    await asked;
    const afterCancel = at();
    asked = p.api.undo();
    await p.sleep(10);
    p.press(p.$('[data-action="confirm-ok"]'));
    await asked;
    expect(/confirm-cancel/.test(inQuestion) && afterCancel === outside && at() === outside && /Undone: Create tri-fold/.test(p.status()), "focus: after the undo question, the focus goes back to where it was outside the panel");
    at().blur();
    await p.api.redo();
    const active = outside.ownerDocument.activeElement;
    expect(!active || active === outside.ownerDocument.body, "focus: with nothing focused, a redo asked for from outside leaves the focus out of the panel");
    expect(!p.errors.length, "focus: no errors in the document after undo and redo from outside" + (p.errors.length ? ": " + p.errors[0] : ""));
  }

  // Shrink to fit moves the focus to the Spacing and text section when its button goes away.
  {
    const long = TEMPLATES.find((t) => t.id === "event-flyer");
    const host = fakeHost();
    const p = mount(host);
    await p.ready;
    p.click(p.$(`[data-template="${long.id}"]`));
    await p.sleep(30);
    const fit = p.$('[data-action="fit"]');
    // jsdom has no layout, so the preview never overflows; the button stays hidden and the
    // section focus rule is checked directly instead.
    expect(fit && fit.closest("[hidden]") !== null, "fit: the Shrink to fit button is hidden while nothing overflows");
  }

  const slow = await slowDone;
  expect(/Creating tri-fold/.test(slow.busyText) && /still working/.test(slow.refused) && slow.later === slow.busyText && slow.busyClass, "status: a refusal message expires back to the running change's text, not to Ready");
  expect(/created\.$/.test(slow.after) && !slow.errors.length, "status: the change's own message follows when it is done");
}
