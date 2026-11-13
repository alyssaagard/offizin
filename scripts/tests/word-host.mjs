/**
 * Word host tests against a small simulated Office.js: Offizin's own undo and redo for creating
 * publications, inserting parts, restyling, and finishing, on Word on the web and on the desktop.
 * The simulation tracks the document body as text, content controls by tag, and settings.
 */

/** The custom document properties a generated .docx carries (the zip is stored, so the XML is plain). */
function readCustomProps(b64) {
  const text = Buffer.from(b64, "base64").toString("utf8");
  const start = text.indexOf('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"');
  if (start < 0) return [];
  const xml = text.slice(start, text.indexOf("</Properties>", start));
  const unescape = (s) => s.replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  return [...xml.matchAll(/<property [^>]*name="([^"]+)"><vt:lpwstr>([^<]*)<\/vt:lpwstr><\/property>/g)].map((m) => ({ name: unescape(m[1]), value: unescape(m[2]) }));
}

/**
 * options.desktop     Word for Windows or Mac with WordApiDesktop 1.3 (page setup and columns).
 * options.api         the highest WordApi version (1.5 by default; 1.3 has no styles API).
 * options.sections    [{ ps, cols, footers }] to start with several sections.
 * options.customProps [{ name, value }] custom document properties the open document carries.
 */
function makeOffice({ desktop, api = 1.5, sections: sectionSpecs, customProps } = {}) {
  const defaultPs = () => ({ pageWidth: 612, pageHeight: 792, topMargin: 72, bottomMargin: 72, leftMargin: 72, rightMargin: 72, gutter: 0, bookFoldPrinting: false, twoPagesOnOne: false, mirrorMargins: false, orientation: "Portrait", headerDistance: 36, footerDistance: 36, differentFirstPageHeaderFooter: false, oddAndEvenPagesHeaderFooter: false });
  const evenWidths = (n, keep = []) => Array.from({ length: n }, (_, i) => keep[i] || { width: Math.round(468 / n), spacing: 36 });
  const makeSection = (spec = {}) => {
    const cols = { count: 1, line: false, even: true, widths: [], ...(spec.cols || {}) };
    cols.widths = evenWidths(cols.count, cols.widths);
    return {
      ps: { ...defaultPs(), ...(spec.ps || {}) },
      cols,
      headers: { Primary: "<hdr/>", FirstPage: "<hdr/>", EvenPages: "<hdr/>", ...(spec.headers || {}) },
      footers: { Primary: "<ftr/>", FirstPage: "<ftr/>", EvenPages: "<ftr/>", ...(spec.footers || {}) },
    };
  };
  // Styles behave like Office.js proxies: load() only queues, context.sync() carries the queue
  // out, and a property that was not loaded cannot be read. Word keeps paragraph formatting only
  // on paragraph styles, so loading paragraphFormat on any other style fails the whole sync.
  const queue = [];
  const styleApi = (s) => {
    const loaded = new Set();
    const mark = (path) => {
      if (/^paragraphFormat\b/.test(path) && s.type !== "Paragraph") throw new Error(`InvalidArgument: ${s.nameLocal} is a ${s.type.toLowerCase()} style and has no paragraphFormat`);
      loaded.add(path);
    };
    const read = (path, value) => {
      if (!loaded.has(path)) throw new Error(`PropertyNotLoaded: ${path} of the style ${s.nameLocal}`);
      return value;
    };
    const group = (name) => {
      const o = { load: (keys) => queue.push(() => String(keys).split(",").forEach((k) => mark(`${name}/${k.trim()}`))) };
      for (const k of Object.keys(s[name])) Object.defineProperty(o, k, { get: () => read(`${name}/${k}`, s[name][k]), set: (v) => (s[name][k] = v), enumerable: true });
      return o;
    };
    const api = { mark, font: group("font"), paragraphFormat: group("paragraphFormat"), delete: () => (state.styles = state.styles.filter((x) => x !== s)) };
    for (const k of ["nameLocal", "type", "builtIn"]) Object.defineProperty(api, k, { get: () => read(k, s[k]), enumerable: true });
    return api;
  };
  const styleCollection = () => {
    const items = state.styles.map(styleApi);
    const load = (paths) =>
      queue.push(() => {
        if (state.stylesFail) throw new Error("GeneralException: the styles could not be read");
        for (const p of String(paths).split(",")) for (const it of items) it.mark(p.trim().replace(/^items\//, ""));
      });
    return { items, load };
  };
  const state = {
    body: "<original document>",
    ccs: [],
    settings: {},
    sections: (sectionSpecs || [{}]).map(makeSection),
    sel: [0], // indexes of the sections the selection touches
    newDocs: 0,
    newDocProps: null, // custom properties of the last document opened with createDocument
    stylesFail: false, // true: Word rejects every read of the styles
    readFails: false, // true: Word rejects reading what a content control holds
    customProps: customProps || [],
    styles: [],
  };
  state.styles = [
    { nameLocal: "Heading 1", type: "Paragraph", builtIn: true, font: { color: "#1B3A5C", name: "Georgia", size: 20, bold: true, italic: false }, paragraphFormat: { alignment: "Left", spaceBefore: 12, spaceAfter: 6, lineSpacing: 24, keepWithNext: true, keepTogether: false, leftIndent: 0, rightIndent: 0, firstLineIndent: 0, outlineLevel: "OutlineLevel1" } },
    { nameLocal: "Default Paragraph Font", type: "Character", builtIn: true, font: { color: "#000000", name: "Calibri", size: 11, bold: false, italic: false }, paragraphFormat: {} },
    { nameLocal: "Table Grid", type: "Table", builtIn: true, font: { color: "#000000", name: "Calibri", size: 11, bold: false, italic: false }, paragraphFormat: {} },
    { nameLocal: "No List", type: "List", builtIn: true, font: {}, paragraphFormat: {} },
  ];
  // The first section's page setup, for the older single-section checks.
  Object.defineProperty(state, "pageSetup", { get: () => state.sections[0].ps });
  const clientResult = (value) => ({ value });
  const readContent = (cc) => {
    if (state.readFails) throw new Error("GeneralException: the content control could not be read");
    return clientResult(cc.content);
  };
  const contentControl = (cc) => ({
    get tag() {
      return cc.tag;
    },
    get appearance() {
      return cc.appearance;
    },
    set appearance(v) {
      cc.appearance = v;
    },
    getOoxml: () => readContent(cc),
    getRange: () => ({ getOoxml: () => readContent(cc) }),
    insertOoxml(x, loc) {
      if (loc !== "Replace") throw new Error("unexpected location " + loc);
      cc.content = x;
    },
    delete(keep) {
      state.ccs = state.ccs.filter((c) => c !== cc);
      if (keep) state.body += cc.content;
    },
  });
  const collection = (list) => ({ items: list, load() {} });
  const textColumns = (sec) => ({
    load() {},
    get items() {
      return sec.cols.widths.map((w) => ({
        get width() {
          return w.width;
        },
        set width(v) {
          w.width = v;
        },
        get spacing() {
          return w.spacing;
        },
        set spacing(v) {
          w.spacing = v;
        },
      }));
    },
    getHasLineBetween: () => clientResult(sec.cols.line),
    getIsEvenlySpaced: () => clientResult(sec.cols.even),
    setCount(n) {
      sec.cols.count = n;
      sec.cols.widths = evenWidths(n);
    },
    setIsEvenlySpaced(v) {
      sec.cols.even = v;
    },
    setHasLineBetween(v) {
      sec.cols.line = v;
    },
  });
  const headerFooter = (store, type) => ({
    getOoxml: () => clientResult(store[type]),
    insertOoxml(x, loc) {
      if (loc !== "Replace") throw new Error("header insert " + loc);
      store[type] = x;
    },
  });
  const sectionObj = (sec, index) => {
    const ps = { textColumns: textColumns(sec), load() {} };
    for (const k of Object.keys(sec.ps)) Object.defineProperty(ps, k, { get: () => sec.ps[k], set: (v) => (sec.ps[k] = v), enumerable: true });
    return {
      pageSetup: ps,
      body: { getRange: () => ({ sectionIndex: index }) },
      getHeader: (type) => headerFooter(sec.headers, type),
      getFooter: (type) => headerFooter(sec.footers, type),
    };
  };
  const sectionList = (indexes) => ({ items: indexes.map((i) => sectionObj(state.sections[i], i)), load() {}, getFirst: () => sectionObj(state.sections[indexes[0]], indexes[0]) });
  // Document.insertFileFromBase64 copies the file's footer and imports its styles (Microsoft Learn).
  const importPublication = (b64) => {
    state.body = "DOCX:" + b64.length;
    state.ccs = [];
    for (const s of state.sections) s.footers.Primary = "<publication footer/>";
    for (const st of state.styles) if (st.type === "Paragraph") st.font.name = "Publication Font";
    const imported = (nameLocal, builtIn, size) => ({ nameLocal, type: "Paragraph", builtIn, font: { color: "#0E7C7B", name: "Publication Font", size, bold: true, italic: false }, paragraphFormat: { alignment: "Left", spaceBefore: 0, spaceAfter: 2, lineSpacing: 12, keepWithNext: true, keepTogether: false, leftIndent: 0, rightIndent: 0, firstLineIndent: 0, outlineLevel: "OutlineLevelBodyText" } });
    // The publication's styles.xml also defines built-in styles (Quote and Caption among them)
    // that a document may not have had before.
    for (const [name, builtIn, size] of [["Offizin Kicker", false, 9], ["Quote", true, 13], ["Caption", true, 8]]) {
      if (!state.styles.some((s) => s.nameLocal === name)) state.styles.push(imported(name, builtIn, size));
    }
  };
  const document = {
    body: {
      getOoxml: () => clientResult(state.body),
      insertOoxml(x, loc) {
        if (loc === "End") {
          const tag = (x.match(/<w:tag w:val="([^"]+)"\/>/) || [])[1];
          if (tag) state.ccs.push({ tag, content: x });
          else state.body += x;
          return;
        }
        if (loc !== "Replace") throw new Error("body insert " + loc);
        state.body = x;
      },
      insertFileFromBase64: importPublication,
      search: (text) => ({ items: state.body.includes(text) ? [{}] : [], load() {} }),
      inlinePictures: { items: [], load() {} },
    },
    insertFileFromBase64: importPublication,
    contentControls: {
      getByTag: (tag) => collection(state.ccs.filter((c) => c.tag === tag).map(contentControl)),
      get items() {
        return state.ccs.map(contentControl);
      },
      load() {},
    },
    getSelection: () => ({
      sections: sectionList(state.sel),
      paragraphs: {
        getLast: () => ({
          getRange: () => ({
            insertOoxml(pkg) {
              const tag = (pkg.match(/<w:tag w:val="([^"]+)"\/>/) || [])[1];
              if (tag) state.ccs.push({ tag, content: pkg });
              else state.body += pkg;
            },
          }),
        }),
      },
      insertOoxml() {},
      compareLocationWith(range) {
        const i = range.sectionIndex;
        if (state.sel.includes(i)) return clientResult(state.sel.length > 1 ? "Contains" : "Inside");
        return clientResult(i < Math.min(...state.sel) ? "After" : "Before");
      },
    }),
    get sections() {
      return sectionList(state.sections.map((s, i) => i));
    },
    getStyles: styleCollection,
    properties: {
      get customProperties() {
        return collection(state.customProps.map((p) => ({ key: p.name, value: p.value })));
      },
    },
  };
  const context = {
    document,
    application: {
      createDocument: (b64) => ({
        open: () => {
          state.newDocs++;
          state.newDocProps = readCustomProps(b64);
        },
      }),
    },
    // Carries out the queued style loads; the first one Word rejects fails the whole sync.
    sync: async () => {
      for (const job of queue.splice(0)) job();
    },
  };
  globalThis.Word = { run: async (fn) => fn(context), PageOrientation: { landscape: "Landscape", portrait: "Portrait" }, StyleType: { paragraph: "Paragraph" } };
  globalThis.Office = {
    PlatformType: { OfficeOnline: "OfficeOnline", PC: "PC" },
    context: {
      platform: desktop ? "PC" : "OfficeOnline",
      requirements: { isSetSupported: (set, v) => (set === "WordApi" ? Number(v) <= api : set === "WordApiDesktop" ? desktop && Number(v) <= 1.3 : false) },
      document: {
        settings: {
          get: (k) => state.settings[k] ?? null,
          set: (k, v) => (state.settings[k] = JSON.parse(JSON.stringify(v))),
          remove: (k) => delete state.settings[k],
          saveAsync: (cb) => cb && cb(),
        },
      },
    },
  };
  // Picture placeholders are drawn on a canvas in the browser.
  globalThis.document = { createElement: () => ({ width: 0, height: 0, getContext: () => new Proxy({}, { get: () => () => {} }), toDataURL: () => "data:image/png;base64,iVBORw0KGgo=" }) };
  return state;
}

export async function run({ fail, ok, load }) {
  const { TEMPLATES, PARTS, findTemplate } = await load("docs/app/core/templates.js");
  const { compose, composePart, restyleMap } = await load("docs/app/core/compose.js");
  const { createWordHost } = await load("docs/app/hosts/word.js");
  const { createHistory } = await load("docs/app/core/history.js");
  const { escapeXml } = await load("docs/app/core/util.js");
  const saved = { Word: globalThis.Word, Office: globalThis.Office, document: globalThis.document };

  try {
    for (const desktop of [false, true]) {
      const where = desktop ? "desktop" : "web";
      const state = makeOffice({ desktop });
      const host = createWordHost();
      const history = createHistory(host);

      // Create, undo, redo.
      const model = compose(findTemplate("trifold"), { paper: "letter" }, "word");
      const created = await host.createPublication(model, "replace");
      if (!created.op || created.op.kind !== "create") fail(`${where}: create returned no op`);
      if (!state.body.startsWith("DOCX:")) fail(`${where}: the publication was not inserted`);
      if (desktop && state.pageSetup.orientation !== "Landscape") fail(`${where}: page setup not applied`);
      if ((await host.getPublication()).mode === "new") fail(`${where}: a publication that replaced the document should not say mode "new"`);
      history.push(created.op);
      await history.undo();
      if (state.body !== "<original document>") fail(`${where}: undo did not restore the document (${state.body.slice(0, 40)})`);
      if (desktop && state.pageSetup.orientation !== "Portrait") fail(`${where}: undo did not restore page setup`);
      await history.redo();
      if (!state.body.startsWith("DOCX:")) fail(`${where}: redo did not rebuild the publication`);

      // Insert a part, undo leaves an anchor, redo fills it, a new action discards the anchor.
      const part = PARTS.find((p) => p.id === "heading");
      const composed = composePart(part, {}, "word", 468);
      const inserted = await host.insertPart(composed, 468, part.id, { label: part.name });
      history.push(inserted.op);
      if (state.ccs.length !== 1) fail(`${where}: the part is not in a content control`);
      await history.undo();
      if (!/w:line="20"/.test(state.ccs[0].content) || /Section label|SECTION LABEL/.test(state.ccs[0].content)) fail(`${where}: undo did not leave an empty anchor`);
      await history.redo();
      if (!/SECTION LABEL/.test(state.ccs[0].content)) fail(`${where}: redo did not put the part back`);
      await history.undo();
      const second = await host.insertPart(composed, 468, part.id, { label: part.name });
      history.push(second.op);
      await new Promise((r) => setTimeout(r, 0));
      if (state.ccs.length !== 1) fail(`${where}: the undone part's anchor was not discarded (${state.ccs.length} controls)`);

      // The person edits a part: Undo keeps that text for Redo, and Redo does not nest a second control.
      {
        const h4 = createHistory(host);
        const coupon = PARTS.find((p) => p.id === "coupon") || part;
        const cp = composePart(coupon, {}, "word", 468);
        const ins = await host.insertPart(cp, 468, coupon.id, { label: coupon.name });
        h4.push(ins.op);
        const ctl = state.ccs.find((c) => c.tag === ins.op.data.tag);
        if (!ctl) fail(`${where}: the coupon was not inserted in a content control`);
        else {
          if (ctl.appearance !== "Hidden") fail(`${where}: the part's marker is not hidden (${ctl.appearance})`);
          if (!/<w15:appearance w15:val="hidden"\/>/.test(ctl.content)) fail(`${where}: the part's package does not ask for the hidden appearance`);
          const sample = escapeXml(cp.samples[0]);
          if (!ctl.content.includes(sample)) fail(`${where}: the coupon sample text is missing from the part`);
          ctl.content = ctl.content.replace(sample, "Buy one loaf, get a second free");
          const u = await h4.undo();
          if (ctl.content.includes("second free")) fail(`${where}: undo did not empty the part`);
          if (!/^Removed: /.test(u.message)) fail(`${where}: unexpected undo message for a part (${u.message})`);
          await h4.redo();
          if (!ctl.content.includes("second free")) fail(`${where}: redo put back the original sample instead of what the person typed`);
          if (ctl.content.includes(sample)) fail(`${where}: redo brought the sample text back`);
          if (/<w:sdt>/.test(ctl.content)) fail(`${where}: redo nested a second content control in the part`);
          if (/could not be read/.test(u.message)) fail(`${where}: undo said the part could not be read when it was (${u.message})`);
          // Word cannot read the part at Undo time: the message says the typing is lost, and Redo
          // puts the part back as Offizin last had it.
          ctl.content = ctl.content.replace("second free", "third free");
          const warn = console.warn;
          console.warn = () => {};
          state.readFails = true;
          let u2;
          try {
            u2 = await h4.undo();
          } finally {
            state.readFails = false;
            console.warn = warn;
          }
          const name = coupon.name.toLowerCase();
          if (u2.message !== `Removed: ${name}. What you typed in it could not be read, so Redo will put back the part as it was before.`) fail(`${where}: undo should say that the typing could not be read (${u2.message})`);
          if (ctl.content.includes("third free")) fail(`${where}: undo did not empty the part when it could not be read`);
          await h4.redo();
          if (ctl.content.includes("third free") || !ctl.content.includes("second free")) fail(`${where}: redo after a failed read should put back the part as Offizin last had it`);
        }
        h4.clear();
        await new Promise((r) => setTimeout(r, 0));
      }

      // Restyle and undo it.
      const map = restyleMap({ colorScheme: "harbor", fontScheme: "editorial" }, { colorScheme: "plum", fontScheme: "bold" }, "word");
      state.body = '<w:color w:val="1B3A5C"/><w:rFonts w:ascii="Georgia"/>';
      const restyled = await host.restyle(map, { colorScheme: "plum", fontScheme: "bold" });
      history.push(restyled.op);
      if (state.body.includes("1B3A5C") || !state.body.includes("5B2A5E")) fail(`${where}: restyle did not change colors`);
      if (state.styles[0].font.color !== "#5B2A5E") fail(`${where}: restyle did not update paragraph styles`);
      await history.undo();
      if (!state.body.includes("1B3A5C") || !state.body.includes('w:ascii="Georgia"')) fail(`${where}: restyle undo did not restore colors and fonts`);

      // Finish removes markers and keeps content.
      const before = state.body.length;
      const msg = await host.finish();
      if (state.ccs.length) fail(`${where}: finish left content controls`);
      if (state.body.length <= before) fail(`${where}: finish lost the part content`);
      if (!/Finished/.test(msg)) fail(`${where}: unexpected finish message`);

      // New document mode makes no op.
      const nd = await host.createPublication(model, "new");
      if (nd.op) fail(`${where}: new document mode should not record an op`);

      // Booklets: book fold on the desktop, page by page on the web.
      const booklet = TEMPLATES.find((t) => t.fold === "booklet");
      if (booklet) {
        const bm = compose(booklet, { paper: "letter" }, "word");
        const res = await host.createPublication(bm, "replace");
        if (desktop && !state.pageSetup.bookFoldPrinting) fail("desktop: booklet did not turn on book fold");
        if (!desktop && !/page by page/.test(res.message)) fail("web: booklet message should explain page by page printing");
        if (!desktop && !(/2507/.test(res.message) && /16\.99\.2/.test(res.message) && /Open in a new document/.test(res.message) && /Custom Margins, Multiple pages/.test(res.message))) {
          fail(`web: booklet message should say where book fold is turned on and what to do elsewhere (${res.message})`);
        }
        if (!desktop && /Offizin turns on book fold there/.test(res.message)) fail("web: booklet message still promises book fold in the desktop app");
        const h2 = createHistory(host);
        h2.push(res.op);
        await h2.undo();
        if (desktop && state.pageSetup.bookFoldPrinting) fail("desktop: undo did not turn book fold off");
      }
      // Undo past the publication and redo it: the part's anchor is gone, so it goes at the end.
      {
        const h3 = createHistory(host);
        const bodyBefore = state.body;
        const pub = await host.createPublication(model, "replace");
        h3.push(pub.op);
        const p3 = await host.insertPart(composed, 468, part.id, { label: part.name });
        h3.push(p3.op);
        await h3.jump(0);
        if (state.body !== bodyBefore) fail(`${where}: a jump to the start did not restore the document`);
        const res = await h3.jump(2);
        const cc = state.ccs.find((c) => c.tag === p3.op.data.tag);
        if (!cc || !/SECTION LABEL/.test(cc.content)) fail(`${where}: the part did not come back after the publication was rebuilt`);
        if (!/end of the document/.test(res.messages[1] || "")) fail(`${where}: the message should say where the part went (${res.messages[1]})`);
        h3.clear();
      }

      // The checker searches every sample it is given, not just the first 80, and reports the
      // picture finding once.
      {
        const all = new Set();
        for (const t of TEMPLATES) for (const s of compose(t, {}, "word").samples) all.add(s);
        for (const pt of PARTS) for (const s of composePart(pt, {}, "word").samples) all.add(s);
        const list = [...all];
        const first = list[0];
        const last = list[list.length - 1];
        state.body = `<w:t>${first}</w:t><w:t>${list[Math.floor(list.length / 2)]}</w:t><w:t>${last}</w:t>`;
        const found = await host.check({ samples: list, infoSamples: [{ key: "org", label: "Business name", sample: "Zzz Example" }] });
        const texts = found.filter((f) => f.kind === "sample").map((f) => f.text);
        if (list.length <= 80) fail(`${where}: expected more than 80 library samples (${list.length})`);
        if (!texts.includes(first) || !texts.includes(last) || texts.length !== 3) fail(`${where}: check searched only part of ${list.length} samples (found ${texts.length})`);
        if (found.filter((f) => f.kind === "picture").length > 1) fail(`${where}: the picture finding was added more than once`);
      }

      // Open in a new document: the file carries the template and theme as custom properties,
      // and Offizin reads them in that document when its settings are empty.
      {
        const themed = compose(findTemplate("trifold"), { paper: "letter", colorScheme: "plum", fontScheme: "bold" }, "word");
        await host.createPublication(themed, "new");
        const props = state.newDocProps || [];
        if (!props.some((p) => p.name === "OffizinTemplate" && p.value === "trifold")) fail(`${where}: the new document carries no template property (${JSON.stringify(props)})`);
        if (props.some((p) => p.value.length > 255)) fail(`${where}: a custom property is longer than 255 characters`);
        const opened = makeOffice({ desktop, customProps: props });
        const host2 = createWordHost();
        const pub = await host2.getPublication();
        if (!pub || pub.templateId !== "trifold") fail(`${where}: getPublication found no record in the new document (${JSON.stringify(pub)})`);
        else {
          if (JSON.stringify(pub.theme) !== JSON.stringify(themed.theme)) fail(`${where}: the theme did not survive the file (${JSON.stringify(pub.theme)})`);
          if (pub.samples) fail(`${where}: the file record should carry no samples; the panel composes them`);
          if (pub.mode !== "new") fail(`${where}: the record of a document made with Open in a new document should say mode "new" (${JSON.stringify(pub.mode)})`);
          // Restyling that document keeps the record and writes the new theme to the settings.
          const map = restyleMap({ colorScheme: "plum", fontScheme: "bold" }, { colorScheme: "harbor", fontScheme: "editorial" }, "word");
          await host2.restyle(map, { colorScheme: "harbor", fontScheme: "editorial" });
          const after = await host2.getPublication();
          if (!after || after.templateId !== "trifold" || !after.theme || after.theme.colorScheme !== "harbor") fail(`${where}: restyle lost the record in the new document (${JSON.stringify(after)})`);
          if (!opened.settings.offizin || opened.settings.offizin.templateId !== "trifold") fail(`${where}: restyle did not save the record to the document settings`);
          if (after && after.mode !== "new") fail(`${where}: restyle lost mode "new" from the record (${JSON.stringify(after.mode)})`);
        }
        const empty = makeOffice({ desktop });
        if ((await createWordHost().getPublication()) !== null) fail(`${where}: a document without a record should give null`);
        void empty;
      }

      ok(`${where}: create, insert, restyle, undo, redo, jumps past a publication, discard, finish, check, and the new document record`);
    }

    // Several sections, Word for Windows and Mac: page setup and columns are recorded and restored
    // by section index, and Undo of a publication brings back every section, the footers, and the
    // style definitions.
    {
      const twoSections = (opts = {}) =>
        makeOffice({
          desktop: true,
          ...opts,
          sections: [{ cols: { count: 2, line: true, even: false, widths: [{ width: 300, spacing: 36 }, { width: 132, spacing: 36 }] }, footers: { Primary: "<my footer/>" } }, { ps: { orientation: "Landscape", pageWidth: 792, pageHeight: 612 } }],
        });
      const A4 = { width: 595, height: 842, orientation: "portrait", margins: { top: 72, bottom: 72, left: 72, right: 72 } };
      const show = (st) => st.sections.map((s, i) => `s${i + 1} ${s.ps.orientation} ${s.ps.pageWidth}x${s.ps.pageHeight} cols=${s.cols.count}${s.cols.line ? "+line" : ""}${s.cols.even ? "" : " uneven"}`).join(" | ");

      // A: a selection over both sections. Each section gets its own values back.
      {
        const st = twoSections();
        const host = createWordHost();
        const h = createHistory(host);
        st.sel = [0, 1];
        const r = await host.applyPageSetup(A4);
        h.push(r.op);
        if (st.sections[1].ps.pageWidth !== 595 || st.sections[0].ps.pageWidth !== 595) fail(`sections A: page setup was not applied to both selected sections (${show(st)})`);
        await h.undo();
        if (st.sections[0].ps.pageWidth !== 612 || st.sections[0].ps.orientation !== "Portrait") fail(`sections A: undo did not restore the first section (${show(st)})`);
        if (st.sections[1].ps.pageWidth !== 792 || st.sections[1].ps.orientation !== "Landscape") fail(`sections A: undo gave the second section the first one's values (${show(st)})`);
      }
      // B: change section 1, move the cursor to section 2, then Undo and Redo.
      {
        const st = twoSections();
        const host = createWordHost();
        const h = createHistory(host);
        st.sel = [0];
        const r = await host.applyPageSetup(A4);
        h.push(r.op);
        if (st.sections[1].ps.pageWidth !== 792) fail(`sections B: page setup touched the unselected section (${show(st)})`);
        st.sel = [1];
        await h.undo();
        if (st.sections[0].ps.pageWidth !== 612 || st.sections[1].ps.pageWidth !== 792) fail(`sections B: undo followed the cursor instead of the changed section (${show(st)})`);
        await h.redo();
        if (st.sections[0].ps.pageWidth !== 595 || st.sections[1].ps.pageWidth !== 792) fail(`sections B: redo followed the cursor instead of the changed section (${show(st)})`);
      }
      // C: columns with a line between and uneven widths come back exactly, in the right section.
      {
        const st = twoSections();
        const host = createWordHost();
        const h = createHistory(host);
        st.sel = [0];
        const r = await host.applyColumns({ count: 3, line: false });
        h.push(r.op);
        if (st.sections[0].cols.count !== 3 || st.sections[0].cols.line || !st.sections[0].cols.even) fail(`sections C: columns were not applied (${show(st)})`);
        await h.undo();
        const c = st.sections[0].cols;
        if (c.count !== 2 || !c.line || c.even || c.widths[0].width !== 300 || c.widths[1].width !== 132) fail(`sections C: undo did not restore the line between and the uneven widths (${show(st)} ${JSON.stringify(c.widths)})`);
        await h.redo();
        st.sel = [1];
        await h.undo();
        if (st.sections[0].cols.count !== 2 || st.sections[1].cols.count !== 1) fail(`sections C: undo with the cursor elsewhere changed the wrong section (${show(st)})`);
      }
      // D: Create, then Undo: every section's page setup, the footer, and the styles are back.
      {
        const st = twoSections();
        const host = createWordHost();
        const h = createHistory(host);
        const m = compose(findTemplate("trifold"), { paper: "letter" }, "word");
        const r = await host.createPublication(m, "replace");
        h.push(r.op);
        if (st.sections[0].footers.Primary !== "<publication footer/>" || !st.styles.some((s) => s.nameLocal === "Offizin Kicker")) fail("sections D: the stand-in did not import the footer and styles");
        if (st.sections[1].ps.orientation !== "Landscape" || st.sections[0].ps.orientation !== "Landscape") fail(`sections D: the tri-fold did not set every section landscape (${show(st)})`);
        const u = await h.undo();
        if (st.sections[0].ps.orientation !== "Portrait" || st.sections[0].ps.pageWidth !== 612) fail(`sections D: undo did not restore the first section (${show(st)})`);
        if (st.sections[1].ps.orientation !== "Landscape" || st.sections[1].ps.pageWidth !== 792) fail(`sections D: undo did not restore the landscape second section (${show(st)})`);
        if (st.sections[0].footers.Primary !== "<my footer/>") fail(`sections D: undo left the publication's footer (${st.sections[0].footers.Primary})`);
        if (st.sections[1].footers.Primary !== "<ftr/>") fail("sections D: undo left the publication's footer in the second section");
        if (st.styles.some((s) => s.nameLocal === "Offizin Kicker")) fail("sections D: undo left the Offizin styles the publication added");
        const h1 = st.styles.find((s) => s.nameLocal === "Heading 1");
        if (!h1 || h1.font.name !== "Georgia" || h1.font.size !== 20 || h1.paragraphFormat.spaceBefore !== 12) fail(`sections D: undo did not restore the Heading 1 definition (${JSON.stringify(h1 && h1.font)})`);
        if (!/back to how it was/.test(u.message) || /could not|does not let/.test(u.message)) fail(`sections D: unexpected undo message: ${u.message}`);
        // Redo takes a fresh snapshot and brings the publication back.
        await h.redo();
        if (st.sections[0].footers.Primary !== "<publication footer/>") fail("sections D: redo did not rebuild the publication");
        // The document had no Quote or Caption style, so the first Undo says they keep the
        // publication's look. By the second Create they were there, so the second Undo is silent.
        if (!u.message.includes("The publication added the Quote and Caption styles, which keep its look.")) fail(`sections D: undo should name the built-in styles the publication added (${u.message})`);
        if (!st.styles.some((s) => s.nameLocal === "Quote" && s.font.name === "Publication Font")) fail("sections D: the stand-in should keep the built-in styles the publication added");
        const u2 = await h.undo();
        if (/publication added/.test(u2.message)) fail(`sections D: a second undo named styles the document already had (${u2.message})`);
      }
      // D2: the document already had its own Quote style, so Undo restores it and names only Caption.
      {
        const st = twoSections();
        st.styles.push({ nameLocal: "Quote", type: "Paragraph", builtIn: true, font: { color: "#333333", name: "My Quote Font", size: 12, bold: false, italic: true }, paragraphFormat: { alignment: "Left", spaceBefore: 6, spaceAfter: 6, lineSpacing: 15, keepWithNext: false, keepTogether: false, leftIndent: 18, rightIndent: 18, firstLineIndent: 0, outlineLevel: "OutlineLevelBodyText" } });
        const host = createWordHost();
        const h = createHistory(host);
        const r = await host.createPublication(compose(findTemplate("trifold"), { paper: "letter" }, "word"), "replace");
        h.push(r.op);
        const quote = st.styles.find((s) => s.nameLocal === "Quote");
        if (quote.font.name !== "Publication Font") fail("sections D2: the stand-in did not restyle the Quote style");
        const u = await h.undo();
        if (quote.font.name !== "My Quote Font" || quote.font.italic !== true) fail(`sections D2: undo did not restore the document's own Quote style (${JSON.stringify(quote.font)})`);
        if (!u.message.includes("The publication added the Caption style, which keeps its look.")) fail(`sections D2: undo should name only the Caption style (${u.message})`);
        if (/Quote/.test(u.message)) fail(`sections D2: undo named a style it restored (${u.message})`);
      }
      // D3: Word has the styles API but the styles could not be read before Create. Undo says
      // that, and does not blame the version of Word.
      {
        const st = twoSections();
        st.stylesFail = true;
        const host = createWordHost();
        const h = createHistory(host);
        const warn = console.warn;
        console.warn = () => {};
        let r;
        try {
          r = await host.createPublication(compose(findTemplate("trifold"), { paper: "letter" }, "word"), "replace");
        } finally {
          console.warn = warn;
        }
        h.push(r.op);
        st.stylesFail = false;
        const u = await h.undo();
        if (!/style definitions could not be read before the publication/.test(u.message) || /This version of Word/.test(u.message)) fail(`sections D3: undo should say the styles could not be read (${u.message})`);
        if (st.sections[0].footers.Primary !== "<my footer/>") fail("sections D3: undo did not put the footer back when the styles could not be read");
      }
      // E: without the styles API (WordApi 1.3), Undo says the style definitions stay.
      {
        const st = twoSections({ api: 1.3 });
        const host = createWordHost();
        const h = createHistory(host);
        const m = compose(findTemplate("trifold"), { paper: "letter" }, "word");
        const r = await host.createPublication(m, "replace");
        h.push(r.op);
        const u = await h.undo();
        if (st.sections[0].footers.Primary !== "<my footer/>") fail("sections E: undo did not put the footer back without the styles API");
        if (!/style definitions/.test(u.message) || !/publication's look/.test(u.message)) fail(`sections E: undo should say the styles could not be put back (${u.message})`);
      }
      // F: Design tab page setup turns two pages per sheet off, and contentWidth halves the
      // width only for book fold or two pages on a landscape sheet.
      {
        const st = makeOffice({ desktop: true });
        const host = createWordHost();
        Object.assign(st.pageSetup, { twoPagesOnOne: true, pageWidth: 792, pageHeight: 612, orientation: "Landscape", leftMargin: 36, rightMargin: 36 });
        if ((await host.contentWidth()) !== 792 / 2 - 72) fail("page setup F: two pages on a landscape sheet should give half the width");
        Object.assign(st.pageSetup, { pageWidth: 612, pageHeight: 792, orientation: "Portrait" });
        if ((await host.contentWidth()) !== 612 - 72) fail(`page setup F: two pages on a portrait sheet keeps the full width (${await host.contentWidth()})`);
        Object.assign(st.pageSetup, { twoPagesOnOne: false, bookFoldPrinting: true, pageWidth: 792, pageHeight: 612 });
        if ((await host.contentWidth()) !== 792 / 2 - 72) fail("page setup F: book fold should give half the width");
        Object.assign(st.pageSetup, { twoPagesOnOne: true, bookFoldPrinting: true });
        await host.applyPageSetup(A4);
        if (st.pageSetup.twoPagesOnOne || st.pageSetup.bookFoldPrinting) fail("page setup F: a page setup without print settings should turn two pages per sheet and book fold off");
        if (st.pageSetup.pageWidth !== 595) fail("page setup F: page size not applied");
      }
      ok("sections: page setup and columns restored by section index; undo of a publication restores sections, footers, and styles, and names the built-in styles it added");
    }
  } finally {
    Object.assign(globalThis, saved);
  }
}
