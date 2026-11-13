/**
 * Google Docs tests: Code.gs against the simulated DocumentApp, with Offizin's own Undo and Redo.
 */
import fs from "node:fs";
import path from "node:path";
import { makeGoogle, invariants, textOf } from "../lib/google-mock.mjs";

const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

function withImages(blocks, walkBlocks) {
  const copy = JSON.parse(JSON.stringify(blocks));
  walkBlocks(copy, (b) => b.t === "image" && (b.image = { mime: "image/png", base64: PNG }));
  return copy;
}

/** A document with real content: headings, runs, links, a list, a table, a picture, a page break. */
function fillDocument(g) {
  const { body, Para, Table, blob } = g;
  body.clear();
  const first = body.getChild(0);
  first.appendText("Annual report");
  first.setHeading("HEADING1");
  first.editAsText().setAttributes(0, 5, { BOLD: true, FOREGROUND_COLOR: "#AA0000" });
  const p2 = body.appendParagraph("Visit our site for details.");
  p2.editAsText().setLinkUrl(10, 13, "https://example.org");
  p2.setAlignment("CENTER");
  p2.setSpacingAfter(9);
  const li1 = body.appendListItem("First point");
  const li2 = body.appendListItem("Second point");
  li2.setListId(li1);
  const t = body.appendTable([["Name", "Score"], ["Ana", "12"]]);
  t.getCell(0, 0).setBackgroundColor("#EEEEEE");
  t.getCell(1, 1).getChild(0).appendText("!");
  t.setColumnWidth(0, 200);
  const p3 = body.appendParagraph("After the table");
  p3.appendInlineImage(blob([1, 2, 3], "image/png", "x")).setWidth(120);
  p3.appendPageBreak();
  body.appendParagraph("Second page");
  body.setPageWidth(500);
  body.setMarginLeft(40);
  body.setHeadingAttributes("HEADING1", { FONT_SIZE: 30 });
  g.doc.addFooter().getChild(0).setText("Old footer");
}

export async function run({ ROOT, fail, ok, load }) {
  const { TEMPLATES, PARTS, findTemplate } = await load("docs/app/core/templates.js");
  const { compose, composePart, walkBlocks, restyleMap, invertRestyleMap } = await load("docs/app/core/compose.js");
  const CODE = path.join(ROOT, "google-docs/Code.gs");
  const check = (body, label) => invariants(body).forEach((m) => fail(`${label}: ${m}`));
  const model = (id, theme = {}, opts = {}) => {
    const m = compose(findTemplate(id), { paper: "letter", ...theme }, "google", opts);
    const { pages, options, ...rest } = m;
    return { ...rest, blocks: withImages(m.blocks, walkBlocks) };
  };

  // 1. Create on an empty document, undo, redo.
  {
    const g = makeGoogle(CODE);
    g.body.clear();
    const res = g.context.ozCreate(model("trifold"));
    if (!res || !res.snapshot || !res.snapshot.empty) fail("create on an empty document should take the fast snapshot");
    check(g.body, "trifold");
    g.context.ozUndoCreate(res.snapshot);
    if (g.body.getNumChildren() !== 1 || g.body.getChild(0).getText() !== "") fail("undo did not bring back the empty document");
    const again = g.context.ozCreate(model("trifold"));
    if (textOf(g.body).indexOf("Why it matters") < 0) fail("redo did not rebuild the publication");
    check(g.body, "trifold after redo");
    if (!again.snapshot.empty) fail("the second snapshot should also be of an empty document");
    ok("create, undo, and redo on an empty document");
  }

  // 2. Create over real content, then undo restores everything Offizin can copy.
  {
    const g = makeGoogle(CODE);
    fillDocument(g);
    const beforeText = textOf(g.body);
    const beforeCount = g.body.getNumChildren();
    const res = g.context.ozCreate(model("newsletter"));
    if (!res.snapshot.chunks) fail("the snapshot was not stored in the cache");
    if (!/Insert, then Page numbers/.test(res.message)) fail("the newsletter message should explain page numbers");
    const footerText = g.getFooter().getText();
    if (/\{page|Page \{/.test(footerText) || !/News/.test(footerText)) fail(`footer text not cleaned: "${footerText}"`);
    if (g.body.getPageWidth() !== 612) fail("page width was not applied");
    if (g.body.getHeadingAttributes("HEADING1").FONT_SIZE === 30) fail("heading styles were not set from the design");
    check(g.body, "newsletter");
    const tables = [];
    const walk = (n) => {
      if (n.type === "TABLE") tables.push(n);
      n.children.forEach(walk);
    };
    walk(g.body);
    const columnTable = tables.find((t) => t.children[0].children.length === 3 && /Member spotlight|MEMBER SPOTLIGHT/.test(textOf(t)));
    if (!columnTable) fail("the newsletter's flowing columns did not become a three cell table");
    g.context.ozUndoCreate(res.snapshot);
    const afterText = textOf(g.body);
    if (afterText !== beforeText) fail(`undo text differs:\n${beforeText}\n---\n${afterText}`);
    if (g.body.getNumChildren() !== beforeCount) fail(`undo has ${g.body.getNumChildren()} elements, expected ${beforeCount}`);
    const h = g.body.getChild(0);
    if (h.getHeading() !== "HEADING1") fail("undo lost the heading");
    if (!h.editAsText().isBold(0) || h.editAsText().getForegroundColor(0) !== "#aa0000") fail("undo lost bold or color runs");
    if (h.editAsText().isBold(8)) fail("undo made unbolded text bold");
    const p2 = g.body.getChild(1);
    if (p2.editAsText().getLinkUrl(11) !== "https://example.org" || p2.getAlignment() !== "CENTER" || p2.getSpacingAfter() !== 9) fail("undo lost a link, alignment, or spacing");
    const li2 = g.body.getChild(3);
    if (li2.getType() !== "LIST_ITEM" || li2.getListId() !== g.body.getChild(2).getListId()) fail("undo did not keep list items together");
    const table = g.body.getChild(4);
    if (table.getType() !== "TABLE" || table.getCell(0, 0).getBackgroundColor() !== "#EEEEEE" || table.getCell(1, 1).getChild(0).getText() !== "12!") fail("undo lost table content or cell color");
    const img = g.body.getImages()[0];
    if (!img || img.getWidth() !== 120) fail("undo lost the picture or its size");
    if (g.body.getPageWidth() !== 500 || g.body.getMarginLeft() !== 40) fail("undo did not restore the page setup");
    if (g.body.getHeadingAttributes("HEADING1").FONT_SIZE !== 30) fail("undo did not restore the heading style");
    if (g.getFooter().getText() !== "Old footer") fail("undo did not restore the footer");
    if (g.cache.size) fail("undo should clear the cached copy");
    check(g.body, "restored document");
    ok("undoing a publication restores text, runs, links, lists, tables, pictures, page setup, styles, and footer");
  }

  // 3. A large document is split across cache values, and an expired copy explains Version history.
  {
    const g = makeGoogle(CODE);
    g.body.clear();
    for (let i = 0; i < 40; i++) g.body.appendParagraph(`Paragraph ${i} ` + "words and élèves ".repeat(200));
    const res = g.context.ozCreate(model("event-flyer"));
    if (res.snapshot.chunks < 2) fail("a large document should use several cache values");
    g.context.ozUndoCreate(res.snapshot);
    if (!/Paragraph 39/.test(textOf(g.body))) fail("a large document did not come back");
    const res2 = g.context.ozCreate(model("event-flyer"));
    g.cache.clear();
    let msg = "";
    try {
      g.context.ozUndoCreate(res2.snapshot);
    } catch (e) {
      msg = e.message;
    }
    if (!/Version history/.test(msg)) fail("an expired copy should point to Version history");
    ok("large documents are split across the cache, and an expired copy is explained");
  }

  // 4. Inserted parts: named ranges, undo leaves an anchor, redo refills it, discard and finish.
  {
    const g = makeGoogle(CODE);
    const part = PARTS.find((p) => p.id === "coupon");
    const blocks = withImages(composePart(part, {}, "google", 468).blocks, walkBlocks);
    g.context.ozInsert(blocks, 468, "a1");
    if (g.doc.getNamedRanges("offizin-op-a1").length !== 1) fail("the part has no named range");
    const withPart = textOf(g.body);
    g.context.ozUndoInsert("a1");
    if (/Save 20%/.test(textOf(g.body))) fail("undo did not remove the part");
    if (g.doc.getNamedRanges("offizin-anchor-a1").length !== 1) fail("undo left no anchor");
    g.context.ozRedoInsert("a1", blocks, 468);
    if (textOf(g.body) !== withPart) fail("redo did not put the part back in the same place");
    g.context.ozUndoInsert("a1");
    g.context.ozDiscardAnchors(["a1"]);
    if (g.doc.getNamedRanges().length) fail("discard left named ranges");
    g.context.ozInsert(blocks, 468, "a2");
    const finish = g.context.ozFinish();
    if (g.doc.getNamedRanges().length || !/Finished/.test(finish)) fail("finish did not remove markers");
    if (!/Save 20%/.test(textOf(g.body))) fail("finish removed content");
    check(g.body, "parts");
    ok("parts: insert, undo to an anchor, redo in place, discard, and finish");
  }

  // 4b. Undo past the publication and redo it: the anchor is gone, and the part returns to its position.
  {
    const g = makeGoogle(CODE);
    g.body.clear();
    const created = g.context.ozCreate(model("trifold"));
    const part = PARTS.find((p) => p.id === "coupon");
    const blocks = withImages(composePart(part, {}, "google", 468).blocks, walkBlocks);
    const ins = g.context.ozInsert(blocks, 468, "j1");
    const withPart = textOf(g.body);
    const undone = g.context.ozUndoInsert("j1");
    if (undone.index !== ins.index) fail(`the anchor should sit where the part started (${undone.index} and ${ins.index})`);
    g.context.ozUndoCreate(created.snapshot);
    g.context.ozCreate(model("trifold"));
    const redone = g.context.ozRedoInsert("j1", blocks, 468, undone.index);
    if (textOf(g.body) !== withPart) fail("the part did not return to the same place after the publication was rebuilt");
    if (!/earlier position/.test(redone.message)) fail("the message should say the part went back to its earlier position");
    if (g.doc.getNamedRanges("offizin-anchor-j1").length) fail("the stale anchor was not removed");
    if (g.doc.getNamedRanges("offizin-op-j1").length !== 1) fail("the part has no named range after redo");
    let threw = "";
    try {
      g.context.ozRedoInsert("nope", blocks, 468, null);
    } catch (e) {
      threw = e.message;
    }
    if (!/cannot find/.test(threw)) fail("a part with no anchor and no position should explain itself");
    check(g.body, "jump");
    ok("parts: after Undo and Redo of the publication, a part returns to its earlier position");
  }

  // 5. Restyle and its undo; page setup and its undo.
  {
    const g = makeGoogle(CODE);
    g.body.clear();
    g.context.ozCreate(model("event-flyer", { colorScheme: "harbor", fontScheme: "editorial" }));
    const map = restyleMap({ colorScheme: "harbor", fontScheme: "editorial" }, { colorScheme: "plum", fontScheme: "bold" }, "google");
    const res = g.context.ozRestyle(map, { colorScheme: "plum", fontScheme: "bold" });
    if (!res.previous || res.previous.colorScheme !== "harbor") fail("restyle did not report the previous design");
    const fonts = new Set();
    const walk = (n) => {
      if (n.type === "PARAGRAPH" && n.text) fonts.add(n.editAsText().getFontFamily(0));
      n.children.forEach(walk);
    };
    walk(g.body);
    if (fonts.has("Georgia") && !fonts.has("Oswald")) fail("restyle did not change fonts");
    g.context.ozRestyle(invertRestyleMap(map), res.previous);
    const fonts2 = new Set();
    const walk2 = (n) => {
      if (n.type === "PARAGRAPH" && n.text) fonts2.add(n.editAsText().getFontFamily(0));
      n.children.forEach(walk2);
    };
    walk2(g.body);
    if (fonts2.has("Oswald")) fail("undoing the restyle left new fonts behind");
    const ps = g.context.ozPageSetup({ width: 400, height: 600, margins: { top: 10, right: 10, bottom: 10, left: 10 } });
    if (!ps.before || ps.before.width !== 612) fail("page setup did not report the previous values");
    g.context.ozPageSetup(ps.before);
    if (g.body.getPageWidth() !== 612) fail("page setup undo did not restore the width");
    ok("restyle and page setup can be undone");
  }

  // 6. Pictures use pixels, and every template still builds.
  {
    const g = makeGoogle(CODE);
    g.body.clear();
    const m = model("event-flyer");
    g.context.ozCreate(m);
    const img = g.body.getImages()[0];
    let expected = 0;
    walkBlocks(m.blocks, (b) => b.t === "image" && !expected && (expected = Math.round(b.width * 96 / 72)));
    if (!img || img.getWidth() !== expected) fail(`picture width should be ${expected} pixels, got ${img && img.getWidth()}`);
    for (const t of TEMPLATES) {
      const g2 = makeGoogle(CODE);
      try {
        const res = g2.context.ozCreate(model(t.id));
        check(g2.body, t.id);
        g2.context.ozUndoCreate(res.snapshot);
        if (textOf(g2.body) !== "Existing text") fail(`${t.id}: undo did not restore the original text`);
      } catch (e) {
        fail(`${t.id}: ${e.message}`);
      }
    }
    ok(`pictures are sized in pixels, and all ${TEMPLATES.length} templates build and undo`);
  }

  /** Makes one method of the simulated document throw the given number of times. */
  const failOnce = (target, method, times = 1) => {
    const real = target[method];
    target[method] = function (...args) {
      if (times-- > 0) throw new Error("Service error: Documents");
      return real.apply(this, args);
    };
  };
  /** Makes the nth call of a method fail on every simulated element of one kind, then puts the method back. */
  const failNthOf = (proto, method, nth) => {
    const real = proto[method];
    let n = 0;
    proto[method] = function (...args) {
      if (++n < nth) return real.apply(this, args);
      proto[method] = real;
      throw new Error("Service error: Documents");
    };
  };
  const threw = (fn) => {
    try {
      fn();
    } catch (e) {
      return e.message;
    }
    return "";
  };
  const formatFooter = (g) => {
    const f = g.doc.addFooter();
    const p1 = f.getChild(0);
    p1.setText("Acme Inc.");
    p1.setAlignment("LEFT");
    p1.setAttributes({ FONT_FAMILY: "Georgia", BOLD: true, FONT_SIZE: 9 });
    const p2 = f.appendParagraph("Confidential");
    p2.setAlignment("RIGHT");
    p2.setAttributes({ ITALIC: true, FONT_SIZE: 8 });
  };
  const describeFooter = (g) =>
    g.getFooter().children.map((p) => {
      const t = p.editAsText();
      return `${p.text}|${p.alignment}|${t.getFontFamily(0)} ${t.getFontSize(0)} ${t.isBold(0) ? "bold" : ""}${t.isItalic(0) ? "italic" : ""}`;
    }).join(" ; ");

  // 7. A failure while building a publication puts the old document back (gs-create-failure-2, page-1, page-3).
  {
    const g = makeGoogle(CODE);
    fillDocument(g);
    const before = textOf(g.body);
    const footerBefore = g.getFooter().getText();
    failOnce(g.body, "insertParagraph", 1);
    const msg = threw(() => g.context.ozCreate(model("newsletter")));
    if (!/back as it was/.test(msg) || !/Service error/.test(msg)) fail(`a failed create should explain itself: "${msg}"`);
    if (textOf(g.body) !== before) fail(`a failed create did not put the document back:\n${textOf(g.body)}`);
    if (g.getFooter().getText() !== footerBefore) fail("a failed create changed the footer");
    if (g.body.getPageWidth() !== 500 || g.body.getHeadingAttributes("HEADING1").FONT_SIZE !== 30) fail("a failed create did not restore page setup or styles");
    if (g.docProps["offizin.pending"]) fail("the pending record should be gone after the document is back");
    if (g.docProps["offizin.publication"]) fail("a failed create must not save a publication record");
    if (g.cache.size) fail("the cached copy should be dropped once the document is back");
    check(g.body, "after a failed create");
    // A bad layout value (negative widths or picture sizes) is one more way the build can fail.
    const bad = model("trifold");
    walkBlocks(bad.blocks, (b) => b.t === "grid" && b.cols && (b.cols[0] = -60));
    const msg2 = threw(() => g.context.ozCreate(bad));
    if (!/back as it was/.test(msg2) || textOf(g.body) !== before) fail("a bad layout value should leave the document as it was");
    ok("a publication that fails half way puts the old document back and says so");
  }

  // 7b. When the script is stopped before it can put the document back, the sidebar recovers from the pending record.
  {
    const g = makeGoogle(CODE);
    fillDocument(g);
    const before = textOf(g.body);
    failOnce(g.body, "insertParagraph", 1);
    const realRestore = g.context.ozRestoreSnapshot;
    let restores = 0;
    g.context.ozRestoreSnapshot = function (...args) {
      if (restores++ === 0) throw new Error("Exceeded maximum execution time");
      return realRestore.apply(this, args);
    };
    const msg = threw(() => g.context.ozCreate(model("event-flyer")));
    if (!/could not put the document back/.test(msg)) fail(`unexpected message: "${msg}"`);
    if (!g.docProps["offizin.pending"]) fail("the pending record should stay when the document could not be put back");
    const rec = g.context.ozRecoverCreate();
    if (!rec || !rec.restored) fail("recovery should report that it restored the document");
    if (textOf(g.body) !== before) fail("recovery did not bring the document back");
    if (g.docProps["offizin.pending"] || g.cache.size) fail("recovery should remove the pending record and the cached copy");
    const again = g.context.ozRecoverCreate();
    if (again.restored) fail("with no pending record there is nothing to recover");
    // A pending record that no sidebar recovered belongs to that run only. A later create that
    // stops before it changes anything must not bring the old copy back over newer work.
    failOnce(g.body, "insertParagraph", 1);
    restores = 0;
    threw(() => g.context.ozCreate(model("event-flyer")));
    if (!g.docProps["offizin.pending"]) fail("the second stopped create should leave a pending record");
    g.body.clear();
    g.body.getChild(0).appendText("Rewritten by hand after the failure");
    g.body.appendParagraph("Two hours of new work");
    const rewritten = textOf(g.body);
    g.context.OZ_MAX_CHUNKS = 0;
    const late = threw(() => g.context.ozCreate(model("event-flyer")));
    if (!/too large for Offizin to keep a copy/.test(late)) fail(`the later create should stop as too large: "${late}"`);
    if (g.docProps["offizin.pending"]) fail("a new create should clear a pending record left by an earlier run");
    const stale = g.context.ozRecoverCreate();
    if (stale.restored) fail("a pending record from an earlier run must not be recovered after a later create");
    if (textOf(g.body) !== rewritten) fail(`recovery after a later failed create replaced the current document:\n${textOf(g.body)}`);
    ok("the sidebar can recover a stopped create from the pending record, and only from its own run");
  }

  // 7c. A pending record carries the id of the attempt that wrote it. A sidebar asking for
  //     another attempt (its own call never reached the server) gets nothing back.
  {
    const g = makeGoogle(CODE);
    fillDocument(g);
    const before = textOf(g.body);
    const stop = () => {
      failOnce(g.body, "insertParagraph", 1);
      const real = g.context.ozRestoreSnapshot;
      let n = 0;
      g.context.ozRestoreSnapshot = function (...args) {
        if (n++ === 0) throw new Error("Exceeded maximum execution time");
        return real.apply(this, args);
      };
    };
    stop();
    threw(() => g.context.ozCreate({ ...model("event-flyer"), run: "first" }));
    if (!g.docProps["offizin.pending"]) fail("the stopped create should leave a pending record");
    const half = textOf(g.body);
    const other = g.context.ozRecoverCreate("second");
    if (other.restored || textOf(g.body) !== half) fail("a record from another attempt must not be restored");
    if (g.docProps["offizin.pending"]) fail("a record from another attempt should be cleared");
    stop();
    threw(() => g.context.ozCreate({ ...model("event-flyer"), run: "third" }));
    const mine = g.context.ozRecoverCreate("third");
    if (!mine.restored) fail("the attempt that wrote the record should get its document back");
    if (before === half) fail("the test document should have changed during the stopped create");
    ok("a pending record is restored only for the attempt that wrote it");
  }

  // 8. Undo of a publication keeps the cached copy until the document is back, and a retry works (gs-undo-cache-3).
  {
    const g = makeGoogle(CODE);
    fillDocument(g);
    const before = textOf(g.body);
    const res = g.context.ozCreate(model("event-flyer"));
    failOnce(g.body, "appendTable", 1);
    const msg = threw(() => g.context.ozUndoCreate(res.snapshot));
    if (!/could not finish bringing back/.test(msg) || !/Try Undo again/.test(msg)) fail(`a failed undo should say to try again: "${msg}"`);
    if (!g.cache.size) fail("a failed undo must keep the cached copy");
    const retry = g.context.ozUndoCreate(res.snapshot);
    if (!/back to how it was/.test(retry) || textOf(g.body) !== before) fail("the retry did not bring the document back");
    if (g.cache.size) fail("a finished undo should drop the cached copy");
    const res2 = g.context.ozCreate(model("event-flyer"));
    g.cache.clear();
    const gone = threw(() => g.context.ozUndoCreate(res2.snapshot));
    if (/six hours have passed/.test(gone) || !/can drop it sooner/.test(gone)) fail(`a missing copy must not claim a time: "${gone}"`);
    ok("a failed undo of a publication keeps the copy for a retry, and a missing copy is explained honestly");
  }

  // 9. The person's own footer survives a publication without a footer, and comes back formatted (gs-footer-4).
  {
    const g = makeGoogle(CODE);
    formatFooter(g);
    const footerBefore = describeFooter(g);
    const res = g.context.ozCreate(model("trifold"));
    if (res.snapshot.footer !== null) fail("a template without a footer should not copy the footer");
    if (describeFooter(g) !== footerBefore) fail("a template without a footer changed the footer");
    g.context.ozUndoCreate(res.snapshot);
    if (describeFooter(g) !== footerBefore) fail(`undo changed a footer the publication never touched:\n${describeFooter(g)}`);
    const res2 = g.context.ozCreate(model("newsletter"));
    if (!res2.snapshot.footer || !res2.snapshot.footer.exists) fail("a template with a footer should copy the footer first");
    if (describeFooter(g) === footerBefore) fail("the newsletter should replace the footer");
    g.context.ozUndoCreate(res2.snapshot);
    if (describeFooter(g) !== footerBefore) fail(`undo did not bring back the footer's paragraphs and formatting:\n${describeFooter(g)}`);
    const g2 = makeGoogle(CODE);
    const res3 = g2.context.ozCreate(model("newsletter"));
    if (!res3.snapshot.footer || res3.snapshot.footer.exists) fail("a document without a footer should remember that");
    g2.context.ozUndoCreate(res3.snapshot);
    if (g2.getFooter() && g2.getFooter().getText() !== "") fail("undo should empty a footer the publication added");
    ok("undo leaves an untouched footer alone and restores a replaced one with its formatting");
  }

  // 10. Undo of a page part keeps what the person typed in it, and Redo puts exactly that back (parts-user-edits-5).
  {
    const g = makeGoogle(CODE);
    const part = PARTS.find((p) => p.id === "coupon");
    const blocks = withImages(composePart(part, {}, "google", 468).blocks, walkBlocks);
    g.context.ozInsert(blocks, 468, "e1");
    const hit = g.body.findText("Save 20%");
    hit.getElement().para.setText("Save 50% this week only");
    const edited = textOf(g.body);
    const undone = g.context.ozUndoInsert("e1");
    if (!Array.isArray(undone.content) || !undone.content.length) fail("undo should hand back the part's content");
    if (/Save 50%/.test(textOf(g.body))) fail("undo did not remove the edited part");
    g.context.ozRedoInsert("e1", blocks, 468, undone.index, undone.content);
    if (textOf(g.body) !== edited) fail(`redo did not put back the edited part:\n${textOf(g.body)}`);
    if (/Save 20%/.test(textOf(g.body))) fail("redo brought back the sample text instead of the typing");
    if (g.doc.getNamedRanges("offizin-op-e1").length !== 1 || g.doc.getNamedRanges("offizin-anchor-e1").length) fail("redo did not name the restored part");
    g.context.ozUndoInsert("e1");
    if (/Save 50%/.test(textOf(g.body))) fail("the restored part could not be undone again");
    check(g.body, "edited part");
    ok("parts: undo keeps the person's typing and redo restores it exactly");
  }

  // 11. A document too large to copy is left alone (gs-large-doc-9).
  {
    const g = makeGoogle(CODE);
    fillDocument(g);
    const before = textOf(g.body);
    g.context.OZ_MAX_CHUNKS = 0;
    const msg = threw(() => g.context.ozCreate(model("event-flyer")));
    if (!/too large for Offizin to keep a copy/.test(msg) || !/Make a copy/.test(msg)) fail(`a large document should stop with a plain message: "${msg}"`);
    if (textOf(g.body) !== before) fail("a large document was changed although no copy could be kept");
    if (g.docProps["offizin.pending"] || g.docProps["offizin.publication"] || g.cache.size) fail("a stopped create should leave nothing behind");
    ok("a document too large to copy is not changed");
  }

  // 12. Undo writes back every style attribute Offizin set, including ones the document left unset (gs-headings-11).
  {
    const g = makeGoogle(CODE);
    g.body.setHeadingAttributes("TITLE", { FONT_SIZE: 26, BOLD: null, FOREGROUND_COLOR: null, FONT_FAMILY: null, ITALIC: null });
    const res = g.context.ozCreate(model("event-flyer"));
    if (g.body.getHeadingAttributes("TITLE").FONT_FAMILY === "Arial") fail("create should set the Title font");
    g.context.ozUndoCreate(res.snapshot);
    const t = g.body.getHeadingAttributes("TITLE");
    const h1 = g.body.getHeadingAttributes("HEADING1");
    if (t.FONT_FAMILY !== "Arial" || t.BOLD !== false || t.ITALIC !== false || t.FOREGROUND_COLOR !== "#000000" || t.FONT_SIZE !== 26) fail(`undo left the Title style changed: ${JSON.stringify(t)}`);
    if (h1.FONT_FAMILY !== "Arial" || h1.BOLD !== false || h1.FONT_SIZE !== 20) fail(`undo left Heading 1 changed: ${JSON.stringify(h1)}`);
    ok("undo restores named styles whose attributes were unset before");
  }

  // 13. A page part that fails half way leaves nothing behind, on insert and on redo (gs-create-failure-2).
  {
    const g = makeGoogle(CODE);
    const part = PARTS.find((p) => p.id === "coupon");
    const blocks = withImages(composePart(part, {}, "google", 468).blocks, walkBlocks);
    const before = textOf(g.body);
    const count = g.body.getNumChildren();
    failOnce(g.body, "insertTable", 1);
    const msg = threw(() => g.context.ozInsert(blocks, 468, "f1"));
    if (!/nothing was added/.test(msg)) fail(`a failed insert should say nothing was added: "${msg}"`);
    if (textOf(g.body) !== before || g.body.getNumChildren() !== count) fail("a failed insert left elements behind");
    if (g.doc.getNamedRanges().length) fail("a failed insert left a named range");
    if (/except/.test(msg)) fail(`with nothing left behind, the message should say only that: "${msg}"`);
    // At the very end of the document Docs keeps the last paragraph, and the message says so.
    {
      const g2 = makeGoogle(CODE);
      failNthOf(g2.body, "insertParagraph", 2);
      const tail = threw(() => g2.context.ozInsert(blocks, 468, "f3"));
      const kids = g2.body.children;
      if (kids.length !== 2 || kids[1].type !== "PARAGRAPH" || kids[1].text !== "" || kids[1].children.length) fail(`a failed insert at the end should leave one empty paragraph at most: ${kids.map((c) => c.type).join(", ")}`);
      if (!/nothing was added except one empty line at the end of the document\. \(Service error/.test(tail)) fail(`a failed insert at the end should name the empty line it leaves: "${tail}"`);
      if (textOf(g2.body) !== before || g2.doc.getNamedRanges().length) fail("a failed insert at the end left content or a named range");
      check(g2.body, "failed part at the end");
    }
    g.context.ozInsert(blocks, 468, "f2");
    const withPart = textOf(g.body);
    const undone = g.context.ozUndoInsert("f2");
    const afterUndo = textOf(g.body);
    failOnce(g.body, "insertTable", 1);
    const msg2 = threw(() => g.context.ozRedoInsert("f2", blocks, 468, undone.index, undone.content));
    if (!/nothing was added/.test(msg2) || textOf(g.body) !== afterUndo) fail("a failed redo should leave the document as it was");
    if (g.doc.getNamedRanges("offizin-anchor-f2").length !== 1) fail("a failed redo should keep the anchor for another try");
    // A restore that stops in the middle: the table is back and the next element fails, or the
    // table itself fails after its first text was put back. Either way nothing may stay behind.
    const shape = () => g.body.children.map((c) => `${c.type}:${textOf(c)}`).join(" | ");
    const shapeAfterUndo = shape();
    const stops = [
      ["its second element", () => failOnce(g.body, "insertParagraph", 1)],
      ["a half filled table", () => failNthOf(g.Para.prototype, "appendText", 2)],
    ];
    for (const [where, arm] of stops) {
      arm();
      const msg3 = threw(() => g.context.ozRedoInsert("f2", blocks, 468, undone.index, undone.content));
      if (!/nothing was added/.test(msg3) || !/Service error/.test(msg3)) fail(`a redo that fails on ${where} should say nothing was added: "${msg3}"`);
      if (shape() !== shapeAfterUndo || g.body.children.some((c) => c.type === "TABLE")) fail(`a redo that fails on ${where} left restored elements behind: ${shape()}`);
      if (g.doc.getNamedRanges("offizin-op-f2").length || g.doc.getNamedRanges("offizin-anchor-f2").length !== 1) fail(`a redo that fails on ${where} should keep the anchor and name nothing`);
    }
    g.context.ozRedoInsert("f2", blocks, 468, undone.index, undone.content);
    if (textOf(g.body) !== withPart) fail("redo after a failed redo did not put the part back");
    check(g.body, "failed part");
    ok("parts: a part that fails half way leaves nothing behind, and redo can be tried again");
  }

  // 13b. Opening the sidebar asks for both permissions again after a partial grant, where Google offers that.
  {
    const g = makeGoogle(CODE);
    const seen = [];
    g.context.HtmlService = { createHtmlOutputFromFile: (name) => ({ setTitle: (title) => ({ name, title }) }) };
    g.context.DocumentApp.getUi = () => ({ showSidebar: (html) => seen.push(`sidebar ${html.name}`) });
    g.context.ozShowSidebar();
    if (seen.join() !== "sidebar Sidebar") fail("the sidebar should open where ScriptApp is not available");
    seen.length = 0;
    g.context.ScriptApp = { AuthMode: { FULL: "FULL" } };
    g.context.ozShowSidebar();
    if (seen.join() !== "sidebar Sidebar") fail("the sidebar should open on a runtime without requireAllScopes");
    seen.length = 0;
    g.context.ScriptApp.requireAllScopes = (mode) => seen.push(`scopes ${mode}`);
    g.context.ozShowSidebar();
    if (seen.join() !== "scopes FULL,sidebar Sidebar") fail(`the sidebar should check every permission first: ${seen.join()}`);
    seen.length = 0;
    g.context.ScriptApp.requireAllScopes = () => {
      throw new Error("Authorization is required to perform that action.");
    };
    const stopped = threw(() => g.context.ozShowSidebar());
    if (!/Authorization is required/.test(stopped) || seen.length) fail("a partial grant should end the run before the sidebar opens, so Google can ask again");
    const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, "google-docs/appsscript.json"), "utf8"));
    if (manifest.oauthScopes.map((s) => s.split("/").pop()).sort().join() !== "documents.currentonly,script.container.ui") fail("Offizin should ask for exactly its two permissions");
    ok("opening the sidebar checks that both permissions were granted");
  }

  // 14. The sidebar host: every sample is searched, and a failed create is explained (checker-1, panel-2, contract d).
  {
    const saved = { document: globalThis.document, google: globalThis.google };
    globalThis.document = { readyState: "loading", addEventListener() {}, createElement: () => ({ width: 0, height: 0, getContext: () => new Proxy({}, { get: () => () => {} }), toDataURL: () => "data:image/png;base64,iVBORw0KGgo=" }) };
    const g = makeGoogle(CODE);
    const calls = [];
    const runner = (ok = () => {}, bad = () => {}) =>
      new Proxy({}, {
        get(t, name) {
          if (name === "withSuccessHandler") return (f) => runner(f, bad);
          if (name === "withFailureHandler") return (f) => runner(ok, f);
          return (...args) => {
            calls.push({ name, args });
            setTimeout(() => {
              try {
                const result = g.context[name](...JSON.parse(JSON.stringify(args)));
                ok(result === undefined ? null : JSON.parse(JSON.stringify(result)));
              } catch (e) {
                bad(e);
              }
            }, 0);
          };
        },
      });
    globalThis.google = { script: { run: runner() } };
    try {
      const { createGoogleHost } = await load("docs/app/hosts/google.js");
      const host = createGoogleHost();
      if (!/typed in it/.test(host.capabilities.historyNote) || /your own typing/.test(host.capabilities.historyNote)) fail("the history note should say what Undo does with typing in a part");
      fillDocument(g);
      const samples = [];
      for (let i = 0; i < 300; i++) samples.push(`Sample ${i}`);
      samples.push("Annual report");
      const found = await host.check({ samples, infoSamples: [{ key: "x", label: "Company", sample: "Visit our site" }] });
      const sent = calls.find((c) => c.name === "ozCheck");
      if (!sent || sent.args[0].length !== samples.length) fail("the host should send every sample, not a slice");
      if (!found.some((f) => f.kind === "sample" && f.text === "Annual report") || !found.some((f) => f.kind === "info")) fail("the checker should find the last sample and the business sample");
      const m = compose(findTemplate("event-flyer"), { paper: "letter" }, "google");
      const before = textOf(g.body);
      failOnce(g.body, "insertParagraph", 1);
      let msg = "";
      try {
        await host.createPublication({ ...m, blocks: withImages(m.blocks, walkBlocks) });
      } catch (e) {
        msg = e.message;
      }
      if (!/back as it was/.test(msg) || textOf(g.body) !== before) fail(`the host should report a failed create plainly: "${msg}"`);
      // When the script was stopped before it could put the document back, the host does it.
      failOnce(g.body, "insertParagraph", 1);
      const realRestore = g.context.ozRestoreSnapshot;
      let restores = 0;
      g.context.ozRestoreSnapshot = function (...args) {
        if (restores++ === 0) throw new Error("Exceeded maximum execution time");
        return realRestore.apply(this, args);
      };
      msg = "";
      try {
        await host.createPublication({ ...m, blocks: withImages(m.blocks, walkBlocks) });
      } catch (e) {
        msg = e.message;
      }
      if (!/back as it was/.test(msg) || textOf(g.body) !== before) fail(`the host should recover a stopped create: "${msg}"`);
      if (!calls.some((c) => c.name === "ozRecoverCreate")) fail("the host should ask Code.gs for the pending record");
      // Redo of a publication is a create too, so a stopped script is recovered the same way.
      const made = await host.createPublication({ ...m, blocks: withImages(m.blocks, walkBlocks) });
      await host.undo(made.op);
      if (textOf(g.body) !== before) fail("undo through the host did not bring the document back");
      const recoveries = () => calls.filter((c) => c.name === "ozRecoverCreate").length;
      const recoveredBefore = recoveries();
      failOnce(g.body, "insertParagraph", 1);
      restores = 0;
      msg = "";
      try {
        await host.redo(made.op);
      } catch (e) {
        msg = e.message;
      }
      if (!/back as it was/.test(msg) || textOf(g.body) !== before) fail(`the host should recover a stopped redo of a publication: "${msg}"`);
      if (recoveries() !== recoveredBefore + 1) fail("a stopped redo should ask Code.gs for the pending record");
      if (g.docProps["offizin.pending"]) fail("a recovered redo should leave no pending record");
      const redone = await host.redo(made.op);
      if (!redone.op.data.snapshot || textOf(g.body) === before) fail("redo after a recovered redo did not build the publication");
      await host.undo(redone.op);
      if (textOf(g.body) !== before) fail("undo after a recovered redo did not bring the document back");
      // Undo of a part hands its content to Redo.
      const part = PARTS.find((p) => p.id === "coupon");
      const composed = composePart(part, {}, "google", 468);
      const ins = await host.insertPart(composed, 468, part.id, { label: part.name });
      g.body.findText("Save 20%").getElement().para.setText("Save 50% this week only");
      const undone = await host.undo(ins.op);
      if (!undone.op.data.content || !undone.op.data.content.length) fail("undo should keep the part's content in the op");
      await host.redo(undone.op);
      if (!/Save 50% this week only/.test(textOf(g.body)) || /Save 20%/.test(textOf(g.body))) fail("redo through the host did not restore the typing");
      ok("the Google host searches every sample, explains a failed create or redo, and keeps typing through undo and redo");
    } finally {
      globalThis.document = saved.document;
      globalThis.google = saved.google;
    }
  }
}
