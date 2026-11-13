/**
 * Offizin for Google Docs
 * Desktop publishing in Google Docs: brochures, newsletters, flyers, cards, and more.
 *
 * This file runs in Google Apps Script, attached to your document. It adds an Offizin menu,
 * opens the sidebar, and builds publications that the sidebar designs. It only works with
 * the document it is attached to, and it never sends your content anywhere.
 *
 * Offizin keeps its own Undo and Redo, separate from the editor's:
 *   - page parts are tracked with named ranges ("offizin-op-<id>"); undoing one leaves a tiny
 *     anchor paragraph ("offizin-anchor-<id>") so Redo can put it back in the same place;
 *   - before a new publication replaces the document, a copy of the document is kept in this
 *     script's cache for your document for up to six hours, so Undo can bring it back; if the
 *     copy cannot be kept, nothing is changed, and if building fails, the copy is put back.
 *
 * Author: Alyssa I. Agard
 * License: MIT
 * Website: https://offizin.alyssaagard.dev
 */

var OZ_VERSION = '2.1.0';
var OZ_SITE = 'https://offizin.alyssaagard.dev';
var OZ_PLACEHOLDER_ALT = 'Offizin picture placeholder';
/** Google Docs measures inline pictures in pixels (96 per inch); Offizin works in points. */
var OZ_PX_PER_PT = 96 / 72;
var OZ_SNAPSHOT_SECONDS = 6 * 60 * 60;
/** Characters per cache value. Each value must stay under 100 KB, even for non-Latin text. */
var OZ_CHUNK = 30000;
var OZ_MAX_CHUNKS = 300;
var OZ_HEADINGS = ['NORMAL', 'TITLE', 'SUBTITLE', 'HEADING1', 'HEADING2'];
/** Named paragraph style for each Offizin role. */
var OZ_ROLE_HEADING = { display: 'TITLE', title: 'TITLE', headline: 'HEADING1', h1: 'HEADING1', h2: 'HEADING2', lead: 'SUBTITLE' };

// ---------------------------------------------------------------------------
// Menu and sidebar
// ---------------------------------------------------------------------------

function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Offizin')
    .addItem('Open Offizin', 'ozShowSidebar')
    .addSeparator()
    .addItem('About Offizin', 'ozShowAbout')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function ozShowSidebar() {
  // Google's consent screen lets people grant only some permissions, and Offizin needs both of
  // its two. When one is missing, Google ends this run here and asks again.
  if (typeof ScriptApp !== 'undefined' && typeof ScriptApp.requireAllScopes === 'function') {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  }
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Offizin');
  DocumentApp.getUi().showSidebar(html);
}

function ozShowAbout() {
  var ui = DocumentApp.getUi();
  ui.alert(
    'Offizin ' + OZ_VERSION,
    'Desktop publishing for Google Docs, by Alyssa I. Agard.\n\nHelp and updates: ' + OZ_SITE,
    ui.ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
// Saved settings
// ---------------------------------------------------------------------------

function ozGetState() {
  var raw = PropertiesService.getUserProperties().getProperty('offizin.state');
  return raw ? JSON.parse(raw) : null;
}

function ozSaveState(json) {
  if (String(json).length > 8500) {
    throw new Error('Your information sets are too large to save. Delete a set or shorten some fields.');
  }
  PropertiesService.getUserProperties().setProperty('offizin.state', json);
  return true;
}

function ozGetPublication() {
  var raw = PropertiesService.getDocumentProperties().getProperty('offizin.publication');
  return raw ? JSON.parse(raw) : null;
}

function ozSavePublication(record) {
  var json = JSON.stringify(record);
  while (json.length > 8500 && record.samples && record.samples.length) {
    record.samples.pop();
    json = JSON.stringify(record);
  }
  PropertiesService.getDocumentProperties().setProperty('offizin.publication', json);
}

// ---------------------------------------------------------------------------
// Creating and inserting
// ---------------------------------------------------------------------------

/**
 * Replaces the document with a publication designed in the sidebar.
 * Returns { message, snapshot } where snapshot lets ozUndoCreate bring the old document back.
 */
function ozCreate(model) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  // A pending record belongs to one run. One left by an earlier run that no sidebar recovered
  // goes first, so that a stop below can never bring back an old copy over newer work.
  ozClearPending();
  // The copy is taken before anything changes, and nothing changes unless the copy is complete.
  var taken = ozTakeSnapshot(doc, !!ozFooterText(model.footer));
  var snapshot = taken.snapshot;
  if (!snapshot.empty && !snapshot.chunks) {
    throw new Error('This document is too large for Offizin to keep a copy for its Undo. Make a copy first (File, then Make a copy), or create the publication in a new document.');
  }
  // While the document is being rebuilt, a small pending record lets the sidebar put it back
  // from the cached copy if Apps Script stops this script before the catch below can run.
  ozSetPending(snapshot, model.run);
  var footerNote;
  try {
    body.clear();
    ozApplyPage(body, model.page);
    ozApplyHeadingStyles(body, model.styles);
    var sink = ozSink(body, 0);
    ozRenderBlocks(sink, model.blocks, ozContentWidthOf(model.page));
    // body.clear() leaves one empty paragraph at the end. Keep it tiny so it never adds a page.
    ozMakeTiny(body.getChild(body.getNumChildren() - 1));
    footerNote = ozApplyFooter(doc, model.footer);
    ozSavePublication({
      templateId: model.templateId,
      theme: model.theme,
      samples: (model.samples || []).slice(0, 80),
      page: model.page,
      createdAt: new Date().toISOString()
    });
  } catch (e) {
    var reason = ozReason(e);
    try {
      ozRestoreSnapshot(doc, snapshot, taken.elements);
    } catch (e2) {
      // The pending record stays, so the sidebar can try again from the cached copy.
      throw new Error('Offizin could not finish the publication (' + reason + ') and could not put the document back (' + ozReason(e2) + '). Use File, then Version history, to go back.');
    }
    ozClearPending();
    ozDropCachedCopy(snapshot);
    throw new Error('Offizin could not finish the publication, so your document is back as it was. (' + reason + ')');
  }
  ozClearPending();
  var message = model.name + ' created.';
  if (footerNote) message += ' ' + footerNote;
  return { message: message, snapshot: snapshot };
}

/** The plain text of an error, for a message to the person. */
function ozReason(e) {
  var m = e && e.message ? String(e.message) : String(e || 'unknown error');
  return m.replace(/\.$/, '');
}

function ozSetPending(snapshot, run) {
  try {
    PropertiesService.getDocumentProperties().setProperty('offizin.pending', JSON.stringify({ run: run || '', snapshot: snapshot }));
  } catch (e) {
    // Too large to keep; ozCreate still puts the document back itself when building fails.
  }
}

function ozClearPending() {
  PropertiesService.getDocumentProperties().deleteProperty('offizin.pending');
}

/**
 * Called by the sidebar when ozCreate failed: puts the document back from the pending record
 * that ozCreate left behind when Apps Script stopped it. Returns { restored } so the sidebar
 * knows whether there was anything to do.
 */
function ozRecoverCreate(run) {
  var raw = PropertiesService.getDocumentProperties().getProperty('offizin.pending');
  if (!raw) return { restored: false };
  var record = JSON.parse(raw);
  var snapshot = record.snapshot || record;
  // A record written by another run is not this sidebar's to restore: its copy is older than
  // the document the person has now.
  if (run && record.run && record.run !== run) {
    ozClearPending();
    return { restored: false };
  }
  try {
    ozRestoreSnapshot(DocumentApp.getActiveDocument(), snapshot, null);
  } catch (e) {
    ozClearPending();
    throw new Error('Offizin could not put the document back (' + ozReason(e) + '). Use File, then Version history, to go back.');
  }
  ozClearPending();
  ozDropCachedCopy(snapshot);
  return { restored: true };
}

/**
 * Removes a page part that failed half way, so nothing unnamed is left behind. Returns the words
 * for the message: Docs never removes the last paragraph of a document, so a part that went in
 * at the very end leaves that one paragraph, emptied.
 */
function ozRemoveTracked(body, track) {
  var left = false;
  for (var i = (track || []).length - 1; i >= 0; i--) {
    try {
      if (track[i].getParent() && !ozRemove(body, track[i])) left = true;
    } catch (e) {
      // Already gone.
    }
  }
  return left ? 'so nothing was added except one empty line at the end of the document.' : 'so nothing was added.';
}

/** Inserts page parts after the paragraph or table that holds the cursor. */
function ozInsert(blocks, width, opId) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var index = body.getNumChildren();
  var el = null;
  var cursor = doc.getCursor();
  if (cursor) {
    el = cursor.getElement();
  } else {
    var selection = doc.getSelection();
    if (selection) {
      var ranges = selection.getRangeElements();
      if (ranges.length) el = ranges[ranges.length - 1].getElement();
    }
  }
  if (el) {
    while (el.getParent() && el.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
      el = el.getParent();
    }
    if (el.getParent()) index = body.getChildIndex(el) + 1;
  }
  var sink = ozSink(body, index);
  sink.track = [];
  try {
    ozRenderBlocks(sink, blocks, width || ozContentWidth());
  } catch (e) {
    throw new Error('Offizin could not insert that page part, ' + ozRemoveTracked(body, sink.track) + ' (' + ozReason(e) + ')');
  }
  if (opId) ozNameRange(doc, 'offizin-op-' + opId, sink.track);
  // The position lets Redo put the part back even if its anchor is gone (see ozRedoInsert).
  return { message: 'Inserted.', index: index };
}

function ozContentWidth() {
  var body = DocumentApp.getActiveDocument().getBody();
  return body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
}

function ozContentWidthOf(page) {
  return page.width - page.margins.left - page.margins.right;
}

function ozReadPage(body) {
  return {
    width: body.getPageWidth(),
    height: body.getPageHeight(),
    margins: { top: body.getMarginTop(), right: body.getMarginRight(), bottom: body.getMarginBottom(), left: body.getMarginLeft() }
  };
}

function ozApplyPage(body, page) {
  body.setPageWidth(page.width);
  body.setPageHeight(page.height);
  body.setMarginTop(page.margins.top);
  body.setMarginBottom(page.margins.bottom);
  body.setMarginLeft(page.margins.left);
  body.setMarginRight(page.margins.right);
}

/** Applies page setup and returns the previous values, for Offizin's Undo. */
function ozPageSetup(setup) {
  var body = DocumentApp.getActiveDocument().getBody();
  var before = ozReadPage(body);
  ozApplyPage(body, setup);
  return { message: 'Page setup applied.', before: before };
}

// ---------------------------------------------------------------------------
// Native paragraph styles and the footer
// ---------------------------------------------------------------------------

/** Sets the document's Normal text, Title, Subtitle, and Heading styles from the design. */
function ozApplyHeadingStyles(body, styles) {
  if (!styles) return;
  var sources = { NORMAL: 'body', TITLE: 'title', SUBTITLE: 'lead', HEADING1: 'h1', HEADING2: 'h2' };
  Object.keys(sources).forEach(function (name) {
    var s = styles[sources[name]];
    if (!s) return;
    var a = {};
    a[DocumentApp.Attribute.FONT_FAMILY] = s.font;
    a[DocumentApp.Attribute.FONT_SIZE] = s.size;
    a[DocumentApp.Attribute.FOREGROUND_COLOR] = s.color;
    a[DocumentApp.Attribute.BOLD] = !!s.bold;
    a[DocumentApp.Attribute.ITALIC] = !!s.italic;
    try {
      body.setHeadingAttributes(DocumentApp.ParagraphHeading[name], a);
    } catch (e) {
      // Older documents may not allow this; paragraphs still carry their own formatting.
    }
  });
}

/** What Docs uses for an attribute that no named style sets. */
var OZ_STYLE_DEFAULTS = { FONT_FAMILY: 'Arial', FONT_SIZE: 11, FOREGROUND_COLOR: '#000000', BOLD: false, ITALIC: false };

/**
 * The current heading style definitions, keeping only plain values so they can be saved.
 * Docs reports an attribute a style leaves unset as null. Offizin sets five of them on every
 * style it touches, so for those five the effective value is recorded (the style's own, else
 * Normal text's, else the Docs default), and Undo can write them back explicitly.
 */
function ozReadHeadingStyles(body) {
  var out = {};
  OZ_HEADINGS.forEach(function (name) {
    try {
      var attrs = body.getHeadingAttributes(DocumentApp.ParagraphHeading[name]) || {};
      var plain = {};
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v === null || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') plain[k] = v;
      });
      out[name] = plain;
    } catch (e) {
      // Skip a heading this document does not report.
    }
  });
  var normal = out.NORMAL || {};
  Object.keys(out).forEach(function (name) {
    Object.keys(OZ_STYLE_DEFAULTS).forEach(function (k) {
      var key = DocumentApp.Attribute[k];
      var v = out[name][key];
      if (v === null || v === undefined) v = normal[key];
      if (v === null || v === undefined) v = OZ_STYLE_DEFAULTS[k];
      out[name][key] = v;
    });
  });
  return out;
}

function ozRestoreHeadingStyles(body, headings) {
  Object.keys(headings || {}).forEach(function (name) {
    var attrs = {};
    Object.keys(headings[name]).forEach(function (k) {
      if (headings[name][k] !== null) attrs[k] = headings[name][k];
    });
    try {
      body.setHeadingAttributes(DocumentApp.ParagraphHeading[name], attrs);
    } catch (e) {
      // Leave this heading as it is.
    }
  });
}

/**
 * Puts the footer text in the document footer. Apps Script cannot insert page number fields, so
 * the page number part is left out and the person is told how to add page numbers in Docs.
 */
function ozApplyFooter(doc, footer) {
  if (!footer) return '';
  var text = footer.runs.map(function (r) { return r.text; }).join('');
  var hadNumbers = /\{pages?\}/.test(text);
  var clean = ozFooterText(footer);
  if (clean) {
    var f = doc.getFooter() || doc.addFooter();
    f.clear();
    var p = f.getNumChildren() ? f.getChild(0).asParagraph() : f.appendParagraph('');
    p.setText(clean);
    var a = {};
    a[DocumentApp.Attribute.FONT_FAMILY] = footer.font;
    a[DocumentApp.Attribute.FONT_SIZE] = footer.size;
    a[DocumentApp.Attribute.FOREGROUND_COLOR] = footer.color;
    p.setAttributes(a);
    p.setAlignment(OZ_ALIGN[footer.align] || OZ_ALIGN.center);
  }
  return hadNumbers ? 'To add page numbers, choose Insert, then Page numbers.' : '';
}

/** The footer text a design puts in the document footer, without the page number part. */
function ozFooterText(footer) {
  if (!footer) return '';
  return footer.runs.map(function (r) { return r.text; }).join('')
    .replace(/\s*(\|\s*)?page \{page\}( of \{pages\})?/gi, '')
    .replace(/\{pages?\}/g, '')
    .replace(/(\s*\|\s*)+$/, '')
    .trim();
}

/**
 * A copy of the footer with its paragraphs, runs and alignment, taken only when a publication
 * is about to replace it. Page numbers and other fields cannot be copied; `unsupported` says so.
 */
function ozReadFooter(doc, unsupported) {
  var f = doc.getFooter();
  if (!f) return { exists: false, elements: [] };
  var elements = [];
  var missing = [];
  for (var i = 0; i < f.getNumChildren(); i++) {
    var o = ozSerialize(f.getChild(i), missing);
    if (o) elements.push(o);
  }
  if (missing.length) unsupported.push('page numbers or other fields in the footer');
  return { exists: true, elements: elements };
}

/** Puts the footer back as it was. Nothing happens when the publication never touched it. */
function ozRestoreFooter(doc, footer) {
  if (!footer) return;
  var f = doc.getFooter();
  if (!footer.exists) {
    if (f) f.clear();
    return;
  }
  if (!f) f = doc.addFooter();
  f.clear();
  ozRestoreElements(f, footer.elements || []);
}

// ---------------------------------------------------------------------------
// Rendering the layout model with DocumentApp
// ---------------------------------------------------------------------------

/**
 * Inserts elements into a body or table cell, one after another, starting at `index`.
 * When `track` is an array, every element inserted into the body is recorded there.
 */
function ozSink(container, index) {
  return {
    container: container,
    index: index,
    lastList: null,
    track: null,
    keep: function (el) {
      if (this.track) this.track.push(el);
      return el;
    },
    paragraph: function (text) {
      this.lastList = null;
      return this.keep(this.container.insertParagraph(this.index++, text));
    },
    listItem: function (text) {
      var item = this.container.insertListItem(this.index++, text);
      if (this.lastList) item.setListId(this.lastList);
      this.lastList = item;
      return this.keep(item);
    },
    table: function (cells) {
      this.lastList = null;
      return this.keep(this.container.insertTable(this.index++, cells));
    },
    image: function (blob) {
      this.lastList = null;
      var img = this.container.insertImage(this.index++, blob);
      this.keep(img.getParent());
      return img;
    },
    pageBreak: function () {
      this.lastList = null;
      if (this.container.getType() !== DocumentApp.ElementType.BODY_SECTION) return null;
      var pb = this.container.insertPageBreak(this.index++);
      this.keep(pb.getParent());
      return pb;
    }
  };
}

function ozRenderBlocks(sink, blocks, width) {
  for (var i = 0; i < blocks.length; i++) {
    var b = blocks[i];
    if (b.t === 'p') ozParagraph(sink, b);
    else if (b.t === 'rule') ozRule(sink, b, width);
    else if (b.t === 'spacer') ozSpacer(sink, b.height);
    else if (b.t === 'image') ozImage(sink, b);
    else if (b.t === 'grid' || b.t === 'columns') {
      if (b.t === 'grid') ozGrid(sink, b, width);
      else ozColumns(sink, b, width);
      // Docs needs a paragraph between a table and whatever follows it.
      ozMakeTiny(sink.paragraph(''));
    } else if (b.t === 'break') {
      var pb = sink.pageBreak();
      if (pb) ozMakeTiny(pb.getParent());
    }
  }
}

function ozStyle(b) {
  var s = {};
  s[DocumentApp.Attribute.FONT_FAMILY] = b.font;
  s[DocumentApp.Attribute.FONT_SIZE] = b.size;
  s[DocumentApp.Attribute.FOREGROUND_COLOR] = b.color;
  s[DocumentApp.Attribute.BOLD] = !!b.bold;
  s[DocumentApp.Attribute.ITALIC] = !!b.italic;
  s[DocumentApp.Attribute.UNDERLINE] = false;
  return s;
}

var OZ_ALIGN = {
  left: DocumentApp.HorizontalAlignment.LEFT,
  center: DocumentApp.HorizontalAlignment.CENTER,
  right: DocumentApp.HorizontalAlignment.RIGHT,
  justify: DocumentApp.HorizontalAlignment.JUSTIFY
};

function ozParagraph(sink, b) {
  var text = b.runs.map(function (r) { return r.text; }).join('');
  var p = b.bullet ? sink.listItem(text) : sink.paragraph(text);
  // The named style comes first, then the exact look, so Docs' outline and styles work too.
  var heading = b.bullet ? 'NORMAL' : OZ_ROLE_HEADING[b.style] || 'NORMAL';
  p.setHeading(DocumentApp.ParagraphHeading[heading]);
  if (b.bullet) {
    p.setGlyphType(DocumentApp.GlyphType.BULLET);
    p.setIndentFirstLine(0);
    p.setIndentStart(b.bullet.indent);
  } else if (b.indent) {
    p.setIndentStart(b.indent.left);
    p.setIndentFirstLine(b.indent.left - b.indent.hanging);
  }
  p.setAttributes(ozStyle(b));
  p.setAlignment(OZ_ALIGN[b.align] || OZ_ALIGN.left);
  p.setSpacingBefore(b.before || 0);
  p.setSpacingAfter(b.after || 0);
  p.setLineSpacing(b.line || 1);
  // Bold words inside the paragraph.
  if (text.length) {
    var t = p.editAsText();
    var at = 0;
    b.runs.forEach(function (r) {
      if (r.text.length && r.bold !== undefined && r.bold !== b.bold) t.setBold(at, at + r.text.length - 1, r.bold);
      at += r.text.length;
    });
  }
  return p;
}

function ozMakeTiny(el) {
  if (!el) return;
  var type = el.getType();
  if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) return;
  var s = {};
  s[DocumentApp.Attribute.FONT_SIZE] = 1;
  el.setAttributes(s);
  el.setSpacingBefore(0);
  el.setSpacingAfter(0);
  el.setLineSpacing(1);
}

function ozSpacer(sink, height) {
  var p = sink.paragraph('');
  ozMakeTiny(p);
  p.setSpacingAfter(Math.max(0, height - 1));
}

/** A colored bar: a one-cell table with a background color. */
function ozRule(sink, b, width) {
  if (b.before) ozSpacer(sink, b.before);
  var w = Math.min(b.width, width);
  if (w < width - 2 && b.align !== 'left') {
    // Docs cannot center a narrow table, so a short centered rule is an underlined run of spaces.
    var p = sink.paragraph(new Array(Math.max(4, Math.round(w / 3))).join(' '));
    var s = {};
    s[DocumentApp.Attribute.UNDERLINE] = true;
    s[DocumentApp.Attribute.FOREGROUND_COLOR] = b.color;
    s[DocumentApp.Attribute.FONT_SIZE] = 6;
    p.setAttributes(s);
    p.setAlignment(OZ_ALIGN[b.align] || OZ_ALIGN.center);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
    p.setLineSpacing(1);
  } else {
    var t = sink.table([['']]);
    t.setBorderWidth(0);
    t.setColumnWidth(0, w);
    var cell = t.getCell(0, 0);
    cell.setBackgroundColor(b.color);
    cell.setPaddingTop(0);
    cell.setPaddingBottom(0);
    cell.setPaddingLeft(0);
    cell.setPaddingRight(0);
    t.getRow(0).setMinimumHeight(Math.max(1, b.weight));
    ozMakeTiny(cell.getChild(0));
    ozMakeTiny(sink.paragraph(''));
  }
  if (b.after) ozSpacer(sink, b.after);
}

function ozImage(sink, b) {
  if (!b.image) {
    var p = sink.paragraph(b.label || 'Picture');
    p.setAlignment(OZ_ALIGN[b.align] || OZ_ALIGN.left);
    return;
  }
  var blob = Utilities.newBlob(Utilities.base64Decode(b.image.base64), b.image.mime, 'offizin-picture.' + (/jpe?g/.test(b.image.mime) ? 'jpg' : 'png'));
  var img = sink.image(blob);
  img.setWidth(Math.round(b.width * OZ_PX_PER_PT));
  img.setHeight(Math.round(b.height * OZ_PX_PER_PT));
  img.setAltDescription(b.alt || OZ_PLACEHOLDER_ALT);
  img.setAltTitle(b.kind === 'logo' ? 'Logo' : 'Picture');
  var para = img.getParent();
  if (para && para.getType() === DocumentApp.ElementType.PARAGRAPH) {
    para.setAlignment(OZ_ALIGN[b.align] || OZ_ALIGN.left);
    para.setSpacingBefore(b.before || 0);
    para.setSpacingAfter(b.after || 0);
    para.setLineSpacing(1);
  }
}

var OZ_VALIGN = {
  top: DocumentApp.VerticalAlignment.TOP,
  middle: DocumentApp.VerticalAlignment.CENTER,
  bottom: DocumentApp.VerticalAlignment.BOTTOM
};

/** The first border found in a grid. Docs sets one border style for a whole table. */
function ozGridBorder(b) {
  for (var r = 0; r < b.rows.length; r++) {
    for (var c = 0; c < b.rows[r].cells.length; c++) {
      var bd = b.rows[r].cells[c].border;
      if (bd) {
        var sides = ['top', 'right', 'bottom', 'left'];
        for (var s = 0; s < sides.length; s++) if (bd[sides[s]]) return bd[sides[s]];
      }
    }
  }
  return null;
}

/** Flowing columns become a table, split where the sidebar balanced the column heights. */
function ozColumns(sink, b, width) {
  var starts = [0].concat(b.split || []).concat([b.blocks.length]);
  var cols = [];
  var cells = [];
  for (var i = 0; i < b.count; i++) {
    if (i) {
      cols.push(b.gap);
      cells.push({ fill: null, pad: [0, 0, 0, 0], valign: 'top', border: null, blocks: [] });
    }
    cols.push(b.colWidth);
    var from = starts[i] === undefined ? b.blocks.length : starts[i];
    var to = starts[i + 1] === undefined ? b.blocks.length : starts[i + 1];
    cells.push({ fill: null, pad: [0, 0, 0, 0], valign: 'top', border: null, blocks: b.blocks.slice(from, to) });
  }
  return ozGrid(sink, { t: 'grid', width: b.width, cols: cols, rows: [{ height: null, exact: false, cells: cells }], align: 'left' }, width);
}

function ozGrid(sink, b, width) {
  // Docs cannot center a narrow table, so it sits in the middle column of a borderless wrapper.
  if (b.align !== 'left' && b.width < width - 2) {
    var side = Math.max(0, width - b.width);
    var left = b.align === 'center' ? side / 2 : side;
    var right = side - left;
    var widths = [];
    var cells = [];
    if (left >= 1) { widths.push(left); cells.push(''); }
    var middle = cells.length;
    widths.push(b.width); cells.push('');
    if (right >= 1) { widths.push(right); cells.push(''); }
    var wrap = sink.table([cells]);
    wrap.setBorderWidth(0);
    for (var w = 0; w < widths.length; w++) {
      wrap.setColumnWidth(w, widths[w]);
      var wc = wrap.getCell(0, w);
      wc.setPaddingTop(0);
      wc.setPaddingBottom(0);
      wc.setPaddingLeft(0);
      wc.setPaddingRight(0);
      ozMakeTiny(wc.getChild(0));
    }
    var inner = ozSink(wrap.getCell(0, middle), 1);
    ozGrid(inner, Object.assign({}, b, { align: 'left' }), b.width);
    ozMakeTiny(inner.paragraph(''));
    ozTrimCell(wrap.getCell(0, middle));
    return wrap;
  }

  var data = b.rows.map(function (r) { return r.cells.map(function () { return ''; }); });
  var table = sink.table(data);
  var border = ozGridBorder(b);
  if (border) {
    table.setBorderWidth(Math.max(0.5, border.weight));
    table.setBorderColor(border.color);
  } else {
    table.setBorderWidth(0);
  }
  for (var c = 0; c < b.cols.length; c++) table.setColumnWidth(c, b.cols[c]);
  for (var r = 0; r < b.rows.length; r++) {
    var row = b.rows[r];
    var trow = table.getRow(r);
    if (row.height) trow.setMinimumHeight(row.height);
    for (var k = 0; k < row.cells.length; k++) {
      var spec = row.cells[k];
      var cell = trow.getCell(k);
      cell.setPaddingTop(spec.pad[0]);
      cell.setPaddingRight(spec.pad[1]);
      cell.setPaddingBottom(spec.pad[2]);
      cell.setPaddingLeft(spec.pad[3]);
      if (spec.fill) cell.setBackgroundColor(spec.fill);
      cell.setVerticalAlignment(OZ_VALIGN[spec.valign] || OZ_VALIGN.top);
      var inner2 = ozSink(cell, 1);
      ozRenderBlocks(inner2, spec.blocks, Math.max(12, b.cols[k] - spec.pad[1] - spec.pad[3]));
      ozTrimCell(cell);
    }
  }
  return table;
}

function ozIsEmptyParagraph(el) {
  if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) return false;
  var para = el.asParagraph();
  var kids = para.getNumChildren();
  return para.getText() === '' && (kids === 0 || (kids === 1 && para.getChild(0).getType() === DocumentApp.ElementType.TEXT));
}

/**
 * A new cell starts with one empty paragraph, and Offizin adds content after it. Once content
 * follows, that first paragraph is removed (never the last one, which Docs keeps), unless a table
 * comes next; then it stays as a tiny paragraph.
 */
function ozTrimCell(cell) {
  var n = cell.getNumChildren();
  var first = cell.getChild(0);
  if (n > 1 && ozIsEmptyParagraph(first)) {
    var nextType = cell.getChild(1).getType();
    if (nextType === DocumentApp.ElementType.PARAGRAPH || nextType === DocumentApp.ElementType.LIST_ITEM) first.removeFromParent();
    else ozMakeTiny(first);
  } else if (n === 1 && ozIsEmptyParagraph(first)) {
    ozMakeTiny(first);
  }
}

// ---------------------------------------------------------------------------
// Offizin's Undo and Redo
// ---------------------------------------------------------------------------

/** Gives a named range to the given body elements. */
function ozNameRange(doc, name, elements) {
  if (!elements || !elements.length) return;
  var builder = doc.newRange();
  elements.forEach(function (el) { builder.addElement(el); });
  doc.addNamedRange(name, builder.build());
}

/** The distinct top level body elements a named range covers, in document order. */
function ozTopLevel(body, range) {
  var seen = {};
  var out = [];
  range.getRangeElements().forEach(function (re) {
    var el = re.getElement();
    while (el && el.getParent() && el.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) el = el.getParent();
    if (!el || !el.getParent()) return;
    var i;
    try {
      i = body.getChildIndex(el);
    } catch (e) {
      return; // No longer in the document, for example after the publication was replaced.
    }
    if (!seen[i]) {
      seen[i] = true;
      out.push({ index: i, el: el });
    }
  });
  out.sort(function (a, b) { return a.index - b.index; });
  return out.map(function (x) { return x.el; });
}

/**
 * Removes an element, or empties it when Docs does not allow removing it (the last paragraph).
 * Returns whether the element is gone.
 */
function ozRemove(body, el) {
  if (body.getChildIndex(el) === body.getNumChildren() - 1 && el.getType() !== DocumentApp.ElementType.TABLE) {
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      el.asParagraph().clear();
      ozMakeTiny(el);
    }
    return false;
  }
  el.removeFromParent();
  return true;
}

/**
 * Undoes an inserted page part, leaving a tiny anchor paragraph for Redo. The part's current
 * content, with whatever the person typed in it, comes back as `content` so Redo puts back
 * exactly what was removed.
 */
function ozUndoInsert(opId) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var ranges = doc.getNamedRanges('offizin-op-' + opId);
  if (!ranges.length) throw new Error('That page part is no longer in the document.');
  var els = ozTopLevel(body, ranges[0].getRange());
  if (!els.length) throw new Error('That page part is no longer in the document.');
  var unsupported = [];
  var content = [];
  els.forEach(function (el) {
    var o = ozSerialize(el, unsupported);
    if (o) content.push(o);
  });
  var at = body.getChildIndex(els[0]);
  var anchor = body.insertParagraph(at, '');
  ozMakeTiny(anchor);
  els.forEach(function (el) { ozRemove(body, el); });
  ranges.forEach(function (r) { r.remove(); });
  ozNameRange(doc, 'offizin-anchor-' + opId, [anchor]);
  return { message: 'Removed.', index: at, content: content, unsupported: unsupported };
}

/**
 * Puts an undone page part back where its anchor is. When the anchor is gone (Undo and Redo of
 * the publication rebuild the whole document), the part goes back to its earlier position:
 * the publication is rebuilt the same way, so the position still fits. With `content` (what
 * Undo removed), that content comes back as it was; otherwise the part is built again from
 * its blocks.
 */
function ozRedoInsert(opId, blocks, width, index, content) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var anchors = doc.getNamedRanges('offizin-anchor-' + opId);
  var anchor = anchors.length ? ozTopLevel(body, anchors[0].getRange())[0] : null;
  var at;
  if (anchor) at = body.getChildIndex(anchor);
  else if (typeof index === 'number' && index >= 0) at = Math.min(index, body.getNumChildren());
  else throw new Error('Offizin cannot find where that page part was.');
  var track = [];
  try {
    if (content && content.length) {
      ozRestoreElements(body, content, at, track);
    } else {
      var sink = ozSink(body, at);
      sink.track = track;
      ozRenderBlocks(sink, blocks, width || ozContentWidth());
    }
  } catch (e) {
    // The anchor and its name stay, so Redo can be tried again.
    throw new Error('Offizin could not put that page part back, ' + ozRemoveTracked(body, track) + ' (' + ozReason(e) + ')');
  }
  anchors.forEach(function (r) { r.remove(); });
  if (anchor) ozRemove(body, anchor);
  ozNameRange(doc, 'offizin-op-' + opId, track);
  return { message: anchor ? 'Inserted again.' : 'Inserted again at its earlier position.', index: at };
}

/** Removes anchors left by undone page parts that can no longer be redone. */
function ozDiscardAnchors(opIds) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  (opIds || []).forEach(function (id) {
    doc.getNamedRanges('offizin-anchor-' + id).forEach(function (r) {
      ozTopLevel(body, r.getRange()).forEach(function (el) { ozRemove(body, el); });
      r.remove();
    });
  });
  return true;
}

/** Removes every Offizin marker. Content stays; anchors of undone parts go. */
function ozFinish() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var count = 0;
  doc.getNamedRanges().forEach(function (r) {
    var name = r.getName();
    if (name.indexOf('offizin-') !== 0) return;
    if (name.indexOf('offizin-anchor-') === 0) ozTopLevel(body, r.getRange()).forEach(function (el) { ozRemove(body, el); });
    r.remove();
    count++;
  });
  return count ? 'Finished. Offizin removed ' + count + ' invisible marker' + (count > 1 ? 's' : '') + '; your content is unchanged.' : 'Finished. There were no Offizin markers to remove.';
}

/** Brings back the document as it was before a publication replaced it. */
function ozUndoCreate(snapshot) {
  var doc = DocumentApp.getActiveDocument();
  try {
    ozRestoreSnapshot(doc, snapshot, null);
  } catch (e) {
    if (e && e.ozFinal) throw e;
    throw new Error('Offizin could not finish bringing back the document (' + ozReason(e) + '). Try Undo again, or use File, then Version history.');
  }
  // The copy is only dropped once the document is back, so a failed attempt can be tried again.
  ozDropCachedCopy(snapshot);
  return 'The document is back to how it was before the publication.';
}

function ozSnapshotKeys(snapshot) {
  var keys = [];
  for (var c = 0; c < (snapshot.chunks || 0); c++) keys.push('oz:' + snapshot.id + ':' + c);
  return keys;
}

function ozDropCachedCopy(snapshot) {
  var keys = ozSnapshotKeys(snapshot);
  var cache = keys.length ? CacheService.getDocumentCache() : null;
  if (cache) cache.removeAll(keys);
}

/** The body elements a snapshot kept in the cache. Throws a final, plain message when they are gone. */
function ozLoadSnapshotElements(snapshot) {
  var gone = function (text) {
    var err = new Error(text);
    err.ozFinal = true;
    return err;
  };
  if (!snapshot.chunks) throw gone('This document was too large for Offizin to keep a copy. Use File, then Version history, to go back.');
  var keys = ozSnapshotKeys(snapshot);
  var cache = CacheService.getDocumentCache();
  var got = cache ? cache.getAll(keys) : {};
  var json = '';
  for (var k = 0; k < keys.length; k++) {
    if (got[keys[k]] === undefined || got[keys[k]] === null) {
      throw gone('Offizin no longer has the earlier copy. Google keeps it for at most six hours and can drop it sooner. Use File, then Version history, to go back.');
    }
    json += got[keys[k]];
  }
  return JSON.parse(json);
}

/**
 * Rebuilds the document from a snapshot: page setup, named styles, the body, and the footer
 * when the publication changed it. `elements` is the serialized body when the caller still
 * holds it (a create that failed); otherwise it comes from the cache.
 */
function ozRestoreSnapshot(doc, snapshot, elements) {
  var body = doc.getBody();
  if (!elements) elements = snapshot.empty ? [] : ozLoadSnapshotElements(snapshot);
  body.clear();
  ozApplyPage(body, snapshot.page);
  ozRestoreHeadingStyles(body, snapshot.headings);
  ozRestoreElements(body, elements);
  ozRestoreFooter(doc, snapshot.footer);
}

/** Restyle: switch colors and fonts everywhere, keep the text. Returns the previous design. */
function ozRestyle(map, theme) {
  var colors = {};
  Object.keys(map.colors || {}).forEach(function (k) { colors[k.toUpperCase()] = map.colors[k]; });
  var fonts = map.fonts || {};
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  ozWalk(body, function (el) {
    var type = el.getType();
    if (type === DocumentApp.ElementType.TEXT) {
      var t = el.asText();
      var len = t.getText().length;
      if (!len) return;
      var starts = t.getTextAttributeIndices();
      for (var i = 0; i < starts.length; i++) {
        var start = starts[i];
        var end = (i + 1 < starts.length ? starts[i + 1] : len) - 1;
        if (end < start) continue;
        var fg = t.getForegroundColor(start);
        if (fg && colors[fg.toUpperCase()]) t.setForegroundColor(start, end, colors[fg.toUpperCase()]);
        var ff = t.getFontFamily(start);
        if (ff && fonts[ff]) t.setFontFamily(start, end, fonts[ff]);
      }
    } else if (type === DocumentApp.ElementType.TABLE_CELL) {
      var cell = el.asTableCell();
      var bg = cell.getBackgroundColor();
      if (bg && colors[bg.toUpperCase()]) cell.setBackgroundColor(colors[bg.toUpperCase()]);
    } else if (type === DocumentApp.ElementType.TABLE) {
      var table = el.asTable();
      var bc = table.getBorderColor();
      if (bc && colors[bc.toUpperCase()]) table.setBorderColor(colors[bc.toUpperCase()]);
    }
  });
  // The document's own style definitions follow too.
  OZ_HEADINGS.forEach(function (name) {
    try {
      var h = DocumentApp.ParagraphHeading[name];
      var attrs = body.getHeadingAttributes(h) || {};
      var change = {};
      var fg2 = attrs[DocumentApp.Attribute.FOREGROUND_COLOR];
      if (fg2 && colors[String(fg2).toUpperCase()]) change[DocumentApp.Attribute.FOREGROUND_COLOR] = colors[String(fg2).toUpperCase()];
      var ff2 = attrs[DocumentApp.Attribute.FONT_FAMILY];
      if (ff2 && fonts[ff2]) change[DocumentApp.Attribute.FONT_FAMILY] = fonts[ff2];
      if (Object.keys(change).length) body.setHeadingAttributes(h, change);
    } catch (e) {
      // Leave this heading as it is.
    }
  });
  var record = ozGetPublication() || {};
  var previous = record.theme || null;
  record.theme = theme;
  ozSavePublication(record);
  return { message: 'Design applied. Pictures keep their colors.', previous: previous };
}

function ozWalk(el, visit) {
  visit(el);
  if (typeof el.getNumChildren === 'function') {
    for (var i = 0; i < el.getNumChildren(); i++) ozWalk(el.getChild(i), visit);
  }
}

// ---------------------------------------------------------------------------
// Snapshots: a copy of the document, kept in the script cache for up to six hours
// ---------------------------------------------------------------------------

function ozIsEmptyBody(body) {
  return body.getNumChildren() === 1 && ozIsEmptyParagraph(body.getChild(0));
}

/**
 * Copies the document before a publication replaces it. Returns { snapshot, elements }: the
 * snapshot (small enough to hand to the sidebar) and the serialized body it stored in the cache.
 * A snapshot that is not `empty` and has no `chunks` could not be stored, and ozCreate stops.
 * The footer is copied only when the publication is about to change it (`withFooter`).
 */
function ozTakeSnapshot(doc, withFooter) {
  var body = doc.getBody();
  var unsupported = [];
  var snapshot = {
    id: Utilities.getUuid(),
    page: ozReadPage(body),
    headings: ozReadHeadingStyles(body),
    footer: withFooter ? ozReadFooter(doc, unsupported) : null,
    empty: false,
    chunks: 0,
    unsupported: unsupported
  };
  var elements = [];
  if (ozIsEmptyBody(body)) {
    snapshot.empty = true;
    return { snapshot: snapshot, elements: elements };
  }
  for (var i = 0; i < body.getNumChildren(); i++) {
    var o = ozSerialize(body.getChild(i), unsupported);
    if (o) elements.push(o);
  }
  var json = JSON.stringify(elements);
  var chunks = Math.ceil(json.length / OZ_CHUNK);
  var cache = CacheService.getDocumentCache();
  if (!cache || chunks > OZ_MAX_CHUNKS) return { snapshot: snapshot, elements: elements };
  var values = {};
  for (var c = 0; c < chunks; c++) values['oz:' + snapshot.id + ':' + c] = json.substr(c * OZ_CHUNK, OZ_CHUNK);
  cache.putAll(values, OZ_SNAPSHOT_SECONDS);
  snapshot.chunks = chunks;
  return { snapshot: snapshot, elements: elements };
}

var OZ_ENUM_NAMES = null;
/** The plain name of a DocumentApp enum value, such as "CENTER". */
function ozEnumName(value) {
  return value === null || value === undefined ? null : String(value);
}

function ozSerialize(el, unsupported) {
  var type = el.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) return ozSerializeParagraph(el.asParagraph(), 'p', unsupported);
  if (type === DocumentApp.ElementType.LIST_ITEM) {
    var li = el.asListItem();
    var o = ozSerializeParagraph(li, 'li', unsupported);
    o.glyph = ozEnumName(li.getGlyphType());
    o.nesting = li.getNestingLevel();
    o.listId = li.getListId();
    return o;
  }
  if (type === DocumentApp.ElementType.TABLE) return ozSerializeTable(el.asTable(), unsupported);
  unsupported.push(String(type).toLowerCase().replace(/_/g, ' '));
  return null;
}

function ozSerializeParagraph(p, kind, unsupported) {
  var o = {
    t: kind,
    heading: ozEnumName(p.getHeading()),
    align: ozEnumName(p.getAlignment()),
    indentStart: p.getIndentStart(),
    indentFirstLine: p.getIndentFirstLine(),
    indentEnd: p.getIndentEnd(),
    before: p.getSpacingBefore(),
    after: p.getSpacingAfter(),
    line: p.getLineSpacing(),
    parts: []
  };
  for (var i = 0; i < p.getNumChildren(); i++) {
    var c = p.getChild(i);
    var ct = c.getType();
    if (ct === DocumentApp.ElementType.TEXT) o.parts.push(ozSerializeText(c.asText()));
    else if (ct === DocumentApp.ElementType.INLINE_IMAGE) o.parts.push(ozSerializeImage(c.asInlineImage()));
    else if (ct === DocumentApp.ElementType.PAGE_BREAK) o.parts.push({ t: 'pb' });
    else if (ct === DocumentApp.ElementType.HORIZONTAL_RULE) o.parts.push({ t: 'hr' });
    else unsupported.push(String(ct).toLowerCase().replace(/_/g, ' '));
  }
  return o;
}

function ozSerializeText(text) {
  var s = text.getText();
  var idx = text.getTextAttributeIndices();
  var runs = [];
  for (var i = 0; i < idx.length; i++) {
    var start = idx[i];
    var end = (i + 1 < idx.length ? idx[i + 1] : s.length) - 1;
    if (end < start) continue;
    runs.push({
      s: start,
      e: end,
      ff: text.getFontFamily(start),
      fs: text.getFontSize(start),
      b: text.isBold(start),
      i: text.isItalic(start),
      u: text.isUnderline(start),
      st: text.isStrikethrough(start),
      fg: text.getForegroundColor(start),
      bg: text.getBackgroundColor(start),
      ln: text.getLinkUrl(start)
    });
  }
  return { t: 'text', text: s, runs: runs };
}

function ozSerializeImage(img) {
  var blob = img.getBlob();
  return {
    t: 'img',
    data: Utilities.base64Encode(blob.getBytes()),
    mime: blob.getContentType(),
    w: img.getWidth(),
    h: img.getHeight(),
    alt: img.getAltDescription(),
    title: img.getAltTitle()
  };
}

function ozSerializeTable(table, unsupported) {
  var o = { t: 'table', bw: table.getBorderWidth(), bc: table.getBorderColor(), rows: [] };
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    var cells = [];
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      var children = [];
      for (var k = 0; k < cell.getNumChildren(); k++) {
        var child = ozSerialize(cell.getChild(k), unsupported);
        if (child) children.push(child);
      }
      cells.push({
        bg: cell.getBackgroundColor(),
        pad: [cell.getPaddingTop(), cell.getPaddingRight(), cell.getPaddingBottom(), cell.getPaddingLeft()],
        va: ozEnumName(cell.getVerticalAlignment()),
        w: cell.getWidth(),
        children: children
      });
    }
    o.rows.push({ h: row.getMinimumHeight(), cells: cells });
  }
  return o;
}

/**
 * Puts serialized elements into a container: at the end (no `at`), or in place from child
 * index `at`, which leaves the container's other children alone. Returns the elements added.
 * A caller that passes its own `added` array finds every element in it from the moment the
 * element is in the container, so it knows what to remove when a later step throws.
 */
function ozRestoreElements(container, elements, at, added) {
  var startCount = container.getNumChildren();
  var inPlace = typeof at === 'number';
  var pos = inPlace ? at : null;
  var lists = {};
  added = added || [];
  elements.forEach(function (o) {
    var el = null;
    if (o.t === 'p' || o.t === 'li') el = ozRestoreParagraph(container, o, lists, pos, added);
    else if (o.t === 'table') el = ozRestoreTable(container, o, pos, added);
    if (el && inPlace) pos++;
  });
  // Remove the empty paragraph a cleared body or a new cell starts with, when a paragraph follows.
  if (!inPlace && elements.length && startCount === 1 && container.getNumChildren() > 1 && ozIsEmptyParagraph(container.getChild(0))) {
    var nextType = container.getChild(1).getType();
    if (nextType === DocumentApp.ElementType.PARAGRAPH || nextType === DocumentApp.ElementType.LIST_ITEM) container.getChild(0).removeFromParent();
    else ozMakeTiny(container.getChild(0));
  }
  return added;
}

function ozRestoreParagraph(container, o, lists, at, added) {
  var p;
  if (typeof at === 'number') p = o.t === 'li' ? container.insertListItem(at, '') : container.insertParagraph(at, '');
  else p = o.t === 'li' ? container.appendListItem('') : container.appendParagraph('');
  // Recorded before anything is set on it (see ozRestoreElements).
  added.push(p);
  if (o.heading && DocumentApp.ParagraphHeading[o.heading]) p.setHeading(DocumentApp.ParagraphHeading[o.heading]);
  if (o.align && DocumentApp.HorizontalAlignment[o.align]) p.setAlignment(DocumentApp.HorizontalAlignment[o.align]);
  if (o.indentStart !== null && o.indentStart !== undefined) p.setIndentStart(o.indentStart);
  if (o.indentFirstLine !== null && o.indentFirstLine !== undefined) p.setIndentFirstLine(o.indentFirstLine);
  if (o.indentEnd !== null && o.indentEnd !== undefined) p.setIndentEnd(o.indentEnd);
  if (o.before !== null && o.before !== undefined) p.setSpacingBefore(o.before);
  if (o.after !== null && o.after !== undefined) p.setSpacingAfter(o.after);
  if (o.line) p.setLineSpacing(o.line);
  if (o.t === 'li') {
    if (o.glyph && DocumentApp.GlyphType[o.glyph]) p.setGlyphType(DocumentApp.GlyphType[o.glyph]);
    p.setNestingLevel(o.nesting || 0);
    if (lists[o.listId]) p.setListId(lists[o.listId]);
    else lists[o.listId] = p;
  }
  o.parts.forEach(function (part) {
    if (part.t === 'text') {
      if (!part.text) return;
      var t = p.appendText(part.text);
      part.runs.forEach(function (r) {
        var a = {};
        if (r.ff) a[DocumentApp.Attribute.FONT_FAMILY] = r.ff;
        if (r.fs) a[DocumentApp.Attribute.FONT_SIZE] = r.fs;
        if (r.b !== null) a[DocumentApp.Attribute.BOLD] = !!r.b;
        if (r.i !== null) a[DocumentApp.Attribute.ITALIC] = !!r.i;
        if (r.u !== null) a[DocumentApp.Attribute.UNDERLINE] = !!r.u;
        if (r.st !== null) a[DocumentApp.Attribute.STRIKETHROUGH] = !!r.st;
        if (r.fg) a[DocumentApp.Attribute.FOREGROUND_COLOR] = r.fg;
        if (r.bg) a[DocumentApp.Attribute.BACKGROUND_COLOR] = r.bg;
        t.setAttributes(r.s, r.e, a);
        if (r.ln) t.setLinkUrl(r.s, r.e, r.ln);
      });
    } else if (part.t === 'img') {
      var img = p.appendInlineImage(Utilities.newBlob(Utilities.base64Decode(part.data), part.mime, 'picture'));
      if (part.w) img.setWidth(part.w);
      if (part.h) img.setHeight(part.h);
      if (part.alt) img.setAltDescription(part.alt);
      if (part.title) img.setAltTitle(part.title);
    } else if (part.t === 'pb') {
      if (container.getType() === DocumentApp.ElementType.BODY_SECTION) p.appendPageBreak();
    } else if (part.t === 'hr') {
      p.appendHorizontalRule();
    }
  });
  return p;
}

function ozRestoreTable(container, o, at, added) {
  var cols = 0;
  o.rows.forEach(function (r) { cols = Math.max(cols, r.cells.length); });
  if (!cols || !o.rows.length) return null;
  var matrix = o.rows.map(function () {
    var row = [];
    for (var i = 0; i < cols; i++) row.push('');
    return row;
  });
  var table = typeof at === 'number' ? container.insertTable(at, matrix) : container.appendTable(matrix);
  // Recorded before anything is set on it (see ozRestoreElements).
  added.push(table);
  if (o.bw !== null && o.bw !== undefined) table.setBorderWidth(o.bw);
  if (o.bc) table.setBorderColor(o.bc);
  o.rows.forEach(function (r, ri) {
    var row = table.getRow(ri);
    if (r.h) row.setMinimumHeight(r.h);
    r.cells.forEach(function (c, ci) {
      var cell = row.getCell(ci);
      if (c.bg) cell.setBackgroundColor(c.bg);
      if (c.pad) {
        cell.setPaddingTop(c.pad[0]);
        cell.setPaddingRight(c.pad[1]);
        cell.setPaddingBottom(c.pad[2]);
        cell.setPaddingLeft(c.pad[3]);
      }
      if (c.va && DocumentApp.VerticalAlignment[c.va]) cell.setVerticalAlignment(DocumentApp.VerticalAlignment[c.va]);
      if (ri === 0 && c.w) table.setColumnWidth(ci, c.w);
      ozRestoreElements(cell, c.children || []);
    });
  });
  return table;
}

// ---------------------------------------------------------------------------
// Design checker
// ---------------------------------------------------------------------------

function ozPattern(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function ozPlaceholders(body) {
  return body.getImages().filter(function (img) { return img.getAltDescription() === OZ_PLACEHOLDER_ALT; });
}

function ozCheck(samples, infoSamples) {
  var body = DocumentApp.getActiveDocument().getBody();
  var findings = [];
  (samples || []).forEach(function (s) {
    if (body.findText(ozPattern(s))) findings.push({ kind: 'sample', text: s, search: s });
  });
  var pics = ozPlaceholders(body).length;
  if (pics) {
    findings.push({ kind: 'picture', text: pics + ' picture placeholder' + (pics > 1 ? 's' : '') + '. Select one and choose Replace image.', picture: true });
  }
  (infoSamples || []).forEach(function (f) {
    if (body.findText(ozPattern(f.sample))) findings.push({ kind: 'info', text: f.label + ' still shows the sample “' + f.sample + '”.', search: f.sample });
  });
  return findings;
}

function ozSelect(finding) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var range;
  if (finding.picture) {
    var pics = ozPlaceholders(body);
    if (!pics.length) throw new Error('No picture placeholders are left.');
    range = doc.newRange().addElement(pics[0]).build();
  } else {
    var found = body.findText(ozPattern(finding.search || finding.text));
    if (!found) throw new Error('That text is no longer in the document.');
    range = doc.newRange().addElement(found.getElement().asText(), found.getStartOffset(), found.getEndOffsetInclusive()).build();
  }
  doc.setSelection(range);
  return true;
}
