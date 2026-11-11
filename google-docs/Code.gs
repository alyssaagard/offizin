/**
 * Offizin for Google Docs
 * Desktop publishing in Google Docs: brochures, newsletters, flyers, cards, and more.
 *
 * This file runs in Google Apps Script, attached to your document. It adds an Offizin menu,
 * opens the sidebar, and builds publications that the sidebar designs. It only works with
 * the document it is attached to, and it never sends your content anywhere.
 *
 * Author: Alyssa I. Agard
 * License: MIT
 * Website: https://offizin.alyssaagard.dev
 */

var OZ_VERSION = '2.0.0';
var OZ_SITE = 'https://offizin.alyssaagard.dev';
var OZ_PLACEHOLDER_ALT = 'Offizin picture placeholder';

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

/** Replaces the document with a publication designed in the sidebar. */
function ozCreate(model) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  body.clear();
  ozApplyPage(body, model.page);
  var sink = ozSink(body, 0);
  ozRenderBlocks(sink, model.blocks, ozContentWidthOf(model.page));
  // body.clear() leaves one empty paragraph at the end. Keep it tiny so it never adds a page.
  ozMakeTiny(body.getChild(body.getNumChildren() - 1));
  ozSavePublication({
    templateId: model.templateId,
    theme: model.theme,
    samples: (model.samples || []).slice(0, 80),
    page: model.page,
    createdAt: new Date().toISOString()
  });
  return model.name + ' created.';
}

/** Inserts page parts after the paragraph or table that holds the cursor. */
function ozInsert(blocks, width) {
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
  ozRenderBlocks(sink, blocks, width || ozContentWidth());
  return 'Inserted.';
}

function ozContentWidth() {
  var body = DocumentApp.getActiveDocument().getBody();
  return body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
}

function ozContentWidthOf(page) {
  return page.width - page.margins.left - page.margins.right;
}

function ozApplyPage(body, page) {
  body.setPageWidth(page.width);
  body.setPageHeight(page.height);
  body.setMarginTop(page.margins.top);
  body.setMarginBottom(page.margins.bottom);
  body.setMarginLeft(page.margins.left);
  body.setMarginRight(page.margins.right);
}

function ozPageSetup(setup) {
  ozApplyPage(DocumentApp.getActiveDocument().getBody(), setup);
  return 'Page setup applied.';
}

// ---------------------------------------------------------------------------
// Rendering the layout model with DocumentApp
// ---------------------------------------------------------------------------

/** Inserts elements into a body or table cell, one after another, starting at `index`. */
function ozSink(container, index) {
  return {
    container: container,
    index: index,
    lastList: null,
    paragraph: function (text) {
      this.lastList = null;
      return this.container.insertParagraph(this.index++, text);
    },
    listItem: function (text) {
      var item = this.container.insertListItem(this.index++, text);
      if (this.lastList) item.setListId(this.lastList);
      this.lastList = item;
      return item;
    },
    table: function (cells) {
      this.lastList = null;
      return this.container.insertTable(this.index++, cells);
    },
    image: function (blob) {
      this.lastList = null;
      return this.container.insertImage(this.index++, blob);
    },
    pageBreak: function () {
      this.lastList = null;
      if (this.container.getType() !== DocumentApp.ElementType.BODY_SECTION) return null;
      return this.container.insertPageBreak(this.index++);
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
    else if (b.t === 'grid') {
      ozGrid(sink, b, width);
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
  if (b.bullet) {
    p.setGlyphType(DocumentApp.GlyphType.BULLET);
    p.setIndentFirstLine(0);
    p.setIndentStart(b.bullet.indent);
  } else {
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
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
  img.setWidth(b.width);
  img.setHeight(b.height);
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
    var inner = ozSink(wrap.getCell(0, middle), 0);
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
      var inner2 = ozSink(cell, 0);
      ozRenderBlocks(inner2, spec.blocks, Math.max(12, b.cols[k] - spec.pad[1] - spec.pad[3]));
      ozTrimCell(cell);
    }
  }
  return table;
}

/** Removes the empty paragraph a new cell starts with, when other content was added before it. */
function ozTrimCell(cell) {
  var n = cell.getNumChildren();
  var last = cell.getChild(n - 1);
  if (n < 2) {
    ozMakeTiny(last);
    return;
  }
  var before = cell.getChild(n - 2).getType();
  var lastIsEmpty = false;
  if (last.getType() === DocumentApp.ElementType.PARAGRAPH) {
    var para = last.asParagraph();
    var kids = para.getNumChildren();
    lastIsEmpty = para.getText() === '' && (kids === 0 || (kids === 1 && para.getChild(0).getType() === DocumentApp.ElementType.TEXT));
  }
  if (lastIsEmpty && (before === DocumentApp.ElementType.PARAGRAPH || before === DocumentApp.ElementType.LIST_ITEM)) {
    last.removeFromParent();
  } else {
    ozMakeTiny(last);
  }
}

// ---------------------------------------------------------------------------
// Restyling: switch colors and fonts, keep the text
// ---------------------------------------------------------------------------

function ozRestyle(map, theme) {
  var colors = {};
  Object.keys(map.colors).forEach(function (k) { colors[k.toUpperCase()] = map.colors[k]; });
  var fonts = map.fonts || {};
  var body = DocumentApp.getActiveDocument().getBody();
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
  var record = ozGetPublication() || {};
  record.theme = theme;
  ozSavePublication(record);
  return 'Design applied. Pictures keep their colors. Use Undo if you want the old look back.';
}

function ozWalk(el, visit) {
  visit(el);
  if (typeof el.getNumChildren === 'function') {
    for (var i = 0; i < el.getNumChildren(); i++) ozWalk(el.getChild(i), visit);
  }
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
