/**
 * More publications for Offizin 2.1: a booklet, name badges, tickets, gift certificates,
 * bookmarks, place cards, envelopes, an invoice, and a photo page.
 *
 * Every publication is Letter and A4 aware. Sheet layouts (badges, tickets, cards) keep their
 * pieces the same size on both papers and center them with sheetMargins, so the cut guides
 * line up with standard stock. Sizes are in points (72 points = 1 inch).
 *
 * Mail merge templates print one sheet (or one envelope) per group of records. When nobody has
 * pasted a list, they fall back to friendly sample records so the preview is never empty.
 */
import { mm, round, MONTHS } from "../util.js";
import { p, list, rule, space, pic, logoBox, band, grid, cell, table, pageBreak, brandMark, contactLines, sheetMargins, CUT_GUIDE, chunk } from "../dsl.js";
import { alt, altShort, altQuote } from "../filler.js";

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

const CARD_STOCK = "Card stock, 80 to 110 lb cover (216 to 300 g/m²).";
const FLAT_SIZES = ["letter", "legal", "tabloid", "halfLetter", "a3", "a4", "a5"];

/** On-screen guides for sheets that are cut apart. */
const cutFolds = () => ({ vertical: [], horizontal: [], cut: true });

/** A whole number from an option, kept between lo and hi. */
function whole(value, lo, hi, fallback) {
  const n = Math.round(Number(value));
  return Number.isFinite(n) ? Math.min(hi, Math.max(lo, n)) : fallback;
}

/** "No. 0001" style numbers. */
const serial = (n) => "No. " + String(n).padStart(4, "0");

/**
 * The text of a text option, falling back to the default when it is blank.
 * Text the user typed is marked fixed, so the design checker only flags untouched samples.
 */
function textOption(ctx, key, fallback) {
  const raw = ctx.opt[key];
  const text = raw == null || String(raw).trim() === "" ? fallback : String(raw).trim();
  return { text, fixed: text !== fallback };
}

/** Splits items into sheets of `per`, with a page break between sheets. */
function sheets(items, per, makeSheet) {
  const out = [];
  chunk(items, per).forEach((group, i) => {
    if (i) out.push(pageBreak());
    out.push(makeSheet(group));
  });
  return out;
}

/** An empty, borderless cell for the unused places on the last sheet. */
const blank = () => cell([], { pad: 0 });

/** A money amount with two decimals and thousands separators. */
function money(symbol, n) {
  const [units, cents] = (Math.round(n * 100) / 100).toFixed(2).split(".");
  return `${symbol}${units.replace(/\B(?=(\d{3})+(?!\d))/g, ",")}.${cents}`;
}

/** A date written out in full, such as "March 4, 2027". */
function longDate(d) {
  return `${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
}

/** A label and a writing line, for "To", "From", and "Expires" on certificates. */
function fillLine(label, width, o = {}) {
  return grid(
    [0.3, 0.7],
    [
      [
        cell([p(label, "label", { after: 0, color: o.labelColor || "muted", fixed: true })], { pad: [0, 6, 2, 0], valign: "bottom" }),
        cell(o.value ? [p(o.value, "small", { after: 0, fixed: o.fixed !== false })] : [], { pad: [0, 0, 2, 0], valign: "bottom", border: { bottom: { color: "dark", weight: 0.5 } } }),
      ],
    ],
    { width, align: "left" }
  );
}

// ---------------------------------------------------------------------------
// Booklet
// ---------------------------------------------------------------------------

/** The page plan for each booklet length: inner pages between the front and back covers. */
const BOOKLET_PLANS = {
  4: ["welcome", "schedule0"],
  8: ["welcome", "contents", "article0", "schedule0", "photos0", "notes"],
  12: ["welcome", "contents", "article0", "article1", "profiles0", "schedule0", "schedule1", "photos0", "article2", "notes"],
  16: ["welcome", "contents", "article0", "article1", "profiles0", "schedule0", "schedule1", "article2", "photos0", "profiles1", "article3", "photos1", "sponsors", "notes"],
};

// Running text is Middle High German filler (filler.js), so it is plainly meant to be replaced.
const ARTICLES = [
  {
    section: "Our story",
    kicker: "Feature",
    title: "How it began",
    // Indexes are chosen so no filler line repeats on one printed side of a 16-page booklet.
    caption: altShort(10),
    lead: alt(1, 9),
    body: [alt(2, 10), alt(1, 12)],
    quote: `“${altQuote(3)}”`,
  },
  {
    section: "This year",
    kicker: "What is new",
    title: "What to look for",
    caption: altShort(5),
    lead: alt(1, 4),
    body: [`**A new stage.** ${alt(1, 5)}`, `**Longer hours.** ${alt(1, 7)}`, `**Family corner.** ${alt(1, 8)}`],
    quote: `“${altQuote(2)}”`,
  },
  {
    section: "Behind the scenes",
    kicker: "People",
    title: "Our volunteers",
    caption: altShort(6),
    lead: alt(1, 13),
    body: [alt(2, 14), alt(1, 3)],
    quote: `“${altQuote(0)}”`,
  },
  {
    section: "Around town",
    kicker: "Guide",
    title: "Eat, park, and rest",
    caption: altShort(9),
    lead: alt(1, 11),
    body: [`**Parking.** ${alt(1, 10)}`, `**Food.** ${alt(1, 14)}`, `**Quiet room.** ${alt(1, 12)}`],
    quote: `“${altQuote(1)}”`,
  },
];

const SCHEDULES = [
  {
    section: "Saturday",
    title: "Saturday schedule",
    rows: [
      ["9:00", "Doors open", "Main lobby"],
      ["9:30", "Welcome", "Main hall"],
      ["10:15", "Session title", "Rooms 1 to 3"],
      ["12:00", "Lunch", "Courtyard"],
      ["1:30", "Session title", "Main hall"],
      ["3:00", "Session title", "Second floor"],
      ["5:00", "Reception", "Garden terrace"],
    ],
    note: alt(1, 14),
  },
  {
    section: "Sunday",
    title: "Sunday schedule",
    rows: [
      ["10:00", "Session title", "Courtyard"],
      ["11:00", "Session title", "Main hall"],
      ["12:30", "Lunch", "Riverside lawn"],
      ["2:00", "Session title", "Main hall"],
      ["4:00", "Closing", "Main hall"],
    ],
    note: alt(1, 10),
  },
];

/** The running head at the top of an inside page: the page number on the outside edge. */
function runningHead(index, section) {
  const n = index + 1;
  const left = n % 2 === 0;
  return [
    p(left ? `${n}   |   ${section}` : `${section}   |   ${n}`, "label", { align: left ? "left" : "right", color: "muted", after: 4, fixed: true }),
    rule({ color: "line", weight: 0.5, before: 0, after: 14 }),
  ];
}

function bookletCover(ctx) {
  const h = ctx.H - 10;
  return [
    grid(
      [ctx.W],
      [
        {
          height: h,
          cells: [
            cell(
              [
                p("{org} presents", "kicker", { align: "center", color: "onPrimary" }),
                p("Event name", "title", { align: "center", color: "onPrimary", size: 28 }),
                rule({ color: "accent", weight: 2, width: 60, align: "center", before: 4, after: 14 }),
                pic("full", round(h * 0.4), { fill: { mix: ["primary", "white", 0.25] }, ink: "onPrimary", after: 16 }),
                p("Saturday and Sunday, May 16 and 17", "h2", { align: "center", color: "onPrimary", after: 2 }),
                p("Riverside Center, 45 River Road", "body", { align: "center", color: "onPrimary", after: 0 }),
              ],
              { pad: 22, fill: "primary", valign: "middle" }
            ),
          ],
        },
      ],
      { abs: true }
    ),
  ];
}

function bookletBack(ctx) {
  return [
    grid(
      [ctx.W],
      [
        {
          height: ctx.H - 10,
          cells: [
            cell(
              [
                p("Thank you for coming", "h1", { align: "center" }),
                p(alt(1, 13), "body", { align: "center", after: 18 }),
                brandMark(ctx, 130, 52, { align: "center" }),
                p("{tagline}", "caption", { align: "center", after: 12 }),
                ...contactLines({ align: "center" }),
                rule({ color: "accent", weight: 1.5, width: 40, align: "center", before: 14, after: 8 }),
                p("{social}", "small", { align: "center", color: "muted", after: 0 }),
              ],
              { pad: [0, 8, 6, 8], valign: "bottom" }
            ),
          ],
        },
      ],
      { abs: true }
    ),
  ];
}

const BOOKLET_PAGES = {
  welcome: {
    section: "Welcome",
    build: (ctx, i, plan) => [
      ...runningHead(i, "Welcome"),
      p("Welcome", "kicker"),
      p("Before we begin", "headline", { size: 18 }),
      pic("full", round(ctx.H * 0.2), { after: 10 }),
      p(alt(1, 2), "lead", { size: 11 }),
      p(alt(2, 0)),
      p(alt(1, 6)),
      plan.length <= 2 ? band([p("Inside", "label", { color: "primary", fixed: true }), list(["Welcome, page 2", "Schedule, page 3", "Contact details, back cover"], { role: "small", gap: 2, after: 0, fixed: true })], { fill: "tint", pad: 10, align: "left" }) : null,
      space(10),
      p("With thanks,", "body", { after: 2 }),
      p("{name}", "h2", { color: "primary", after: 0, fixed: true }),
      p("{title}, {org}", "caption", { after: 0 }),
    ],
  },
  contents: {
    section: "Contents",
    build: (ctx, i, plan, entries) => {
      // Longer booklets have more entries, so the rows tighten and the extra box steps aside.
      const long = entries.length > 8;
      const pad = long ? [4, 0, 4, 0] : [6, 0, 6, 0];
      const line = { bottom: { color: "line", weight: 0.5 } };
      return [
      ...runningHead(i, "Contents"),
      p("Inside", "kicker"),
      p("Contents", "headline", { size: 22, fixed: true, after: 12 }),
      grid(
        [1, 0.16],
        entries.map((e) => [
          cell([p(e.title, "body", { after: 0, fixed: true })], { pad, border: line }),
          cell([p(String(e.page), "h2", { align: "right", color: "primary", after: 0, fixed: true })], { pad, border: line }),
        ])
      ),
      long ? null : space(16),
      long ? null : band(
        [
          p("Good to know", "h2", { color: "primary" }),
          list(["Wi-Fi: Network name and password", "First aid: Main lobby", "Lost and found: Welcome desk"], { role: "small", gap: 2, after: 0 }),
        ],
        { fill: "tint", pad: 12, align: "left" }
      ),
      ];
    },
  },
  article: {
    build: (ctx, i, plan, entries, a) => [
      ...runningHead(i, a.section),
      p(a.kicker, "kicker"),
      p(a.title, "h1", { size: 16 }),
      pic("full", round(ctx.H * 0.22), { after: 4 }),
      p(a.caption, "caption", { after: 8 }),
      p(a.lead, "body", { italic: true }),
      ...a.body.map((t) => p(t)),
      rule({ color: "accent", weight: 1.25, width: 40, before: 4, after: 6 }),
      p(a.quote, "quote", { size: 13, after: 0 }),
    ],
  },
  schedule: {
    build: (ctx, i, plan, entries, s) => [
      ...runningHead(i, s.section),
      p("Schedule", "kicker", { fixed: true }),
      p(s.title, "headline", { size: 20, fixed: true }),
      table(["Time", "Event", "Place"], s.rows, { widths: [0.2, 0.48, 0.32], firstBold: true, firstColor: "primary" }),
      space(14),
      band(
        [
          p("Please note", "h2", { color: "primary" }),
          p(s.note, "small", { after: 0 }),
        ],
        { fill: "tint", pad: 12, align: "left" }
      ),
    ],
  },
  profiles: {
    build: (ctx, i, plan, entries, n) => [
      ...runningHead(i, n ? "Artists" : "Speakers"),
      p(n ? "Artists" : "Speakers", "kicker", { fixed: true }),
      p(n ? "Meet this year's artists" : "Meet the speakers", "headline", { size: 18, fixed: true }),
      ...[1, 2, 3].map((k) =>
        grid(
          [0.27, 0.73],
          [
            [
              cell([pic("full", 74, { after: 0, label: "Photo" })], { pad: [0, 0, 12, 0] }),
              cell(
                [
                  p("Full name", "h2", { after: 0 }),
                  p("Role or title", "caption", { after: 3 }),
                  p(alt(1, 4 * n + 3 * k), "small", { after: 0 }),
                ],
                { pad: [0, 0, 12, 0] }
              ),
            ],
          ],
          { gap: 12 }
        )
      ),
    ],
  },
  photos: {
    build: (ctx, i, plan, entries, n) => {
      // Deep top and bottom margins can leave very little height, so the pictures never go below 36 pt.
      const h = Math.max(36, round((ctx.H - 150) / 2));
      const shot = (text) => cell([pic("full", h, { after: 3 }), p(text, "caption", { after: 0 })], { pad: [0, 0, 10, 0] });
      return [
        ...runningHead(i, "Snapshots"),
        p(n ? "More moments" : "Snapshots", "kicker", { fixed: true }),
        p(n ? "Moments from the weekend" : "Highlights from last year", "headline", { size: 18 }),
        grid([1, 1], [[shot(altShort(4 * n)), shot(altShort(4 * n + 1))], [shot(altShort(4 * n + 2)), shot(altShort(4 * n + 3))]], { gap: 10 }),
      ];
    },
  },
  sponsors: {
    section: "Sponsors",
    build: (ctx, i) => {
      // A picture placeholder, not a logo box: the composer would fill a logo box with the organization's own logo.
      const logo = () => cell([pic("full", 46, { label: "Sponsor logo", after: 0 })], { pad: [0, 0, 10, 0] });
      return [
        ...runningHead(i, "Sponsors"),
        p("With thanks", "kicker", { fixed: true }),
        p("Our sponsors and partners", "headline", { size: 18 }),
        p(alt(1, 6), "body", { after: 12 }),
        grid([1, 1], [[logo(), logo()], [logo(), logo()], [logo(), logo()]], { gap: 12 }),
        p("Sponsor name, Sponsor name, Sponsor name, and Sponsor name", "small", { align: "center", color: "muted", before: 4 }),
      ];
    },
  },
  notes: {
    section: "Notes",
    build: (ctx, i) => {
      const lines = Math.max(4, Math.floor((ctx.H - 90) / 24));
      return [
        ...runningHead(i, "Notes"),
        p("Notes", "headline", { size: 18, fixed: true, after: 14 }),
        ...Array.from({ length: lines }, () => rule({ color: "line", weight: 0.5, before: 0, after: 23 })),
      ];
    },
  },
};

/** The section name and builder for one entry of a booklet plan, such as "article2". */
function bookletPage(key) {
  const m = /^([a-z]+)(\d*)$/.exec(key);
  const kind = m[1];
  const n = m[2] === "" ? 0 : Number(m[2]);
  const def = BOOKLET_PAGES[kind];
  if (kind === "article") return { section: ARTICLES[n].section, title: ARTICLES[n].section, build: (ctx, i, plan, e) => def.build(ctx, i, plan, e, ARTICLES[n]) };
  if (kind === "schedule") return { section: SCHEDULES[n].section, title: SCHEDULES[n].title, build: (ctx, i, plan, e) => def.build(ctx, i, plan, e, SCHEDULES[n]) };
  if (kind === "profiles") return { section: n ? "Artists" : "Speakers", title: n ? "Meet this year's artists" : "Meet the speakers", build: (ctx, i, plan, e) => def.build(ctx, i, plan, e, n) };
  if (kind === "photos") return { section: "Snapshots", title: n ? "More moments" : "Snapshots", build: (ctx, i, plan, e) => def.build(ctx, i, plan, e, n) };
  return { section: def.section, title: def.section, build: def.build };
}

// ---------------------------------------------------------------------------
// Address fields (the same fields as the address labels)
// ---------------------------------------------------------------------------

const ADDRESS_FIELDS = [
  { key: "name", label: "Name", aliases: ["full name", "recipient", "addressee"] },
  { key: "company", label: "Company", aliases: ["organization", "organisation"] },
  { key: "address", label: "Street address", aliases: ["address 1", "address line 1", "street"] },
  { key: "address2", label: "Address line 2", aliases: ["apt", "suite", "unit", "address 2"] },
  { key: "city", label: "City", aliases: ["town", "city state zip"] },
  { key: "state", label: "State or region", aliases: ["province", "region", "county"] },
  { key: "postal", label: "Postal code", aliases: ["zip", "zip code", "postcode"] },
  { key: "country", label: "Country" },
];

/** Envelope sizes: a #10 envelope for Letter regions and DL for A4 regions. */
function envelopeSpec(paper) {
  return paper === "a4"
    ? { width: round(mm(220)), height: round(mm(110)), margin: 24, toLeft: round(mm(95)), toTop: round(mm(48)) }
    : { width: 684, height: 297, margin: 27, toLeft: 288, toTop: 144 };
}

// ---------------------------------------------------------------------------
// Publications
// ---------------------------------------------------------------------------

export const MORE_TEMPLATES = [
  {
    id: "booklet",
    name: "Booklet",
    category: "programs",
    description: "A folded and stapled booklet of 4 to 16 pages with a cover, welcome note, contents, stories, schedule, photos, and notes. Offizin puts the pages in printer order for you.",
    sizeLabel: "Half letter or A5 pages, folded and stapled",
    fold: "booklet",
    options: [
      {
        key: "pages",
        label: "Number of pages",
        type: "select",
        choices: [
          { value: "4", label: "4 pages (1 sheet)" },
          { value: "8", label: "8 pages (2 sheets)" },
          { value: "12", label: "12 pages (3 sheets)" },
          { value: "16", label: "16 pages (4 sheets)" },
        ],
        default: "8",
        help: "A booklet always has a multiple of four pages, because each sheet holds four pages.",
      },
    ],
    print: {
      duplex: "short",
      paper: "Plain 24 lb (90 g/m²) paper for the inside. Use 65 to 80 lb cover (176 to 216 g/m²) for a sturdier cover.",
      finishing: "Fold the stack in half and staple twice on the fold.",
    },
    customize: { paper: false, orientation: false, margins: true, spacing: true, text: true },
    page: (paper) => ({ size: paper === "a4" ? "a5" : "halfLetter", orientation: "portrait", margins: { top: 34, right: 32, bottom: 34, left: 32 } }),
    build(ctx) {
      const count = [4, 8, 12, 16].includes(Number(ctx.opt.pages)) ? Number(ctx.opt.pages) : 8;
      const plan = BOOKLET_PLANS[count].map(bookletPage);
      // Contents lists every inside page after the contents page itself, with its real page number.
      const entries = plan.map((pg, k) => ({ title: pg.title, page: k + 2 })).filter((e, k) => !["Welcome", "Contents"].includes(plan[k].section));
      entries.push({ title: "Back cover: contact details", page: count });
      const pages = [bookletCover(ctx)];
      plan.forEach((pg, k) => pages.push(pg.build(ctx, k + 1, BOOKLET_PLANS[count], entries)));
      pages.push(bookletBack(ctx));
      return { pages };
    },
  },

  {
    id: "name-badges",
    name: "Name badges",
    category: "labels",
    description: "Six 4 x 3 inch badges per sheet, one for each person on your list, with cut guides.",
    sizeLabel: "6 per sheet, 4 x 3 in",
    fold: "cards",
    // The badges stay 4 x 3 inches on A4 to match holders and insert sheets, which puts them 3.4 mm from the side edges.
    print: (ctx) => ({
      paper: CARD_STOCK + " Badge insert sheets with six 4 x 3 inch badges also work.",
      finishing:
        "Cut on the dashed guides and slide each badge into a holder." +
        (ctx.paper === "a4" ? " On A4, the badges come within 3.4 mm of the side edges. If your printer cannot print that close to the edge, the outer cut guides may not show, so cut each badge 4 inches (101.6 mm) wide." : ""),
    }),
    customize: { paper: false, margins: false },
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 288, 216, 2, 3) }),
    folds: cutFolds,
    merge: {
      fields: [
        { key: "name", label: "Name", aliases: ["full name", "attendee", "guest"] },
        { key: "role", label: "Role", aliases: ["title", "job title", "position"] },
        { key: "organization", label: "Organization", aliases: ["company", "organisation", "affiliation"] },
      ],
      sample: [
        { name: "Avery Johnson", role: "Volunteer Coordinator", organization: "Lakeside Library" },
        { name: "Sam Okafor", role: "Guest Speaker", organization: "Northfield College" },
        { name: "Priya Raman", role: "Board Member", organization: "Riverside Arts" },
        { name: "Mateo Alvarez", role: "Workshop Leader", organization: "Studio Nine" },
        { name: "Grace Liu", role: "Event Staff", organization: "Riverside Arts" },
        { name: "Daniel Brooks", role: "Sponsor", organization: "Harbor Point Bank" },
      ],
      note: "Paste a list with a name, role, and organization for each person. Each sheet holds six badges.",
    },
    build(ctx) {
      const sample = !ctx.hasRecords;
      const strip = (r) =>
        r
          ? cell(
              [ctx.hasLogo ? logoBox(96, 24, { align: "center", after: 0, fill: "white" }) : p("{org}", "label", { color: "onPrimary", align: "center", after: 0 })],
              { fill: "primary", pad: [8, 12, 8, 12], valign: "middle", border: { top: CUT_GUIDE.all, left: CUT_GUIDE.all, right: CUT_GUIDE.all } }
            )
          : blank();
      const body = (r) =>
        r
          ? cell(
              [
                p("Hello, my name is", "label", { align: "center", color: "muted", after: 6, fixed: true }),
                p(r.name || "", "display", { align: "center", size: 26, line: 1.05, after: 6, fixed: !sample }),
                rule({ color: "accent", weight: 1.5, width: 36, align: "center", before: 0, after: 8 }),
                r.role ? p(r.role, "lead", { align: "center", size: 12, after: 2, fixed: !sample }) : null,
                r.organization ? p(r.organization, "caption", { align: "center", size: 10, after: 0, fixed: !sample }) : null,
              ],
              { pad: [12, 16, 10, 16], valign: "middle", border: { left: CUT_GUIDE.all, right: CUT_GUIDE.all, bottom: CUT_GUIDE.all } }
            )
          : blank();
      return sheets(ctx.records, 6, (group) => {
        const rows = [];
        for (let i = 0; i < 3; i++) {
          const a = group[i * 2];
          const b = group[i * 2 + 1];
          rows.push({ height: 40, exact: true, cells: [strip(a), strip(b)] });
          rows.push({ height: 176, exact: true, cells: [body(a), body(b)] });
        }
        return grid([288, 288], rows, { abs: true, align: "center" });
      });
    },
  },

  {
    id: "tickets",
    name: "Tickets",
    category: "labels",
    description: "Ten numbered tickets per sheet, each with a tear-off stub. Paste a guest list to print a named ticket for each person.",
    sizeLabel: "10 per sheet, 3.5 x 2 in",
    fold: "cards",
    print: { paper: CARD_STOCK, finishing: "Cut on the outer dashed guides. Fold along the line between each ticket and its stub, so the stub tears off cleanly." },
    customize: { paper: false, margins: false },
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 252, 144, 2, 5) }),
    folds: cutFolds,
    options: [
      { key: "event", label: "Event name", type: "text", default: "Event name" },
      { key: "date", label: "Date and time", type: "text", default: "Saturday, April 18, 7 PM" },
      { key: "place", label: "Place", type: "text", default: "Riverside Hall" },
      { key: "price", label: "Price", type: "text", default: "$25" },
      { key: "startNumber", label: "First ticket number", type: "number", default: 1, min: 1, max: 99999 },
      { key: "sheets", label: "Sheets to print", type: "number", default: 1, min: 1, max: 20, help: "Each sheet holds ten tickets. With a guest list, Offizin prints one ticket per person instead." },
    ],
    merge: {
      fields: [
        { key: "name", label: "Name", aliases: ["full name", "guest", "attendee"] },
        { key: "seat", label: "Seat or table", aliases: ["table", "seat number", "row"] },
      ],
      sample: [],
      note: "Optional. Paste a list of names to print one numbered ticket per person. Leave it empty to print blank numbered tickets.",
    },
    build(ctx) {
      const start = whole(ctx.opt.startNumber, 1, 99999, 1);
      const count = ctx.hasRecords ? ctx.records.length : whole(ctx.opt.sheets, 1, 20, 1) * 10;
      const event = textOption(ctx, "event", "Event name");
      const date = textOption(ctx, "date", "Saturday, April 18, 7 PM");
      const place = textOption(ctx, "place", "Riverside Hall");
      const price = textOption(ctx, "price", "$25");
      const tickets = Array.from({ length: count }, (_, i) => ({ n: start + i, r: ctx.hasRecords ? ctx.records[i] : null }));
      const guide = CUT_GUIDE.all;
      const perforation = { color: "muted", weight: 0.75, style: "dashed" };
      const main = (t) =>
        t
          ? cell(
              [
                p("{org} presents", "kicker", { after: 2 }),
                p(event.text, "h2", { size: 13, color: "primary", after: 4, fixed: event.fixed }),
                p(date.text, "small", { after: 0, fixed: date.fixed }),
                p(place.text, "small", { after: 4, fixed: place.fixed }),
                t.r && (t.r.name || t.r.seat) ? p([t.r.name, t.r.seat].filter(Boolean).join("  |  "), "small", { bold: true, after: 3, fixed: true }) : null,
                p(`**${price.text}**   ${serial(t.n)}`, "small", { color: "dark", after: 0, fixed: true }),
              ],
              { pad: [10, 10, 8, 12], valign: "middle", border: { top: guide, left: guide, bottom: guide, right: perforation } }
            )
          : blank();
      const stub = (t) =>
        t
          ? cell(
              [
                p("Admit one", "label", { align: "center", color: "primary", after: 6, fixed: true }),
                p(serial(t.n), "small", { align: "center", bold: true, after: 4, fixed: true }),
                p(price.text, "caption", { align: "center", after: 0, fixed: true }),
              ],
              { fill: "tint", pad: [8, 4, 8, 4], valign: "middle", border: { top: guide, right: guide, bottom: guide } }
            )
          : blank();
      return sheets(tickets, 10, (group) => {
        const rows = [];
        for (let i = 0; i < 5; i++) {
          const a = group[i * 2];
          const b = group[i * 2 + 1];
          rows.push({ height: 144, exact: true, cells: [main(a), stub(a), main(b), stub(b)] });
        }
        return grid([180, 72, 180, 72], rows, { abs: true, align: "center" });
      });
    },
  },

  {
    id: "gift-certificates",
    name: "Gift certificates",
    category: "calendars",
    description: "Three numbered gift certificates per sheet, with the amount, who it is for, and an expiry line.",
    sizeLabel: "3 per sheet, 7.5 x 3.25 in",
    fold: "cards",
    print: { paper: CARD_STOCK + " A textured or cream stock feels special.", finishing: "Cut on the dashed guides. Write the recipient, the date, and your initials before you hand one out." },
    customize: { paper: false, margins: false },
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 540, 234, 1, 3) }),
    folds: cutFolds,
    options: [
      { key: "amount", label: "Amount", type: "text", default: "$50" },
      { key: "startNumber", label: "First certificate number", type: "number", default: 1, min: 1, max: 99999 },
      { key: "sheets", label: "Sheets to print", type: "number", default: 1, min: 1, max: 20, help: "Each sheet holds three certificates, numbered in order." },
    ],
    build(ctx) {
      const amount = textOption(ctx, "amount", "$50");
      const start = whole(ctx.opt.startNumber, 1, 99999, 1);
      const count = whole(ctx.opt.sheets, 1, 20, 1) * 3;
      const guide = CUT_GUIDE.all;
      const left = (n) =>
        cell(
          [
            p("Gift certificate", "kicker", { color: "onPrimary", align: "center", after: 8, fixed: true }),
            p(amount.text, "display", { color: "onPrimary", align: "center", size: 44, after: 6, fixed: amount.fixed }),
            rule({ color: "accent", weight: 1.5, width: 40, align: "center", before: 0, after: 8 }),
            p("{web}", "small", { color: "onPrimary", align: "center", after: 0 }),
          ],
          { fill: "primary", pad: [16, 14, 16, 14], valign: "middle", border: { top: guide, left: guide, bottom: guide } }
        );
      const right = (n) =>
        cell(
          [
            grid(
              [0.6, 0.4],
              [
                [
                  cell([brandMark(ctx, 120, 36, { after: 0, size: 14 })], { valign: "middle" }),
                  cell([p(serial(n), "label", { align: "right", color: "accent", after: 0, fixed: true })], { valign: "middle" }),
                ],
              ]
            ),
            p("A gift for you", "h1", { size: 17, before: 6, after: 10 }),
            fillLine("To", 312),
            space(6),
            fillLine("From", 312),
            space(6),
            fillLine("Expires", 312),
            p("Redeem at {address}, {city}, or at {web}. Not exchangeable for cash.", "caption", { before: 8, after: 0 }),
          ],
          { pad: [14, 20, 12, 20], valign: "middle", border: { top: guide, right: guide, bottom: guide } }
        );
      const numbers = Array.from({ length: count }, (_, i) => start + i);
      return sheets(numbers, 3, (group) =>
        grid(
          [180, 360],
          group.map((n) => ({ height: 234, exact: true, cells: [left(n), right(n)] })),
          { abs: true, align: "center" }
        )
      );
    },
  },

  {
    id: "bookmarks",
    name: "Bookmarks",
    category: "cards",
    description: "Four tall bookmarks per sheet with a picture, a quotation, and your organization.",
    sizeLabel: "4 per sheet, 2 x 7 in",
    fold: "cards",
    print: { paper: CARD_STOCK, finishing: "Cut on the dashed guides. Punch a hole near the top and add a ribbon or tassel if you like." },
    customize: { paper: false, margins: false },
    // 2 x 7 inches on Letter. On A4, 50 x 178 mm (nearly the same size) keeps a printable side margin.
    page: (paper) => (paper === "a4" ? { size: "paper", orientation: "portrait", margins: sheetMargins(paper, round(mm(50)), round(mm(178)), 4, 1) } : { size: "paper", orientation: "portrait", margins: sheetMargins(paper, 144, 504, 4, 1) }),
    folds: cutFolds,
    build(ctx) {
      const a4 = ctx.paper === "a4";
      const w = a4 ? round(mm(50)) : 144;
      const h = a4 ? round(mm(178)) : 504;
      const mark = () =>
        cell(
          [
            pic("full", round(h * 0.34), { fill: { mix: ["primary", "white", 0.25] }, ink: "onPrimary", after: 16 }),
            p("Read", "kicker", { align: "center", after: 4, fixed: true }),
            p(`“${altQuote(3)}”`, "quote", { align: "center", size: 14, after: 6 }),
            p("A favorite saying", "caption", { align: "center", after: 18 }),
            rule({ color: "accent", weight: 1.5, width: 30, align: "center", before: 0, after: 12 }),
            brandMark(ctx, 90, 36, { align: "center", after: 2, size: 11 }),
            p("{web}", "small", { align: "center", color: "muted", size: 7.5, after: 0 }),
          ],
          { pad: [12, 12, 12, 12], valign: "top", border: CUT_GUIDE }
        );
      return [grid([w, w, w, w], [{ height: h, exact: true, cells: [mark(), mark(), mark(), mark()] }], { abs: true, align: "center" })];
    },
  },

  {
    id: "place-cards",
    name: "Place cards",
    category: "cards",
    description: "Four folded tent cards per sheet with each guest's name on the half that faces them.",
    sizeLabel: "4 per sheet, folds to 3.75 x 2.5 in",
    fold: "tent",
    print: { paper: CARD_STOCK, finishing: "Cut on the dashed guides, then fold each card in half across the middle with the name on the outside." },
    customize: { paper: false, margins: false },
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 270, 360, 2, 2) }),
    folds: (ctx) => {
      const top = ctx.page.margins.top;
      const h = ctx.page.height;
      return { vertical: [], horizontal: [round((top + 180) / h, 4), round((top + 540) / h, 4)], cut: true };
    },
    merge: {
      fields: [
        { key: "name", label: "Guest name", aliases: ["name", "full name", "guest"] },
        { key: "table", label: "Table", aliases: ["table number", "seat"] },
      ],
      sample: [
        { name: "Avery Johnson", table: "Table 1" },
        { name: "Sam Okafor", table: "Table 1" },
        { name: "Priya Raman", table: "Table 2" },
        { name: "Mateo Alvarez", table: "Table 2" },
      ],
      note: "Paste one guest name per line, with an optional table number. Each sheet holds four cards.",
    },
    build(ctx) {
      const sample = !ctx.hasRecords;
      const guide = CUT_GUIDE.all;
      // The top half becomes the back of the tent, so it stays plain: text there would stand upside down.
      const back = (r) => (r ? cell([], { pad: 0, border: { top: guide, left: guide, right: guide } }) : blank());
      const front = (r) =>
        r
          ? cell(
              [
                p(r.name || "", "title", { align: "center", size: 26, line: 1.05, after: 6, fixed: !sample }),
                rule({ color: "accent", weight: 1.5, width: 40, align: "center", before: 0, after: 6 }),
                r.table ? p(r.table, "label", { align: "center", color: "muted", after: 0, fixed: !sample }) : null,
              ],
              { pad: [14, 18, 14, 18], valign: "middle", border: { left: guide, right: guide, bottom: guide } }
            )
          : blank();
      return sheets(ctx.records, 4, (group) => {
        const rows = [];
        for (let i = 0; i < 2; i++) {
          const a = group[i * 2];
          const b = group[i * 2 + 1];
          rows.push({ height: 180, exact: true, cells: [back(a), back(b)] });
          rows.push({ height: 180, exact: true, cells: [front(a), front(b)] });
        }
        return grid([270, 270], rows, { abs: true, align: "center" });
      });
    },
  },

  {
    id: "envelope",
    name: "Envelope",
    category: "stationery",
    description: "A #10 business envelope (DL for A4) with your return address. Paste a list to print one envelope per person.",
    sizeLabel: "#10 (9.5 x 4.125 in) or DL (220 x 110 mm)",
    fold: "envelope",
    print: { paper: "#10 envelopes (9.5 x 4.125 in) in Letter regions, or DL envelopes (220 x 110 mm) in A4 regions.", finishing: "Feed each envelope the way your printer's envelope icon shows. Print one test envelope first." },
    customize: { paper: false, margins: false },
    page: (paper) => {
      const e = envelopeSpec(paper);
      return { width: e.width, height: e.height, orientation: "landscape", margins: e.margin };
    },
    merge: {
      fields: ADDRESS_FIELDS,
      sample: [],
      note: "Paste a list to print one envelope per person. Leave it empty to see a sample recipient.",
    },
    build(ctx) {
      const e = envelopeSpec(ctx.paper);
      const m = ctx.page.margins;
      const leftW = round(e.toLeft - m.left);
      const topH = round(e.toTop - m.top);
      const returnAddress = [
        ctx.hasLogo ? logoBox(110, 30, { after: 4 }) : null,
        p("{org}", "small", { bold: true, color: "primary", after: 1 }),
        p("{address}\n{city}", "small", { after: 0, line: 1.2 }),
      ];
      const recipient = (r) => {
        // The sample name stands on its own line, so the design checker can find and flag it.
        if (!r) return [p("Recipient Name", "body", { size: 11, line: 1.25, after: 0 }), p("Company Name\n456 Oak Avenue\nSpringfield, ST 12345", "body", { size: 11, line: 1.25, after: 0, fixed: true })];
        const place = [r.city, [r.state, r.postal].filter(Boolean).join(" ")].filter(Boolean).join(", ");
        const lines = [r.name, r.company, r.address, r.address2, place, r.country].filter(Boolean);
        return [p(lines.join("\n"), "body", { size: 11, line: 1.25, after: 0, fixed: true })];
      };
      const envelope = (r) =>
        grid(
          [leftW, round(ctx.W - leftW)],
          [
            { height: topH, cells: [cell(returnAddress, { pad: 0 }), cell([], { pad: 0 })] },
            { cells: [cell([], { pad: 0 }), cell(recipient(r), { pad: 0 })] },
          ],
          { abs: true }
        );
      const people = ctx.hasRecords ? ctx.records : [null];
      const out = [];
      people.forEach((r, i) => {
        if (i) out.push(pageBreak());
        out.push(envelope(r));
      });
      return out;
    },
  },

  {
    id: "invoice",
    name: "Invoice",
    category: "stationery",
    description: "A clean business invoice with your logo, the client, an itemized table with totals, payment terms, and notes.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    print: { paper: "Plain white paper. To send it by email, save it as a PDF." },
    customize: { paper: ["letter", "a4", "legal"], margins: true, spacing: true, text: true },
    page: { size: "paper", orientation: "portrait", margins: { top: 44, right: 50, bottom: 44, left: 50 } },
    options: [
      { key: "number", label: "Invoice number", type: "text", default: "1001" },
      { key: "currency", label: "Currency symbol", type: "text", default: "$" },
      { key: "tax", label: "Tax rate (percent)", type: "number", default: 8, min: 0, max: 50 },
      { key: "due", label: "Days until payment is due", type: "number", default: 30, min: 0, max: 365 },
    ],
    build(ctx) {
      const number = textOption(ctx, "number", "1001");
      const cur = textOption(ctx, "currency", "$").text;
      const taxRate = Math.min(50, Math.max(0, Number(ctx.opt.tax) || 0));
      const dueDays = whole(ctx.opt.due, 0, 365, 30);
      const today = new Date();
      const due = new Date(today.getFullYear(), today.getMonth(), today.getDate() + dueDays);
      const items = [
        ["Item or service", 1, 600],
        ["Item or service", 6, 150],
        ["Item or service", 8, 65],
        ["Item or service", 1, 180],
      ];
      const subtotal = items.reduce((s, [, q, r]) => s + q * r, 0);
      const tax = Math.round(subtotal * taxRate) / 100;
      const total = subtotal + tax;
      const meta = (label, value, fixed = true) => [
        cell([p(label, "label", { align: "right", after: 0, fixed: true })], { pad: [2, 8, 2, 0], valign: "middle" }),
        cell([p(value, "small", { align: "right", bold: true, after: 0, fixed })], { pad: [2, 0, 2, 0], valign: "middle" }),
      ];
      const sum = (label, value, strong) => [
        cell([p(label, strong ? "label" : "small", { color: strong ? "onPrimary" : "dark", after: 0, fixed: true })], { pad: [5, 8, 5, 8], fill: strong ? "primary" : null, border: strong ? null : { bottom: { color: "line", weight: 0.5 } } }),
        cell([p(value, strong ? "h2" : "small", { align: "right", color: strong ? "onPrimary" : "dark", after: 0, fixed: true })], { pad: [5, 8, 5, 8], fill: strong ? "primary" : null, border: strong ? null : { bottom: { color: "line", weight: 0.5 } } }),
      ];
      return [
        grid(
          [0.56, 0.44],
          [
            [
              cell([brandMark(ctx, 150, 54, { after: 6, size: 18 }), p("{address}\n{city}", "small", { after: 2 }), p("{phone}  |  {email}\n{web}", "small", { after: 0 })]),
              cell(
                [
                  p("Invoice", "title", { align: "right", size: 30, after: 6, fixed: true }),
                  grid([0.55, 0.45], [meta("Invoice number", number.text, number.fixed), meta("Invoice date", longDate(today)), meta("Due date", longDate(due))]),
                ],
                { valign: "top" }
              ),
            ],
          ]
        ),
        rule({ color: "accent", weight: 2, before: 10, after: 16 }),
        grid(
          [1, 1],
          [
            [
              cell([p("Bill to", "label", { color: "accent", fixed: true }), p("Client Name, Company Name", "body", { after: 0 }), p("456 Oak Avenue\nSpringfield, ST 12345", "body", { after: 0, fixed: true })]),
              cell([p("Project", "label", { color: "accent", fixed: true }), p("Project name", "body", { after: 8 }), p("Purchase order", "label", { color: "accent", fixed: true }), p("PO 4521", "body", { after: 0, sample: true })]),
            ],
          ],
          { gap: 24 }
        ),
        space(16),
        table(
          ["Description", "Quantity", "Rate", "Amount"],
          items.map(([d, q, r]) => [d, String(q), money(cur, r), money(cur, q * r)]),
          { widths: [0.52, 0.14, 0.17, 0.17] }
        ),
        space(10),
        grid([0.55, 0.45], [sum("Subtotal", money(cur, subtotal)), sum(`Tax (${taxRate}%)`, money(cur, tax)), sum("Total due", money(cur, total), true)], { width: round(ctx.W * 0.44), align: "right" }),
        space(20),
        grid(
          [1, 1],
          [
            [
              cell([
                p("Payment terms", "h2", { color: "primary" }),
                p(`${dueDays === 0 ? "Payment is due on receipt." : `Payment is due within ${dueDays} day${dueDays === 1 ? "" : "s"} of the invoice date.`} Please include the invoice number with your payment.`, "small", { fixed: true }),
                p("Bank transfer: Bank name, account name, and account number. Checks payable to {org}.", "small", { after: 0 }),
              ]),
              cell([p("Notes", "h2", { color: "primary" }), p("Thank you for your business. For questions, call {phone} or email {email}.", "small", { after: 0 })], { fill: "tint", pad: 10 }),
            ],
          ],
          { gap: 24 }
        ),
        space(22),
        rule({ color: "line", weight: 0.75, before: 0, after: 6 }),
        p("{org}  |  {address}, {city}  |  {web}", "small", { align: "center", color: "muted", after: 0 }),
      ];
    },
  },

  {
    id: "photo-page",
    name: "Photo page",
    category: "flyers",
    description: "A photo collage page with a title and six pictures with captions.",
    sizeLabel: "1 page, 6 pictures",
    fold: "none",
    print: { paper: "Matte or glossy photo paper gives the richest colors. Plain paper works for proofs." },
    customize: { paper: FLAT_SIZES, orientation: true },
    page: { size: "paper", orientation: "portrait", margins: 40 },
    build(ctx) {
      const across = ctx.W > ctx.H ? 3 : 2;
      const down = 6 / across;
      const small = ctx.W < 400;
      const gap = small ? 10 : 14;
      const header = small ? 86 : 104;
      const captionH = 26;
      const picH = round(Math.max(60, (ctx.H - header - down * captionH - (down - 1) * 4) / down));
      const shot = (k) =>
        cell(
          [
            pic("full", picH, { after: 4 }),
            p(altShort(k - 1), "caption", { after: 0 }),
          ],
          { pad: [0, 0, 4, 0] }
        );
      const rows = Array.from({ length: down }, (_, r) => ({ cells: Array.from({ length: across }, (_, c) => shot(r * across + c + 1)) }));
      return [
        p("{org}", "kicker", { align: "center" }),
        p("A Year in Pictures", "title", { align: "center", size: small ? 24 : 30, after: 4 }),
        p(alt(1, 7), "lead", { align: "center", color: "muted", after: 12 }),
        grid(Array(across).fill(1), rows, { gap }),
      ];
    },
  },
];

// ---------------------------------------------------------------------------
// Print and fold guides
// ---------------------------------------------------------------------------

export const MORE_FOLD_GUIDES = {
  booklet: {
    title: "Printing and folding a booklet",
    steps: [
      "In Word, Offizin turns on book fold, so Word prints the pages in booklet order. In Google Docs, Offizin lays out the printer spreads for you, with two pages on each side of the sheet.",
      "Print on both sides of the paper and flip on the short edge. If your printer cannot print on both sides, print the front sides first, put the stack back in the way your printer's manual shows, and then print the back sides.",
      "Print one copy first and check that the pages run in order before you print the rest.",
      "Keep the sheets in the order they came out of the printer, stack them neatly, and fold the stack in half.",
      "Staple twice along the fold, about a quarter of the way in from the top and the bottom.",
    ],
    tip: "A long-reach stapler makes stapling on the fold easy. With more sheets, the inner pages stick out a little at the open edge, so trim that edge with a paper cutter for a clean finish.",
  },
  tent: {
    title: "Folding place cards and table tents",
    steps: [
      "Print at actual size (100%) on card stock. Turn off any “fit to page” or “shrink” option.",
      "Cut along the light dashed guides with a paper trimmer.",
      "Fold each card in half across the middle, with the printed side out, and crease the fold.",
      "Stand the card on the table so the name faces the guest. The plain half is the back.",
    ],
    tip: "For a crisp fold, score the fold line first with a dull knife or an empty ballpoint pen against a ruler.",
  },
  envelope: {
    title: "Printing envelopes",
    steps: [
      "Find the envelope icon on your printer or in its manual. It shows which side faces up and which edge goes in first.",
      "Load one envelope in the envelope tray or manual feed slot, and choose the same envelope size in the print dialog: #10 in Letter regions or DL in A4 regions.",
      "Print one test envelope and check where the addresses land before you print the rest.",
      "With a pasted list, Offizin makes one page for each person, so the whole stack prints in a single run.",
    ],
    tip: "If the addresses land in the wrong place, the print dialog is probably using a different paper size. Choose the envelope size there and print again.",
  },
};
