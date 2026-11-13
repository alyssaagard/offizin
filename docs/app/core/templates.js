/**
 * The Offizin publication library and page parts.
 *
 * Each template describes one publication with the blocks in dsl.js:
 *   p (paragraph), list, rule, spacer, break (new page), image, band, grid, table, columns, scope.
 * Text can use business information fields such as {org} and {phone}, which the composer
 * fills from the user's business information (or friendly samples until it is filled in).
 * "**bold**" marks bold words inside a paragraph.
 * Running sample text (paragraphs, captions, quotes, bullets) is Middle High German filler from
 * filler.js. Headings stay short plain English, and names, dates, prices, and contact details
 * stay plain placeholders.
 *
 * Sizes are in points (72 points = 1 inch). ctx.W and ctx.H are the usable width and height
 * of the page inside its margins. See compose.js for everything the composer accepts.
 */
import { inch, mm, round, MONTHS, WEEKDAYS } from "./util.js";
import { p, list, rule, space, pageBreak, pic, logoBox, band, grid, cell, table, columns, scope, detail, contactLines, brandMark, panelFolds, panelPad, PANEL_PAD, sheetMargins, CUT_GUIDE, chunk } from "./dsl.js";
import { alt, altShort, altList, altQuote } from "./filler.js";
import { RESUME_TEMPLATES } from "./templates/resumes.js";
import { MORE_TEMPLATES, MORE_FOLD_GUIDES } from "./templates/more.js";

// ---------------------------------------------------------------------------
// Publications
// ---------------------------------------------------------------------------

export const CATEGORIES = [
  { id: "all", name: "All" },
  { id: "brochures", name: "Brochures" },
  { id: "newsletters", name: "Newsletters" },
  { id: "flyers", name: "Flyers, posters, and signs" },
  { id: "cards", name: "Cards and invitations" },
  { id: "programs", name: "Programs, booklets, and menus" },
  { id: "stationery", name: "Stationery and envelopes" },
  { id: "labels", name: "Labels, badges, and tickets" },
  { id: "resumes", name: "Resumes and CVs" },
  { id: "calendars", name: "Calendars and certificates" },
];

const FLAT_SIZES = ["letter", "legal", "tabloid", "halfLetter", "a3", "a4", "a5"];

/** One certificate page, for a merge record or for the sample recipient. */
function certificatePage(ctx, r) {
  const recipient = r && r.recipient ? r.recipient : "Recipient Name";
  // A pasted list without a reason column leaves that line blank, as the merge status says.
  const reason = r ? r.reason || "" : alt(1, 6);
  const date = r && r.date ? r.date : "Date";
      const inner = [
        p("{org}", "kicker", { align: "center", after: 14 }),
        p("Certificate of Achievement", "title", { align: "center", size: 34 }),
        rule({ color: "accent", weight: 1.5, width: 80, align: "center", before: 4, after: 16 }),
        p("Presented to", "body", { align: "center", italic: true, color: "muted", after: 10 }),
        p(recipient, "display", { align: "center", size: 40, color: "accent", italic: true, after: reason ? 12 : 42, fixed: !!(r && r.recipient) }),
        reason ? p(reason, "body", { align: "center", size: 12, after: 30, fixed: !!(r && r.reason) }) : null,
        grid(
          [1, 0.3, 1],
          [
            [
              cell([rule({ color: "dark", weight: 0.75, before: 0, after: 4 }), p("{name}, {title}", "small", { align: "center", after: 0 })]),
              cell([]),
              cell([rule({ color: "dark", weight: 0.75, before: 0, after: 4 }), p(date, "small", { align: "center", after: 0, fixed: true })]),
            ],
          ],
          { width: round(ctx.W * 0.7) }
        ),
      ];
      return [
        grid(
          [ctx.W],
          [
            {
              height: ctx.H - 10,
              cells: [
                cell([grid([ctx.W - 16], [{ height: ctx.H - 26, cells: [cell(inner, { pad: 30, valign: "middle", border: { all: { color: "accent", weight: 1 } } })] }], { abs: true })], {
                  pad: 7,
                  border: { all: { color: "primary", weight: 4 } },
                }),
              ],
            },
          ],
          { abs: true }
        ),
      ];

}

const CORE_TEMPLATES = [
  {
    id: "trifold",
    name: "Tri-fold brochure",
    category: "brochures",
    description: "Six panels on one sheet, printed on both sides and folded in thirds.",
    sizeLabel: "Landscape, 2 sides, folds in thirds",
    fold: "trifold",
    print: { duplex: "short", paper: "24 to 32 lb (90 to 120 g/m\u00B2) paper folds cleanly. Use card stock for a sturdier piece.", finishing: "Fold the right panel in first, then fold the left panel over it." },
    customize: { paper: ["letter", "legal", "a4"] },
    page: { size: "paper", orientation: "landscape", margins: 27 },
    folds: panelFolds(3),
    build(ctx) {
      const w = ctx.panels(3);
      const h = ctx.H - 10;
      const outside = grid(
        w,
        [
          {
            height: h,
            cells: [
              cell(
                [
                  p("Inside this brochure", "kicker"),
                  p("Why it matters", "h1"),
                  p(alt(1, 0)),
                  p(`“${altQuote(0)}”`, "quote", { size: 13, before: 6 }),
                  p("Name of the person quoted", "caption", { after: 12 }),
                  list(altList(3, 0)),
                ],
                { pad: panelPad(0, 3) }
              ),
              cell(
                [
                  brandMark(ctx, 120, 60, { align: "center", after: ctx.hasLogo ? 8 : 2 }),
                  ctx.hasLogo ? p("{org}", "h2", { align: "center", after: 2 }) : null,
                  p("{tagline}", "caption", { align: "center", after: 14 }),
                  ...contactLines({ align: "center" }),
                  space(18),
                  p(altShort(3), "caption", { align: "center" }),
                ],
                { pad: panelPad(1, 3), valign: "bottom" }
              ),
              cell(
                [
                  p("{tagline}", "kicker", { color: "onPrimary", align: "center" }),
                  p("{org}", "title", { color: "onPrimary", align: "center", size: 26 }),
                  rule({ color: "accent", weight: 2, width: 60, align: "center", before: 2, after: 12 }),
                  p(altShort(10), "lead", { color: "onPrimary", align: "center" }),
                  pic("full", 150, { fill: { mix: ["primary", "white", 0.25] }, ink: "onPrimary", before: 8, after: 12, align: "center" }),
                  p("{web}", "small", { color: "onPrimary", align: "center" }),
                ],
                { pad: [PANEL_PAD, PANEL_PAD, PANEL_PAD, PANEL_PAD], fill: "primary", valign: "middle" }
              ),
            ],
          },
        ],
        { abs: true }
      );
      const inside = grid(
        w,
        [
          {
            height: h,
            cells: [
              cell(
                [
                  p("About us", "kicker"),
                  p("Who we are", "h1"),
                  p(alt(1, 1)),
                  p(alt(1, 8)),
                  pic("full", 110, { before: 4 }),
                  p(altShort(5), "caption"),
                ],
                { pad: panelPad(0, 3) }
              ),
              cell(
                [
                  p("What we do", "kicker"),
                  p("Programs and services", "h1"),
                  p("Program name", "h2"),
                  p(alt(1, 2)),
                  p("Program name", "h2"),
                  p(alt(1, 10)),
                  p("Program name", "h2"),
                  p(alt(1, 12)),
                ],
                { pad: panelPad(1, 3) }
              ),
              cell(
                [
                  p("Get involved", "kicker"),
                  p("Join us", "h1"),
                  list(altList(3, 6)),
                  space(6),
                  band([p("Visit {web}", "button", { align: "center" })], { fill: "accent", pad: 9 }),
                  space(10),
                  ...contactLines(),
                ],
                { pad: panelPad(2, 3), fill: "tint" }
              ),
            ],
          },
        ],
        { abs: true }
      );
      return [outside, pageBreak(), inside];
    },
  },

  {
    id: "bifold",
    name: "Bi-fold brochure",
    category: "brochures",
    description: "A four page booklet from one sheet, folded in half.",
    sizeLabel: "Landscape, 2 sides, folds in half",
    fold: "halffold",
    print: { duplex: "short", paper: "24 to 32 lb (90 to 120 g/m\u00B2) paper folds cleanly. Use card stock for a sturdier piece.", finishing: "Fold in half with side 2 inside." },
    customize: { paper: ["letter", "legal", "tabloid", "a4", "a3"] },
    page: { size: "paper", orientation: "landscape", margins: 27 },
    folds: panelFolds(2),
    build(ctx) {
      const w = ctx.panels(2);
      const h = ctx.H - 10;
      return [
        grid(
          w,
          [
            {
              height: h,
              cells: [
                cell(
                  [
                    p("Thank you for reading", "kicker", { align: "center" }),
                    brandMark(ctx, 140, 70, { align: "center" }),
                    p("{tagline}", "caption", { align: "center", after: 16 }),
                    ...contactLines({ align: "center" }),
                  ],
                  { pad: panelPad(0, 2), valign: "bottom" }
                ),
                cell(
                  [
                    pic("full", 210, { fill: { mix: ["primary", "white", 0.25] }, ink: "onPrimary", after: 18 }),
                    p("{org}", "kicker", { color: "onPrimary" }),
                    p("Brochure title", "title", { color: "onPrimary", size: 26 }),
                    p(altShort(1), "lead", { color: "onPrimary" }),
                  ],
                  { pad: [28, 28, 28, 28], fill: "primary", valign: "middle" }
                ),
              ],
            },
          ],
          { abs: true }
        ),
        pageBreak(),
        grid(
          w,
          [
            {
              height: h,
              cells: [
                cell(
                  [
                    p("Our mission", "kicker"),
                    p("Why we do it", "headline"),
                    p(alt(2, 0), "lead"),
                    p(alt(1, 9)),
                    p(`“${altQuote(1)}”`, "quote", { before: 8 }),
                    p("Name, how they know you", "caption"),
                  ],
                  { pad: panelPad(0, 2) }
                ),
                cell(
                  [
                    pic("full", 160, { after: 12 }),
                    p("What we offer", "h1"),
                    grid(
                      [1, 1],
                      [
                        [
                          cell([p("Service name", "h2"), p(alt(1, 5), "small")], { pad: [0, 8, 6, 0] }),
                          cell([p("Service name", "h2"), p(alt(1, 7), "small")], { pad: [0, 0, 6, 8] }),
                        ],
                        [
                          cell([p("Service name", "h2"), p(alt(1, 10), "small")], { pad: [0, 8, 0, 0] }),
                          cell([p("Service name", "h2"), p(alt(1, 11), "small")], { pad: [0, 0, 0, 8] }),
                        ],
                      ]
                    ),
                    space(10),
                    band([p("Call {phone}", "button", { align: "center" })], { fill: "accent", pad: 9 }),
                  ],
                  { pad: panelPad(1, 2) }
                ),
              ],
            },
          ],
          { abs: true }
        ),
      ];
    },
  },

  {
    id: "newsletter",
    name: "Newsletter",
    category: "newsletters",
    description: "A two page newsletter with a masthead, lead story, sidebar, and contact strip.",
    sizeLabel: "Portrait, 2 pages",
    fold: "none",
    print: { duplex: "long", paper: "Plain copy paper works well. Use 24 lb (90 g/m\u00B2) for a heavier feel." },
    customize: { paper: ["letter", "legal", "a4"] },
    footer: { text: "{org} News  |  Page {page} of {pages}", align: "center" },
    page: { size: "paper", orientation: "portrait", margins: 36 },
    build(ctx) {
      const now = new Date();
      ctx.setVar("issueMonth", MONTHS[now.getMonth()]);
      ctx.setVar("issueYear", now.getFullYear());
      return [
        band(
          [
            p("Volume 1, Issue 1  |  {issueMonth} {issueYear}", "kicker", { color: "onPrimary", after: 2 }),
            p("{org} News", "title", { color: "onPrimary", after: 0 }),
            p("{tagline}", "caption", { color: "onPrimary", italic: true, after: 0 }),
          ],
          { fill: "primary", pad: [16, 18, 16, 18], valign: "middle" }
        ),
        space(14),
        grid(
          [0.64, 0.36],
          [
            [
              cell([
                p("Lead story", "kicker"),
                p("Story headline", "headline"),
                p("By the author's name", "caption", { after: 8 }),
                pic("full", 170, { after: 4 }),
                p(altShort(7), "caption", { after: 8 }),
                p(alt(2, 0)),
                p(alt(1, 9)),
                p(alt(1, 2)),
              ]),
              cell(
                [
                  p("In this issue", "h2", { color: "primary" }),
                  list(["Lead story", "Member spotlight", "Upcoming dates", "From the director"], { role: "small", gap: 2 }),
                  rule({ color: "line", before: 6, after: 8 }),
                  p("Upcoming", "h2", { color: "primary" }),
                  p("**June 14**  Event name", "small"),
                  p("**July 2**  Event name", "small"),
                  p("**July 19**  Event name", "small", { after: 8 }),
                  rule({ color: "line", before: 0, after: 8 }),
                  p(`“${altQuote(3)}”`, "quote", { size: 13 }),
                ],
                { fill: "tint", pad: 12 }
              ),
            ],
          ],
          { gap: 18 }
        ),
        pageBreak(),
        // Text flows from the first column into the second, like linked text boxes.
        columns(
          2,
          [
            p("Member spotlight", "kicker"),
            p("Meet a member", "h1"),
            pic("full", 130, { after: 8 }),
            p(alt(1, 4)),
            p(alt(1, 8)),
            p("From the director", "kicker", { before: 8 }),
            p("The season ahead", "h1"),
            p(alt(1, 6)),
            p(alt(1, 12)),
            p("{name}", "h2", { before: 6, after: 0 }),
            p("{title}", "caption"),
          ],
          { gap: 18, rule: true }
        ),
        space(14),
        grid(
          [0.5, 0.5],
          [
            [
              cell([p("Save the date", "h2", { color: "primary" }), table(null, [["June 14", "Event name"], ["July 2", "Event name"], ["July 19", "Event name"]], { widths: [0.35, 0.65], firstBold: true, firstColor: "primary" })], { pad: [0, 9, 0, 0] }),
              cell([p("Get in touch", "h2", { color: "primary" }), ...contactLines()], { pad: [0, 0, 0, 9] }),
            ],
          ]
        ),
        space(14),
        band([p("{org}  |  {web}  |  {phone}", "small", { color: "onPrimary", align: "center", after: 0 })], { fill: "primary", pad: 8 }),
      ];
    },
  },

  {
    id: "event-flyer",
    name: "Event flyer",
    category: "flyers",
    description: "A bold one page flyer with a picture, event details, and a call to action.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    customize: { paper: FLAT_SIZES },
    page: { size: "paper", orientation: "portrait", margins: 36 },
    build(ctx) {
      return [
        pic("full", round(ctx.H * 0.36), { after: 16 }),
        p("{org} presents", "kicker", { align: "center" }),
        p("Event name", "display", { align: "center" }),
        p(altShort(3), "lead", { align: "center", after: 16 }),
        grid(
          [1, 1, 1],
          [[cell(detail("Date", "Saturday, June 14"), { fill: "tint", pad: 10 }), cell(detail("Time", "6 to 9 PM"), { fill: "tint", pad: 10 }), cell(detail("Place", "Riverside Park"), { fill: "tint", pad: 10 })]],
          { gap: 10 }
        ),
        space(14),
        p(alt(1, 1), "body", { align: "center", size: 11 }),
        space(8),
        band([p("Reserve your seat at {web}", "button", { align: "center" })], { fill: "accent", pad: 11, width: round(ctx.W * 0.7) }),
        space(10),
        p("Questions? Call {phone} or email {email}", "small", { align: "center", color: "muted" }),
      ];
    },
  },

  {
    id: "sale-flyer",
    name: "Sale flyer with coupon",
    category: "flyers",
    description: "An announcement with offers and a cut-out coupon.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    customize: { paper: FLAT_SIZES },
    page: { size: "paper", orientation: "portrait", margins: 36 },
    build(ctx) {
      return [
        band(
          [
            p("{org}", "kicker", { color: "onPrimary", align: "center" }),
            p("Spring sale", "display", { color: "onPrimary", align: "center", sample: true }),
            p(alt(1, 5), "lead", { color: "onPrimary", align: "center", after: 0 }),
          ],
          { fill: "primary", pad: [26, 20, 26, 20] }
        ),
        space(16),
        grid(
          [0.48, 0.52],
          [
            [
              cell([pic("full", 220)], { pad: [0, 9, 0, 0] }),
              cell(
                [
                  p("This weekend's offers", "h1"),
                  list(["**30% off** all spring stock", "**Buy two, get one** on accessories", "**Free gift wrap** with any purchase", "**Members** shop an hour early on Friday"], { gap: 5 }),
                  p("Friday to Sunday, 9 AM to 6 PM", "h2", { before: 6 }),
                  p("{address}, {city}", "small"),
                ],
                { pad: [0, 0, 0, 9] }
              ),
            ],
          ]
        ),
        space(18),
        grid(
          [ctx.W * 0.72],
          [
            [
              cell(
                [
                  p("Coupon", "kicker", { align: "center" }),
                  p("Take an extra 20% off one item", "headline", { align: "center" }),
                  p("One coupon per customer. Expires at closing on Sunday.", "small", { align: "center" }),
                  p("Code: SPRING20", "label", { align: "center", color: "dark", after: 0 }),
                ],
                { pad: 14, border: { all: { color: "dark", weight: 1.5, style: "dashed" } } }
              ),
            ],
          ],
          { abs: true, align: "center" }
        ),
        space(14),
        p("{web}  |  {phone}", "small", { align: "center", color: "muted" }),
      ];
    },
  },

  {
    id: "poster",
    name: "Poster",
    category: "flyers",
    description: "A large format poster for walls and windows.",
    sizeLabel: "Tabloid or A3, portrait",
    fold: "none",
    print: { paper: "For posters larger than your printer, save as PDF and send it to a print shop." },
    customize: { paper: ["tabloid", "a3", "letter", "legal", "a4"] },
    page: (paper) => ({ size: paper === "a4" ? "a3" : "tabloid", orientation: "portrait", margins: 48 }),
    build(ctx) {
      return [
        p("{org}", "kicker", { size: 13, align: "center" }),
        p("Event name", "display", { size: 72, align: "center", after: 14 }),
        pic("full", round(ctx.H * 0.42), { after: 20 }),
        p(alt(1, 11), "lead", { size: 22, align: "center", after: 22 }),
        grid(
          [1, 1, 1],
          [[cell(detail("When", "Saturday, May 3", { size: 22 }), { pad: 6 }), cell(detail("Time", "9 AM to noon", { size: 22 }), { pad: 6 }), cell(detail("Meet at", "City Hall steps", { size: 22 }), { pad: 6 })]],
          { gap: 12 }
        ),
        space(22),
        band([p("Sign up at {web}", "button", { align: "center", size: 20 })], { fill: "accent", pad: 16, width: round(ctx.W * 0.75) }),
      ];
    },
  },

  {
    id: "sign",
    name: "Sign or notice",
    category: "flyers",
    description: "A clear landscape sign for doors, desks, and events.",
    sizeLabel: "Landscape, 1 page",
    fold: "none",
    customize: { paper: FLAT_SIZES, orientation: true },
    page: { size: "paper", orientation: "landscape", margins: 40 },
    build(ctx) {
      return [
        band([p("{org}", "h2", { color: "onPrimary", align: "center", after: 0 })], { fill: "primary", pad: 10 }),
        space(round(ctx.H * 0.14)),
        p("Use the side entrance", "display", { size: 54, align: "center", color: "dark", after: 16 }),
        p(alt(1, 3), "lead", { size: 18, align: "center" }),
        space(round(ctx.H * 0.12)),
        rule({ color: "accent", weight: 3, width: 120, align: "center" }),
        p("Questions? {phone}", "body", { align: "center", size: 13, color: "muted" }),
      ];
    },
  },

  {
    id: "business-cards",
    name: "Business cards",
    category: "cards",
    description: "Ten 3.5 x 2 inch cards per sheet, filled from your business information.",
    sizeLabel: "10 per sheet, with cut guides",
    fold: "cards",
    print: { paper: "Card stock, 80 to 110 lb cover (216 to 300 g/m\u00B2).", finishing: "Cut on the dashed guides with a paper trimmer." },
    customize: { paper: false, margins: false },
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 252, 144, 2, 5) }),
    folds: () => ({ vertical: [], horizontal: [], cut: true }),
    build(ctx) {
      const guide = { all: { color: "line", weight: 0.25, style: "dashed" } };
      const card = () =>
        cell(
          [
            ctx.hasLogo ? logoBox(90, 26, { after: 6 }) : p("{org}", "label", { color: "accent", after: 8 }),
            p("{name}", "h2", { color: "primary", size: 13, after: 0, fixed: true }),
            p("{title}", "caption", { after: 10 }),
            p("{phone}\n{email}\n{web}", "small", { size: 7.5, after: 3, line: 1.25 }),
            p("{address}, {city}", "small", { size: 7, color: "muted", after: 0 }),
          ],
          { pad: [16, 18, 10, 18], border: guide }
        );
      const rows = Array.from({ length: 5 }, () => ({ height: 144, exact: true, cells: [card(), card()] }));
      return [grid([252, 252], rows, { abs: true, align: "center" })];
    },
  },

  {
    id: "postcards",
    name: "Postcards",
    category: "cards",
    description: "Two 6 x 4 inch postcards per sheet, with fronts on side 1 and mailing backs on side 2.",
    sizeLabel: "2 per sheet, 2 sides",
    fold: "cards",
    print: { duplex: "long", paper: "Card stock, 80 to 110 lb cover (216 to 300 g/m\u00B2).", finishing: "Cut on the dashed guides." },
    customize: { paper: false, margins: false },
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 432, 288, 1, 2) }),
    folds: () => ({ vertical: [], horizontal: [], cut: true }),
    build(ctx) {
      const guide = { all: { color: "line", weight: 0.25, style: "dashed" } };
      const front = () =>
        cell(
          [
            p("{org}", "kicker", { color: "onPrimary" }),
            p("You are invited", "title", { color: "onPrimary", size: 30 }),
            p("Saturday, October 4, 1 to 4 PM", "lead", { color: "onPrimary" }),
            p("{web}", "small", { color: "onPrimary", after: 0 }),
          ],
          { pad: [28, 30, 28, 30], fill: "primary", valign: "middle", border: guide }
        );
      const back = () =>
        cell(
          [
            grid(
              [0.54, 0.46],
              [
                [
                  cell(
                    [
                      p("Open house", "h2", { color: "primary", sample: true }),
                      p(alt(1, 2), "small"),
                      space(6),
                      p("{org}", "small", { bold: true, after: 0 }),
                      p("{address}\n{city}", "small", { color: "muted" }),
                    ],
                    { pad: [0, 12, 0, 0], border: { right: { color: "line", weight: 0.75 } } }
                  ),
                  cell(
                    [
                      grid([54], [[cell([p("Place stamp here", "label", { align: "center", after: 0, fixed: true })], { pad: 6, border: { all: { color: "line", weight: 0.75 } }, valign: "middle" })]], { abs: true, align: "right" }),
                      space(36),
                      rule({ color: "line", weight: 0.75, before: 0, after: 20 }),
                      rule({ color: "line", weight: 0.75, before: 0, after: 20 }),
                      rule({ color: "line", weight: 0.75, before: 0, after: 20 }),
                    ],
                    { pad: [0, 0, 0, 14] }
                  ),
                ],
              ]
            ),
          ],
          { pad: [22, 24, 18, 24], border: guide }
        );
      return [
        grid([432], [{ height: 288, exact: true, cells: [front()] }, { height: 288, exact: true, cells: [front()] }], { abs: true, align: "center" }),
        pageBreak(),
        grid([432], [{ height: 288, exact: true, cells: [back()] }, { height: 288, exact: true, cells: [back()] }], { abs: true, align: "center" }),
      ];
    },
  },

  {
    id: "greeting-card",
    name: "Greeting card",
    category: "cards",
    description: "A folded card: front and back on side 1, the message inside on side 2.",
    sizeLabel: "Landscape, 2 sides, folds in half",
    fold: "halffold",
    print: { duplex: "short", paper: "Card stock, 80 to 110 lb cover (216 to 300 g/m\u00B2).", finishing: "Fold in half with side 2 inside." },
    customize: { paper: ["letter", "a4"] },
    page: { size: "paper", orientation: "landscape", margins: 27 },
    folds: panelFolds(2),
    build(ctx) {
      const w = ctx.panels(2);
      const h = ctx.H - 10;
      return [
        grid(
          w,
          [
            {
              height: h,
              cells: [
                cell([ctx.hasLogo ? logoBox(90, 40, { align: "center", after: 4 }) : null, p("Made with care by {org}", "caption", { align: "center" })], { pad: panelPad(0, 2), valign: "bottom" }),
                cell(
                  [
                    pic("full", round(h * 0.55), { after: 22 }),
                    p("With Gratitude", "title", { align: "center", color: "primary" }),
                    rule({ color: "accent", weight: 1.5, width: 50, align: "center", before: 4, after: 10 }),
                    p(altShort(3), "lead", { align: "center", italic: true, color: "muted" }),
                  ],
                  { pad: panelPad(1, 2), valign: "middle" }
                ),
              ],
            },
          ],
          { abs: true }
        ),
        pageBreak(),
        grid(
          w,
          [
            {
              height: h,
              cells: [
                cell([p(`“${altQuote(0)}”`, "quote", { align: "center", size: 16 }), p("A favorite saying", "caption", { align: "center" })], { pad: panelPad(0, 2), valign: "middle" }),
                cell(
                  [
                    p("Dear friend,", "h2", { color: "primary", after: 10 }),
                    p(alt(2, 2), "lead"),
                    p("With thanks,", "body", { before: 10, after: 2 }),
                    p("{name}", "h2", { color: "primary", fixed: true }),
                  ],
                  { pad: panelPad(1, 2), valign: "middle" }
                ),
              ],
            },
          ],
          { abs: true }
        ),
      ];
    },
  },

  {
    id: "invitation",
    name: "Invitation",
    category: "cards",
    description: "A framed invitation on a 5 x 7 inch card (A5 for A4 paper).",
    sizeLabel: "5 x 7 in or A5",
    fold: "none",
    print: { paper: "Card stock, 80 to 110 lb cover (216 to 300 g/m\u00B2)." },
    customize: { paper: ["card5x7", "a5", "halfLetter", "letter", "a4"] },
    page: (paper) => ({ size: paper === "a4" ? "a5" : "card5x7", orientation: "portrait", margins: 22 }),
    build(ctx) {
      return [
        grid(
          [ctx.W],
          [
            {
              height: ctx.H - 10,
              cells: [
                cell(
                  [
                    p("You are invited to", "kicker", { align: "center", after: 10 }),
                    p("Event name", "title", { align: "center", size: 26 }),
                    rule({ color: "accent", weight: 1.5, width: 50, align: "center", before: 6, after: 14 }),
                    p(alt(1, 1), "lead", { align: "center", italic: true, size: 11 }),
                    space(10),
                    p("Saturday, November 22", "h2", { align: "center", after: 2 }),
                    p("Seven o'clock in the evening", "body", { align: "center" }),
                    p("The Glass House, 45 River Road", "body", { align: "center", after: 16 }),
                    p("Kindly reply by November 10", "small", { align: "center", color: "muted", after: 2 }),
                    p("{email}  |  {phone}", "small", { align: "center", color: "muted", after: 0 }),
                  ],
                  { pad: 20, valign: "middle", border: { all: { color: "accent", weight: 1.5 } } }
                ),
              ],
            },
          ],
          { abs: true }
        ),
      ];
    },
  },

  {
    id: "program",
    name: "Event program",
    category: "programs",
    description: "A folded program with a cover, welcome, schedule, and thanks.",
    sizeLabel: "Landscape, 2 sides, folds in half",
    fold: "halffold",
    print: { duplex: "short", paper: "24 to 32 lb (90 to 120 g/m\u00B2) paper folds cleanly. Use card stock for a sturdier piece.", finishing: "Fold in half with side 2 inside." },
    customize: { paper: ["letter", "legal", "tabloid", "a4", "a3"] },
    page: { size: "paper", orientation: "landscape", margins: 27 },
    folds: panelFolds(2),
    build(ctx) {
      const w = ctx.panels(2);
      const h = ctx.H - 10;
      return [
        grid(
          w,
          [
            {
              height: h,
              cells: [
                cell(
                  [
                    p("With thanks", "h1", { align: "center" }),
                    p(alt(1, 6), "body", { align: "center", after: 12 }),
                    p("Sponsor name\nSponsor name\nSponsor name", "body", { align: "center", bold: true, after: 22 }),
                    brandMark(ctx, 120, 50, { align: "center" }),
                    ...contactLines({ align: "center" }),
                  ],
                  { pad: panelPad(0, 2), valign: "middle" }
                ),
                cell(
                  [
                    p("{org} presents", "kicker", { align: "center", color: "onPrimary" }),
                    p("Event name", "title", { align: "center", color: "onPrimary" }),
                    rule({ color: "accent", weight: 2, width: 60, align: "center", before: 4, after: 14 }),
                    pic("full", round(h * 0.38), { fill: { mix: ["primary", "white", 0.25] }, ink: "onPrimary", after: 16 }),
                    p("Friday, March 20  |  Six o'clock", "h2", { align: "center", color: "onPrimary", after: 2 }),
                    p("Grand Hall, Riverside Center", "body", { align: "center", color: "onPrimary" }),
                  ],
                  { pad: [26, 26, 26, 26], fill: "primary", valign: "middle" }
                ),
              ],
            },
          ],
          { abs: true }
        ),
        pageBreak(),
        grid(
          w,
          [
            {
              height: h,
              cells: [
                cell(
                  [
                    p("Welcome", "kicker"),
                    p("Before we begin", "h1"),
                    p(alt(1, 2)),
                    p(alt(1, 5)),
                    p("{name}, {title}", "caption", { after: 14 }),
                    p("Program", "h1"),
                    table(
                      null,
                      [
                        ["6:00", "Reception and music"],
                        ["6:45", "Welcome and opening remarks"],
                        ["7:00", "Dinner is served"],
                        ["7:45", "Awards presentation"],
                        ["8:30", "Closing toast"],
                      ],
                      { widths: [0.22, 0.78], firstBold: true, firstColor: "primary" }
                    ),
                  ],
                  { pad: panelPad(0, 2) }
                ),
                cell(
                  [
                    p("Honorees", "kicker"),
                    p("Tonight we recognize", "h1"),
                    p("Honoree name", "h2", { after: 0 }),
                    p(altShort(4), "small", { after: 8 }),
                    p("Honoree name", "h2", { after: 0 }),
                    p(altShort(8), "small", { after: 8 }),
                    p("Honoree name", "h2", { after: 0 }),
                    p(altShort(6), "small", { after: 14 }),
                    band([p("Notes", "label", { after: 30, fixed: true })], { fill: "tint", pad: 10, valign: "top" }),
                  ],
                  { pad: panelPad(1, 2) }
                ),
              ],
            },
          ],
          { abs: true }
        ),
      ];
    },
  },

  {
    id: "menu",
    name: "Menu",
    category: "programs",
    description: "A one page menu with sections, descriptions, and prices.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    print: { paper: "Heavy matte paper or card stock holds up best at the table." },
    customize: { paper: FLAT_SIZES },
    page: { size: "paper", orientation: "portrait", margins: 42 },
    build(ctx) {
      const item = (name, desc, price) => [
        cell([p(name, "h2", { after: 1 }), p(desc, "small", { color: "muted", after: 0 })], { pad: [0, 8, 9, 0] }),
        cell([p(price, "h2", { align: "right", color: "primary", after: 0, fixed: true })], { pad: [0, 0, 9, 0] }),
      ];
      const section = (title, items) => [
        p(title, "kicker", { size: 9, before: 6 }),
        rule({ color: "line", before: 0, after: 8 }),
        grid([0.82, 0.18], items.map((i) => item(...i))),
      ];
      return [
        // The title line below already carries the name, so only a real logo goes above it.
        ctx.hasLogo ? logoBox(150, 60, { align: "center", after: 2 }) : null,
        p("{org}", "title", { align: "center", after: 2 }),
        p("{tagline}", "caption", { align: "center", after: 14 }),
        ...section("To start", [
          ["Dish name", altShort(0), "8"],
          ["Dish name", altShort(1), "9"],
          ["Dish name", altShort(2), "11"],
        ]),
        ...section("Mains", [
          ["Dish name", altShort(3), "22"],
          ["Dish name", altShort(4), "24"],
          ["Dish name", altShort(5), "19"],
        ]),
        ...section("To finish", [
          ["Dish name", altShort(6), "9"],
          ["Dish name", altShort(7), "9"],
        ]),
        space(6),
        p(altShort(11), "caption", { align: "center" }),
        p("{address}, {city}  |  {phone}", "small", { align: "center", color: "muted" }),
      ];
    },
  },

  {
    id: "letterhead",
    name: "Letterhead",
    category: "stationery",
    description: "Branded letter paper with your logo, address, and a sample letter.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    print: { paper: "24 lb (90 g/m\u00B2) bond paper gives letters a professional feel." },
    customize: { paper: ["letter", "legal", "a4"] },
    page: { size: "paper", orientation: "portrait", margins: { top: 40, right: 54, bottom: 40, left: 54 } },
    build(ctx) {
      return [
        grid(
          [0.5, 0.5],
          [
            [
              cell([brandMark(ctx, 150, 54, { after: 0, size: 18 })], { valign: "middle" }),
              cell([p("{address}\n{city}", "small", { align: "right", after: 2 }), p("{phone}  |  {email}\n{web}", "small", { align: "right", after: 0 })], { valign: "middle" }),
            ],
          ]
        ),
        rule({ color: "accent", weight: 2, before: 8, after: 28 }),
        p("Month Day, Year", "body", { after: 14 }),
        p("Recipient Name\nCompany\nStreet Address\nCity, State Postal Code", "body", { after: 14 }),
        p("Dear Recipient Name,", "body", { after: 10 }),
        p(alt(2, 2), "body", { size: 10.5, line: 1.4, after: 10 }),
        p(alt(1, 8), "body", { size: 10.5, line: 1.4, after: 10 }),
        p(alt(1, 6), "body", { size: 10.5, line: 1.4, after: 18 }),
        p("Sincerely,", "body", { after: 30 }),
        p("{name}", "body", { bold: true, after: 0 }),
        p("{title}, {org}", "body", { color: "muted" }),
      ];
    },
  },

  {
    id: "labels",
    name: "Address labels",
    category: "labels",
    description: "Return address labels, or one label per person from a pasted list: 30 per Letter sheet (2.625 x 1 in) or 21 per A4 sheet (63.5 x 38.1 mm).",
    sizeLabel: "30 per sheet (Letter), 21 per sheet (A4)",
    fold: "cards",
    print: { paper: "Label sheets with the same layout: 30 labels of 2.625 x 1 in (Letter) or 21 labels of 63.5 x 38.1 mm (A4).", finishing: "Print a test on plain paper first and hold it over a label sheet to check alignment." },
    customize: { paper: false, margins: false },
    page: (paper) =>
      paper === "a4"
        ? { size: "paper", orientation: "portrait", margins: { top: round(mm(15.15)), bottom: round(mm(15.15)), left: round(mm(7.21)), right: round(mm(7.21)) } }
        : { size: "paper", orientation: "portrait", margins: { top: 36, bottom: 36, left: 13.5, right: 13.5 } },
    folds: () => ({ vertical: [], horizontal: [], cut: true }),
    merge: {
      fields: [
        { key: "name", label: "Name", aliases: ["full name", "recipient", "addressee"] },
        { key: "company", label: "Company", aliases: ["organization", "organisation"] },
        { key: "address", label: "Street address", aliases: ["address 1", "address line 1", "street"] },
        { key: "address2", label: "Address line 2", aliases: ["apt", "suite", "unit", "address 2"] },
        { key: "city", label: "City", aliases: ["town", "city state zip"] },
        { key: "state", label: "State or region", aliases: ["province", "region", "county"] },
        { key: "postal", label: "Postal code", aliases: ["zip", "zip code", "postcode"] },
        { key: "country", label: "Country" },
      ],
      sample: [],
      note: "Paste a list to print one label per person. Leave it empty to print your own return address on every label.",
    },
    build(ctx) {
      const a4 = ctx.paper === "a4";
      const w = a4 ? round(mm(63.5)) : 189;
      const h = a4 ? round(mm(38.1)) : 72;
      const gap = a4 ? round(mm(2.54)) : 9;
      const down = a4 ? 7 : 10;
      const perSheet = down * 3;
      const returnLabel = () =>
        cell([p("{org}", "small", { bold: true, color: "primary", after: 1 }), p("{address}\n{city}", "small", { after: 0, line: 1.2 })], { pad: [4, 10, 4, 12], valign: "middle" });
      const mailLabel = (r) => {
        const place = [r.city, [r.state, r.postal].filter(Boolean).join(" ")].filter(Boolean).join(", ");
        const lines = [r.company, r.address, r.address2, place, r.country].filter(Boolean);
        // A Letter label is 1 inch tall, so a six-line address is set tighter there. The address stays at 8 pt or more.
        const tight = !a4 && lines.length > 4;
        return cell(
          [
            p(r.name || "", "small", { bold: true, color: "dark", after: 1, fixed: true, line: tight ? 1.1 : undefined }),
            lines.length ? p(lines.join("\n"), "small", { after: 0, line: tight ? 1 : 1.15, size: tight ? 8 : undefined, fixed: true }) : null,
          ],
          { pad: [4, 10, 4, 12], valign: "middle" }
        );
      };
      const sheet = (cells) => {
        const rows = [];
        for (let i = 0; i < down; i++) rows.push({ height: h, exact: true, cells: cells.slice(i * 3, i * 3 + 3) });
        return grid([w, w, w], rows, { abs: true, gap, align: "center" });
      };
      if (!ctx.hasRecords) return [sheet(Array.from({ length: perSheet }, returnLabel))];
      const out = [];
      chunk(ctx.records, perSheet).forEach((group, i) => {
        if (i) out.push(pageBreak());
        const cells = group.map(mailLabel);
        while (cells.length < perSheet) cells.push(cell([], { pad: 0 }));
        out.push(sheet(cells));
      });
      return out;
    },
  },

  {
    id: "calendar",
    name: "Monthly calendar",
    category: "calendars",
    description: "A month at a glance, generated for any month and year.",
    sizeLabel: "Landscape, 1 page",
    fold: "none",
    customize: { paper: FLAT_SIZES, orientation: true },
    page: { size: "paper", orientation: "landscape", margins: 30 },
    options: [
      { key: "month", label: "Month", type: "month", default: () => new Date().getMonth() + 1 },
      { key: "year", label: "Year", type: "number", default: () => new Date().getFullYear() },
      {
        key: "weekStart",
        label: "Weeks start on",
        type: "select",
        choices: [
          { value: "sunday", label: "Sunday" },
          { value: "monday", label: "Monday" },
        ],
        default: "sunday",
      },
    ],
    build(ctx) {
      const month = Math.min(12, Math.max(1, Number(ctx.opt.month) || 1));
      const year = Number(ctx.opt.year) || new Date().getFullYear();
      ctx.setVar("calMonth", MONTHS[month - 1]);
      ctx.setVar("calYear", year);
      const startDay = ctx.opt.weekStart === "monday" ? 1 : 0;
      const first = (new Date(year, month - 1, 1).getDay() - startDay + 7) % 7;
      const dayNames = WEEKDAYS.slice(startDay).concat(WEEKDAYS.slice(0, startDay));
      const days = new Date(year, month, 0).getDate();
      const weeks = Math.ceil((first + days) / 7);
      const headerH = 62;
      const dayRowH = 20;
      const rowH = round((ctx.H - headerH - dayRowH - 16) / weeks);
      const border = { all: { color: "line", weight: 0.5 } };
      const rows = [
        { cells: dayNames.map((d) => cell([p(d, "label", { color: "onPrimary", align: "center", after: 0, fixed: true })], { fill: "primary", pad: 5 })) },
      ];
      let n = 1 - first;
      for (let wk = 0; wk < weeks; wk++) {
        const cells = [];
        for (let d = 0; d < 7; d++, n++) {
          const inMonth = n >= 1 && n <= days;
          const sunday = dayNames[d] === "Sunday";
          cells.push(cell(inMonth ? [p(String(n), "h2", { color: sunday ? "accent" : "dark", after: 0, size: 11, fixed: true })] : [], { pad: [4, 6, 4, 6], border, fill: inMonth ? null : "tint" }));
        }
        rows.push({ height: rowH, cells });
      }
      return [
        grid(
          [0.7, 0.3],
          [
            {
              height: headerH,
              cells: [
                cell([p("{calMonth} {calYear}", "display", { size: 38, after: 0 })], { valign: "bottom", pad: [0, 0, 8, 0] }),
                cell([brandMark(ctx, 120, 40, { align: "right", after: 0 }), p("{web}", "small", { align: "right", color: "muted", after: 0 })], { valign: "bottom", pad: [0, 0, 10, 0] }),
              ],
            },
          ]
        ),
        grid([1, 1, 1, 1, 1, 1, 1], rows),
      ];
    },
  },

  {
    id: "certificate",
    name: "Certificate",
    category: "calendars",
    description: "A framed certificate of achievement with signature lines.",
    sizeLabel: "Landscape, 1 page",
    fold: "none",
    print: { paper: "Card stock, 80 to 110 lb cover (216 to 300 g/m\u00B2). Parchment or bright white stock looks best." },
    customize: { paper: ["letter", "legal", "a4"] },
    page: { size: "paper", orientation: "landscape", margins: 30 },
    merge: {
      fields: [
        { key: "recipient", label: "Recipient name", aliases: ["name", "full name", "student", "employee"] },
        { key: "reason", label: "Reason or achievement", aliases: ["achievement", "for", "award", "description"] },
        { key: "date", label: "Date", aliases: ["awarded", "date awarded"] },
      ],
      sample: [],
      note: "Paste a list of names to print one certificate per person. Reason and date are optional.",
    },
    build(ctx) {
      const people = ctx.hasRecords ? ctx.records : [null];
      const out = [];
      people.forEach((r, i) => {
        if (i) out.push(pageBreak());
        out.push(...certificatePage(ctx, r));
      });
      return out;
    },
  },

  {
    id: "resume",
    name: "Designed resume",
    category: "resumes",
    description: "A designed one page resume with a shaded sidebar, for printing and handing out. For online applications, use a plain resume: screening software reads it more reliably.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    customize: { paper: ["letter", "a4"] },
    page: { size: "paper", orientation: "portrait", margins: 36 },
    build(ctx) {
      const job = (role, place, dates, text) => [p(role, "h2", { after: 0 }), p(`${place}  |  ${dates}`, "caption", { after: 3 }), p(text, "small", { after: 9 })];
      return [
        p("{name}", "title", { size: 32, after: 0, fixed: true }),
        p("{title}", "lead", { color: "accent", after: 10 }),
        rule({ color: "primary", weight: 2, before: 0, after: 14 }),
        grid(
          [0.66, 0.34],
          [
            [
              cell(
                [
                  p("Profile", "kicker"),
                  p(alt(2, 7), "body", { after: 12 }),
                  p("Experience", "kicker"),
                  ...job("Job title", "Organization name, City", "2022 to present", alt(1, 0)),
                  ...job("Job title", "Organization name, City", "2018 to 2022", alt(1, 11)),
                  ...job("Job title", "Organization name, City", "2015 to 2018", alt(1, 14)),
                  p("Education", "kicker", { before: 4 }),
                  p("Degree", "h2", { after: 0, sample: true }),
                  p("School name, City  |  2014", "caption"),
                ],
                { pad: [0, 16, 0, 0] }
              ),
              cell(
                [
                  p("Contact", "kicker"),
                  p("{phone}\n{email}\n{web}\n{city}", "small", { after: 12 }),
                  p("Skills", "kicker"),
                  list(["Project planning", "Writing and editing", "Budgets and reporting", "Team leadership", "Public speaking"], { role: "small", gap: 2, after: 12 }),
                  p("Languages", "kicker"),
                  p("English (native)\nSpanish (conversational)", "small", { after: 12 }),
                  p("Interests", "kicker"),
                  p(altShort(2), "small"),
                ],
                { fill: "tint", pad: 12 }
              ),
            ],
          ]
        ),
      ];
    },
  },
];

export const TEMPLATES = [...CORE_TEMPLATES, ...RESUME_TEMPLATES, ...MORE_TEMPLATES];

export function findTemplate(id) {
  return TEMPLATES.find((t) => t.id === id) || null;
}

// ---------------------------------------------------------------------------
// Page parts: building blocks to insert at the cursor
// ---------------------------------------------------------------------------

export const PARTS = [
  {
    id: "heading",
    name: "Heading and text",
    description: "A small label, a heading, and a paragraph.",
    build: () => [p("Section label", "kicker"), p("Section heading", "h1"), p(alt(1, 0))],
  },
  {
    id: "pull-quote",
    name: "Pull quote",
    description: "A large quotation set between two rules.",
    build: (ctx) => [
      rule({ color: "accent", weight: 1.5, before: 6, after: 8 }),
      p(`“${altQuote(0)}”`, "quote", { align: "center" }),
      p("Name of the person quoted", "caption", { align: "center", after: 6 }),
      rule({ color: "accent", weight: 1.5, before: 0, after: 10 }),
    ],
  },
  {
    id: "sidebar",
    name: "Sidebar box",
    description: "A shaded box for tips, facts, or a short list.",
    build: () => [band([p("Did you know?", "h2", { color: "primary" }), p(alt(1, 7), "small"), list(altList(3, 9), { role: "small", gap: 2, after: 0 })], { fill: "tint", pad: 12, valign: "top", align: "left" })],
  },
  {
    id: "coupon",
    name: "Coupon",
    description: "A dashed cut-out box with an offer, terms, and a code.",
    build: (ctx) => [
      grid(
        [Math.min(ctx.W, 340)],
        [[cell([p("Coupon", "kicker", { align: "center" }), p("Save 20% on your next visit", "headline", { align: "center", size: 17 }), p("One per customer. Expires December 31.", "small", { align: "center" }), p("Code: SAVE20", "label", { align: "center", color: "dark", after: 0 })], { pad: 12, border: { all: { color: "dark", weight: 1.5, style: "dashed" } } })]],
        { abs: true, align: "center" }
      ),
      space(8),
    ],
  },
  {
    id: "attention",
    name: "Attention getter",
    description: "A bold colored strip that announces something new.",
    build: () => [band([p("Free admission this weekend", "headline", { color: "onAccent", align: "center", after: 0 })], { fill: "accent", pad: 12 }), space(8)],
  },
  {
    id: "event-details",
    name: "Event details",
    description: "Date, time, and place in three shaded boxes.",
    build: () => [grid([1, 1, 1], [[cell(detail("Date", "Saturday, June 14"), { fill: "tint", pad: 9 }), cell(detail("Time", "6 to 9 PM"), { fill: "tint", pad: 9 }), cell(detail("Place", "Riverside Park"), { fill: "tint", pad: 9 })]], { gap: 8 }), space(8)],
  },
  {
    id: "contact",
    name: "Contact block",
    description: "Your business information in two columns.",
    build: () => [grid([1, 1], [[cell([p("{org}", "h2", { color: "primary" }), p("{address}\n{city}", "small", { after: 0 })]), cell([p("{phone}\n{email}\n{web}", "small", { after: 0 })], { valign: "bottom" })]], { gap: 12 }), space(8)],
  },
  {
    id: "button",
    name: "Call to action",
    description: "A colored button-like strip that tells readers what to do.",
    build: (ctx) => [band([p("Register at {web}", "button", { align: "center" })], { fill: "accent", pad: 10, width: Math.round(ctx.W * 0.6) }), space(8)],
  },
  {
    id: "picture",
    name: "Picture with caption",
    description: "A picture placeholder you can replace, with a caption.",
    build: (ctx) => [pic("full", Math.round(ctx.W * 0.5), { after: 4 }), p(altShort(3), "caption", { after: 8 })],
  },
  {
    id: "logo",
    name: "Logo",
    description: "Your logo from business information, or a placeholder.",
    build: () => [logoBox(140, 60, { after: 8 })],
  },
  {
    id: "two-columns",
    name: "Two columns of text",
    description: "Side by side stories with headings.",
    build: () => [grid([1, 1], [[cell([p("Story headline", "h2", { color: "primary" }), p(alt(1, 2))]), cell([p("Story headline", "h2", { color: "primary" }), p(alt(1, 3))])]], { gap: 18 })],
  },
  {
    id: "schedule",
    name: "Schedule table",
    description: "Times and activities with a colored header row.",
    build: () => [table(["Time", "Activity", "Location"], [["9:00", "Welcome and coffee", "Lobby"], ["9:30", "Opening session", "Main hall"], ["11:00", "Workshops", "Rooms 1 to 3"], ["12:30", "Lunch", "Courtyard"]], { widths: [0.2, 0.5, 0.3], firstBold: true, firstColor: "primary" }), space(8)],
  },
  {
    id: "divider",
    name: "Divider",
    description: "A colored rule to separate sections.",
    build: () => [rule({ color: "primary", weight: 2, before: 6, after: 10 })],
  },
  {
    id: "testimonial",
    name: "Testimonial",
    description: "A short review with the reviewer's name.",
    build: () => [band([p(`“${altQuote(3)}”`, "quote", { size: 13 }), p("Customer name, City", "caption", { after: 0 })], { fill: "accentTint", pad: 12, align: "left" }), space(8)],
  },
];

export function findPart(id) {
  return PARTS.find((x) => x.id === id) || null;
}

/** Plain words for the fold types, used by the print and fold guide. */
const BASE_FOLD_GUIDES = {
  trifold: {
    title: "Folding a tri-fold brochure",
    steps: [
      "Print side 1, then put the sheet back in the printer and print side 2. Use the printer's two-sided setting if it has one, flipping on the short edge.",
      "Lay the sheet with side 2 (the inside) facing up.",
      "Fold the right panel in toward the center so it covers the middle panel.",
      "Fold the left panel over the top. The front cover now faces you.",
    ],
    tip: "Crease each fold with the edge of a ruler. Heavier paper (about 28 lb or 105 g/m²) folds cleanly and feels more professional.",
  },
  halffold: {
    title: "Folding a half-fold publication",
    steps: [
      "Print side 1, then print side 2 on the back. With two-sided printing, flip on the short edge.",
      "Lay the sheet with side 2 (the inside) facing up.",
      "Fold the sheet in half from left to right and crease it. The front cover is on the outside right.",
    ],
    tip: "For a booklet with more pages, print several sheets the same way, nest them, and staple along the fold.",
  },
  cards: {
    title: "Cutting cards and labels",
    steps: [
      "Print at actual size (100%). Turn off any “fit to page” or “shrink” option.",
      "For cards, print on card stock and cut along the light dashed guides with a paper trimmer.",
      "For labels, use matching label sheets. Print one test page on plain paper and hold it over a label sheet against a window to check the alignment.",
    ],
    tip: "If everything is shifted the same way, adjust the page margins by that amount and print again.",
  },
  none: {
    title: "Printing a flat publication",
    steps: ["Print at actual size (100%).", "For the sharpest result, save as PDF and print from the PDF, or send the PDF to a print shop."],
    tip: "Most home printers cannot print to the very edge of the paper. Offizin keeps important content inside a safe margin.",
  },
};

export const FOLD_GUIDES = { ...BASE_FOLD_GUIDES, ...MORE_FOLD_GUIDES };
