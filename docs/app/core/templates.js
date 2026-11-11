/**
 * The Offizin publication library and page parts.
 *
 * Each template describes one publication with simple blocks:
 *   p (paragraph), list, rule, spacer, break (new page), image, band, grid, table.
 * Text can use business information fields such as {org} and {phone}, which the composer
 * fills from the user's business information (or friendly samples until it is filled in).
 * "**bold**" marks bold words inside a paragraph.
 *
 * Sizes are in points (72 points = 1 inch). ctx.W and ctx.H are the usable width and height
 * of the page inside its margins. See compose.js for everything the composer accepts.
 */
import { inch, mm, round, MONTHS, WEEKDAYS } from "./util.js";

// ---------------------------------------------------------------------------
// Block helpers
// ---------------------------------------------------------------------------

const p = (text, role = "body", o = {}) => ({ t: "p", text, role, ...o });
const list = (items, o = {}) => ({ t: "list", items, role: "body", ...o });
const rule = (o = {}) => ({ t: "rule", ...o });
const space = (height) => ({ t: "spacer", height });
const pageBreak = () => ({ t: "break" });
const pic = (width, height, o = {}) => ({ t: "image", kind: "photo", width, height, ...o });
const logoBox = (width, height, o = {}) => ({ t: "image", kind: "logo", width, height, ...o });
const band = (blocks, o = {}) => ({ t: "band", blocks, ...o });
const grid = (widths, rows, o = {}) => ({ t: "grid", widths, rows, ...o });
const cell = (blocks, o = {}) => ({ blocks, ...o });
const table = (header, rows, o = {}) => ({ t: "table", header, rows, ...o });

/** A label with the value under it, used for dates, times, and places. */
const detail = (label, value, o = {}) => [p(label, "label", { align: o.align || "center", color: o.labelColor || "accent", fixed: true }), p(value, "h2", { align: o.align || "center", after: 0, color: o.color, size: o.size })];

/** Contact lines from business information. */
const contactLines = (o = {}) => [
  p("{address}\n{city}", "small", { align: o.align, color: o.color, after: 4 }),
  p("{phone}\n{email}\n{web}", "small", { align: o.align, color: o.color, after: 0 }),
];

/** Logo when the user has one, otherwise the organization name set as a wordmark. */
const brandMark = (ctx, w, h, o = {}) =>
  ctx.hasLogo ? logoBox(w, h, { align: o.align || "left", after: o.after ?? 8 }) : p("{org}", "h2", { align: o.align || "left", color: o.color || "primary", after: o.after ?? 6, size: o.size });

const panelFolds = (n) => () => ({ vertical: Array.from({ length: n - 1 }, (_, i) => (i + 1) / n), horizontal: [] });

const PANEL_PAD = 20;
const panelPad = (i, n) => [PANEL_PAD, i === n - 1 ? 4 : PANEL_PAD, PANEL_PAD, i === 0 ? 4 : PANEL_PAD];

/** Sheet layout for cards and labels: where a grid of equal pieces sits on the page. */
function sheetMargins(paper, pieceW, pieceH, across, down, gapX = 0, gapY = 0) {
  const page = paper === "a4" ? { width: mm(210), height: mm(297) } : { width: 612, height: 792 };
  const usedW = across * pieceW + (across - 1) * gapX;
  const usedH = down * pieceH + (down - 1) * gapY;
  const side = round((page.width - usedW) / 2);
  const top = round((page.height - usedH) / 2);
  return { top, bottom: top, left: side, right: side };
}

// ---------------------------------------------------------------------------
// Publications
// ---------------------------------------------------------------------------

export const CATEGORIES = [
  { id: "all", name: "All" },
  { id: "brochures", name: "Brochures" },
  { id: "newsletters", name: "Newsletters" },
  { id: "flyers", name: "Flyers and posters" },
  { id: "cards", name: "Cards and invitations" },
  { id: "programs", name: "Programs and menus" },
  { id: "stationery", name: "Stationery and labels" },
  { id: "calendars", name: "Calendars and certificates" },
];

export const TEMPLATES = [
  {
    id: "trifold",
    name: "Tri-fold brochure",
    category: "brochures",
    description: "Six panels on one sheet, printed on both sides and folded in thirds.",
    sizeLabel: "Landscape, 2 sides, folds in thirds",
    fold: "trifold",
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
                  p("Open with the single most important reason a reader should care. This panel is the first thing people see when they unfold the brochure."),
                  p("“The best sentence a happy client ever said about your work goes here.”", "quote", { size: 13, before: 6 }),
                  p("Name of the person quoted", "caption", { after: 12 }),
                  list(["A quick fact that builds trust", "A number that shows your impact", "A promise you always keep"]),
                ],
                { pad: panelPad(0, 3) }
              ),
              cell(
                [
                  brandMark(ctx, 120, 60, { align: "center" }),
                  p("{org}", "h2", { align: "center", after: 2 }),
                  p("{tagline}", "caption", { align: "center", after: 14 }),
                  ...contactLines({ align: "center" }),
                  space(18),
                  p("Mail or hand this side to your reader.", "caption", { align: "center" }),
                ],
                { pad: panelPad(1, 3), valign: "bottom" }
              ),
              cell(
                [
                  p("{tagline}", "kicker", { color: "onPrimary", align: "center" }),
                  p("{org}", "title", { color: "onPrimary", align: "center", size: 26 }),
                  rule({ color: "accent", weight: 2, width: 60, align: "center", before: 2, after: 12 }),
                  p("A clear promise in one short line", "lead", { color: "onPrimary", align: "center" }),
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
                  p("Tell your story in two or three short paragraphs. Say who you serve, what you believe, and what makes your approach different."),
                  p("Keep sentences short and concrete. Readers skim brochures, so lead each paragraph with its point."),
                  pic("full", 110, { before: 4 }),
                  p("A caption that explains the picture", "caption"),
                ],
                { pad: panelPad(0, 3) }
              ),
              cell(
                [
                  p("What we do", "kicker"),
                  p("Programs and services", "h1"),
                  p("First program", "h2"),
                  p("Describe what it is, who it is for, and the result people can expect."),
                  p("Second program", "h2"),
                  p("Explain the value in plain words. One or two sentences is plenty."),
                  p("Third program", "h2"),
                  p("Close with something that invites the reader to take part."),
                ],
                { pad: panelPad(1, 3) }
              ),
              cell(
                [
                  p("Get involved", "kicker"),
                  p("Join us", "h1"),
                  list(["Volunteer an afternoon a month", "Become a member or sponsor", "Bring a friend to our next event"]),
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
                    p("A brochure title that makes people curious", "title", { color: "onPrimary", size: 26 }),
                    p("A subtitle that says what readers will learn inside", "lead", { color: "onPrimary" }),
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
                    p("The work we do and why we do it", "headline"),
                    p("Use this page to explain the problem you solve and the approach you take. A bi-fold gives you room for a little more depth than a flyer.", "lead"),
                    p("Add a second paragraph with a short story or a result. Specific details make a brochure believable."),
                    p("“A short quote from someone you have helped.”", "quote", { before: 8 }),
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
                          cell([p("First offer", "h2"), p("One or two sentences about it.", "small")], { pad: [0, 8, 6, 0] }),
                          cell([p("Second offer", "h2"), p("One or two sentences about it.", "small")], { pad: [0, 0, 6, 8] }),
                        ],
                        [
                          cell([p("Third offer", "h2"), p("One or two sentences about it.", "small")], { pad: [0, 8, 0, 0] }),
                          cell([p("Fourth offer", "h2"), p("One or two sentences about it.", "small")], { pad: [0, 0, 0, 8] }),
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
                p("A headline that tells readers the news in eight words or fewer", "headline"),
                p("By the author's name", "caption", { after: 8 }),
                pic("full", 170, { after: 4 }),
                p("A caption that adds a fact the picture cannot show", "caption", { after: 8 }),
                p("Start with the most important news. Who, what, when, and why should all appear in the first two sentences, because many readers stop there."),
                p("Then add details in order of importance: a quote, a number, a next step. Short paragraphs look friendlier in narrow columns."),
                p("End with what readers can do, such as signing up, attending, or replying with their ideas."),
              ]),
              cell(
                [
                  p("In this issue", "h2", { color: "primary" }),
                  list(["Lead story", "Member spotlight", "Upcoming dates", "A note from the director"], { role: "small", gap: 2 }),
                  rule({ color: "line", before: 6, after: 8 }),
                  p("Upcoming", "h2", { color: "primary" }),
                  p("**June 14**  Summer concert", "small"),
                  p("**July 2**  Volunteer training", "small"),
                  p("**July 19**  Open studio day", "small", { after: 8 }),
                  rule({ color: "line", before: 0, after: 8 }),
                  p("“A memorable line from this issue, pulled out to catch the eye.”", "quote", { size: 13 }),
                ],
                { fill: "tint", pad: 12 }
              ),
            ],
          ],
          { gap: 18 }
        ),
        pageBreak(),
        grid(
          [1, 1],
          [
            [
              cell([
                p("Member spotlight", "kicker"),
                p("Meet someone who makes it happen", "h1"),
                pic("full", 130, { after: 8 }),
                p("Profile a member, volunteer, or client. Ask three questions and share the best answers in their own words."),
                p("A short profile builds community and gives readers a reason to open the next issue."),
              ]),
              cell([
                p("From the director", "kicker"),
                p("A note on the season ahead", "h1"),
                p("Use this column for a personal message. Thank people, name what is changing, and say what you hope to accomplish together."),
                p("Keep it warm and specific. Mention one achievement and one challenge, and invite readers to help with the challenge."),
                p("Your name", "h2", { before: 6, after: 0 }),
                p("Your title", "caption"),
              ]),
            ],
          ],
          { gap: 18 }
        ),
        space(14),
        grid(
          [0.5, 0.5],
          [
            [
              cell([p("Save the date", "h2", { color: "primary" }), table(null, [["June 14", "Summer concert"], ["July 2", "Volunteer training"], ["July 19", "Open studio day"]], { widths: [0.35, 0.65], firstBold: true, firstColor: "primary" })], { pad: [0, 9, 0, 0] }),
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
    page: { size: "paper", orientation: "portrait", margins: 36 },
    build(ctx) {
      return [
        pic("full", round(ctx.H * 0.36), { after: 16 }),
        p("{org} presents", "kicker", { align: "center" }),
        p("Summer Concert Night", "display", { align: "center" }),
        p("Live music, local food, and a sky full of stars", "lead", { align: "center", after: 16 }),
        grid(
          [1, 1, 1],
          [[cell(detail("Date", "Saturday, June 14"), { fill: "tint", pad: 10 }), cell(detail("Time", "6 to 9 PM"), { fill: "tint", pad: 10 }), cell(detail("Place", "Riverside Park"), { fill: "tint", pad: 10 })]],
          { gap: 10 }
        ),
        space(14),
        p("Bring a blanket and your friends. Admission is free for members and children under twelve. Food trucks open at 5:30, and the first act starts at six.", "body", { align: "center", size: 11 }),
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
    page: { size: "paper", orientation: "portrait", margins: 36 },
    build(ctx) {
      return [
        band(
          [
            p("{org}", "kicker", { color: "onPrimary", align: "center" }),
            p("The Big Spring Sale", "display", { color: "onPrimary", align: "center" }),
            p("Three days only. Everything in the shop is marked down.", "lead", { color: "onPrimary", align: "center", after: 0 }),
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
                  p("Bring this coupon to the register. One coupon per customer. Expires at closing on Sunday.", "small", { align: "center" }),
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
    page: (paper) => ({ size: paper === "a4" ? "a3" : "tabloid", orientation: "portrait", margins: 48 }),
    build(ctx) {
      return [
        p("{org}", "kicker", { size: 13, align: "center" }),
        p("Community Clean-Up Day", "display", { size: 72, align: "center", after: 14 }),
        pic("full", round(ctx.H * 0.42), { after: 20 }),
        p("Gloves, bags, and coffee provided. Bring your neighbors.", "lead", { size: 22, align: "center", after: 22 }),
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
    page: { size: "paper", orientation: "landscape", margins: 40 },
    build(ctx) {
      return [
        band([p("{org}", "h2", { color: "onPrimary", align: "center", after: 0 })], { fill: "primary", pad: 10 }),
        space(round(ctx.H * 0.14)),
        p("Please use the side entrance", "display", { size: 54, align: "center", color: "dark", after: 16 }),
        p("The main doors are closed for repairs until Friday. Thank you for your patience.", "lead", { size: 18, align: "center" }),
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
    page: (paper) => ({ size: "paper", orientation: "portrait", margins: sheetMargins(paper, 432, 288, 1, 2) }),
    folds: () => ({ vertical: [], horizontal: [], cut: true }),
    build(ctx) {
      const guide = { all: { color: "line", weight: 0.25, style: "dashed" } };
      const front = () =>
        cell(
          [
            p("{org}", "kicker", { color: "onPrimary" }),
            p("You are invited", "title", { color: "onPrimary", size: 30 }),
            p("Our open house is on Saturday, October 4, from 1 to 4 PM.", "lead", { color: "onPrimary" }),
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
                      p("Come see what is new", "h2", { color: "primary" }),
                      p("Tour the studios, meet the teachers, and try a free mini class. Refreshments all afternoon.", "small"),
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
                cell([brandMark(ctx, 90, 40, { align: "center", after: 4 }), p("Made with care by {org}", "caption", { align: "center" })], { pad: panelPad(0, 2), valign: "bottom" }),
                cell(
                  [
                    pic("full", round(h * 0.55), { after: 22 }),
                    p("With Gratitude", "title", { align: "center", color: "primary" }),
                    rule({ color: "accent", weight: 1.5, width: 50, align: "center", before: 4, after: 10 }),
                    p("for everything you do", "lead", { align: "center", italic: true, color: "muted" }),
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
                cell([p("“Gratitude turns what we have into enough.”", "quote", { align: "center", size: 16 }), p("A favorite saying", "caption", { align: "center" })], { pad: panelPad(0, 2), valign: "middle" }),
                cell(
                  [
                    p("Dear friend,", "h2", { color: "primary", after: 10 }),
                    p("Write your message here. A few warm, specific sentences mean more than a long letter. Mention a moment you remember and what it meant to you.", "lead"),
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
                    p("An Evening of Celebration", "title", { align: "center", size: 26 }),
                    rule({ color: "accent", weight: 1.5, width: 50, align: "center", before: 6, after: 14 }),
                    p("Join us as we mark ten wonderful years together", "lead", { align: "center", italic: true, size: 11 }),
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
                    p("This event is possible because of our sponsors, volunteers, and friends.", "body", { align: "center", after: 12 }),
                    p("Sponsor name\nSponsor name\nSponsor name", "body", { align: "center", bold: true, after: 22 }),
                    brandMark(ctx, 120, 50, { align: "center" }),
                    ...contactLines({ align: "center" }),
                  ],
                  { pad: panelPad(0, 2), valign: "middle" }
                ),
                cell(
                  [
                    p("{org} presents", "kicker", { align: "center", color: "onPrimary" }),
                    p("Annual Awards Evening", "title", { align: "center", color: "onPrimary" }),
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
                    p("A note before we begin", "h1"),
                    p("Thank you for joining us. Tonight we celebrate the people who made this year remarkable, and we look ahead to the next one."),
                    p("Please silence your phones during the ceremony. Photographs are welcome during the reception."),
                    p("Your name, Your title", "caption", { after: 14 }),
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
                    p("Award title and one line about why", "small", { after: 8 }),
                    p("Honoree name", "h2", { after: 0 }),
                    p("Award title and one line about why", "small", { after: 8 }),
                    p("Honoree name", "h2", { after: 0 }),
                    p("Award title and one line about why", "small", { after: 14 }),
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
        brandMark(ctx, 150, 60, { align: "center", after: 2 }),
        p("{org}", "title", { align: "center", after: 2 }),
        p("{tagline}", "caption", { align: "center", after: 14 }),
        ...section("To start", [
          ["Garden soup", "Seasonal vegetables, fresh herbs, toasted bread", "8"],
          ["Warm flatbread", "Olive oil, sea salt, house dips", "9"],
          ["Market salad", "Greens, citrus, shaved fennel, toasted seeds", "11"],
        ]),
        ...section("Mains", [
          ["Roast chicken", "Herb butter, crushed potatoes, greens", "22"],
          ["River trout", "Brown butter, almonds, lemon, green beans", "24"],
          ["Mushroom risotto", "Parmesan, thyme, wild mushrooms", "19"],
        ]),
        ...section("To finish", [
          ["Apple crumble", "Oat topping, vanilla cream", "9"],
          ["Chocolate pot", "Dark chocolate, sea salt, shortbread", "9"],
        ]),
        space(6),
        p("Please tell us about any allergies. Prices include tax.", "caption", { align: "center" }),
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
        p("Type your letter here. Your letterhead keeps every letter consistent and professional, and it uses the same colors and fonts as the rest of your publications.", "body", { size: 10.5, line: 1.4, after: 10 }),
        p("A second paragraph can give details, answer a question, or describe next steps. Keep paragraphs short so the letter is easy to read.", "body", { size: 10.5, line: 1.4, after: 10 }),
        p("Thank you for your time. I look forward to hearing from you.", "body", { size: 10.5, line: 1.4, after: 18 }),
        p("Sincerely,", "body", { after: 30 }),
        p("{name}", "body", { bold: true, after: 0 }),
        p("{title}, {org}", "body", { color: "muted" }),
      ];
    },
  },

  {
    id: "labels",
    name: "Address labels",
    category: "stationery",
    description: "Return address labels: 30 per Letter sheet (2.625 x 1 in) or 21 per A4 sheet (63.5 x 38.1 mm).",
    sizeLabel: "30 per sheet (Letter), 21 per sheet (A4)",
    fold: "cards",
    page: (paper) =>
      paper === "a4"
        ? { size: "paper", orientation: "portrait", margins: { top: round(mm(15.15)), bottom: round(mm(15.15)), left: round(mm(7.21)), right: round(mm(7.21)) } }
        : { size: "paper", orientation: "portrait", margins: { top: 36, bottom: 36, left: 13.5, right: 13.5 } },
    folds: () => ({ vertical: [], horizontal: [], cut: true }),
    build(ctx) {
      const a4 = ctx.paper === "a4";
      const w = a4 ? round(mm(63.5)) : 189;
      const h = a4 ? round(mm(38.1)) : 72;
      const gap = a4 ? round(mm(2.54)) : 9;
      const down = a4 ? 7 : 10;
      const label = () =>
        cell([p("{org}", "small", { bold: true, color: "primary", after: 1 }), p("{address}\n{city}", "small", { after: 0, line: 1.2 })], { pad: [4, 10, 4, 12], valign: "middle" });
      const rows = Array.from({ length: down }, () => ({ height: h, exact: true, cells: [label(), label(), label()] }));
      return [grid([w, w, w], rows, { abs: true, gap, align: "center" })];
    },
  },

  {
    id: "calendar",
    name: "Monthly calendar",
    category: "calendars",
    description: "A month at a glance, generated for any month and year.",
    sizeLabel: "Landscape, 1 page",
    fold: "none",
    page: { size: "paper", orientation: "landscape", margins: 30 },
    options: [
      { key: "month", label: "Month", type: "month", default: () => new Date().getMonth() + 1 },
      { key: "year", label: "Year", type: "number", default: () => new Date().getFullYear() },
    ],
    build(ctx) {
      const month = Math.min(12, Math.max(1, Number(ctx.opt.month) || 1));
      const year = Number(ctx.opt.year) || new Date().getFullYear();
      ctx.setVar("calMonth", MONTHS[month - 1]);
      ctx.setVar("calYear", year);
      const first = new Date(year, month - 1, 1).getDay();
      const days = new Date(year, month, 0).getDate();
      const weeks = Math.ceil((first + days) / 7);
      const headerH = 62;
      const dayRowH = 20;
      const rowH = round((ctx.H - headerH - dayRowH - 16) / weeks);
      const border = { all: { color: "line", weight: 0.5 } };
      const rows = [
        { cells: WEEKDAYS.map((d) => cell([p(d, "label", { color: "onPrimary", align: "center", after: 0, fixed: true })], { fill: "primary", pad: 5 })) },
      ];
      let n = 1 - first;
      for (let wk = 0; wk < weeks; wk++) {
        const cells = [];
        for (let d = 0; d < 7; d++, n++) {
          const inMonth = n >= 1 && n <= days;
          cells.push(cell(inMonth ? [p(String(n), "h2", { color: d === 0 ? "accent" : "dark", after: 0, size: 11, fixed: true })] : [], { pad: [4, 6, 4, 6], border, fill: inMonth ? null : "tint" }));
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
    page: { size: "paper", orientation: "landscape", margins: 30 },
    build(ctx) {
      const inner = [
        p("{org}", "kicker", { align: "center", after: 14 }),
        p("Certificate of Achievement", "title", { align: "center", size: 34 }),
        rule({ color: "accent", weight: 1.5, width: 80, align: "center", before: 4, after: 16 }),
        p("This certificate is proudly presented to", "body", { align: "center", italic: true, color: "muted", after: 10 }),
        p("Recipient Name", "display", { align: "center", size: 40, color: "accent", italic: true, after: 12 }),
        p("in recognition of outstanding dedication, leadership, and service during the past year.", "body", { align: "center", size: 12, after: 30 }),
        grid(
          [1, 0.3, 1],
          [
            [
              cell([rule({ color: "dark", weight: 0.75, before: 0, after: 4 }), p("{name}, {title}", "small", { align: "center", after: 0 })]),
              cell([]),
              cell([rule({ color: "dark", weight: 0.75, before: 0, after: 4 }), p("Date", "small", { align: "center", after: 0, fixed: true })]),
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
    },
  },

  {
    id: "resume",
    name: "Resume",
    category: "stationery",
    description: "A clean one page resume with a sidebar for skills and contact details.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
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
                  p("Two or three sentences about the work you do best and the value you bring. Write it for the job you want next.", "body", { after: 12 }),
                  p("Experience", "kicker"),
                  ...job("Current role", "Organization", "2022 to now", "Describe a result you achieved, with a number if you can. Then name the skills it shows."),
                  ...job("Previous role", "Organization", "2018 to 2022", "Lead with outcomes, not duties. Keep each description to two or three lines."),
                  ...job("Earlier role", "Organization", "2015 to 2018", "Older roles can be shorter. One strong line is enough."),
                  p("Education", "kicker", { before: 4 }),
                  p("Degree or certificate", "h2", { after: 0 }),
                  p("School name  |  Year", "caption"),
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
                  p("Community theater, trail running, and bread baking", "small"),
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
    build: () => [p("Section label", "kicker"), p("A heading for this section", "h1"), p("Write a short paragraph here. Lead with the point, then add one or two supporting details.")],
  },
  {
    id: "pull-quote",
    name: "Pull quote",
    description: "A large quotation set between two rules.",
    build: (ctx) => [
      rule({ color: "accent", weight: 1.5, before: 6, after: 8 }),
      p("“A memorable sentence from the story, pulled out to catch a reader's eye.”", "quote", { align: "center" }),
      p("Name of the person quoted", "caption", { align: "center", after: 6 }),
      rule({ color: "accent", weight: 1.5, before: 0, after: 10 }),
    ],
  },
  {
    id: "sidebar",
    name: "Sidebar box",
    description: "A shaded box for tips, facts, or a short list.",
    build: () => [band([p("Did you know?", "h2", { color: "primary" }), p("Use a sidebar for helpful extras that do not belong in the main text.", "small"), list(["A quick tip", "A useful fact", "A related link"], { role: "small", gap: 2, after: 0 })], { fill: "tint", pad: 12, valign: "top", align: "left" })],
  },
  {
    id: "coupon",
    name: "Coupon",
    description: "A dashed cut-out box with an offer, terms, and a code.",
    build: (ctx) => [
      grid(
        [Math.min(ctx.W, 340)],
        [[cell([p("Coupon", "kicker", { align: "center" }), p("Save 20% on your next visit", "headline", { align: "center", size: 17 }), p("Present this coupon at checkout. One per customer. Expires December 31.", "small", { align: "center" }), p("Code: SAVE20", "label", { align: "center", color: "dark", after: 0 })], { pad: 12, border: { all: { color: "dark", weight: 1.5, style: "dashed" } } })]],
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
    build: (ctx) => [band([p("Register today at {web}", "button", { align: "center" })], { fill: "accent", pad: 10, width: Math.round(ctx.W * 0.6) }), space(8)],
  },
  {
    id: "picture",
    name: "Picture with caption",
    description: "A picture placeholder you can replace, with a caption.",
    build: (ctx) => [pic("full", Math.round(ctx.W * 0.5), { after: 4 }), p("A caption that adds a fact the picture cannot show", "caption", { after: 8 })],
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
    build: () => [grid([1, 1], [[cell([p("First story", "h2", { color: "primary" }), p("Write a short story or update here. Two columns make short items easy to scan.")]), cell([p("Second story", "h2", { color: "primary" }), p("Write a second item here. Keep both columns about the same length.")])]], { gap: 18 })],
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
    build: () => [band([p("“The kindest, most helpful team we have worked with. We tell everyone about them.”", "quote", { size: 13 }), p("Customer name, City", "caption", { after: 0 })], { fill: "accentTint", pad: 12, align: "left" }), space(8)],
  },
];

export function findPart(id) {
  return PARTS.find((x) => x.id === id) || null;
}

/** Plain words for the fold types, used by the print and fold guide. */
export const FOLD_GUIDES = {
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
