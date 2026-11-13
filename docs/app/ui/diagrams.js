/**
 * Small folding and cutting diagrams for the print guide. Colors come from CSS custom
 * properties so they follow the panel's colors. Every diagram has a text alternative.
 */

const SHEET = "fill:var(--oz-sheet);stroke:var(--oz-line);stroke-width:1";
const SHEET_BACK = "fill:var(--oz-accent-wash);stroke:var(--oz-line);stroke-width:1";
const FOLD = "stroke:var(--oz-fold);stroke-width:1.2;stroke-dasharray:4 3";
const CUT = "stroke:var(--oz-fold);stroke-width:1;stroke-dasharray:2 2";
const ARROW = "fill:none;stroke:var(--oz-accent-strong);stroke-width:1.6";
const INK = "stroke:var(--oz-ink-soft);stroke-width:2.4;stroke-linecap:round";
const LINE = "stroke:var(--oz-muted);stroke-width:1.2;stroke-linecap:round";
const TEXT = "font:600 8px system-ui,sans-serif;fill:var(--oz-muted)";
const BIG = "font:700 13px system-ui,sans-serif;fill:var(--oz-ink-soft)";

/** Arrow heads pointing up (at the end of a curved arrow) and to the right. */
const arrowHead = (x, y) => `<path d="M${x - 3} ${y + 5} L${x} ${y} L${x + 3} ${y + 5}" style="${ARROW}"/>`;
const arrowRight = (x, y) => `<path d="M${x - 5} ${y - 3} L${x} ${y} L${x - 5} ${y + 3}" style="${ARROW}"/>`;
const straightArrow = (x1, x2, y) => `<path d="M${x1} ${y} H${x2}" style="${ARROW}"/>${arrowRight(x2, y)}`;

export const DIAGRAMS = {
  trifold: `<svg viewBox="0 0 260 104" role="img" aria-label="A sheet divided into three panels. The right panel folds in first, then the left panel folds over it.">
    <rect x="10" y="10" width="240" height="64" rx="2" style="${SHEET}"/>
    <line x1="90" y1="10" x2="90" y2="74" style="${FOLD}"/><line x1="170" y1="10" x2="170" y2="74" style="${FOLD}"/>
    <text x="50" y="46" text-anchor="middle" style="${TEXT}">Panel 1</text><text x="130" y="46" text-anchor="middle" style="${TEXT}">Panel 2</text><text x="210" y="46" text-anchor="middle" style="${TEXT}">Panel 3</text>
    <path d="M214 78 C214 96 150 96 150 80" style="${ARROW}"/>${arrowHead(150, 78)}<text x="190" y="101" text-anchor="middle" style="${TEXT}">1. fold in</text>
    <path d="M46 78 C46 96 110 96 110 80" style="${ARROW}"/>${arrowHead(110, 78)}<text x="70" y="101" text-anchor="middle" style="${TEXT}">2. fold over</text>
  </svg>`,
  halffold: `<svg viewBox="0 0 260 104" role="img" aria-label="A sheet divided into two halves, folded in half from right to left.">
    <rect x="10" y="10" width="240" height="64" rx="2" style="${SHEET}"/>
    <line x1="130" y1="10" x2="130" y2="74" style="${FOLD}"/>
    <text x="70" y="46" text-anchor="middle" style="${TEXT}">Inside left</text><text x="190" y="46" text-anchor="middle" style="${TEXT}">Inside right</text>
    <path d="M200 78 C200 100 60 100 60 80" style="${ARROW}"/>${arrowHead(60, 78)}<text x="130" y="101" text-anchor="middle" style="${TEXT}">fold in half</text>
  </svg>`,
  booklet: `<svg viewBox="0 0 260 104" role="img" aria-label="A booklet sheet holds two pages on each side, such as page 8 next to page 1. The sheets are folded in half, nested inside each other, and stapled twice along the fold.">
    <rect x="8" y="14" width="120" height="70" rx="2" style="${SHEET}"/>
    <line x1="68" y1="14" x2="68" y2="84" style="${FOLD}"/>
    <text x="38" y="54" text-anchor="middle" style="${BIG}">8</text><text x="98" y="54" text-anchor="middle" style="${BIG}">1</text>
    <text x="68" y="98" text-anchor="middle" style="${TEXT}">sheet 1, front side</text>
    ${straightArrow(136, 160, 49)}
    <rect x="184" y="20" width="56" height="70" rx="2" style="${SHEET_BACK}"/>
    <rect x="180" y="17" width="56" height="70" rx="2" style="${SHEET_BACK}"/>
    <rect x="176" y="14" width="56" height="70" rx="2" style="${SHEET}"/>
    <line x1="176" y1="26" x2="176" y2="36" style="${INK}"/><line x1="176" y1="62" x2="176" y2="72" style="${INK}"/>
    <text x="204" y="54" text-anchor="middle" style="${BIG}">1</text>
    <text x="208" y="100" text-anchor="middle" style="${TEXT}">fold, nest, and staple</text>
  </svg>`,
  tent: `<svg viewBox="0 0 260 104" role="img" aria-label="A place card folded in half across the middle. Only the lower half is printed, so the name reads the right way up when the card stands like a tent.">
    <rect x="18" y="6" width="100" height="86" rx="2" style="${SHEET}"/>
    <line x1="18" y1="49" x2="118" y2="49" style="${FOLD}"/>
    <text x="68" y="31" text-anchor="middle" transform="rotate(180 68 28)" style="${TEXT}">Back</text>
    <text x="68" y="74" text-anchor="middle" style="${TEXT}">Front</text>
    <text x="68" y="102" text-anchor="middle" style="${TEXT}">fold across the middle</text>
    ${straightArrow(128, 152, 49)}
    <line x1="160" y1="88" x2="252" y2="88" style="${LINE}"/>
    <path d="M172 88 L206 18 L240 88" style="fill:var(--oz-accent-wash);stroke:var(--oz-ink-soft);stroke-width:1.4;stroke-linejoin:round"/>
    <line x1="206" y1="18" x2="206" y2="24" style="${FOLD}"/>
    <text x="226" y="58" text-anchor="middle" style="${TEXT}">Front</text>
    <text x="206" y="102" text-anchor="middle" style="${TEXT}">stands on the table</text>
  </svg>`,
  envelope: `<svg viewBox="0 0 260 104" role="img" aria-label="An envelope with the return address in the top left corner, room for a stamp in the top right corner, and the recipient's address in the middle.">
    <rect x="44" y="12" width="172" height="76" rx="3" style="${SHEET}"/>
    <line x1="52" y1="22" x2="92" y2="22" style="${LINE}"/><line x1="52" y1="28" x2="86" y2="28" style="${LINE}"/><line x1="52" y1="34" x2="90" y2="34" style="${LINE}"/>
    <rect x="188" y="18" width="20" height="24" rx="1" style="fill:none;${CUT}"/>
    <line x1="110" y1="54" x2="170" y2="54" style="${INK}"/><line x1="110" y1="62" x2="162" y2="62" style="${INK}"/><line x1="110" y1="70" x2="166" y2="70" style="${INK}"/>
    <text x="72" y="48" text-anchor="middle" style="${TEXT}">return address</text>
    <text x="198" y="52" text-anchor="middle" style="${TEXT}">stamp</text>
    <text x="130" y="100" text-anchor="middle" style="${TEXT}">feed it the way your printer's envelope icon shows</text>
  </svg>`,
  cards: `<svg viewBox="0 0 260 104" role="img" aria-label="A sheet of cards with dashed cut lines between them.">
    <rect x="70" y="4" width="120" height="96" rx="2" style="${SHEET}"/>
    ${[0, 1, 2, 3].map((i) => `<line x1="70" y1="${4 + 19.2 * (i + 1)}" x2="190" y2="${4 + 19.2 * (i + 1)}" style="${FOLD}"/>`).join("")}
    <line x1="130" y1="4" x2="130" y2="100" style="${FOLD}"/>
    <text x="200" y="54" style="${TEXT}">cut on the</text><text x="200" y="64" style="${TEXT}">dashed lines</text>
  </svg>`,
  none: `<svg viewBox="0 0 260 104" role="img" aria-label="A single flat sheet.">
    <rect x="95" y="6" width="70" height="92" rx="2" style="${SHEET}"/>
    <rect x="103" y="14" width="54" height="76" style="fill:none;${CUT}"/>
    <text x="175" y="54" style="${TEXT}">keep text inside</text><text x="175" y="64" style="${TEXT}">the safe margin</text>
  </svg>`,
};
