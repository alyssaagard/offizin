/**
 * Small folding and cutting diagrams for the print guide. Colors come from CSS custom
 * properties so they follow the panel's light or dark appearance.
 */

const SHEET = "fill:var(--oz-sheet);stroke:var(--oz-line);stroke-width:1";
const FOLD = "stroke:var(--oz-fold);stroke-width:1.2;stroke-dasharray:4 3";
const ARROW = "fill:none;stroke:var(--oz-accent-strong);stroke-width:1.6";
const TEXT = "font:600 8px system-ui,sans-serif;fill:var(--oz-muted)";

const arrowHead = (x, y) => `<path d="M${x - 3} ${y + 5} L${x} ${y} L${x + 3} ${y + 5}" style="${ARROW}"/>`;

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
  cards: `<svg viewBox="0 0 260 104" role="img" aria-label="A sheet of cards with dashed cut lines between them.">
    <rect x="70" y="4" width="120" height="96" rx="2" style="${SHEET}"/>
    ${[0, 1, 2, 3].map((i) => `<line x1="70" y1="${4 + 19.2 * (i + 1)}" x2="190" y2="${4 + 19.2 * (i + 1)}" style="${FOLD}"/>`).join("")}
    <line x1="130" y1="4" x2="130" y2="100" style="${FOLD}"/>
    <text x="200" y="54" style="${TEXT}">cut on the</text><text x="200" y="64" style="${TEXT}">dashed lines</text>
  </svg>`,
  none: `<svg viewBox="0 0 260 104" role="img" aria-label="A single flat sheet.">
    <rect x="95" y="6" width="70" height="92" rx="2" style="${SHEET}"/>
    <rect x="103" y="14" width="54" height="76" style="fill:none;stroke:var(--oz-fold);stroke-width:1;stroke-dasharray:2 2"/>
    <text x="175" y="54" style="${TEXT}">keep text inside</text><text x="175" y="64" style="${TEXT}">the safe margin</text>
  </svg>`,
};
