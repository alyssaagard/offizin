/**
 * Filler text in Middle High German (about 1150 to 1350), used the way Word and Google Docs use
 * Latin "lorem ipsum": running text that is meant to be replaced. Nobody mistakes it for finished
 * copy, so it is hard to print by accident, and the design checker finds every line of it.
 *
 * The lines come from public domain poems, in the usual normalized spelling:
 *   the Nibelungenlied; Walther von der Vogelweide ("Under der linden", "Ir sult sprechen
 *   willekomen", "Ich saz ûf eime steine", and the Elegy); Hartmann von Aue (Iwein); Wolfram von
 *   Eschenbach (Parzival); Der von Kürenberg (the falcon song); and the Tegernsee love letters
 *   ("Dû bist mîn").
 *
 * Only running text is filler: paragraphs, descriptions, captions, quotes, and bullet points.
 * Headings and labels stay short plain English. Names, emails, phone numbers, addresses, web
 * addresses, dates, times, prices, and other numbers stay plain placeholders, because their
 * format matters (business information tokens such as {phone} fill most of them).
 */

/** Sentences for paragraphs, in a fixed order so publications always look the same. */
export const ALT_DEUTSCH = [
  "Uns ist in alten mæren wunders vil geseit von helden lobebæren, von grôzer arebeit.",
  "Von fröuden, hôchgezîten, von weinen und von klagen muget ir nu wunder hœren sagen.",
  "Ir sult sprechen willekomen: der iu mære bringet, daz bin ich.",
  "Allez, daz ir habt vernomen, daz ist gar ein wint: nû frâget mich.",
  "Under der linden an der heide, dâ muget ir vinden schône beide gebrochen bluomen unde gras.",
  "Vor dem walde in einem tal schône sanc diu nahtegal.",
  "Swer an rehte güete wendet sîn gemüete, dem volget sælde und êre.",
  "Ich saz ûf eime steine und dahte bein mit beine.",
  "Dô dâhte ich mir vil ange, wie man zer welte solte leben.",
  "Ez wuohs in Burgonden ein vil edel magedîn, daz in allen landen niht schœners mohte sîn.",
  "Ich zôch mir einen valken mêre danne ein jâr.",
  "Er huop sich ûf vil hôhe und floug in anderiu lant.",
  "Ist zwîvel herzen nâchgebûr, daz muoz der sêle werden sûr.",
  "Dû bist mîn, ich bin dîn: des solt dû gewis sîn.",
  "Owê war sint verswunden alliu mîniu jâr!",
  "Ist mir mîn leben getroumet, oder ist ez wâr?",
];

/** Short phrases for captions, list items, and other single lines. */
export const ALT_SHORT = [
  "Under der linden an der heide",
  "Ir sult sprechen willekomen",
  "Gebrochen bluomen unde gras",
  "Schône sanc diu nahtegal",
  "Von helden lobebæren",
  "Dem volget sælde und êre",
  "Ich saz ûf eime steine",
  "Vor dem walde in einem tal",
  "Er floug in anderiu lant",
  "Wunders vil geseit",
  "Dû bist mîn, ich bin dîn",
  "Nû frâget mich",
];

/** Lines for pull quotes and testimonials. */
export const ALT_QUOTES = [
  "Swer an rehte güete wendet sîn gemüete, dem volget sælde und êre.",
  "Dû bist mîn, ich bin dîn: des solt dû gewis sîn.",
  "Ir sult sprechen willekomen: der iu mære bringet, daz bin ich.",
  "Ist mir mîn leben getroumet, oder ist ez wâr?",
];

const at = (list, i) => list[((i % list.length) + list.length) % list.length];

/** `count` filler sentences as one paragraph, starting at sentence `start` (wrapping around). */
export function alt(count = 1, start = 0) {
  const n = Math.max(1, Math.floor(count));
  return Array.from({ length: n }, (_, i) => at(ALT_DEUTSCH, start + i)).join(" ");
}

/** One short filler phrase. */
export function altShort(i = 0) {
  return at(ALT_SHORT, i);
}

/** `count` short filler phrases, for bullet points. */
export function altList(count = 3, start = 0) {
  return Array.from({ length: Math.max(1, Math.floor(count)) }, (_, i) => at(ALT_SHORT, start + i));
}

/** One filler quotation. */
export function altQuote(i = 0) {
  return at(ALT_QUOTES, i);
}
