/**
 * What search engines and link previews show for each public page of the website.
 *
 * `npm run build` writes these into every page between the "seo:start" and "seo:end" comments,
 * and writes docs/sitemap.xml, docs/robots.txt, and docs/site.webmanifest. Edit the text here,
 * never between those comments.
 *
 * Keep each title under about 60 characters and each description between 70 and 160, since
 * search results cut longer ones off. Write descriptions as plain sentences that say what the
 * page is for; search engines ignore keyword lists.
 */

/** The name search results show for the whole site. */
export const SITE_NAME = "Offizin";

/** The picture link previews show (1280 by 640 pixels), and its description for screen readers. */
export const SOCIAL_IMAGE = { path: "assets/brand/offizin-social.png", width: 1280, height: 640, alt: "Offizin: desktop publishing for Word and Google Docs" };

/**
 * Public pages, in sitemap order. `path` is the address under the site ("" is the home page)
 * and `file` is the page inside docs/. `schema` adds structured data: "site" describes the
 * website and the software, and "demo" describes the live demo as a web application.
 */
export const SEO_PAGES = [
  {
    path: "",
    file: "index.html",
    title: "Offizin | Desktop publishing for Word and Google Docs",
    description:
      "Free desktop publishing for Microsoft Word and Google Docs. Make brochures, newsletters, flyers, cards, and resumes from templates. A Publisher alternative.",
    schema: "site",
    priority: "1.0",
  },
  {
    path: "demo/",
    file: "demo/index.html",
    title: "Offizin live demo | Make a brochure in your browser",
    description:
      "Try Offizin in your browser with no sign-up: pick a brochure, flyer, newsletter, or resume, change its colors, fonts, and details, and download a Word file.",
    schema: "demo",
    priority: "0.9",
  },
  {
    path: "word/",
    file: "word/index.html",
    title: "Offizin for Word | Install the free publishing add-in",
    description: "Install Offizin in Microsoft Word on the web, Windows, or Mac, or deploy it to your whole organization from the Microsoft 365 admin center.",
    priority: "0.8",
  },
  {
    path: "google-docs/",
    file: "google-docs/index.html",
    title: "Offizin for Google Docs | Install the free add-on",
    description: "Add Offizin to Google Docs in a few minutes by pasting three files into Apps Script, or by copying a starter document.",
    priority: "0.8",
  },
  {
    path: "support/",
    file: "support/index.html",
    title: "Offizin help | Questions and answers",
    description: "Answers to common questions about Offizin for Word and Google Docs: installing, printing, folding, undo, mail merge, and how to report a problem.",
    priority: "0.6",
  },
  {
    path: "privacy/",
    file: "privacy/index.html",
    title: "Offizin | Privacy policy",
    description: "Offizin has no server and collects no personal data. This policy explains what stays in your document and what stays in your browser.",
    priority: "0.3",
  },
  {
    path: "terms/",
    file: "terms/index.html",
    title: "Offizin | Terms of use",
    description: "The terms for using Offizin, a free and open source add-in for Microsoft Word and Google Docs, released under the MIT License.",
    priority: "0.3",
  },
];
