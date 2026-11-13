/**
 * Builds the files that are generated from the source in docs/app and google-docs:
 *
 *   docs/word/manifest.xml          the Word add-in manifest, from scripts/site.config.mjs
 *   docs/google-docs/Sidebar.html   the whole panel bundled into one file for Apps Script
 *   docs/google-docs/Code.gs        copied from google-docs/Code.gs
 *   docs/google-docs/appsscript.json
 *   docs/version.json
 *   docs/download/offizin-word.zip         the Word package: the manifest and a plain text guide
 *   docs/download/offizin-google-docs.zip  the Google Docs package: the three Apps Script files
 *                                          and a plain text guide
 *   docs/sitemap.xml, docs/robots.txt, docs/site.webmanifest, and the search and sharing tags
 *   in the head of every public page, from scripts/seo.config.mjs
 *
 * Usage: npm run build   (Node 18 or later; no packages needed)
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { SITE_URL, ADDIN_ID, VERSION, PROVIDER, REPO_URL, GOOGLE_STARTER_COPY_URL, RELEASE_DATE } from "./site.config.mjs";
import { SITE_NAME, SOCIAL_IMAGE, SEO_PAGES } from "./seo.config.mjs";
import { zip } from "../docs/app/render/zip.js";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_SITE = "https://offizin.alyssaagard.dev";
const rel = (p) => path.relative(ROOT, p);

function write(file, text) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, text);
  console.log("  wrote " + rel(file));
}

// ---------------------------------------------------------------------------
// A tiny bundler for Offizin's own ES modules
// ---------------------------------------------------------------------------

/**
 * Apps Script sidebars cannot load modules, so the sidebar gets one script. Each module is
 * wrapped in its own function scope, its imports become reads from the modules it depends on,
 * and its exports become the scope's return value. This only needs to handle the plain
 * `import { a, b } from "./x.js"` and `export function/const/class` forms Offizin uses.
 */
export function bundle(entry) {
  const order = [];
  const seen = new Set();
  const ids = new Map();
  const idFor = (file) => {
    if (!ids.has(file)) ids.set(file, "__oz_" + rel(file).replace(/^docs\/app\//, "").replace(/[^A-Za-z0-9]/g, "_"));
    return ids.get(file);
  };
  const importRe = /^import\s*\{([^}]*)\}\s*from\s*"([^"]+)";?[ \t]*$/gm;

  function visit(file) {
    if (seen.has(file)) return;
    seen.add(file);
    const src = fs.readFileSync(file, "utf8");
    for (const m of src.matchAll(importRe)) visit(path.resolve(path.dirname(file), m[2]));
    order.push(file);
  }
  visit(entry);

  const parts = order.map((file) => {
    let src = fs.readFileSync(file, "utf8");
    if (/^export\s+default\b/m.test(src)) throw new Error(`${rel(file)}: default exports are not supported by the bundler`);
    if (/^import\s+(?!\{)/m.test(src)) throw new Error(`${rel(file)}: only named imports are supported by the bundler`);
    src = src.replace(importRe, (all, names, from) => {
      const target = idFor(path.resolve(path.dirname(file), from));
      const fields = names
        .split(",")
        .map((n) => n.trim())
        .filter(Boolean)
        .map((n) => n.replace(/\s+as\s+/, ": "));
      return `const { ${fields.join(", ")} } = ${target};`;
    });
    const exported = [];
    src = src.replace(/^export\s+((?:async\s+)?(?:function\*?|const|let|class))\s+([A-Za-z_$][\w$]*)/gm, (all, kind, name) => {
      exported.push(name);
      return `${kind} ${name}`;
    });
    src = src.replace(/^export\s*\{([^}]*)\};?[ \t]*$/gm, (all, names) => {
      names.split(",").map((n) => n.trim()).filter(Boolean).forEach((n) => exported.push(n));
      return "";
    });
    if (/^\s*export\b/m.test(src)) throw new Error(`${rel(file)}: an export form the bundler does not understand`);
    return `// ${rel(file)}\nconst ${idFor(file)} = (() => {\n${src}\nreturn { ${exported.join(", ")} };\n})();`;
  });
  return `(function () {\n"use strict";\n${parts.join("\n\n")}\n})();`;
}

// ---------------------------------------------------------------------------
// Word manifest
// ---------------------------------------------------------------------------

const BUTTONS = [
  { id: "Create", tab: "create", label: "Publications", tip: "Choose a brochure, newsletter, flyer, card, or other publication and build it in your colors." },
  { id: "Design", tab: "design", label: "Design", tip: "Pick color and font schemes, set up the page, and restyle a publication." },
  { id: "Insert", tab: "insert", label: "Page Parts", tip: "Insert pull quotes, sidebars, coupons, event details, and other building blocks at the cursor." },
  { id: "Info", tab: "info", label: "Business Info", tip: "Enter your name, organization, contact details, and logo once and use them everywhere." },
  { id: "Check", tab: "check", label: "Check and Print", tip: "Find leftover sample text and pictures, and get folding and printing help." },
];

export function manifestXml(site = SITE_URL) {
  const icon = (name, size) => `${site}/assets/icons/${name}-${size}.png`;
  const images = [];
  const iconRefs = (name) => {
    for (const size of [16, 32, 80]) images.push(`<bt:Image id="Icon.${name}.${size}" DefaultValue="${icon(name === "Offizin" ? "icon" : "ribbon-" + name.toLowerCase(), size)}"/>`);
    return `<Icon><bt:Image size="16" resid="Icon.${name}.16"/><bt:Image size="32" resid="Icon.${name}.32"/><bt:Image size="80" resid="Icon.${name}.80"/></Icon>`;
  };
  const groupIcon = iconRefs("Offizin");
  const controls = BUTTONS.map(
    (b) => `
              <Control xsi:type="Button" id="Offizin.${b.id}">
                <Label resid="Label.${b.id}"/>
                <Supertip><Title resid="Label.${b.id}"/><Description resid="Tip.${b.id}"/></Supertip>
                ${iconRefs(b.id)}
                <Action xsi:type="ShowTaskpane">
                  <TaskpaneId>OffizinPane</TaskpaneId>
                  <SourceLocation resid="Url.${b.id}"/>
                </Action>
              </Control>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Offizin for Microsoft Word. Generated by scripts/build.mjs from scripts/site.config.mjs. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_ID}</Id>
  <Version>${VERSION}.0</Version>
  <ProviderName>${PROVIDER}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Offizin"/>
  <Description DefaultValue="Desktop publishing in Word: brochures, newsletters, flyers, cards, and more, built from templates in your own colors and fonts."/>
  <IconUrl DefaultValue="${icon("icon", 32)}"/>
  <HighResolutionIconUrl DefaultValue="${icon("icon", 64)}"/>
  <SupportUrl DefaultValue="${site}/support/"/>
  <AppDomains>
    <AppDomain>${site}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="WordApi" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${site}/word/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Url.LearnMore"/>
          </GetStarted>
          <FunctionFile resid="Url.Commands"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Offizin.Tab">
              <Group id="Offizin.Group">
                <Label resid="Label.Group"/>
                ${groupIcon}${controls}
              </Group>
              <Label resid="Label.Tab"/>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        ${images.join("\n        ")}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Url.LearnMore" DefaultValue="${site}/word/"/>
        <bt:Url id="Url.Commands" DefaultValue="${site}/word/commands.html"/>
        ${BUTTONS.map((b) => `<bt:Url id="Url.${b.id}" DefaultValue="${site}/word/taskpane.html?tab=${b.tab}"/>`).join("\n        ")}
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Label.Tab" DefaultValue="Offizin"/>
        <bt:String id="Label.Group" DefaultValue="Publishing"/>
        <bt:String id="GetStarted.Title" DefaultValue="Offizin is ready"/>
        ${BUTTONS.map((b) => `<bt:String id="Label.${b.id}" DefaultValue="${b.label}"/>`).join("\n        ")}
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Open the Offizin tab and choose Publications to make your first brochure, flyer, or newsletter."/>
        ${BUTTONS.map((b) => `<bt:String id="Tip.${b.id}" DefaultValue="${b.tip}"/>`).join("\n        ")}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

// ---------------------------------------------------------------------------
// Download packages
// ---------------------------------------------------------------------------

/** The guide inside the Word package. Plain text, so it opens anywhere. */
export function wordReadme(site = SITE_URL) {
  return `Offizin ${VERSION} for Microsoft Word
Desktop publishing inside Word: brochures, newsletters, flyers, cards, labels,
resumes, and more. By ${PROVIDER}. ${site}

WHAT IS IN THIS PACKAGE

  offizin-manifest.xml  The add-in manifest. It tells Word where Offizin
                        lives and which buttons to add. It holds no program
                        code. Offizin has no server, no account, and no
                        tracking.
  README.txt            This guide.

Word installs add-ins from the manifest file itself. Unzip this package first,
then give Word the file offizin-manifest.xml (not the zip).

WORD ON THE WEB (personal, work, or school accounts)

  1. Open any document in Word on the web at office.com.
  2. On the Home tab, select Add-ins, then More Settings (or More Add-ins).
  3. Choose Upload My Add-in, browse to offizin-manifest.xml, and select
     Upload.
  4. The Offizin tab appears in the ribbon. Select Publications to start.

WORD FOR MAC

  1. Quit Word.
  2. In Finder, choose Go, then Go to Folder, and open this folder (create the
     wef folder if it does not exist):
     ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
  3. Move offizin-manifest.xml into the wef folder.
  4. Open Word and any document. Select Home, then Add-ins, and choose Offizin.

WORD FOR WINDOWS

  Word for Windows installs single add-ins from a shared folder.
  1. Put offizin-manifest.xml in a folder, for example C:\\OffizinAddin. In the
     folder's Properties, open Sharing, select Share, share it with yourself,
     and copy its network path, such as \\\\YOUR-PC\\OffizinAddin.
  2. In Word, choose File, Options, Trust Center, Trust Center Settings, then
     Trusted Add-in Catalogs. Paste the network path into Catalog Url, select
     Add catalog, check Show in Menu, and select OK. Restart Word.
  3. Select Home, then Add-ins, then Advanced (or More Add-ins). Open Shared
     Folder, choose Offizin, and select Add.

A WHOLE ORGANIZATION (Microsoft 365 administrators)

  1. In the Microsoft 365 admin center, go to Settings, then Integrated apps.
  2. Select Upload custom apps and choose Office Add-in as the app type.
  3. Upload offizin-manifest.xml from this package, or provide the link
     ${site}/word/manifest.xml instead.
  4. Choose who gets Offizin (just you, chosen groups, or everyone). Select
     Next, accept the permission request if Microsoft asks, select Next again,
     and then select Finish deployment.
  5. Offizin usually appears in Word on the web, Windows, and Mac within a day,
     but Microsoft says it can take up to 72 hours, and people may need to
     restart Word.

UPDATES AND HELP

  Word loads Offizin from ${site} each time it opens, so there is nothing to
  reinstall when a new version comes out. Help: ${site}/support/
`;
}

/** The guide inside the Google Docs package. */
export function googleReadme(site = SITE_URL) {
  return `Offizin ${VERSION} for Google Docs
Desktop publishing inside Google Docs: brochures, newsletters, flyers, cards,
labels, resumes, and more. By ${PROVIDER}. ${site}

WHAT IS IN THIS PACKAGE

  Code.gs          The Offizin menu and the code that builds publications.
  Sidebar.html     The Offizin panel.
  appsscript.json  The project settings. Offizin asks only for access to the
                   document it is in.
  README.txt       This guide.

Google Docs has no upload box for add-ins outside the Google Workspace
Marketplace, so these files go into a document's own Apps Script project. You
do this once, in a starter document, and then copy that document.

ADD OFFIZIN TO A DOCUMENT (about three minutes)

  1. Open a Google Doc, or a new one at docs.new. If it is pageless, choose
     File, Page setup, then Pages.
  2. Choose Extensions, then Apps Script. Rename the project Offizin.
  3. Open Code.gs, select everything, and replace it with the contents of
     Code.gs from this package.
  4. Next to Files, select the plus sign, then HTML, and name the file Sidebar.
     Replace everything in it with the contents of Sidebar.html.
  5. Open Project Settings (the gear icon), check Show "appsscript.json"
     manifest file in editor, then open appsscript.json in the editor and
     replace everything in it with the contents of appsscript.json.
  6. Save, close the Apps Script tab, and reload the document. Choose Offizin,
     then Open Offizin, and allow access when Google asks. If Google lists the
     two permissions with checkboxes, check both; Offizin needs both.

  Keep this document empty as your starter. Each time you need a publication,
  choose File, then Make a copy: Google copies Offizin along with the document.
  Each copy asks for permission the first time Offizin opens in it. Settings
  and business details stay with the document they were entered in, so save a
  brand kit once (Info tab, Save brand kit) and open it in each copy.

UPLOAD WITH CLASP (for IT teams and developers)

  Google's command line tool, clasp, can upload these files for you.
  1. Turn on the Google Apps Script API at script.google.com/home/usersettings.
  2. Install and sign in:  npm install -g @google/clasp   then   clasp login
  3. In an empty folder, create a starter document with its own project:
     clasp create --type docs --title "Offizin starter"
  4. Copy Code.gs, Sidebar.html, and appsscript.json from this package into
     that folder. If clasp made a Code.js file, delete it, and let this
     package's appsscript.json replace the one clasp made.
  5. Run  clasp push --force  and open the new document from Google Drive.

UPDATES AND HELP

  Copies made this way do not update themselves. When a new version comes out,
  paste the new files into your starter document (or run clasp push again);
  documents copied from it afterward get the new version.
  Help: ${site}/support/
`;
}

/** Fixed timestamps (the release date), so a package only changes when its files do. */
function releaseDate() {
  const d = new Date(`${RELEASE_DATE}T12:00:00`);
  return Number.isNaN(d.getTime()) ? new Date(2026, 0, 1, 12) : d;
}

export function wordPackage(manifest, site = SITE_URL) {
  return zip(
    [
      { name: "offizin-manifest.xml", data: manifest },
      { name: "README.txt", data: wordReadme(site) },
    ],
    { date: releaseDate() }
  );
}

export function googlePackage(files, site = SITE_URL) {
  return zip(
    [
      { name: "Code.gs", data: files.code },
      { name: "Sidebar.html", data: files.sidebar },
      { name: "appsscript.json", data: files.manifest },
      { name: "README.txt", data: googleReadme(site) },
    ],
    { date: releaseDate() }
  );
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Search engines and link previews
// ---------------------------------------------------------------------------

const SEO_START = "<!-- seo:start (written by npm run build from scripts/seo.config.mjs; edit that file) -->";
const SEO_END = "<!-- seo:end -->";
const attr = (text) => String(text).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Structured data (schema.org) for a page, as an object, or null. */
export function seoSchema(page, site = SITE_URL) {
  const author = { "@type": "Person", "@id": `${site}/#author`, name: PROVIDER };
  const free = { "@type": "Offer", price: "0", priceCurrency: "USD" };
  if (page.schema === "site") {
    return {
      "@context": "https://schema.org",
      "@graph": [
        { "@type": "WebSite", "@id": `${site}/#website`, url: `${site}/`, name: SITE_NAME, description: page.description, inLanguage: "en", publisher: { "@id": `${site}/#author` } },
        author,
        {
          "@type": "SoftwareApplication",
          "@id": `${site}/#software`,
          name: SITE_NAME,
          description: page.description,
          url: `${site}/`,
          image: `${site}/assets/icons/icon-512.png`,
          applicationCategory: "DesignApplication",
          operatingSystem: "Microsoft Word for Windows, Mac, and the web; Google Docs",
          softwareVersion: VERSION,
          offers: free,
          isAccessibleForFree: true,
          license: "https://opensource.org/licenses/MIT",
          author: { "@id": `${site}/#author` },
          sameAs: [REPO_URL],
        },
      ],
    };
  }
  if (page.schema === "demo") {
    return {
      "@context": "https://schema.org",
      "@type": "WebApplication",
      name: `${SITE_NAME} live demo`,
      description: page.description,
      url: `${site}/${page.path}`,
      applicationCategory: "DesignApplication",
      operatingSystem: "Any",
      browserRequirements: "Requires JavaScript",
      offers: free,
      isAccessibleForFree: true,
      isPartOf: { "@id": `${site}/#website` },
      author: { "@type": "Person", name: PROVIDER },
    };
  }
  return null;
}

/** The tags between the seo comments in a page's head: title, description, address, icons, and previews. */
export function seoBlock(page, site = SITE_URL) {
  const up = "../".repeat(page.path.split("/").filter(Boolean).length);
  const url = `${site}/${page.path}`;
  const schema = seoSchema(page, site);
  const lines = [
    SEO_START,
    `<title>${attr(page.title)}</title>`,
    `<meta name="description" content="${attr(page.description)}">`,
    `<meta name="author" content="${attr(PROVIDER)}">`,
    `<meta name="robots" content="index, follow, max-image-preview:large">`,
    `<link rel="canonical" href="${url}">`,
    `<meta property="og:type" content="website">`,
    `<meta property="og:site_name" content="${attr(SITE_NAME)}">`,
    `<meta property="og:title" content="${attr(page.title)}">`,
    `<meta property="og:description" content="${attr(page.description)}">`,
    `<meta property="og:url" content="${url}">`,
    `<meta property="og:image" content="${site}/${SOCIAL_IMAGE.path}">`,
    `<meta property="og:image:width" content="${SOCIAL_IMAGE.width}">`,
    `<meta property="og:image:height" content="${SOCIAL_IMAGE.height}">`,
    `<meta property="og:image:alt" content="${attr(SOCIAL_IMAGE.alt)}">`,
    `<meta name="twitter:card" content="summary_large_image">`,
    `<meta name="theme-color" content="#F4F8FC">`,
    `<link rel="icon" href="${up}favicon.ico" sizes="any">`,
    `<link rel="icon" href="${up}favicon.svg" type="image/svg+xml">`,
    `<link rel="icon" href="${up}assets/icons/icon-96.png" type="image/png" sizes="96x96">`,
    `<link rel="icon" href="${up}assets/icons/icon-192.png" type="image/png" sizes="192x192">`,
    `<link rel="apple-touch-icon" href="${up}apple-touch-icon.png">`,
    `<link rel="manifest" href="${up}site.webmanifest">`,
  ];
  if (schema) lines.push(`<script type="application/ld+json">${JSON.stringify(schema).replace(/</g, "\\u003c")}</script>`);
  lines.push(SEO_END);
  return lines.join("\n  ");
}

/** A page with its seo block replaced by a fresh one. The page must already have the two comments. */
export function applySeo(html, page, site = SITE_URL) {
  const a = html.indexOf("<!-- seo:start");
  const b = html.indexOf(SEO_END);
  if (a < 0 || b < a) throw new Error(`${page.file} has no seo:start and seo:end comments in its head`);
  return html.slice(0, a) + seoBlock(page, site) + html.slice(b + SEO_END.length);
}

export function sitemapXml(site = SITE_URL) {
  const urls = SEO_PAGES.map((p) => `  <url>\n    <loc>${site}/${p.path}</loc>\n    <lastmod>${RELEASE_DATE}</lastmod>\n    <priority>${p.priority}</priority>\n  </url>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls}\n</urlset>\n`;
}

export function robotsTxt(site = SITE_URL) {
  return `# Offizin: every public page may be read and indexed.\nUser-agent: *\nAllow: /\n\nSitemap: ${site}/sitemap.xml\n`;
}

export function webManifest() {
  const icon = (size) => ({ src: `assets/icons/icon-${size}.png`, sizes: `${size}x${size}`, type: "image/png" });
  const manifest = {
    name: `${SITE_NAME}: desktop publishing for Word and Google Docs`,
    short_name: SITE_NAME,
    description: SEO_PAGES[0].description,
    start_url: "./",
    scope: "./",
    display: "browser",
    background_color: "#F4F8FC",
    theme_color: "#F4F8FC",
    icons: [icon(192), icon(512)],
  };
  return JSON.stringify(manifest, null, 2) + "\n";
}

async function main() {
  console.log(`Building Offizin ${VERSION} for ${SITE_URL}`);
  const manifest = manifestXml();
  write(path.join(ROOT, "docs/word/manifest.xml"), manifest);

  const { googleFontsHref } = await import(pathToFileURL(path.join(ROOT, "docs/app/core/schemes.js")).href);
  const swapSite = (text) => (SITE_URL === DEFAULT_SITE ? text : text.split(DEFAULT_SITE).join(SITE_URL));
  const js = swapSite(bundle(path.join(ROOT, "docs/app/hosts/google.js")));
  const css = fs.readFileSync(path.join(ROOT, "docs/app/ui/panel.css"), "utf8");
  const template = fs.readFileSync(path.join(ROOT, "google-docs/Sidebar.template.html"), "utf8");
  // Function replacements, so "$" characters in the code are inserted literally.
  const sidebar = template
    .replace("{{GOOGLE_FONTS}}", () => googleFontsHref().replace(/&/g, "&amp;"))
    .replace("{{PANEL_CSS}}", () => css)
    .replace("{{BUNDLE}}", () => js.replace(/<\/script/gi, "<\\/script"));
  const code = swapSite(fs.readFileSync(path.join(ROOT, "google-docs/Code.gs"), "utf8"));
  const appsscript = fs.readFileSync(path.join(ROOT, "google-docs/appsscript.json"), "utf8");
  write(path.join(ROOT, "docs/google-docs/Sidebar.html"), sidebar);
  write(path.join(ROOT, "docs/google-docs/Code.gs"), code);
  write(path.join(ROOT, "docs/google-docs/appsscript.json"), appsscript);
  write(path.join(ROOT, "docs/download/offizin-word.zip"), wordPackage(manifest));
  write(path.join(ROOT, "docs/download/offizin-google-docs.zip"), googlePackage({ code, sidebar, manifest: appsscript }));
  write(
    path.join(ROOT, "docs/version.json"),
    JSON.stringify({ name: "Offizin", version: VERSION, site: SITE_URL, repository: REPO_URL, googleStarterCopyUrl: GOOGLE_STARTER_COPY_URL }, null, 2) + "\n"
  );
  for (const page of SEO_PAGES) {
    const file = path.join(ROOT, "docs", page.file);
    write(file, applySeo(fs.readFileSync(file, "utf8"), page));
  }
  write(path.join(ROOT, "docs/sitemap.xml"), sitemapXml());
  write(path.join(ROOT, "docs/robots.txt"), robotsTxt());
  write(path.join(ROOT, "docs/site.webmanifest"), webManifest());
  console.log("Done.");
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((e) => {
    console.error(e);
    process.exit(1);
  });
}
