/**
 * Builds the files that are generated from the source in docs/app and google-docs:
 *
 *   docs/word/manifest.xml          the Word add-in manifest, from scripts/site.config.mjs
 *   docs/google-docs/Sidebar.html   the whole panel bundled into one file for Apps Script
 *   docs/google-docs/Code.gs        copied from google-docs/Code.gs
 *   docs/google-docs/appsscript.json
 *   docs/version.json
 *
 * Usage: npm run build   (Node 18 or later; no packages needed)
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { SITE_URL, ADDIN_ID, VERSION, PROVIDER, REPO_URL, GOOGLE_STARTER_COPY_URL } from "./site.config.mjs";

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
// Build
// ---------------------------------------------------------------------------

async function main() {
  console.log(`Building Offizin ${VERSION} for ${SITE_URL}`);
  write(path.join(ROOT, "docs/word/manifest.xml"), manifestXml());

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
  write(path.join(ROOT, "docs/google-docs/Sidebar.html"), sidebar);
  write(path.join(ROOT, "docs/google-docs/Code.gs"), swapSite(fs.readFileSync(path.join(ROOT, "google-docs/Code.gs"), "utf8")));
  write(path.join(ROOT, "docs/google-docs/appsscript.json"), fs.readFileSync(path.join(ROOT, "google-docs/appsscript.json"), "utf8"));
  write(
    path.join(ROOT, "docs/version.json"),
    JSON.stringify({ name: "Offizin", version: VERSION, site: SITE_URL, repository: REPO_URL, googleStarterCopyUrl: GOOGLE_STARTER_COPY_URL }, null, 2) + "\n"
  );
  console.log("Done.");
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((e) => {
    console.error(e);
    process.exit(1);
  });
}
