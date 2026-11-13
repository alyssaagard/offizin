/**
 * Search engine and link preview checks for the website.
 *
 *  1. Every public page has one title, one description, and one address of its own, of a length
 *     search results show in full, and its tags match a fresh build.
 *  2. The pictures and icons the tags point to exist, and the structured data is valid JSON.
 *  3. The sitemap lists exactly the public pages, robots.txt points to it, and every other
 *     page asks to be left out of search results.
 *
 * Run with: node scripts/check.mjs --only=seo
 */
import fs from "node:fs";
import path from "node:path";

export async function run({ ROOT, fail, ok }) {
  const { SEO_PAGES, SOCIAL_IMAGE } = await import(path.join(ROOT, "scripts/seo.config.mjs"));
  const build = await import(path.join(ROOT, "scripts/build.mjs"));
  const { SITE_URL } = await import(path.join(ROOT, "scripts/site.config.mjs"));
  const docs = path.join(ROOT, "docs");
  const read = (f) => fs.readFileSync(path.join(docs, f), "utf8");
  const count = (text, re) => (text.match(re) || []).length;

  // 1 and 2. Tags on every public page ---------------------------------------------------
  {
    const problems = [];
    const titles = new Set();
    const descriptions = new Set();
    for (const page of SEO_PAGES) {
      const html = read(page.file);
      const head = html.slice(0, html.indexOf("</head>"));
      const where = page.file;
      if (count(head, /<title>/g) !== 1) problems.push(`${where}: needs exactly one title`);
      if (count(head, /<meta name="description"/g) !== 1) problems.push(`${where}: needs exactly one description`);
      if (count(head, /<link rel="canonical"/g) !== 1 || !head.includes(`<link rel="canonical" href="${SITE_URL}/${page.path}">`)) problems.push(`${where}: the canonical address should be ${SITE_URL}/${page.path}`);
      if (page.title.length > 60) problems.push(`${where}: the title has ${page.title.length} characters; search results show about 60`);
      if (page.description.length < 70 || page.description.length > 160) problems.push(`${where}: the description has ${page.description.length} characters; keep it between 70 and 160`);
      if (titles.has(page.title) || descriptions.has(page.description)) problems.push(`${where}: its title or description repeats another page's`);
      titles.add(page.title);
      descriptions.add(page.description);
      if (!head.includes(build.seoBlock(page))) problems.push(`${where}: its search tags are out of date; run npm run build`);
      if (/name="robots" content="[^"]*noindex/.test(head)) problems.push(`${where}: a public page must not ask to be left out of search results`);
      for (const m of head.matchAll(/<link rel="(?:icon|apple-touch-icon|manifest)" href="([^"]+)"/g)) {
        if (!fs.existsSync(path.join(docs, path.dirname(page.file), m[1]))) problems.push(`${where}: ${m[1]} does not exist`);
      }
      const ld = head.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/);
      if (page.schema) {
        try {
          const data = JSON.parse(ld[1]);
          if (data["@context"] !== "https://schema.org") problems.push(`${where}: structured data needs the schema.org context`);
        } catch (e) {
          problems.push(`${where}: structured data is not valid JSON`);
        }
      }
    }
    if (!fs.existsSync(path.join(docs, SOCIAL_IMAGE.path))) problems.push(`the link preview picture ${SOCIAL_IMAGE.path} does not exist`);
    const manifest = JSON.parse(read("site.webmanifest"));
    for (const icon of manifest.icons) if (!fs.existsSync(path.join(docs, icon.src))) problems.push(`site.webmanifest: ${icon.src} does not exist`);
    problems.forEach((p) => fail(`seo: ${p}`));
    if (!problems.length) ok(`seo: ${SEO_PAGES.length} public pages each have their own title, description, address, icons, and link preview`);
  }

  // 3. Sitemap, robots.txt, and pages kept out of search results -----------------------------
  {
    const problems = [];
    const sitemap = read("sitemap.xml");
    const locs = [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map((m) => m[1]);
    const want = SEO_PAGES.map((p) => `${SITE_URL}/${p.path}`);
    if (locs.join("\n") !== want.join("\n")) problems.push("sitemap.xml should list exactly the public pages, in order");
    if (sitemap !== build.sitemapXml() || read("robots.txt") !== build.robotsTxt() || read("site.webmanifest") !== build.webManifest()) problems.push("sitemap.xml, robots.txt, or site.webmanifest is out of date; run npm run build");
    if (!read("robots.txt").includes(`Sitemap: ${SITE_URL}/sitemap.xml`)) problems.push("robots.txt should point to the sitemap");
    // Every other HTML page on the site asks to be left out.
    const pages = [];
    const walk = (dir) => {
      for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, e.name);
        if (e.isDirectory()) walk(full);
        else if (e.name.endsWith(".html")) pages.push(path.relative(docs, full).split(path.sep).join("/"));
      }
    };
    walk(docs);
    const publicFiles = new Set(SEO_PAGES.map((p) => p.file));
    for (const f of pages) {
      if (publicFiles.has(f)) continue;
      if (!/<meta name="robots" content="noindex/.test(read(f))) problems.push(`${f} is not in scripts/seo.config.mjs, so it should carry <meta name="robots" content="noindex">`);
    }
    problems.forEach((p) => fail(`seo: ${p}`));
    if (!problems.length) ok(`seo: the sitemap lists the public pages, and the ${pages.length - publicFiles.size} other pages stay out of search results`);
  }
}
