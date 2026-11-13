/**
 * The download packages in docs/download: each one opens as a normal zip file, holds exactly the
 * files its guide names, matches the files published next to it, and has a guide in plain words.
 *
 * Run with: node scripts/check.mjs --only=packages
 */
import fs from "node:fs";
import path from "node:path";

/** Names and contents of a zip written with stored (uncompressed) entries. */
function readZip(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.length - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06054b50) end--;
  if (end < 0) throw new Error("no end of central directory");
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const out = {};
  for (let i = 0; i < count; i++) {
    if (view.getUint32(at, true) !== 0x02014b50) throw new Error("bad central directory entry");
    const method = view.getUint16(at + 10, true);
    const size = view.getUint32(at + 24, true);
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    const local = view.getUint32(at + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(at + 46, at + 46 + nameLen));
    if (method !== 0) throw new Error(`${name} is compressed`);
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    out[name] = bytes.subarray(start, start + size);
    at += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

export async function run({ ROOT, fail, ok, load }) {
  const build = await load("scripts/build.mjs");
  const { VERSION } = await load("scripts/site.config.mjs");
  const read = (p) => fs.readFileSync(path.join(ROOT, p));
  const text = (bytes) => new TextDecoder().decode(bytes);
  const expect = (cond, msg) => (cond ? ok(msg) : fail(msg));
  const dash = String.fromCharCode(0x2014);
  // The example folder and its network path, exactly as the Word install page shows them.
  const WINDOWS_PATHS = ["C:\\OffizinAddin", "\\\\YOUR-PC\\OffizinAddin"];
  const packages = [
    { file: "docs/download/offizin-word.zip", files: { "offizin-manifest.xml": "docs/word/manifest.xml" }, guideSays: [...WINDOWS_PATHS, "Finish deployment", "72 hours"] },
    {
      file: "docs/download/offizin-google-docs.zip",
      files: { "Code.gs": "docs/google-docs/Code.gs", "Sidebar.html": "docs/google-docs/Sidebar.html", "appsscript.json": "docs/google-docs/appsscript.json" },
      guideSays: ["check both", "brand kit once", "Each copy asks for permission"],
    },
  ];
  for (const pkg of packages) {
    if (!fs.existsSync(path.join(ROOT, pkg.file))) {
      fail(`${pkg.file} is missing; run npm run build`);
      continue;
    }
    let entries;
    try {
      entries = readZip(read(pkg.file));
    } catch (e) {
      fail(`${pkg.file}: ${e.message}`);
      continue;
    }
    const names = Object.keys(entries).sort();
    const expected = [...Object.keys(pkg.files), "README.txt"].sort();
    if (names.join() !== expected.join()) fail(`${pkg.file} holds ${names.join(", ")}, expected ${expected.join(", ")}`);
    for (const [inside, published] of Object.entries(pkg.files)) {
      if (entries[inside] && Buffer.compare(Buffer.from(entries[inside]), read(published)) !== 0) fail(`${pkg.file}: ${inside} differs from ${published}; run npm run build`);
    }
    const guide = entries["README.txt"] ? text(entries["README.txt"]) : "";
    if (!guide.includes(`Offizin ${VERSION}`)) fail(`${pkg.file}: the guide does not name version ${VERSION}`);
    for (const inside of Object.keys(pkg.files)) if (!guide.includes(inside)) fail(`${pkg.file}: the guide does not mention ${inside}`);
    if (guide.includes(dash)) fail(`${pkg.file}: em dash in the guide`);
    const long = guide.split("\n").filter((l) => l.length > 80 && !/https?:|~\//.test(l));
    if (long.length) fail(`${pkg.file}: guide lines longer than 80 characters: ${long[0]}`);
    for (const must of pkg.guideSays || []) if (!guide.includes(must)) fail(`${pkg.file}: the guide does not say ${JSON.stringify(must)}`);
  }
  // The Windows paths in the Word guide: template literals drop lone backslashes, so check the source too.
  const wordGuide = build.wordReadme();
  for (const p of WINDOWS_PATHS) if (!wordGuide.includes(p)) fail(`wordReadme() does not contain the Windows path ${JSON.stringify(p)}`);
  // The Whole organization steps: Microsoft's instructions name the Choose File button but give no
  // label for the upload option, so the install page and the guide describe it without quoting one.
  const page = text(read("docs/word/index.html"));
  const orgAt = page.indexOf('id="p-org"');
  const orgSteps = orgAt < 0 ? "" : page.slice(orgAt, page.indexOf("</section>", orgAt));
  expect(
    orgSteps.includes("<strong>Choose File</strong>") && orgSteps.includes("<code>offizin-manifest.xml</code>"),
    "install page: the organization steps say to select Choose File and pick offizin-manifest.xml"
  );
  expect(!page.includes("Upload manifest file") && !wordGuide.includes("Upload manifest file"), "install page and Word guide: no label is quoted for the upload option");
  const word = Buffer.from(build.wordPackage(text(read("docs/word/manifest.xml"))));
  if (Buffer.compare(word, read("docs/download/offizin-word.zip")) !== 0) fail("the Word package is out of date; run npm run build");
  const google = Buffer.from(
    build.googlePackage({ code: text(read("docs/google-docs/Code.gs")), sidebar: text(read("docs/google-docs/Sidebar.html")), manifest: text(read("docs/google-docs/appsscript.json")) })
  );
  if (Buffer.compare(google, read("docs/download/offizin-google-docs.zip")) !== 0) fail("the Google Docs package is out of date; run npm run build");
  // The panel's in-document tests run in jsdom, which npm install provides. There is no lockfile,
  // so package.json pins every development package to an exact version.
  const dev = JSON.parse(text(read("package.json"))).devDependencies || {};
  const exact = (version) => /^\d+\.\d+\.\d+$/.test(version || "");
  expect(exact(dev.jsdom) && Object.values(dev).every(exact), "package.json: jsdom and the other development packages are pinned to exact versions");
  expect(
    text(read("MAINTAINING.md")).includes("OFFIZIN_JSDOM") && text(read("scripts/tests/panel.mjs")).includes("process.env.OFFIZIN_JSDOM"),
    "MAINTAINING.md: names OFFIZIN_JSDOM, the variable the panel tests read to find jsdom"
  );
  ok("download packages open as zip files, match the published files and a fresh build, and include plain guides");
}
