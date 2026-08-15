// Builds a single self-contained HTML file from index.html + app.js + vendor/.
//
// Why: the normal app loads five files over HTTP. A single file can be opened
// straight from a phone's Downloads folder, emailed, or published as a hosted
// preview page for manual testing, with no web server involved.
//
// Usage: node tools/build-standalone.mjs [outfile] [--fragment]
//
// --fragment emits <title> + <style> + body content only, with no
// <html>/<head>/<body> wrapper, for hosts that supply their own document
// skeleton (the Artifact viewer does this).

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const positional = args.filter((a) => !a.startsWith("--"));
const out = positional[0] || path.join(root, "dist", "docgen-standalone.html");
const read = (p) => fs.readFileSync(path.join(root, p), "utf8");

// An inline <script> ends at the first literal "</script>", even inside a
// string. Escaping the slash is inert to JS and keeps the parser out.
//
// The docx bundle also carries four literal U+FFFD characters inside string
// literals (its UTF-8 decoder's error marker). Some hosts reject a payload
// containing raw replacement characters, and � is the identical value to
// the JS engine. A raw U+FFFD outside a string would already be a syntax
// error, so this can only ever rewrite string contents.
const inlineSafe = (source) => source.replaceAll("</script", "<\\/script").replaceAll("�", "\\uFFFD");

// ---- vendor ---------------------------------------------------------------
// The three UMD bundles are plain scripts: inlining them upfront means the
// app's lazy loaders find window.JSZip / window.docx / window.html2pdf already
// present and never inject a <script src>.
const umd = ["vendor/jszip-3.10.1.min.js", "vendor/docx-preview-0.4.0.min.js", "vendor/html2pdf-0.10.2.bundle.min.js"]
  .map((file) => '<script>/* ' + file + ' */\n' + inlineSafe(read(file)) + "\n</script>")
  .join("\n");

// docx is an ES module, and an inline module cannot be imported by another
// inline module. Rewrite its final `export { ... }` into an object literal
// published on globalThis, which the app module then picks up.
let docx = read("vendor/docx-8.2.2.js");
const exportIndex = docx.lastIndexOf("\nexport {");
if (exportIndex === -1) throw new Error("docx bundle: no export block found");
const exportBlock = docx
  .slice(exportIndex)
  .replace("\nexport {", "\nglobalThis.__docxLib = {")
  // `File as Document` is not valid in an object literal; `Document: File` is.
  .replace(/(\w+\$?\d*) as (\w+)/g, "$2: $1");
docx = docx.slice(0, exportIndex) + exportBlock;

// ---- app ------------------------------------------------------------------
let app = read("app.js");
const importLine = 'import * as docxLib from "./vendor/docx-8.2.2.js";';
if (!app.includes(importLine)) throw new Error("app.js: docx import line not found");
app = app.replace(importLine, "const docxLib = globalThis.__docxLib;");

// ---- page -----------------------------------------------------------------
const html = read("index.html");
const scriptTag = /<script type="module" src="\.\/app\.js[^"]*"><\/script>/;
if (!scriptTag.test(html)) throw new Error("index.html: app.js script tag not found");

const replacement = [
  umd,
  '<script type="module">/* vendor/docx-8.2.2.js */\n' + inlineSafe(docx) + "\n</script>",
  '<script type="module">/* app.js */\n' + inlineSafe(app) + "\n</script>",
].join("\n");

// A function replacer, not a string: minified vendor code contains `$&` and
// `$'`, which String.replace would expand as match references.
let bundle = html.replace(scriptTag, () => replacement);

// In a hosted preview frame the page cannot start a download itself, so an
// <a download> click silently does nothing. Route through the host's save API
// when it is present, and tell the tester plainly when a format is refused
// rather than leaving them staring at a button that appears to do nothing.
// Only the preview build is patched; the deployed app downloads normally.
const HOSTED_DOWNLOAD_SHIM = `
async function downloadBlob(blob, filename) {
  const host = typeof claude !== "undefined" && claude.use ? await claude.use("downloads").catch(() => null) : null;
  if (!host) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.style.display = "none"; a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
    return;
  }
  try {
    await host.save({ filename, data: await blob.arrayBuffer() });
    showToast("Saved", "ok");
  } catch (error) {
    const code = error && error.code;
    if (code === "declined") return;
    if (code === "rejected_extension" || code === "extension_not_enabled") {
      showToast("This preview cannot save " + filename.split(".").pop().toUpperCase() + " files. Use the deployed app.", "er");
      return;
    }
    showToast("Save failed: " + ((error && error.message) || code || "unknown"), "er");
  }
}
`;

if (args.includes("--fragment")) {
  const original = /function downloadBlob\(blob, filename\) \{[\s\S]*?\n\}/;
  if (!original.test(bundle)) throw new Error("downloadBlob not found — cannot patch the preview build");
  bundle = bundle.replace(original, () => HOSTED_DOWNLOAD_SHIM.trim());
}

if (args.includes("--fragment")) {
  // Keep <title> and <style> from the head, drop the document scaffolding.
  const title = (bundle.match(/<title>[\s\S]*?<\/title>/) || [""])[0];
  const style = (bundle.match(/<style>[\s\S]*?<\/style>/) || [""])[0];
  const body = bundle.slice(bundle.indexOf("<body>") + "<body>".length, bundle.lastIndexOf("</body>"));
  bundle = [title, style, body].join("\n");
}

fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, bundle);
console.log("wrote " + out + " (" + (Buffer.byteLength(bundle) / 1048576).toFixed(2) + " MB)");
