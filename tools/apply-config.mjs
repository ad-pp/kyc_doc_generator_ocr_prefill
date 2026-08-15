// Edits the repo's config files in place so setup.sh doesn't have to do it
// with fragile sed expressions.
//
//   node tools/apply-config.mjs --kv-id <id>
//   node tools/apply-config.mjs --api-url <url>

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const arg = (name) => {
  const index = args.indexOf("--" + name);
  return index === -1 ? null : args[index + 1];
};

function edit(file, transform) {
  const full = path.join(root, file);
  const before = fs.readFileSync(full, "utf8");
  const after = transform(before);
  if (after === before) return false;
  fs.writeFileSync(full, after);
  return true;
}

const kvId = arg("kv-id");
if (kvId) {
  const changed = edit("worker/wrangler.toml", (text) => {
    if (/^\[\[kv_namespaces\]\]/m.test(text)) {
      // Block already live — just refresh the id.
      return text.replace(/(\[\[kv_namespaces\]\][\s\S]*?id\s*=\s*)"[^"]*"/, '$1"' + kvId + '"');
    }
    return text.replace(
      /# \[\[kv_namespaces\]\]\n# binding = "USAGE"\n# id = "[^"]*"/,
      '[[kv_namespaces]]\nbinding = "USAGE"\nid = "' + kvId + '"'
    );
  });
  console.log(changed ? "wrangler.toml: KV namespace bound (" + kvId + ")" : "wrangler.toml: unchanged — check the kv_namespaces block by hand");
}

const apiUrl = arg("api-url");
if (apiUrl) {
  const clean = apiUrl.replace(/\/+$/, "");
  const wired = edit("app.js", (text) =>
    text.replace(/const API_BASE_URL = "[^"]*";/, 'const API_BASE_URL = "' + clean + '";')
  );
  console.log(wired ? "app.js: API_BASE_URL set to " + clean : "app.js: API_BASE_URL unchanged — set it by hand");

  // Phones cache aggressively; a new query string forces the new build.
  const bumped = edit("index.html", (text) =>
    text.replace(/(<script type="module" src="\.\/app\.js\?v=)(\d+)(")/, (m, a, version, c) => a + (Number(version) + 1) + c)
  );
  console.log(bumped ? "index.html: cache version bumped" : "index.html: cache version unchanged");
}

if (!kvId && !apiUrl) {
  console.error("nothing to do — pass --kv-id and/or --api-url");
  process.exit(1);
}
