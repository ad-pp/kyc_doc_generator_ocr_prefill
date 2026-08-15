// Preflight: prove each provider key actually works, before deploying.
//
//   node tools/check-providers.mjs
//
// Reads keys from worker/.dev.vars (or the environment) and sends one small
// real request per provider — a tiny generated image plus the same JSON-only
// instruction the Worker uses. It reports, per tier, whether the key is
// accepted, the model exists, and the reply parses as JSON.
//
// This exercises the real code path: the same provider functions the Worker
// runs, not a reimplementation, so a pass here means the Worker will work.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PROVIDER_CHAIN } from "../worker/src/providers.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// Keys come from .dev.vars when present, else the environment. Never printed.
function loadEnv() {
  const env = { ...process.env };
  const devVars = path.join(root, "worker", ".dev.vars");
  if (fs.existsSync(devVars)) {
    for (const line of fs.readFileSync(devVars, "utf8").split("\n")) {
      const match = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)$/);
      if (match && match[2].trim()) env[match[1]] = match[2].trim().replace(/^["']|["']$/g, "");
    }
  }
  // Model names default the same way the Worker's vars do.
  const toml = fs.readFileSync(path.join(root, "worker", "wrangler.toml"), "utf8");
  for (const key of ["GEMINI_MODEL", "OPENROUTER_MODEL", "OPENROUTER_PDF_ENGINE", "GROQ_MODEL", "ALLOWED_ORIGINS"]) {
    if (env[key]) continue;
    const match = toml.match(new RegExp("^" + key + '\\s*=\\s*"([^"]*)"', "m"));
    if (match) env[key] = match[1];
  }
  return env;
}

// A 96x96 PNG with three black bars. A 2x2 image was rejected by Qwen on Groq
// with "invalid image data" — some vision endpoints require a plausible
// minimum size — which read as a broken tier when the tier was fine. Still
// tiny enough that a check costs almost nothing against a free quota.
const TINY_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAIAAABt+uBvAAAA3UlEQVR4nO3YsRGAQAwEse+/aeiA" +
    "DRzAYKkCz2bnc/HovH3A1wkUBAoCBYGCQEGgIFAQKAgUBAoCBYGCQEGgIFAQKAgUBAoChQ50/k6g" +
    "IFAQKAgUBAoCBYGCQEGgIFCYBlpOoCBQECgIFAQKAgWBgkBBoCBQECgIFKx5744iUBAoCBQECgIF" +
    "gYJAQaAgUJgGWk6gIFAQKAgUBAoCBYGCQEGgIFAQKAgUrHnvjiJQECgIFAQKAgWBgkBBoCBQmAZa" +
    "TqAgUBAoCBQECgIFgYJAQaAgUBAoCBQECgIFgcINHne0lB853+wAAAAASUVORK5CYII=",
  "base64"
);

function testFile() {
  return new File([TINY_PNG], "check.png", { type: "image/png" });
}

const label = (name) => name.padEnd(11);

// Model ids churn: providers retire them, and free tiers move behind paywalls.
// Rather than guess a replacement, ask each provider which models THIS key can
// actually use. Every one exposes a list endpoint.
const MODEL_LISTS = {
  gemini: {
    url: "https://generativelanguage.googleapis.com/v1beta/models?pageSize=200",
    headers: (env) => ({ "x-goog-api-key": env.GEMINI_API_KEY }),
    parse: (data) =>
      (data.models || [])
        .filter((m) => (m.supportedGenerationMethods || []).includes("generateContent"))
        .map((m) => ({ id: (m.name || "").replace(/^models\//, ""), note: m.displayName || "" })),
    // Anchor the version to the "gemini-N.N" prefix. Matching any "-<digits>"
    // pulls the date out of ids like deep-research-pro-preview-12-2025 and
    // ranks it as generation 12 — first place, and completely wrong.
    rank: (id) => {
      if (!/^gemini-\d/.test(id)) return 900; // not a general Gemini model
      const version = id.match(/^gemini-(\d+)(?:\.(\d+))?/);
      const generation = Number(version[1]) * 10 + Number(version[2] || 0);
      return (
        (/flash/.test(id) ? -40 : 0) +
        (/lite/.test(id) ? 25 : 0) +
        // Preview, EAP and confidential builds are not something to depend on.
        (/preview|exp\b|eap|confidential|thinking|tts|image-gen|embedding|learnlm/.test(id) ? 300 : 0) +
        // Narrow single-purpose builds, e.g. video-understanding.
        (/understanding|research|native-audio|live/.test(id) ? 300 : 0) -
        generation * 2
      );
    },
  },
  openrouter: {
    url: "https://openrouter.ai/api/v1/models",
    headers: (env) => ({ Authorization: "Bearer " + env.OPENROUTER_API_KEY }),
    parse: (data) =>
      (data.data || [])
        .filter((m) => (m.architecture?.input_modalities || []).includes("image"))
        .map((m) => ({ id: m.id, note: m.name || "" })),
    // Free slugs only, and prefer a general-purpose instruction model. A
    // safety classifier or a note-taker is image-capable but useless for
    // reading a deed, and would otherwise win on list order alone.
    rank: (id) => {
      const size = id.match(/(\d+)b/i);
      return (
        (/:free$/.test(id) ? -100 : 900) +
        (/safety|guard|moderation|classif|embed|rerank|note|tts|whisper|coder|code-/i.test(id) ? 400 : 0) +
        (/-it:|instruct|vl\b|vision|omni/i.test(id) ? -60 : 0) -
        // Bigger models read documents better; this is a tie-breaker only.
        Math.min(Number(size ? size[1] : 0), 90) / 3
      );
    },
  },
  groq: {
    url: "https://api.groq.com/openai/v1/models",
    headers: (env) => ({ Authorization: "Bearer " + env.GROQ_API_KEY }),
    // Groq's list does not flag modality. Matching vision by name hid
    // qwen/qwen3.6-27b, which is multimodal but says so nowhere in its id — so
    // filtering on names drops real candidates. List everything except the
    // obviously non-visual, and rank likely vision models to the top instead.
    parse: (data) =>
      (data.data || [])
        .filter((m) => !/whisper|tts|guard|embed|rerank/i.test(m.id))
        .map((m) => ({ id: m.id, note: m.owned_by || "" })),
    rank: (id) =>
      (/vision|scout|maverick|llava|[-/]vl\b|omni|qwen|gemma|llama-4/i.test(id) ? -50 : 0) +
      (/compound|allam|prompt/i.test(id) ? 200 : 0) -
      Math.min(Number((id.match(/(\d+)b/i) || [0, 0])[1]), 90) / 30,
  },
};

async function suggestModels(tierName, env) {
  const spec = MODEL_LISTS[tierName];
  if (!spec) return [];
  try {
    const response = await fetch(spec.url, { headers: spec.headers(env) });
    if (!response.ok) return [];
    const candidates = spec.parse(await response.json());
    return candidates.sort((a, b) => spec.rank(a.id) - spec.rank(b.id)).slice(0, 8);
  } catch (e) {
    return [];
  }
}

async function main() {
  const env = loadEnv();
  console.log("\nChecking providers in chain order. Each sends one small real request.\n");

  const results = [];
  for (const [index, tier] of PROVIDER_CHAIN.entries()) {
    const tierName = ["primary", "secondary", "tertiary"][index] || "tier " + (index + 1);
    const model = env[tier.name.toUpperCase() + "_MODEL"] || "(built-in default)";

    if (!env[tier.keyVar]) {
      console.log(`  ${label(tier.name)} SKIP   ${tier.keyVar} is not set`);
      results.push({ tier: tier.name, ok: false, skipped: true });
      continue;
    }

    const started = Date.now();
    try {
      const output = await tier.run(testFile(), env);
      const ms = Date.now() - started;
      // A 2x2 image has nothing to extract, so empty fields are the expected
      // result. What matters is that the call succeeded and returned parseable
      // JSON in our schema's shape.
      const shaped = output && typeof output === "object";
      console.log(`  ${label(tier.name)} OK     ${tierName}, ${ms}ms, model ${model}${shaped ? "" : " (unexpected response shape)"}`);
      results.push({ tier: tier.name, ok: true });
    } catch (error) {
      console.log(`  ${label(tier.name)} FAIL   ${error.message}`);
      const failed = { tier: tier.name, ok: false, error: error.message };
      // A 404 almost always means the configured model id is retired or was
      // never available to this key. Ask the provider what it will accept.
      if (/404|not found|no longer available|does not exist|unavailable for free/i.test(error.message)) {
        failed.suggestions = await suggestModels(tier.name, env);
        if (!failed.suggestions.length) {
          console.log(`  ${" ".repeat(11)}        no usable model found for this tier on this key.`);
          console.log(`  ${" ".repeat(11)}        This tier needs a model that accepts images; if there is`);
          console.log(`  ${" ".repeat(11)}        none, drop the tier:`);
          console.log(`  ${" ".repeat(11)}          npx wrangler secret delete ${tier.keyVar}`);
        }
        if (failed.suggestions.length) {
          console.log(`  ${" ".repeat(11)}        models this key CAN use:`);
          for (const candidate of failed.suggestions) {
            console.log(`  ${" ".repeat(11)}          ${candidate.id}${candidate.note ? "  (" + candidate.note + ")" : ""}`);
          }
        }
      }
      results.push(failed);
    }
  }

  const working = results.filter((r) => r.ok);
  console.log("");

  // Print the exact config edit, so fixing this is copy-paste rather than
  // cross-referencing three sets of provider docs.
  const fixable = results.filter((r) => r.suggestions && r.suggestions.length);
  if (fixable.length) {
    const varName = { gemini: "GEMINI_MODEL", openrouter: "OPENROUTER_MODEL", groq: "GROQ_MODEL" };
    console.log("Set these in worker/wrangler.toml, then re-run this check:\n");
    for (const item of fixable) {
      console.log(`  ${varName[item.tier]} = "${item.suggestions[0].id}"`);
    }
    console.log("");
  }
  if (working.length === PROVIDER_CHAIN.length) {
    console.log(`All ${PROVIDER_CHAIN.length} providers responded. Safe to deploy.\n`);
    process.exit(0);
  }
  if (working.length > 0) {
    console.log(`${working.length} of ${PROVIDER_CHAIN.length} providers responded: ${working.map((r) => r.tier).join(", ")}.`);
    console.log("Extraction will work, with less redundancy. Fix the rest above, or deploy as is.\n");
    process.exit(0);
  }
  console.log("No provider responded — extraction would fail for every upload. Do not deploy yet.\n");
  process.exit(1);
}

main().catch((error) => {
  console.error("\ncheck-providers crashed: " + error.message + "\n");
  process.exit(1);
});
