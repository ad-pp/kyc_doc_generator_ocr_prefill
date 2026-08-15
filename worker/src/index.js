// Merchant Onboarding Document Generator — API Worker
// ---------------------------------------------------
// The static app has no login and ships its own source, so it can hold no
// secrets. This Worker is the only place OCR/LLM keys exist. Keys are set once
// with `wrangler secret put` and are never seen by an agent or by the browser.
//
// Endpoints:
//   POST /api/extract  multipart: file, agentMobile  -> extraction JSON
//   POST /api/log      JSON usage event              -> { ok: true }
//   GET  /api/health                                 -> provider readiness
//
// Provider order is decided here, not by the client: the primary chain is
// tried first and the secondary chain runs only if it fails.

import { extractWithGeminiVision, runSecondaryChain } from "./providers.js";

const MAX_UPLOAD_BYTES = 8 * 1024 * 1024;
const ALLOWED_TYPES = ["application/pdf", "image/jpeg", "image/png", "image/webp"];

function corsHeaders(request, env) {
  const allowed = (env.ALLOWED_ORIGINS || "").split(",").map((value) => value.trim()).filter(Boolean);
  const origin = request.headers.get("Origin") || "";
  // With no allowlist configured, refuse to echo an arbitrary origin.
  const allowOrigin = allowed.includes(origin) ? origin : allowed[0] || "";
  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400",
    Vary: "Origin",
  };
}

function json(body, status, request, env) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(request, env) },
  });
}

function todayKey() {
  return new Date().toISOString().slice(0, 10);
}

// Quota control without any login: a per-day cap per agent mobile, plus a
// per-day cap per client IP so a missing/spoofed mobile cannot bypass it.
async function checkAndCountQuota(env, agentMobile, ip) {
  if (!env.USAGE) return { ok: true, counted: false };
  const perAgent = Number(env.DAILY_LIMIT_PER_AGENT || 40);
  const perIp = Number(env.DAILY_LIMIT_PER_IP || 120);
  const day = todayKey();
  const agentKey = "q:agent:" + day + ":" + (agentMobile || "unknown");
  const ipKey = "q:ip:" + day + ":" + ip;

  const [agentRaw, ipRaw] = await Promise.all([env.USAGE.get(agentKey), env.USAGE.get(ipKey)]);
  const agentCount = Number(agentRaw || 0);
  const ipCount = Number(ipRaw || 0);
  if (agentCount >= perAgent) return { ok: false, reason: "Daily extraction limit reached for this agent number." };
  if (ipCount >= perIp) return { ok: false, reason: "Daily extraction limit reached for this connection." };

  // Two-day TTL so a counter written just before midnight still expires.
  await Promise.all([
    env.USAGE.put(agentKey, String(agentCount + 1), { expirationTtl: 172800 }),
    env.USAGE.put(ipKey, String(ipCount + 1), { expirationTtl: 172800 }),
  ]);
  return { ok: true, counted: true, agentCount: agentCount + 1 };
}

async function recordEvent(env, event) {
  if (!env.USAGE) return;
  const key = "log:" + new Date().toISOString() + ":" + crypto.randomUUID().slice(0, 8);
  // 180-day retention; enough for a POC review without unbounded growth.
  await env.USAGE.put(key, JSON.stringify(event), { expirationTtl: 15552000 });
}

async function handleExtract(request, env, ctx) {
  const contentType = request.headers.get("Content-Type") || "";
  if (!contentType.includes("multipart/form-data")) {
    return json({ error: "Expected a multipart/form-data upload." }, 400, request, env);
  }
  const form = await request.formData();
  const file = form.get("file");
  const agentMobile = String(form.get("agentMobile") || "").replace(/\D/g, "").slice(0, 10);
  if (!file || typeof file === "string") return json({ error: "No file was uploaded." }, 400, request, env);
  if (file.size > MAX_UPLOAD_BYTES) {
    return json({ error: "File is larger than 8 MB. Retake the photo or split the PDF." }, 413, request, env);
  }
  if (file.type && !ALLOWED_TYPES.includes(file.type)) {
    return json({ error: "Unsupported file type. Upload a PDF, JPG, PNG or WebP." }, 415, request, env);
  }

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  const quota = await checkAndCountQuota(env, agentMobile, ip);
  if (!quota.ok) return json({ error: quota.reason }, 429, request, env);

  const warnings = [];
  const started = Date.now();
  let result = null;
  let provider = "";
  // Kept so a total failure can say WHY. Provider messages are status codes
  // and human-readable text; no key is ever echoed back in them.
  const failures = [];

  if (env.GEMINI_API_KEY) {
    try {
      result = await extractWithGeminiVision(file, env);
      provider = "gemini-vision";
    } catch (error) {
      warnings.push("Primary extraction was unavailable, so a backup provider was used.");
      failures.push("primary: " + error.message);
      console.log("primary failed: " + error.message);
    }
  } else {
    failures.push("primary: no Gemini key configured");
  }
  if (!result) {
    try {
      result = await runSecondaryChain(file, env);
      provider = "ocrspace-groq";
    } catch (error) {
      failures.push("secondary: " + error.message);
      console.log("secondary failed: " + error.message);
      ctx.waitUntil(
        recordEvent(env, {
          type: "extract",
          ok: false,
          agentMobile,
          failures,
          fileType: file.type || "",
          fileKb: Math.round(file.size / 1024),
          at: new Date().toISOString(),
        })
      );
      return json(
        {
          error: "Document reading is temporarily unavailable. Fill the form manually and try the upload again later.",
          // Surfaced so a failure is diagnosable from the phone that hit it,
          // rather than only from `wrangler tail` on someone's laptop.
          detail: failures.join(" | "),
          fileKb: Math.round(file.size / 1024),
          fileType: file.type || "",
        },
        502,
        request,
        env
      );
    }
  }

  if (result && result.sourceQuality && typeof result.sourceQuality === "object") {
    result.sourceQuality.warnings = (result.sourceQuality.warnings || []).concat(warnings);
  }
  ctx.waitUntil(
    recordEvent(env, {
      type: "extract",
      ok: true,
      provider,
      agentMobile,
      ms: Date.now() - started,
      at: new Date().toISOString(),
    })
  );
  return json(result, 200, request, env);
}

async function handleLog(request, env, ctx) {
  let body = {};
  try { body = await request.json(); } catch (e) { /* tolerate a bad body */ }
  ctx.waitUntil(
    recordEvent(env, {
      type: "generate",
      agentMobile: String(body.agentMobile || "").replace(/\D/g, "").slice(0, 10),
      merchant: String(body.merchant || "").slice(0, 64),
      merchantStatus: String(body.merchantStatus || "").slice(0, 16),
      at: new Date().toISOString(),
    })
  );
  return json({ ok: true }, 200, request, env);
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: corsHeaders(request, env) });

    if (url.pathname === "/api/health") {
      return json(
        {
          ok: true,
          primary: Boolean(env.GEMINI_API_KEY),
          secondary: Boolean(env.OCRSPACE_API_KEY && env.GROQ_API_KEY),
        },
        200,
        request,
        env
      );
    }
    if (request.method !== "POST") return json({ error: "Method not allowed." }, 405, request, env);
    if (url.pathname === "/api/extract") return handleExtract(request, env, ctx);
    if (url.pathname === "/api/log") return handleLog(request, env, ctx);
    return json({ error: "Not found." }, 404, request, env);
  },
};
