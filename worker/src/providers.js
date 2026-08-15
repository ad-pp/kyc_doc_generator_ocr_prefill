import { buildTextPrompt, buildVisionPrompt } from "./prompt.js";

// Models sometimes fence their JSON despite being told not to.
export function parseModelJson(text, label) {
  const cleaned = String(text || "")
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/```$/, "")
    .trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    throw new Error(label + " returned a response that was not valid JSON.");
  }
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  }
  return btoa(binary);
}

async function withTimeout(promise, ms, label) {
  let timer;
  try {
    return await Promise.race([
      promise,
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error(label + " timed out after " + ms + "ms")), ms);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

// ---- PRIMARY CHAIN: Gemini reads the document directly ----------------------
// One call does both OCR and extraction, which is fewer moving parts and
// noticeably better than OCR-then-LLM on poor phone-camera scans.
export async function extractWithGeminiVision(file, env) {
  const model = env.GEMINI_MODEL || "gemini-2.5-flash";
  const bytes = new Uint8Array(await file.arrayBuffer());
  const response = await withTimeout(
    fetch("https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(model) + ":generateContent", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
      body: JSON.stringify({
        generationConfig: { responseMimeType: "application/json", temperature: 0 },
        contents: [
          {
            role: "user",
            parts: [
              { text: buildVisionPrompt() },
              { inline_data: { mime_type: file.type || "application/pdf", data: bytesToBase64(bytes) } },
            ],
          },
        ],
      }),
    }),
    45000,
    "Gemini"
  );
  if (!response.ok) throw new Error("Gemini failed with status " + response.status);
  const payload = await response.json();
  const text = payload?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error("Gemini returned an empty response.");
  return parseModelJson(text, "Gemini");
}

// ---- SECONDARY CHAIN: dedicated OCR, then a different LLM ------------------
// Deliberately a different vendor at both stages, so one provider's outage or
// quota exhaustion cannot take down both attempts.
// Tuned against https://ocr.space/ocrapi#ocrengine for a free-plan key.
//
// Engine 3 is the primary: it has the highest accuracy, recognises tables and
// returns them as Markdown, which matters here because partner name/share
// tables are exactly what we extract. Its free quota is 2,500 conversions a
// month, separate from the 25,000 shared by Engines 1/2 — ample for a path
// that only runs when the Gemini primary has already failed.
//
// Engine 2 is the fallback: the documented all-rounder, strong on noisy photo
// backgrounds and rotated text, drawing on the larger 25,000 pool.
//
// Every knob is a wrangler var, so this can be retuned without a code change.
const ENGINE_SUPPORTS_AUTO_LANGUAGE = ["2", "3"];

function ocrSpaceForm(file, env, engine) {
  const form = new FormData();
  form.append("file", file, file.name || "upload");
  // Engines 2 and 3 auto-detect; Engine 1 has no "auto" and needs a real code.
  const configured = env.OCRSPACE_LANGUAGE || "auto";
  const language = configured === "auto" && !ENGINE_SUPPORTS_AUTO_LANGUAGE.includes(String(engine))
    ? "eng"
    : configured;
  form.append("language", language);
  form.append("isOverlayRequired", "false");
  // The API defaults scale to false; enabling it upscales internally and the
  // docs call out a significant gain on low-resolution scans.
  form.append("scale", "true");
  form.append("detectOrientation", "true"); // deed photographed sideways
  // Engine 3 already emits Markdown tables, so forcing line-by-line output is
  // only useful on the other engines or for heavily tabular deeds.
  form.append("isTable", env.OCRSPACE_TABLE || "false");
  form.append("OCREngine", String(engine));
  return form;
}

async function callOcrSpace(file, env, engine) {
  const response = await withTimeout(
    fetch("https://api.ocr.space/parse/image", {
      method: "POST",
      headers: { apikey: env.OCRSPACE_API_KEY },
      body: ocrSpaceForm(file, env, engine),
    }),
    40000,
    "OCR.space"
  );
  if (!response.ok) throw new Error("OCR.space failed with status " + response.status);
  const payload = await response.json();

  // OCRExitCode: 1 success, 2 partial success, 3 failed, 4 fatal.
  const exitCode = Number(payload.OCRExitCode || 0);
  if (payload.IsErroredOnProcessing || exitCode >= 3) {
    const message = [payload.ErrorMessage, payload.ErrorDetails]
      .flat()
      .filter(Boolean)
      .join(" ");
    // The free plan reads only the first 3 pages of a PDF, which is the most
    // likely cause of a PDF failing here and is not obvious from the raw error.
    const hint = (file.type || "").includes("pdf") ? " (free plan reads at most 3 PDF pages)" : "";
    throw new Error("OCR.space error: " + (message || "exit code " + exitCode) + hint);
  }
  const pages = (payload.ParsedResults || []).map((result, index) => ({
    page: index + 1,
    text: result.ParsedText || "",
  }));
  const text = pages.map((page) => page.text).join("\n\n").trim();
  if (!text) throw new Error("OCR.space found no readable text.");
  return { pages, text, partial: exitCode === 2 };
}

export async function ocrWithOcrSpace(file, env) {
  // Free plans cap upload size well below the Worker's own limit, and an
  // oversized file fails with a message the agent cannot act on.
  const maxKb = Number(env.OCRSPACE_MAX_KB || 1024);
  if (file.size > maxKb * 1024) {
    throw new Error("Scan is " + Math.round(file.size / 1024) + " KB, above the OCR service limit of " + maxKb + " KB.");
  }
  const engine = (env.OCRSPACE_ENGINE || "3").trim();
  // ?? not ||, so setting the var to "" genuinely disables the retry.
  const fallback = (env.OCRSPACE_ENGINE_FALLBACK ?? "2").trim();
  try {
    return await callOcrSpace(file, env, engine);
  } catch (error) {
    // Engine availability differs by plan, and one engine can reject a file
    // the other reads. Trying the alternate costs one call and often works.
    if (!fallback || fallback === engine) throw error;
    console.log("ocr engine " + engine + " failed (" + error.message + "), retrying with " + fallback);
    return callOcrSpace(file, env, fallback);
  }
}

export async function extractWithGroq(ocrResult, env) {
  const model = env.GROQ_MODEL || "llama-3.3-70b-versatile";
  const response = await withTimeout(
    fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + env.GROQ_API_KEY },
      body: JSON.stringify({
        model,
        temperature: 0,
        response_format: { type: "json_object" },
        messages: [{ role: "user", content: buildTextPrompt(ocrResult) }],
      }),
    }),
    40000,
    "Groq"
  );
  if (!response.ok) throw new Error("Groq failed with status " + response.status);
  const payload = await response.json();
  const text = payload?.choices?.[0]?.message?.content;
  if (!text) throw new Error("Groq returned an empty response.");
  return parseModelJson(text, "Groq");
}

export async function runSecondaryChain(file, env) {
  if (!env.OCRSPACE_API_KEY) throw new Error("Secondary OCR is not configured.");
  if (!env.GROQ_API_KEY) throw new Error("Secondary LLM is not configured.");
  const ocrResult = await ocrWithOcrSpace(file, env);
  return extractWithGroq(ocrResult, env);
}
