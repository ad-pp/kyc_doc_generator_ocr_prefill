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
export async function ocrWithOcrSpace(file, env) {
  const form = new FormData();
  form.append("file", file, file.name || "upload");
  form.append("language", "eng");
  form.append("isOverlayRequired", "false");
  form.append("scale", "true");
  form.append("OCREngine", "2");
  const response = await withTimeout(
    fetch("https://api.ocr.space/parse/image", { method: "POST", headers: { apikey: env.OCRSPACE_API_KEY }, body: form }),
    40000,
    "OCR.space"
  );
  if (!response.ok) throw new Error("OCR.space failed with status " + response.status);
  const payload = await response.json();
  if (payload.IsErroredOnProcessing) {
    const message = Array.isArray(payload.ErrorMessage) ? payload.ErrorMessage.join(" ") : payload.ErrorMessage;
    throw new Error("OCR.space error: " + (message || "unknown"));
  }
  const pages = (payload.ParsedResults || []).map((result, index) => ({ page: index + 1, text: result.ParsedText || "" }));
  const text = pages.map((page) => page.text).join("\n\n").trim();
  if (!text) throw new Error("OCR.space found no readable text.");
  return { pages, text };
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
