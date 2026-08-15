import { buildVisionPrompt } from "./prompt.js";

// Both chains read the document directly with a multimodal model. No OCR
// service sits in front of either one.
//
// This is deliberate. A dedicated free-tier OCR API caps file size and page
// count (OCR.space free: 1 MB, 3 pages), and a real partnership deed breaches
// both — so the fallback could never handle the documents that matter, and
// only failed on the days the primary was already down. A model that reads the
// PDF itself has neither limit, and tends to do better on phone photos, where
// skew and shadow cost classical OCR more than they cost a vision model.

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

async function fileToBase64(file) {
  return bytesToBase64(new Uint8Array(await file.arrayBuffer()));
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

// Read as much of the provider's error body as is useful. A bare status code
// rarely says whether the cause is the key, the model name, or the file.
async function describeFailure(response, label) {
  let body = "";
  try {
    body = (await response.text()).slice(0, 300).replace(/\s+/g, " ").trim();
  } catch (e) { /* body already consumed or absent */ }
  return new Error(label + " failed with status " + response.status + (body ? ": " + body : ""));
}

// ---- PRIMARY: Gemini reads the document directly ---------------------------
export async function extractWithGeminiVision(file, env) {
  const model = env.GEMINI_MODEL || "gemini-3.7-flash";
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
              { inline_data: { mime_type: file.type || "application/pdf", data: await fileToBase64(file) } },
            ],
          },
        ],
      }),
    }),
    60000,
    "Gemini"
  );
  if (!response.ok) throw await describeFailure(response, "Gemini");
  const payload = await response.json();
  const text = payload?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    // A blocked or truncated response has no text part; say which.
    const reason = payload?.candidates?.[0]?.finishReason || payload?.promptFeedback?.blockReason || "no text in response";
    throw new Error("Gemini returned no output (" + reason + ").");
  }
  return parseModelJson(text, "Gemini");
}

// ---- SECONDARY: OpenRouter, a different vendor, also multimodal ------------
// OpenRouter fronts many models behind one OpenAI-compatible endpoint, so
// switching the fallback model — when a free one is retired, rate-limited, or
// simply reads deeds better — is an OPENROUTER_MODEL change, not a code change.
export async function extractWithOpenRouter(file, env) {
  const model = env.OPENROUTER_MODEL || "google/gemma-4-31b-it:free";
  const dataUrl = "data:" + (file.type || "application/pdf") + ";base64," + (await fileToBase64(file));
  const isPdf = (file.type || "").includes("pdf");

  const content = [{ type: "text", text: buildVisionPrompt() }];
  if (isPdf) {
    // PDFs go as a file part; OpenRouter's parser plugin turns them into
    // something the model can read. Engine is configurable because the
    // available parsers and their costs change.
    content.push({ type: "file", file: { filename: file.name || "document.pdf", file_data: dataUrl } });
  } else {
    content.push({ type: "image_url", image_url: { url: dataUrl } });
  }

  const body = {
    model,
    temperature: 0,
    response_format: { type: "json_object" },
    messages: [{ role: "user", content }],
  };
  if (isPdf) {
    body.plugins = [{ id: "file-parser", pdf: { engine: env.OPENROUTER_PDF_ENGINE || "pdf-text" } }];
  }

  const response = await withTimeout(
    fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + env.OPENROUTER_API_KEY,
        // OpenRouter attributes traffic by these; harmless if unset.
        "HTTP-Referer": (env.ALLOWED_ORIGINS || "").split(",")[0].trim(),
        "X-Title": "Merchant Onboarding Document Generator",
      },
      body: JSON.stringify(body),
    }),
    60000,
    "OpenRouter"
  );
  if (!response.ok) throw await describeFailure(response, "OpenRouter");
  const payload = await response.json();
  if (payload?.error) throw new Error("OpenRouter error: " + (payload.error.message || JSON.stringify(payload.error)));
  const text = payload?.choices?.[0]?.message?.content;
  if (!text) throw new Error("OpenRouter returned an empty response.");
  return parseModelJson(text, "OpenRouter");
}

// ---- TERTIARY: Groq, a third vendor -----------------------------------------
// Groq's vision models take images, not PDFs. That makes it a genuine last
// resort for photo uploads — which is what agents in the field actually
// produce — and a clear, immediate failure for PDFs rather than a confusing
// provider error.
export async function extractWithGroq(file, env) {
  const model = env.GROQ_MODEL || "qwen/qwen3.6-27b";
  if ((file.type || "").includes("pdf")) {
    throw new Error("Groq accepts images only, and this upload is a PDF");
  }
  const dataUrl = "data:" + (file.type || "image/jpeg") + ";base64," + (await fileToBase64(file));
  const response = await withTimeout(
    fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + env.GROQ_API_KEY },
      body: JSON.stringify({
        model,
        temperature: 0,
        response_format: { type: "json_object" },
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: buildVisionPrompt() },
              { type: "image_url", image_url: { url: dataUrl } },
            ],
          },
        ],
      }),
    }),
    60000,
    "Groq"
  );
  if (!response.ok) throw await describeFailure(response, "Groq");
  const payload = await response.json();
  if (payload?.error) throw new Error("Groq error: " + (payload.error.message || JSON.stringify(payload.error)));
  const text = payload?.choices?.[0]?.message?.content;
  if (!text) throw new Error("Groq returned an empty response.");
  return parseModelJson(text, "Groq");
}

// The ordered chain. Each tier is a different vendor, so one outage, quota
// exhaustion or model retirement cannot take out the ones behind it.
export const PROVIDER_CHAIN = [
  { name: "gemini", keyVar: "GEMINI_API_KEY", run: extractWithGeminiVision },
  { name: "openrouter", keyVar: "OPENROUTER_API_KEY", run: extractWithOpenRouter },
  { name: "groq", keyVar: "GROQ_API_KEY", run: extractWithGroq },
];
