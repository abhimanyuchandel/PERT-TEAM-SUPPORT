#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const http = require("http");
const crypto = require("crypto");

const HOST = process.env.HOST || "127.0.0.1";
const PORT = Number(process.env.PORT || 3000);
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const TOKEN_TTL_SECONDS = Number(process.env.TOKEN_TTL_SECONDS || 1800);
const TOKEN_TTL_MS = Math.max(60, TOKEN_TTL_SECONDS) * 1000;
const ALLOWED_MODELS = new Set(["gpt-4.1-mini", "gpt-4.1"]);

const tokenStore = new Map();
const workspaceRoot = process.cwd();

function cleanupExpiredTokens() {
  const now = Date.now();
  for (const [token, expiresAt] of tokenStore.entries()) {
    if (expiresAt <= now) {
      tokenStore.delete(token);
    }
  }
}

setInterval(cleanupExpiredTokens, 60000).unref();

function issueToken() {
  const token = crypto.randomBytes(24).toString("hex");
  tokenStore.set(token, Date.now() + TOKEN_TTL_MS);
  return token;
}

function validateToken(token) {
  if (!token) return false;
  const expiresAt = tokenStore.get(token);
  if (!expiresAt) return false;
  if (expiresAt <= Date.now()) {
    tokenStore.delete(token);
    return false;
  }
  tokenStore.set(token, Date.now() + TOKEN_TTL_MS);
  return true;
}

function json(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(body)
  });
  res.end(body);
}

function text(res, statusCode, message) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    "Cache-Control": "no-store"
  });
  res.end(message);
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > 2 * 1024 * 1024) {
        reject(new Error("Request body too large"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => {
      if (!chunks.length) {
        resolve({});
        return;
      }
      const raw = Buffer.concat(chunks).toString("utf8");
      try {
        resolve(JSON.parse(raw));
      } catch (err) {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function extractAuthToken(req) {
  const auth = req.headers.authorization || "";
  const match = auth.match(/^Bearer\s+(.+)$/i);
  if (match && match[1]) return match[1].trim();
  return "";
}

function extractResponseText(payload) {
  if (!payload) return "";
  if (typeof payload.output_text === "string") return payload.output_text;
  if (Array.isArray(payload.output_text)) return payload.output_text.join("\n");
  if (Array.isArray(payload.output)) {
    const chunks = [];
    for (const item of payload.output) {
      if (!item || !Array.isArray(item.content)) continue;
      for (const part of item.content) {
        if (part && typeof part.text === "string") {
          chunks.push(part.text);
        }
      }
    }
    if (chunks.length) return chunks.join("\n");
  }
  if (
    payload.choices &&
    payload.choices[0] &&
    payload.choices[0].message &&
    typeof payload.choices[0].message.content === "string"
  ) {
    return payload.choices[0].message.content;
  }
  return "";
}

function parseBullets(text) {
  const lines = (text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const bullets = lines
    .map((line) => line.replace(/^[-*•]\s+/, "").replace(/^\d+\.\s+/, ""))
    .filter(Boolean);
  if (!bullets.length && text.trim().length) {
    return [text.trim()];
  }
  return bullets.slice(0, 6);
}

function contentTypeFor(ext) {
  if (ext === ".html") return "text/html; charset=utf-8";
  if (ext === ".js") return "text/javascript; charset=utf-8";
  if (ext === ".css") return "text/css; charset=utf-8";
  if (ext === ".json") return "application/json; charset=utf-8";
  if (ext === ".txt") return "text/plain; charset=utf-8";
  if (ext === ".pdf") return "application/pdf";
  if (ext === ".docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (ext === ".pptx") return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  return "application/octet-stream";
}

function safePathFromUrl(urlPath) {
  const relative = urlPath === "/" ? "/index.html" : urlPath;
  const normalized = path.normalize(relative).replace(/^(\.\.[/\\])+/, "");
  const resolved = path.resolve(workspaceRoot, `.${normalized}`);
  if (!resolved.startsWith(workspaceRoot)) return null;
  return resolved;
}

async function handleToken(req, res) {
  if (!OPENAI_API_KEY) {
    json(res, 503, { error: "OPENAI_API_KEY is not configured on server." });
    return;
  }
  const token = issueToken();
  json(res, 200, {
    token,
    expiresInSeconds: Math.floor(TOKEN_TTL_MS / 1000)
  });
}

async function handleAiAddendum(req, res) {
  if (!OPENAI_API_KEY) {
    json(res, 503, { error: "OPENAI_API_KEY is not configured on server." });
    return;
  }

  const token = extractAuthToken(req);
  if (!validateToken(token)) {
    json(res, 401, { error: "Invalid or expired session token." });
    return;
  }

  let body;
  try {
    body = await parseBody(req);
  } catch (err) {
    json(res, 400, { error: err.message });
    return;
  }

  const model = ALLOWED_MODELS.has(body.model) ? body.model : "gpt-4.1-mini";
  const narrative = typeof body.narrative === "string" ? body.narrative.trim() : "";
  const profile = body.profile && typeof body.profile === "object" ? body.profile : {};

  if (!narrative) {
    json(res, 400, { error: "Narrative is required." });
    return;
  }

  const prompt =
    `Structured profile:\n` +
    `- Category: ${profile.category || "n/a"}\n` +
    `- Profile: ${profile.descriptor || "n/a"}\n` +
    `- Diagnosis status: ${profile.diagnosisStatus || "n/a"}\n` +
    `- Hemodynamics: persistent hypotension=${profile.hemodynamics?.persistentHypotension ? "yes" : "no"}, transient hypotension=${profile.hemodynamics?.transientHypotension ? "yes" : "no"}, MAP=${profile.hemodynamics?.map ?? "n/a"}, lactate=${profile.hemodynamics?.lactate ?? "n/a"}, vasopressors=${profile.hemodynamics?.vasopressors ?? "n/a"}\n` +
    `- Respiratory support: ${profile.respiratory?.oxygenSupport ?? "n/a"}, RR=${profile.respiratory?.rr ?? "n/a"}\n` +
    `- Contraindications: anticoag=${profile.contraindications?.anticoagulation ? "yes" : "no"}, thrombolysis=${profile.contraindications?.thrombolysis ? "yes" : "no"}, high bleeding risk=${profile.contraindications?.highBleedingRisk ? "yes" : "no"}\n` +
    `- Special populations: pregnancy=${profile.specialPopulations?.pregnancy ? "yes" : "no"}, breastfeeding=${profile.specialPopulations?.breastfeeding ? "yes" : "no"}, APS=${profile.specialPopulations?.aps ? "yes" : "no"}, severe CKD=${profile.specialPopulations?.severeCKD ? "yes" : "no"}\n` +
    `- Existing immediate strategy items: ${Array.isArray(profile.immediateStrategy) ? profile.immediateStrategy.join(" | ") : "n/a"}\n` +
    `- Existing medication strategy items: ${Array.isArray(profile.medicationStrategy) ? profile.medicationStrategy.join(" | ") : "n/a"}\n\n` +
    `De-identified case narrative:\n${narrative}\n\n` +
    "Return only bullet points for immediate care (0-24 hours), avoiding any mention of patient identifiers.";

  try {
    const openaiResp = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model,
        input: [
          {
            role: "system",
            content:
              "You are a pulmonary embolism clinical support assistant. Use the structured and narrative inputs to produce immediate, actionable recommendations for frontline clinicians. Do not include patient identifiers. Output 3 to 6 concise bullet points only."
          },
          {
            role: "user",
            content: prompt
          }
        ],
        temperature: 0.2
      })
    });

    if (!openaiResp.ok) {
      const errText = await openaiResp.text();
      json(res, openaiResp.status, {
        error: "OpenAI request failed.",
        details: errText.slice(0, 1000)
      });
      return;
    }

    const payload = await openaiResp.json();
    const bullets = parseBullets(extractResponseText(payload));
    if (!bullets.length) {
      json(res, 502, { error: "OpenAI response did not contain recommendation text." });
      return;
    }
    json(res, 200, { bullets });
  } catch (err) {
    json(res, 500, { error: `AI request failed: ${err.message}` });
  }
}

function handleStatic(req, res, pathname) {
  const filePath = safePathFromUrl(pathname);
  if (!filePath) {
    text(res, 403, "Forbidden");
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (err.code === "ENOENT") {
        text(res, 404, "Not found");
        return;
      }
      text(res, 500, "Server error");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
    res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
    res.writeHead(200, {
      "Content-Type": contentTypeFor(ext),
      "Cache-Control": ext === ".html" ? "no-cache" : "public, max-age=300"
    });
    res.end(data);
  });
}

const server = http.createServer(async (req, res) => {
  const reqUrl = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  const pathname = reqUrl.pathname;

  if (req.method === "OPTIONS") {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
    res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
    res.writeHead(204, {
      Allow: "GET,POST,OPTIONS"
    });
    res.end();
    return;
  }

  if (req.method === "POST" && pathname === "/api/token") {
    await handleToken(req, res);
    return;
  }

  if (req.method === "POST" && pathname === "/api/ai-addendum") {
    await handleAiAddendum(req, res);
    return;
  }

  if (req.method === "GET" || req.method === "HEAD") {
    handleStatic(req, res, pathname);
    return;
  }

  text(res, 405, "Method not allowed");
});

server.listen(PORT, HOST, () => {
  const keyStatus = OPENAI_API_KEY ? "configured" : "missing";
  console.log(`PE tool server listening on http://${HOST}:${PORT} (OPENAI_API_KEY ${keyStatus})`);
});
