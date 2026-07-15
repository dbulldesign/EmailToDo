import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export function loadProjects() {
  const raw = fs.readFileSync(
    path.join(__dirname, "..", "projects.json"),
    "utf8"
  );
  return JSON.parse(raw);
}

// --- Rules-based: keyword + sender matching, scored ---
export function categorizeByRules(email, config) {
  const hay = `${email.subject} ${email.preview}`.toLowerCase();
  const from = email.from.toLowerCase();

  let best = { name: config.fallback || "Uncategorized", score: 0 };

  for (const project of config.projects) {
    let score = 0;
    for (const kw of project.keywords || []) {
      if (hay.includes(kw.toLowerCase())) score += 2;
    }
    for (const s of project.senders || []) {
      if (from.includes(s.toLowerCase())) score += 3;
    }
    if (score > best.score) best = { name: project.name, score };
  }

  return best.name;
}

// --- AI-based: single Anthropic call classifies into known projects ---
export async function categorizeByAI(email, config) {
  const names = config.projects.map((p) => p.name);
  const system = `You classify emails into exactly one project category. ` +
    `Valid categories: ${names.join(", ")}, ${config.fallback || "Uncategorized"}. ` +
    `Respond with ONLY the category name, nothing else.`;

  const userMsg =
    `From: ${email.fromName} <${email.from}>\n` +
    `Subject: ${email.subject}\n` +
    `Preview: ${email.preview}`;

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 20,
      system,
      messages: [{ role: "user", content: userMsg }],
    }),
  });

  if (!res.ok) {
    throw new Error(`Anthropic API ${res.status}: ${await res.text()}`);
  }

  const data = await res.json();
  const answer = (data.content?.[0]?.text || "").trim();

  const valid = [...names, config.fallback || "Uncategorized"];
  const match = valid.find((v) => v.toLowerCase() === answer.toLowerCase());
  return match || config.fallback || "Uncategorized";
}

export async function categorize(email, config, mode) {
  if (mode === "ai") return categorizeByAI(email, config);
  return categorizeByRules(email, config);
}
