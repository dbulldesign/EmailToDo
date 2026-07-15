import "dotenv/config";
import express from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { getAccessToken } from "./auth.js";
import { fetchInbox, applyCategory } from "./graph.js";
import { loadProjects, categorize } from "./categorize.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = Number(process.env.PORT || 3000);

app.use(express.static(path.join(__dirname, "..", "public")));

app.get("/api/emails", async (req, res) => {
  try {
    const config = loadProjects();
    const mode = process.env.CATEGORIZE_MODE || "rules";
    const count = Number(process.env.FETCH_COUNT || 50);
    const apply = String(process.env.APPLY_CATEGORIES).toLowerCase() === "true";

    const token = await getAccessToken();
    const emails = await fetchInbox(token, count);

    const buckets = {};
    for (const email of emails) {
      const project = await categorize(email, config, mode);
      (buckets[project] ||= []).push({
        subject: email.subject,
        from: email.fromName || email.from,
        received: email.received,
        webLink: email.webLink,
        preview: email.preview.slice(0, 140),
      });
      if (apply) {
        try {
          await applyCategory(token, email.id, project);
        } catch {
          /* non-fatal */
        }
      }
    }

    res.json({ mode, applied: apply, buckets });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Dashboard running at http://localhost:${PORT}`);
  console.log("Note: the /callback route is handled by the sign-in flow on first load.");
});
