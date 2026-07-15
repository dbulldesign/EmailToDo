import "dotenv/config";
import { getAccessToken } from "./auth.js";
import { fetchInbox, applyCategory } from "./graph.js";
import { loadProjects, categorize } from "./categorize.js";

function requireConfig() {
  if (!process.env.CLIENT_ID || process.env.CLIENT_ID.startsWith("your-")) {
    console.error(
      "Missing CLIENT_ID. Copy .env.example to .env and fill in your Azure app details (see README)."
    );
    process.exit(1);
  }
}

async function run() {
  requireConfig();
  const command = process.argv[2] || "categorize";
  const config = loadProjects();

  if (command === "list-projects") {
    console.log("Configured projects:");
    for (const p of config.projects) console.log("  - " + p.name);
    console.log("  - " + (config.fallback || "Uncategorized") + " (fallback)");
    return;
  }

  const mode = process.env.CATEGORIZE_MODE || "rules";
  const count = Number(process.env.FETCH_COUNT || 50);
  const apply = String(process.env.APPLY_CATEGORIES).toLowerCase() === "true";

  console.log("Signing in to Microsoft Graph...");
  const token = await getAccessToken();

  console.log(`Fetching ${count} most recent inbox messages...`);
  const emails = await fetchInbox(token, count);
  console.log(`Got ${emails.length} messages. Categorizing (mode: ${mode})...\n`);

  const buckets = {};
  for (const email of emails) {
    const project = await categorize(email, config, mode);
    (buckets[project] ||= []).push(email);

    if (apply) {
      try {
        await applyCategory(token, email.id, project);
      } catch (err) {
        console.warn(`  ! Could not tag "${email.subject}": ${err.message}`);
      }
    }
  }

  for (const [project, list] of Object.entries(buckets)) {
    console.log(`\n=== ${project} (${list.length}) ===`);
    for (const e of list) {
      const date = new Date(e.received).toLocaleDateString();
      console.log(`  [${date}] ${e.fromName || e.from}: ${e.subject}`);
    }
  }

  if (apply) {
    console.log(
      "\nDone. Project names were written as Outlook categories on each message."
    );
  } else {
    console.log(
      "\nDone (read-only). Set APPLY_CATEGORIES=true in .env to tag messages in Outlook."
    );
  }
}

run().catch((err) => {
  console.error("\nError:", err.message);
  process.exit(1);
});
