import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";

function client(accessToken) {
  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
}

// Pull the most recent N messages from the inbox.
export async function fetchInbox(accessToken, count = 50) {
  const graph = client(accessToken);
  const res = await graph
    .api("/me/mailFolders/inbox/messages")
    .top(Math.min(count, 100))
    .select([
      "id",
      "subject",
      "bodyPreview",
      "from",
      "receivedDateTime",
      "categories",
      "webLink",
    ])
    .orderby("receivedDateTime DESC")
    .get();

  return (res.value || []).map((m) => ({
    id: m.id,
    subject: m.subject || "(no subject)",
    preview: m.bodyPreview || "",
    from: m.from?.emailAddress?.address || "",
    fromName: m.from?.emailAddress?.name || "",
    received: m.receivedDateTime,
    categories: m.categories || [],
    webLink: m.webLink,
  }));
}

// Write a project name into the message's Outlook categories.
export async function applyCategory(accessToken, messageId, projectName) {
  const graph = client(accessToken);
  const current = await graph
    .api(`/me/messages/${messageId}`)
    .select("categories")
    .get();

  const set = new Set(current.categories || []);
  set.add(projectName);

  await graph
    .api(`/me/messages/${messageId}`)
    .patch({ categories: [...set] });
}
