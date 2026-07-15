import { PublicClientApplication } from "@azure/msal-node";
import express from "express";
import open from "open";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CACHE_PATH = path.join(__dirname, "..", ".token-cache.json");

// Scopes we need: read mail, and write only if applying categories.
export const SCOPES = ["Mail.Read", "Mail.ReadWrite", "User.Read"];

function buildApp() {
  const beforeCacheAccess = (ctx) => {
    if (fs.existsSync(CACHE_PATH)) {
      ctx.tokenCache.deserialize(fs.readFileSync(CACHE_PATH, "utf8"));
    }
  };
  const afterCacheAccess = (ctx) => {
    if (ctx.cacheHasChanged) {
      fs.writeFileSync(CACHE_PATH, ctx.tokenCache.serialize());
    }
  };

  return new PublicClientApplication({
    auth: {
      clientId: process.env.CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.TENANT_ID || "common"}`,
    },
    cache: { cachePlugin: { beforeCacheAccess, afterCacheAccess } },
  });
}

// Try silent token from cache first; fall back to interactive auth-code flow.
export async function getAccessToken() {
  const pca = buildApp();
  const accounts = await pca.getTokenCache().getAllAccounts();

  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: SCOPES,
      });
      if (result?.accessToken) return result.accessToken;
    } catch {
      // fall through to interactive
    }
  }

  return interactiveLogin(pca);
}

function interactiveLogin(pca) {
  const port = Number(process.env.PORT || 3000);
  const redirectUri =
    process.env.REDIRECT_URI || `http://localhost:${port}/callback`;

  return new Promise((resolve, reject) => {
    const app = express();
    let server;

    app.get("/callback", async (req, res) => {
      try {
        const result = await pca.acquireTokenByCode({
          code: req.query.code,
          scopes: SCOPES,
          redirectUri,
        });
        res.send(
          "<h2>Signed in.</h2><p>You can close this tab and return to the terminal.</p>"
        );
        server.close();
        resolve(result.accessToken);
      } catch (err) {
        res.status(500).send("Auth failed: " + err.message);
        server.close();
        reject(err);
      }
    });

    server = app.listen(port, async () => {
      const authUrl = await pca.getAuthCodeUrl({ scopes: SCOPES, redirectUri });
      console.log("\nOpening browser to sign in...");
      console.log("If it doesn't open, visit:\n" + authUrl + "\n");
      open(authUrl).catch(() => {});
    });
  });
}
