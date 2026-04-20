import { OAuth2Client } from "google-auth-library";
import { readFileSync } from "node:fs";
import { join } from "node:path";

interface CredentialsFile {
  client_id: string;
  client_secret: string;
  refresh_token: string;
  token?: string;
  token_uri?: string;
  scopes?: string[];
}

export function loadAuth(slug: string, tokenDir: string): OAuth2Client {
  const credPath = join(tokenDir, `google-${slug}-credentials.json`);

  let raw: string;
  try {
    raw = readFileSync(credPath, "utf-8");
  } catch {
    throw new Error(
      `Credentials file not found: ${credPath}\nRun add-google-account.sh to create it.`
    );
  }

  const creds: CredentialsFile = JSON.parse(raw);

  if (!creds.client_id || !creds.client_secret || !creds.refresh_token) {
    throw new Error(
      `Credentials file missing required fields (client_id, client_secret, refresh_token): ${credPath}`
    );
  }

  const client = new OAuth2Client(creds.client_id, creds.client_secret);
  client.setCredentials({
    refresh_token: creds.refresh_token,
    access_token: creds.token || undefined,
  });

  return client;
}
