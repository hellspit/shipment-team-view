// Shared Microsoft Graph (app-only) helpers.
//
// Azure app needs this Graph APPLICATION permission + admin consent:
//   Sites.ReadWrite.All — read and update the "Shipment Requests" list

let tokenCache: { token: string; exp: number } | null = null;

export async function getGraphToken(): Promise<string> {
  if (tokenCache && Date.now() < tokenCache.exp - 60_000) return tokenCache.token;
  const body = new URLSearchParams({
    client_id: process.env.AZURE_CLIENT_ID ?? "",
    client_secret: process.env.AZURE_CLIENT_SECRET ?? "",
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );
  const j = await res.json();
  if (!res.ok) throw new Error(j.error_description || "Graph token error");
  tokenCache = { token: j.access_token, exp: Date.now() + j.expires_in * 1000 };
  return tokenCache.token;
}

export type GraphError = Error & { status?: number };

/** Call Graph with the app-only token. Throws a GraphError carrying the HTTP status. */
export async function graph<T = unknown>(
  path: string,
  init: RequestInit = {}
): Promise<T> {
  const token = await getGraphToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init.body ? { "Content-Type": "application/json" } : {}),
      ...(init.headers ?? {}),
    },
  });
  const text = await res.text();
  const parsed = text ? JSON.parse(text) : {};
  if (!res.ok) {
    const err = new Error(parsed?.error?.message || `Graph ${res.status}`) as GraphError;
    err.status = res.status;
    throw err;
  }
  return parsed as T;
}
