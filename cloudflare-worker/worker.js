const ALLOWED_ORIGINS = new Set([
  "https://valeriyameteleva.github.io",
  "http://127.0.0.1:8765",
  "http://127.0.0.1:8766",
  "http://localhost:8765",
  "http://localhost:8766"
]);

const ALLOWED_ACTIONS = new Set([
  "parentLogin",
  "childLogin",
  "readState",
  "verifyPin",
  "syncState",
  "addMember"
]);
const MAX_BODY_BYTES = 256 * 1024;

function jsonResponse(body, status, origin) {
  const headers = new Headers({
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Vary": "Origin"
  });
  if (origin && ALLOWED_ORIGINS.has(origin)) {
    headers.set("Access-Control-Allow-Origin", origin);
    headers.set("Access-Control-Allow-Methods", "POST, OPTIONS");
    headers.set("Access-Control-Allow-Headers", "Content-Type");
  }
  return new Response(JSON.stringify(body), { status, headers });
}

function isValidScriptUrl(value) {
  try {
    const url = new URL(value);
    return url.protocol === "https:" &&
      url.hostname === "script.google.com" &&
      /^\/macros\/s\/[^/]+\/exec\/?$/.test(url.pathname);
  } catch (error) {
    return false;
  }
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    if (request.method === "GET") {
      const configured = isValidScriptUrl(env.APPS_SCRIPT_URL);
      return jsonResponse({
        ok: configured,
        service: "family-chores-proxy",
        message: configured
          ? "Worker is deployed and the Apps Script URL is configured."
          : "Set the APPS_SCRIPT_URL secret."
      }, configured ? 200 : 503, origin);
    }

    if (!ALLOWED_ORIGINS.has(origin)) {
      return jsonResponse({ ok: false, error: "Origin is not allowed." }, 403, "");
    }

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          "Access-Control-Allow-Origin": origin,
          "Access-Control-Allow-Methods": "POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type",
          "Access-Control-Max-Age": "86400",
          "Vary": "Origin"
        }
      });
    }

    if (request.method !== "POST") {
      return jsonResponse({ ok: false, error: "Use POST for API requests." }, 405, origin);
    }
    if (!isValidScriptUrl(env.APPS_SCRIPT_URL)) {
      return jsonResponse({ ok: false, error: "Apps Script URL is not configured." }, 503, origin);
    }

    const declaredLength = Number(request.headers.get("Content-Length") || 0);
    if (declaredLength > MAX_BODY_BYTES) {
      return jsonResponse({ ok: false, error: "Request is too large." }, 413, origin);
    }

    let requestBody;
    try {
      const bodyText = await request.text();
      if (new TextEncoder().encode(bodyText).byteLength > MAX_BODY_BYTES) {
        return jsonResponse({ ok: false, error: "Request is too large." }, 413, origin);
      }
      requestBody = JSON.parse(bodyText);
    } catch (error) {
      return jsonResponse({ ok: false, error: "Request body must be valid JSON." }, 400, origin);
    }

    if (!requestBody || typeof requestBody.action !== "string" ||
        !ALLOWED_ACTIONS.has(requestBody.action) ||
        !requestBody.payload || typeof requestBody.payload !== "object" ||
        Array.isArray(requestBody.payload)) {
      return jsonResponse({ ok: false, error: "Invalid API request." }, 400, origin);
    }

    let upstream;
    try {
      upstream = await fetch(env.APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
      });
    } catch (error) {
      return jsonResponse({ ok: false, error: "Could not reach Apps Script." }, 502, origin);
    }

    let result;
    try {
      result = await upstream.json();
    } catch (error) {
      return jsonResponse({ ok: false, error: "Apps Script returned an invalid response." }, 502, origin);
    }

    if (!result || typeof result.ok !== "boolean") {
      return jsonResponse({ ok: false, error: "Apps Script returned an invalid response." }, 502, origin);
    }
    return jsonResponse(result, result.ok ? 200 : 400, origin);
  }
};
