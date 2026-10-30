"use strict";

const https = require("https");
const crypto = require("crypto");

const OAUTH_HOST = "ngw.devices.sberbank.ru";
const OAUTH_PATH = "/api/v2/oauth";
const API_HOST = "gigachat.devices.sberbank.ru";
const API_PATH = "/api/v1/chat/completions";
const SCOPE = process.env.GIGACHAT_SCOPE || "GIGACHAT_API_PERS";
const MODEL = "GigaChat-2";
const MAX_TOKENS = 4096;

const SITE_ORIGIN = "https://ztrpgz.github.io";
const ALLOWED_ORIGIN = [
  /^https:\/\/[a-z0-9-]+\.netlify\.app$/i,
  /^https:\/\/[a-z0-9-]+\.vercel\.app$/i,
  /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i,
];

let token = "";
let tokenExpires = 0;

function authKey() {
  const raw = String(process.env.GIGACHAT_API_KEY || "").trim();
  if (!raw) return "";
  if (raw.indexOf(":") > 0) return Buffer.from(raw).toString("base64");
  return raw;
}

function request(options, body) {
  return new Promise(function (resolve, reject) {
    const req = https.request(
      Object.assign({ rejectUnauthorized: false, family: 4 }, options),
      function (res) {
        let data = "";
        res.setEncoding("utf8");
        res.on("data", function (chunk) {
          data += chunk;
        });
        res.on("end", function () {
          resolve({ status: res.statusCode || 0, body: data });
        });
      }
    );

    req.on("error", reject);
    req.setTimeout(options.timeout || 60000, function () {
      req.destroy(new Error("timeout"));
    });
    if (body) req.write(body);
    req.end();
  });
}

async function getToken(force) {
  const key = authKey();
  if (!key) throw new Error("GigaChat не настроен: нет GIGACHAT_API_KEY");

  if (!force && token && Date.now() < tokenExpires - 60000) return token;

  const body = "scope=" + encodeURIComponent(SCOPE);
  let response;
  try {
    response = await request(
      {
        host: OAUTH_HOST,
        port: 9443,
        path: OAUTH_PATH,
        method: "POST",
        timeout: 30000,
        headers: {
          Authorization: "Basic " + key,
          RqUID: crypto.randomUUID(),
          "Content-Type": "application/x-www-form-urlencoded",
          "Content-Length": Buffer.byteLength(body),
        },
      },
      body
    );
  } catch (error) {
    throw new Error("OAuth GigaChat недоступен: " + error.message);
  }

  if (response.status !== 200) {
    throw new Error("OAuth GigaChat вернул HTTP " + response.status + ": " + response.body.slice(0, 200));
  }

  const data = JSON.parse(response.body);
  token = data.access_token;
  tokenExpires = data.expires_at || Date.now() + 25 * 60 * 1000;
  return token;
}

async function callChat(payload) {
  const access = await getToken(false);
  const body = JSON.stringify(payload);
  const headers = {
    Authorization: "Bearer " + access,
    "Content-Type": "application/json",
    "X-Request-ID": crypto.randomUUID(),
    "Content-Length": Buffer.byteLength(body),
  };

  let response;
  try {
    response = await request(
      { host: API_HOST, path: API_PATH, method: "POST", timeout: 60000, headers: headers },
      body
    );
  } catch (error) {
    throw new Error("GigaChat недоступен: " + error.message);
  }

  if (response.status === 401) {
    token = "";
    const fresh = await getToken(true);
    headers.Authorization = "Bearer " + fresh;
    response = await request(
      { host: API_HOST, path: API_PATH, method: "POST", timeout: 60000, headers: headers },
      body
    );
  }

  return response;
}

function normalizePayload(input) {
  const source = input && typeof input === "object" ? input : {};
  const messages = Array.isArray(source.messages) ? source.messages : null;
  if (!messages || !messages.length) {
    return { error: "В запросе нет messages." };
  }

  const payload = {
    model: MODEL,
    messages: messages,
    temperature: Math.min(2, Math.max(0, Number(source.temperature) || 0.7)),
    max_tokens: Math.min(MAX_TOKENS, Math.max(1, Number(source.max_tokens) || MAX_TOKENS)),
  };

  if (source.response_format && source.response_format.type) {
    payload.response_format = source.response_format;
  }
  return { payload: payload };
}

function allowedOrigin(origin) {
  if (origin === SITE_ORIGIN) return origin;
  return ALLOWED_ORIGIN.some(function (pattern) {
    return pattern.test(origin);
  })
    ? origin
    : SITE_ORIGIN;
}

function corsHeaders(origin) {
  return {
    "Access-Control-Allow-Origin": allowedOrigin(String(origin || "")),
    Vary: "Origin",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
    "Access-Control-Max-Age": "86400",
    "Cache-Control": "no-store",
  };
}

module.exports = {
  MODEL: MODEL,
  authKey: authKey,
  corsHeaders: corsHeaders,
  normalizePayload: normalizePayload,
  callChat: callChat,
};
