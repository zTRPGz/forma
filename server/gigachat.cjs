"use strict";

const https = require("https");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const tls = require("tls");

const OAUTH_HOST = "ngw.devices.sberbank.ru";
const OAUTH_PATH = "/api/v2/oauth";
const API_HOST = "gigachat.devices.sberbank.ru";
const API_PATH = "/api/v1/chat/completions";
const SCOPE = process.env.GIGACHAT_SCOPE || "GIGACHAT_API_PERS";
const MODEL = "GigaChat-2";
const MAX_TOKENS = 4096;
const MAX_MESSAGES = 8;
const MAX_MESSAGE_CHARS = 8000;
const MAX_TOTAL_CHARS = 24000;
const ROLES = { system: true, user: true, assistant: true };

const CERT_DIR = path.join(__dirname, "certs");
const TRUSTED_CA = ["russian_trusted_root_ca_pem.crt", "russian_trusted_sub_ca_pem.crt"]
  .map(function (name) {
    return fs.readFileSync(path.join(CERT_DIR, name));
  })
  .concat(tls.rootCertificates);

const SITE_ORIGIN = "https://ztrpgz.github.io";
const ALLOWED_ORIGIN = [
  /^https:\/\/forma-docs\.duckdns\.org$/i,
  /^https:\/\/forma(?:-[a-z0-9]+)*-trpg4\.vercel\.app$/i,
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
      Object.assign({ ca: TRUSTED_CA, minVersion: "TLSv1.2", family: 4 }, options),
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
  if (!key) throw new Error("Сервис ИИ не настроен: нет ключа доступа");

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
        timeout: 10000,
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
    throw new Error("Сервис ИИ недоступен: " + error.message);
  }

  if (response.status !== 200) {
    throw new Error("Сервис ИИ вернул HTTP " + response.status + ": " + response.body.slice(0, 200));
  }

  let data;
  try {
    data = JSON.parse(response.body);
  } catch (error) {
    throw new Error("Сервис ИИ вернул не JSON: " + response.body.slice(0, 200));
  }
  if (!data || typeof data.access_token !== "string" || !data.access_token) {
    throw new Error("Сервис ИИ не выдал access_token: " + response.body.slice(0, 200));
  }
  token = data.access_token;
  tokenExpires = data.expires_at || Date.now() + 25 * 60 * 1000;
  return token;
}

function sendChat(headers, body, timeout) {
  return request(
    { host: API_HOST, path: API_PATH, method: "POST", timeout: timeout, headers: headers },
    body
  );
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
    response = await sendChat(headers, body, 40000);
  } catch (error) {
    if (String(error.message) === "timeout") {
      throw new Error("Сервис ИИ не ответил за 40 секунд.");
    }
    try {
      response = await sendChat(headers, body, 15000);
    } catch (retryError) {
      throw new Error("Сервис ИИ недоступен: " + retryError.message);
    }
  }

  if (response.status === 401) {
    token = "";
    const fresh = await getToken(true);
    headers.Authorization = "Bearer " + fresh;
    response = await sendChat(headers, body, 25000);
  }

  return response;
}

function clampNumber(value, fallback, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(max, Math.max(min, number));
}

function checkMessages(messages) {
  if (!Array.isArray(messages) || !messages.length) {
    return { error: "В запросе нет messages." };
  }
  if (messages.length > MAX_MESSAGES) {
    return { error: "Слишком много сообщений в запросе." };
  }

  const result = [];
  let total = 0;
  for (let index = 0; index < messages.length; index += 1) {
    const message = messages[index];
    if (!message || typeof message !== "object" || !ROLES[message.role] || typeof message.content !== "string") {
      return { error: "Сообщение " + (index + 1) + " некорректно." };
    }
    const content = message.content.trim();
    if (!content) {
      return { error: "Сообщение " + (index + 1) + " пустое." };
    }
    if (content.length > MAX_MESSAGE_CHARS) {
      return { error: "Сообщение " + (index + 1) + " длиннее " + MAX_MESSAGE_CHARS + " знаков." };
    }
    total += content.length;
    if (total > MAX_TOTAL_CHARS) {
      return { error: "Запрос длиннее " + MAX_TOTAL_CHARS + " знаков." };
    }
    result.push({ role: message.role, content: content });
  }
  return { messages: result };
}

function normalizePayload(input) {
  const source = input && typeof input === "object" ? input : {};
  const checked = checkMessages(source.messages);
  if (checked.error) return { error: checked.error };

  const payload = {
    model: MODEL,
    messages: checked.messages,
    temperature: clampNumber(source.temperature, 0.7, 0, 2),
    max_tokens: Math.floor(clampNumber(source.max_tokens, MAX_TOKENS, 1, MAX_TOKENS)),
  };

  if (source.response_format && source.response_format.type === "json_object") {
    payload.response_format = { type: "json_object" };
  }
  return { payload: payload };
}

function allowedOrigin(origin) {
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
  corsHeaders: corsHeaders,
  normalizePayload: normalizePayload,
  callChat: callChat,
};
