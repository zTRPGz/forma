"use strict";

// Локальный OpenAI-совместимый прокси к GigaChat.
// Браузер не может обращаться к GigaChat напрямую: сертификат Сбера не входит
// в доверенные, а OAuth-сервер не отдаёт CORS-заголовки. Прокси делает OAuth
// на этой машине и пересылает запросы, ключ остаётся в локальном файле.

const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT || 8787);
const OAUTH_URL = process.env.GIGACHAT_OAUTH_URL || "https://ngw.devices.sberbank.ru:9443/api/v2/oauth";
const API_URL = process.env.GIGACHAT_API_URL || "https://gigachat.devices.sberbank.ru/api/v1/chat/completions";
const SCOPE = process.env.GIGACHAT_SCOPE || "GIGACHAT_API_PERS";
const MAX_TOKENS_LIMIT = 4096;

function readKey() {
  const fromEnv = String(process.env.GIGACHAT_API_KEY || "").trim();
  if (fromEnv) return fromEnv;

  const file = path.join(__dirname, "gigachat-key.txt");
  if (!fs.existsSync(file)) return "";

  const value = fs
    .readFileSync(file, "utf8")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith("#"))
    .join("");

  return value;
}

const rawKey = readKey();
const authKey = rawKey.includes(":") ? Buffer.from(rawKey).toString("base64") : rawKey;

function request(url, { method = "POST", headers = {}, body = null } = {}) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      url,
      {
        method,
        headers,
        timeout: 120000,
        // GigaChat использует сертификаты российского УЦ, которых нет в
        // системном хранилище Node. Проверка отключена только для этого
        // соединения, наружу уходят лишь запросы к api.gigachat.
        rejectUnauthorized: false,
      },
      (res) => {
        let data = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => (data += chunk));
        res.on("end", () => resolve({ status: res.statusCode, body: data }));
      },
    );
    req.on("timeout", () => req.destroy(new Error("таймаут запроса к GigaChat")));
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

let token = "";
let tokenExpiresAt = 0;

async function getToken(force) {
  if (!authKey) {
    throw new Error("нет ключа GigaChat: положите его в proxy/gigachat-key.txt или задайте GIGACHAT_API_KEY");
  }
  if (!force && token && Date.now() < tokenExpiresAt - 60000) return token;

  const body = "scope=" + encodeURIComponent(SCOPE);
  const res = await request(OAUTH_URL, {
    headers: {
      Authorization: "Basic " + authKey,
      RqUID: crypto.randomUUID(),
      "Content-Type": "application/x-www-form-urlencoded",
      "Content-Length": Buffer.byteLength(body),
    },
    body,
  });

  if (res.status !== 200) {
    throw new Error("GigaChat OAuth HTTP " + res.status + ": " + res.body.slice(0, 200));
  }

  const data = JSON.parse(res.body);
  if (!data.access_token) throw new Error("GigaChat OAuth вернул ответ без access_token");

  token = data.access_token;
  tokenExpiresAt = Number(data.expires_at) || Date.now() + 25 * 60 * 1000;
  return token;
}

function chat(accessToken, payload) {
  const body = JSON.stringify(payload);
  return request(API_URL, {
    headers: {
      Authorization: "Bearer " + accessToken,
      "Content-Type": "application/json",
      "X-Request-ID": crypto.randomUUID(),
      "Content-Length": Buffer.byteLength(body),
    },
    body,
  });
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
  });
  res.end(body);
}

function applyCors(req, res) {
  res.setHeader("Access-Control-Allow-Origin", req.headers.origin || "*");
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "authorization, content-type, x-request-id");
  res.setHeader("Access-Control-Allow-Private-Network", "true");
  res.setHeader("Access-Control-Max-Age", "600");
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > 2 * 1024 * 1024) {
        reject(new Error("слишком большой запрос"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

const server = http.createServer(async (req, res) => {
  applyCors(req, res);
  const started = Date.now();
  const route = (req.url || "").split("?")[0];
  let status = 500;

  try {
    if (req.method === "OPTIONS") {
      status = 204;
      res.writeHead(204);
      res.end();
      return;
    }

    if (req.method === "GET" && route === "/health") {
      status = 200;
      sendJson(res, 200, { ok: true, key: Boolean(authKey), model: "GigaChat-2" });
      return;
    }

    if (req.method === "GET" && route === "/v1/models") {
      status = 200;
      sendJson(res, 200, {
        object: "list",
        data: [
          { id: "GigaChat-2", object: "model", owned_by: "sberbank" },
          { id: "GigaChat-2-Pro", object: "model", owned_by: "sberbank" },
        ],
      });
      return;
    }

    if (req.method === "POST" && route === "/v1/chat/completions") {
      const raw = await readBody(req);
      let payload;
      try {
        payload = JSON.parse(raw);
      } catch (error) {
        status = 400;
        sendJson(res, 400, { error: { message: "Некорректный JSON в запросе" } });
        return;
      }

      payload.model = payload.model || "GigaChat-2";
      if (typeof payload.max_tokens === "number") {
        payload.max_tokens = Math.min(payload.max_tokens, MAX_TOKENS_LIMIT);
      }

      let accessToken = await getToken(false);
      let answer = await chat(accessToken, payload);
      if (answer.status === 401) {
        accessToken = await getToken(true);
        answer = await chat(accessToken, payload);
      }

      status = answer.status;
      res.writeHead(answer.status, { "Content-Type": "application/json; charset=utf-8" });
      res.end(answer.body);
      return;
    }

    status = 404;
    sendJson(res, 404, { error: { message: "Не найдено: " + route } });
  } catch (error) {
    status = 502;
    sendJson(res, 502, { error: { message: error.message || String(error) } });
  } finally {
    console.log(
      new Date().toLocaleTimeString("ru-RU") +
        " " +
        req.method +
        " " +
        route +
        " -> " +
        status +
        " (" +
        (Date.now() - started) +
        " мс)",
    );
  }
});

server.on("error", (error) => {
  if (error.code === "EADDRINUSE") {
    console.error("Порт " + PORT + " занят. Закройте другой экземпляр прокси или задайте PORT.");
  } else {
    console.error("Ошибка сервера:", error.message);
  }
  process.exit(1);
});

server.listen(PORT, HOST, () => {
  console.log("GigaChat-прокси: http://" + HOST + ":" + PORT + "/v1");
  console.log("Модель: GigaChat-2, область: " + SCOPE);
  console.log(authKey ? "Ключ загружен." : "Ключ не найден: положите его в proxy/gigachat-key.txt");
  console.log("Окно можно свернуть. Закройте его, чтобы остановить прокси.");
});
