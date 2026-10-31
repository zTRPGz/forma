"use strict";

const giga = require("../server/gigachat.cjs");

module.exports = async function handler(req, res) {
  const headers = giga.corsHeaders(req.headers.origin);
  Object.keys(headers).forEach(function (name) {
    res.setHeader(name, headers[name]);
  });
  res.setHeader("Content-Type", "application/json; charset=utf-8");

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }
  if (req.method !== "POST") {
    res.status(405).json({ error: { message: "Нужен метод POST." } });
    return;
  }

  let input;
  try {
    input = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  } catch (error) {
    res.status(400).json({ error: { message: "Некорректный JSON в запросе." } });
    return;
  }

  const prepared = giga.normalizePayload(input);
  if (prepared.error) {
    res.status(400).json({ error: { message: prepared.error } });
    return;
  }

  try {
    const answer = await giga.callChat(prepared.payload);
    res.status(answer.status).send(answer.body);
  } catch (error) {
    res.status(502).json({ error: { message: error.message || "Сервис ИИ недоступен." } });
  }
};
