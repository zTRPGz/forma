"use strict";

const giga = require("../server/gigachat.cjs");

module.exports = function handler(req, res) {
  const headers = giga.corsHeaders(req.headers.origin);
  Object.keys(headers).forEach(function (name) {
    res.setHeader(name, headers[name]);
  });
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.status(200).json({
    ok: true,
    model: giga.MODEL,
    key: Boolean(String(process.env.GIGACHAT_API_KEY || "").trim()),
  });
};
