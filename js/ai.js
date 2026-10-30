(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const STORAGE_KEY = "forma.llm.v1";
  const DEFAULTS = {
    baseUrl: "https://api.openai.com/v1",
    model: "gpt-4o-mini",
    apiKey: "",
  };
  const MAX_TOKENS = 4096;
  const REPAIR_TOKENS = 5000;
  const TIMEOUT_MS = 90000;

  const FORMAT_NAMES = { docx: "Word-документ", pdf: "PDF-записка", pptx: "презентация" };

  const SYSTEM_PROMPT = [
    "Ты редактор деловых документов. Пишешь по-русски ясно, конкретно и без канцелярита.",
    "Ты собираешь структуру документа: заголовок, подзаголовок и 3–6 разделов.",
    "В каждом разделе допустимы 1–3 абзаца, 0–6 пунктов списка и необязательная таблица.",
    "Заголовок раздела — до 8 слов. Пункт списка — до 20 слов. Абзац — до 700 знаков.",
    "Таблица: 2–4 столбца, до 6 строк, короткие ячейки. Таблицу добавляй только там, где она действительно помогает.",
    "Не используй markdown, звёздочки, решётки, эмодзи и нумерацию внутри текста.",
    "Не выдумывай статистику, даты и цитаты. Если данных нет — пиши качественно, но без ложных фактов.",
    "Верни только корректный JSON без пояснений и без обёртки из ```.",
  ].join(" ");

  function schemaHint(format) {
    return [
      'Формат: {"title": "…", "subtitle": "…", "author": "", "sections": [',
      '{"heading": "…", "paragraphs": ["…"], "bullets": ["…"], "table": {"headers": ["…"], "rows": [["…"]]}}',
      "]}",
      "Поле table необязательно: если таблица не нужна, не добавляй его вовсе.",
      "Формат документа: " + FORMAT_NAMES[format] + ".",
    ].join("\n");
  }

  function cleanKey(value) {
    return String(value || "").trim().replace(/^Bearer\s+/i, "");
  }

  function load() {
    let stored = {};
    try {
      stored = JSON.parse(global.localStorage.getItem(STORAGE_KEY) || "{}") || {};
    } catch (error) {
      stored = {};
    }
    return {
      baseUrl: String(stored.baseUrl || DEFAULTS.baseUrl).trim().replace(/\/+$/, ""),
      model: String(stored.model || DEFAULTS.model).trim(),
      apiKey: cleanKey(stored.apiKey),
    };
  }

  function save(settings) {
    const payload = {
      baseUrl: String(settings.baseUrl || DEFAULTS.baseUrl).trim().replace(/\/+$/, ""),
      model: String(settings.model || DEFAULTS.model).trim() || DEFAULTS.model,
      apiKey: cleanKey(settings.apiKey),
    };
    try {
      global.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (error) {
      /* хранилище недоступно — настройки живут до перезагрузки страницы */
    }
    return payload;
  }

  function clear() {
    try {
      global.localStorage.removeItem(STORAGE_KEY);
    } catch (error) {
      /* хранилище недоступно */
    }
  }

  function stripFences(text) {
    let value = String(text || "").trim();
    value = value.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim();
    if (value.startsWith("{")) return value;
    const start = value.indexOf("{");
    const end = value.lastIndexOf("}");
    if (start >= 0 && end > start) return value.slice(start, end + 1);
    return value;
  }

  function parseContent(content) {
    const text = String(content || "").trim();
    if (!text) return null;
    try {
      const parsed = JSON.parse(stripFences(text));
      return parsed && typeof parsed === "object" ? parsed : null;
    } catch (error) {
      return null;
    }
  }

  function textFromContent(content) {
    if (typeof content === "string") return content;
    if (Array.isArray(content)) {
      return content
        .map(function (part) {
          if (typeof part === "string") return part;
          if (part && typeof part.text === "string") return part.text;
          return "";
        })
        .join("");
    }
    return "";
  }

  function friendlyError(status, detail) {
    if (status === 401) return "Ключ не подошёл. Проверьте API-ключ в настройках.";
    if (status === 403) {
      if (/country|region|territory/i.test(detail)) {
        return "API недоступен из вашего региона. Укажите прокси или другой OpenAI-совместимый адрес в настройках.";
      }
      return "Доступ запрещён. Проверьте ключ и права доступа к модели.";
    }
    if (status === 404) return "Модель или адрес API не найдены. Проверьте настройки.";
    if (status === 408 || status === 504) return "Сервис модели не ответил вовремя. Попробуйте ещё раз.";
    if (status === 413) return "Запрос оказался слишком большим для этой модели.";
    if (status === 429) return "API просит подождать: слишком много запросов или закончилась квота.";
    if (status >= 500) return "Сервис модели временно недоступен.";
    return detail || "Модель вернула ошибку.";
  }

  function buildMessages(topic, format) {
    return [
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content: "Тема: «" + topic + "»\n\nСобери документ на эту тему.\n\n" + schemaHint(format),
      },
    ];
  }

  function repairPrompt(answer) {
    const draft = String(answer.content || "").trim().slice(0, 2000);
    const lines = [
      "Предыдущий ответ не удалось разобрать как JSON" +
        (answer.finish === "length" ? " — он оборван по лимиту длины" : "") +
        ".",
    ];
    lines.push(draft ? "Вот он:\n\n" + draft : "Он пришёл пустым.");
    lines.push("Верни только исправленный JSON по схеме из первого сообщения, без пояснений и без ```.");
    return lines.join("\n\n");
  }

  async function post(settings, payload, signal) {
    const response = await fetch(settings.baseUrl + "/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + settings.apiKey,
      },
      body: JSON.stringify(payload),
      signal: signal,
    });

    if (!response.ok) {
      const raw = await response.text().catch(function () {
        return "";
      });
      let detail = "";
      try {
        const parsed = JSON.parse(raw);
        detail = parsed.error && parsed.error.message ? parsed.error.message : "";
      } catch (error) {
        detail = raw.slice(0, 200);
      }
      const failure = new Error(friendlyError(response.status, detail));
      failure.status = response.status;
      failure.detail = detail;
      throw failure;
    }

    let data;
    try {
      data = await response.json();
    } catch (error) {
      throw new Error("API вернул ответ не в формате JSON.");
    }

    const choice = data && data.choices && data.choices[0];
    const message = choice && choice.message;
    return {
      content: textFromContent(message && message.content),
      finish: choice && choice.finish_reason,
    };
  }

  async function complete(settings, messages, options, signal) {
    const payload = {
      model: settings.model,
      temperature: options.temperature,
      max_tokens: options.maxTokens,
      messages: messages,
    };
    if (options.json) payload.response_format = { type: "json_object" };

    try {
      return await post(settings, payload, signal);
    } catch (error) {
      if ((error.status === 400 || error.status === 422) && /max_tokens/i.test(error.detail || "")) {
        const fallback = Object.assign({}, payload);
        delete fallback.max_tokens;
        return post(settings, fallback, signal);
      }
      throw error;
    }
  }

  async function requestCompletion(settings, topic, format) {
    const controller = new AbortController();
    const timer = setTimeout(function () {
      controller.abort();
    }, TIMEOUT_MS);
    const messages = buildMessages(topic, format);

    try {
      let answer;
      let jsonMode = true;
      try {
        answer = await complete(
          settings,
          messages,
          { json: true, temperature: 0.7, maxTokens: MAX_TOKENS },
          controller.signal
        );
      } catch (error) {
        if (error.status !== 400 && error.status !== 422) throw error;
        jsonMode = false;
        answer = await complete(
          settings,
          messages,
          { json: false, temperature: 0.7, maxTokens: MAX_TOKENS },
          controller.signal
        );
      }

      let parsed = parseContent(answer.content);

      if (!parsed) {
        const repair = await complete(
          settings,
          messages.concat([{ role: "user", content: repairPrompt(answer) }]),
          { json: jsonMode, temperature: 0.2, maxTokens: REPAIR_TOKENS },
          controller.signal
        );
        parsed = parseContent(repair.content);
      }

      if (!parsed) {
        const error = new Error("Модель вернула ответ не в формате JSON. Попробуйте сгенерировать ещё раз.");
        error.code = "bad-json";
        throw error;
      }

      return parsed;
    } finally {
      clearTimeout(timer);
    }
  }

  async function generateContent(topic, format) {
    const settings = load();
    const demoDoc = Forma.normalize(Forma.demo(topic, format), format);

    if (!settings.apiKey) {
      return {
        doc: demoDoc,
        provider: "демо-режим",
        model: "",
        note: "Ключ API не задан — показан демонстрационный текст. Добавьте ключ в настройках, чтобы получить материал по вашей теме.",
      };
    }

    try {
      const raw = await requestCompletion(settings, topic, format);
      const doc = Forma.normalize(raw, format);
      if (raw.author) doc.author = Forma.clean(raw.author, 80);
      if (raw.date) doc.date = Forma.clean(raw.date, 40);
      return {
        doc: doc,
        provider: "ИИ",
        model: settings.model,
        note: "",
      };
    } catch (error) {
      let reason;
      if (error && error.name === "AbortError") {
        reason = "Модель не ответила за 90 секунд.";
      } else if (error && error.name === "TypeError") {
        reason = "Не удалось связаться с API — проверьте адрес и доступ к сети.";
      } else {
        reason = (error && error.message) || "Не удалось получить ответ модели.";
      }
      return {
        doc: demoDoc,
        provider: "демо-режим",
        model: "",
        note: reason + " Показан демонстрационный документ; ключ и адрес API можно проверить в настройках.",
      };
    }
  }

  Object.assign(Forma, {
    ai: {
      DEFAULTS: DEFAULTS,
      load: load,
      save: save,
      clear: clear,
      generateContent: generateContent,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);
