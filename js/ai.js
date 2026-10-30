(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const REMOTE_API = "https://forma-docs.duckdns.org";
  const REMOTE_HOST = "forma-docs.duckdns.org";
  const MAX_TOKENS = 4096;
  const REPAIR_TOKENS = 5000;
  const TIMEOUT_MS = 90000;

  const FORMAT_NAMES = { docx: "Word-документ", pdf: "PDF-записка", pptx: "презентация" };

  function apiBase() {
    const host = global.location ? global.location.hostname : "";
    if (host === REMOTE_HOST) return "";
    return REMOTE_API;
  }

  const SYSTEM_PROMPT = [
    "Ты опытный редактор деловых документов на русском языке.",
    "Пиши конкретно: действия, сроки, ответственные, критерии и риски. Без вводных фраз вроде «в современном мире», без канцелярита, клише и саморекламы.",
    "Не выдумывай числа, даты, проценты, суммы, исследования и цитаты. Если точных данных нет, формулируй без числа: «сократит время подготовки отчётов», «заметно ускорит согласование». Заглушку «___» ставь только там, где без числа фраза теряет смысл. Никогда не подставляй правдоподобные цифры от себя.",
    "Структура: заголовок документа до 10 слов, подзаголовок до 12 слов, 4–6 разделов с содержательными заголовками до 8 слов.",
    "В разделе — 1–3 абзаца по 400–700 знаков каждый. Абзац — это 3–5 законченных предложений с конкретикой, а не набор коротких фраз.",
    "Список — 3–6 пунктов, каждый пункт — законченная мысль до 20 слов. Таблица только там, где помогает сравнению или плану: 2–4 столбца, до 6 строк.",
    "Каждый раздел раскрывает свою мысль: никаких повторов, пустых обещаний и общих слов.",
    "Оформление: только чистый текст. Без markdown, звёздочек, решёток, эмодзи, нумерации внутри абзацев и подписей вида «Раздел 1».",
    "Верни только корректный JSON без пояснений и без обёртки из ```. Не обрывай ответ: JSON должен быть полностью закрыт.",
  ].join(" ");

  const FORMAT_BRIEF = {
    docx: "Деловой документ: цель и контекст, затем шаги, роли, сроки и риски. Таблица уместна для плана или сравнения вариантов. Объём — 1800–3000 знаков текста.",
    pdf: "Аналитическая записка: проблема, наблюдения, выводы и рекомендации. Тон сдержанный, выводы проверяемые. Объём — 1500–2500 знаков текста.",
    pptx: "Презентация: 4–6 разделов-слайдов, в каждом 2–4 коротких пункта до 15 слов; абзацы короткие, не больше одного на слайд. Объём — 700–1200 знаков: слайды не должны быть перегружены текстом.",
  };

  function maxTokensFor(format) {
    return format === "pptx" ? 3000 : MAX_TOKENS;
  }

  function schemaHint(format) {
    return [
      'Формат: {"title": "…", "subtitle": "…", "author": "", "sections": [',
      '{"heading": "…", "paragraphs": ["…"], "bullets": ["…"], "table": {"headers": ["…"], "rows": [["…"]]}}',
      "]}",
      "Поле table необязательно: если таблица не нужна, не добавляй его вовсе.",
      "Формат документа: " + FORMAT_NAMES[format] + ". " + (FORMAT_BRIEF[format] || ""),
    ].join("\n");
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

  function relaxJson(text) {
    let inString = false;
    let escaped = false;
    let out = "";
    for (const ch of text) {
      if (inString) {
        if (escaped) {
          out += ch;
          escaped = false;
          continue;
        }
        if (ch === "\\") {
          out += ch;
          escaped = true;
          continue;
        }
        if (ch === '"') {
          out += ch;
          inString = false;
          continue;
        }
        if (ch === "\n") {
          out += "\\n";
          continue;
        }
        if (ch === "\r") continue;
        if (ch === "\t") {
          out += "\\t";
          continue;
        }
        out += ch;
        continue;
      }
      if (ch === '"') inString = true;
      out += ch;
    }
    return out.replace(/,\s*([}\]])/g, "$1");
  }

  function parseContent(content) {
    const text = String(content || "").trim();
    if (!text) return null;
    const cleaned = stripFences(text);
    try {
      const parsed = JSON.parse(cleaned);
      return parsed && typeof parsed === "object" ? parsed : null;
    } catch (error) {
      try {
        const parsed = JSON.parse(relaxJson(cleaned));
        return parsed && typeof parsed === "object" ? parsed : null;
      } catch (relaxError) {
        return null;
      }
    }
  }

  function documentWeight(doc) {
    const sections = doc && Array.isArray(doc.sections) ? doc.sections : [];
    let filled = 0;
    let chars = 0;
    sections.forEach(function (section) {
      if (!section || typeof section !== "object") return;
      const paragraphs = Array.isArray(section.paragraphs) ? section.paragraphs : [];
      const bullets = Array.isArray(section.bullets) ? section.bullets : [];
      const hasText =
        paragraphs.some(function (part) {
          return String(part || "").trim().length > 40;
        }) ||
        bullets.length >= 3 ||
        Boolean(section.table);
      if (section.heading && hasText) filled += 1;
      paragraphs.concat(bullets).forEach(function (part) {
        chars += String(part || "").length;
      });
    });
    return { filled: filled, chars: chars };
  }

  function isThin(doc, format) {
    const weight = documentWeight(doc);
    const minChars = format === "pptx" ? 500 : 900;
    return weight.filled < 4 || weight.chars < minChars;
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
    if (status === 404) return "Сервис ИИ не найден по адресу " + (apiBase() || "сайта") + ". Проверьте подключение к сети.";
    if (status === 408 || status === 504) return "ИИ не ответил вовремя. Попробуйте ещё раз.";
    if (status === 413) return "Запрос оказался слишком большим для модели.";
    if (status === 429) return "ИИ просит подождать: слишком много запросов или закончилась квота.";
    if (status >= 500) return "Сервис ИИ временно недоступен.";
    return detail || "ИИ вернул ошибку.";
  }

  function buildMessages(topic, format) {
    const facts = /\d/.test(topic)
      ? ""
      : "\n\nВ теме не указано ни одного числа, поэтому в документе не должно быть процентов, сумм, статистики и сроков «через N дней» — только качественные формулировки.";
    return [
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content:
          "Тема: «" +
          topic +
          "»\n\nСобери документ на эту тему. Читатель — занятый руководитель: ему нужны суть, действия и риски, а не общие рассуждения." +
          facts +
          "\n\n" +
          schemaHint(format),
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

  async function post(payload, signal) {
    const response = await fetch(apiBase() + "/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
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
      throw new Error("Сервис ИИ вернул ответ не в формате JSON.");
    }

    const choice = data && data.choices && data.choices[0];
    const message = choice && choice.message;
    return {
      content: textFromContent(message && message.content),
      finish: choice && choice.finish_reason,
    };
  }

  async function complete(messages, options, signal) {
    const payload = {
      temperature: options.temperature,
      max_tokens: options.maxTokens,
      messages: messages,
    };
    if (options.json) payload.response_format = { type: "json_object" };

    try {
      return await post(payload, signal);
    } catch (error) {
      if ((error.status === 400 || error.status === 422) && /max_tokens/i.test(error.detail || "")) {
        const fallback = Object.assign({}, payload);
        delete fallback.max_tokens;
        return post(fallback, signal);
      }
      throw error;
    }
  }

  async function requestCompletion(topic, format) {
    const controller = new AbortController();
    const timer = setTimeout(function () {
      controller.abort();
    }, TIMEOUT_MS);
    const messages = buildMessages(topic, format);
    const budget = maxTokensFor(format);

    try {
      let answer;
      let jsonMode = true;
      try {
        answer = await complete(messages, { json: true, temperature: 0.6, maxTokens: budget }, controller.signal);
      } catch (error) {
        if (error.status !== 400 && error.status !== 422) throw error;
        jsonMode = false;
        answer = await complete(messages, { json: false, temperature: 0.6, maxTokens: budget }, controller.signal);
      }

      let parsed = parseContent(answer.content);

      if (!parsed) {
        const repair = await complete(
          messages.concat([{ role: "user", content: repairPrompt(answer) }]),
          { json: jsonMode, temperature: 0.2, maxTokens: REPAIR_TOKENS },
          controller.signal
        );
        parsed = parseContent(repair.content);
      }

      if (!parsed) {
        const retryMessages = messages.concat([
          {
            role: "user",
            content:
              "Ответ не разобрался как JSON. Собери документ заново и верни строго один JSON-объект по схеме из первого сообщения. Никакого текста, пояснений и markdown до или после JSON.",
          },
        ]);
        const retry = await complete(
          retryMessages,
          { json: jsonMode, temperature: 0.3, maxTokens: budget },
          controller.signal
        );
        parsed = parseContent(retry.content);

        if (!parsed) {
          const repair = await complete(
            messages.concat([{ role: "user", content: repairPrompt(retry) }]),
            { json: jsonMode, temperature: 0.2, maxTokens: REPAIR_TOKENS },
            controller.signal
          );
          parsed = parseContent(repair.content);
        }
      }

      if (!parsed) {
        const error = new Error("ИИ вернул ответ не в формате JSON. Попробуйте сгенерировать ещё раз.");
        error.code = "bad-json";
        throw error;
      }

      if (isThin(parsed, format)) {
        const expanded = await complete(
          messages.concat([
            {
              role: "user",
              content:
                "Содержания слишком мало: нужно 4–6 разделов с конкретикой — действиями, сроками, ответственными и рисками. В каждом разделе 1–3 абзаца по 400–700 знаков. Расширь документ и верни полный JSON по схеме из первого сообщения, без пояснений.",
            },
          ]),
          { json: jsonMode, temperature: 0.4, maxTokens: budget },
          controller.signal
        );
        const fuller = parseContent(expanded.content);
        if (fuller && !isThin(fuller, format)) parsed = fuller;
      }

      return parsed;
    } finally {
      clearTimeout(timer);
    }
  }

  async function checkHealth() {
    const controller = new AbortController();
    const timer = setTimeout(function () {
      controller.abort();
    }, 6000);

    try {
      const response = await fetch(apiBase() + "/health", {
        method: "GET",
        cache: "no-store",
        signal: controller.signal,
      });
      if (!response.ok) return { ok: false, reason: friendlyError(response.status, "") };
      const data = await response.json();
      if (!data || data.ok !== true) return { ok: false, reason: "Сервис ИИ ответил неожиданно." };
      if (!data.key) return { ok: false, reason: "На сервере не настроен ключ GigaChat." };
      return { ok: true, model: String(data.model || "") };
    } catch (error) {
      return { ok: false, reason: "Не удалось связаться с сервисом ИИ." };
    } finally {
      clearTimeout(timer);
    }
  }

  async function generateContent(topic, format) {
    const demoDoc = Forma.normalize(Forma.demo(topic, format), format);

    try {
      const raw = await requestCompletion(topic, format);
      const doc = Forma.normalize(raw, format);
      if (raw.author) doc.author = Forma.clean(raw.author, 80);
      if (raw.date) doc.date = Forma.clean(raw.date, 40);
      return { doc: doc, provider: "ИИ", model: "GigaChat-2", note: "" };
    } catch (error) {
      let reason;
      if (error && error.name === "AbortError") {
        reason = "ИИ не ответил за 90 секунд.";
      } else if (error && error.name === "TypeError") {
        reason = "Не удалось связаться с сервисом ИИ — проверьте подключение к сети.";
      } else {
        reason = (error && error.message) || "Не удалось получить ответ ИИ.";
      }
      return {
        doc: demoDoc,
        provider: "демо-режим",
        model: "",
        note: reason + " Показан демонстрационный документ — попробуйте сгенерировать ещё раз.",
      };
    }
  }

  Object.assign(Forma, {
    ai: {
      apiBase: apiBase,
      health: checkHealth,
      generateContent: generateContent,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);
