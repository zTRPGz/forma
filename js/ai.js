(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const REMOTE_API = "https://forma-docs.duckdns.org";
  const REMOTE_HOST = "forma-docs.duckdns.org";
  const MAX_TOKENS = 4096;
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
    "Структура: заголовок документа до 10 слов, подзаголовок до 12 слов, содержательные заголовки разделов до 8 слов. Число разделов задано в описании формата ниже.",
    "В разделе — 1–3 абзаца по 400–700 знаков каждый. Абзац — это 3–5 законченных предложений с конкретикой, а не набор коротких фраз.",
    "Список — 3–6 пунктов, каждый пункт — законченная мысль до 20 слов. Таблица только там, где помогает сравнению или плану: 2–4 столбца, до 6 строк.",
    "Перечисления действий, условий и рекомендаций оформляй полем bullets, а не отдельными короткими абзацами: абзац — это связный текст из нескольких предложений.",
    "Каждый раздел раскрывает свою мысль: никаких повторов, пустых обещаний и общих слов.",
    "Оформление: только чистый текст. Без markdown, звёздочек, решёток, эмодзи, нумерации внутри абзацев и подписей вида «Раздел 1».",
    "Верни только корректный JSON без пояснений и без обёртки из ```. Не обрывай ответ: JSON должен быть полностью закрыт.",
  ].join(" ");

  const FORMAT_BRIEF = {
    docx: "Деловой документ из 4–6 разделов: цель и контекст, затем шаги, роли, сроки и риски. Таблица уместна для плана или сравнения вариантов. Объём — 1800–3000 знаков текста.",
    pdf: "Аналитическая записка из 4–6 разделов: проблема, наблюдения, выводы и рекомендации. Тон сдержанный, выводы проверяемые. Объём — 1500–2500 знаков текста.",
    pptx: "Презентация по сюжету: обложка с темой, затем содержательные слайды с пунктами и финальный слайд с темой. В каждом разделе вводный абзац 180–240 знаков и 4–5 пунктов по 10–18 слов: подробно, но читаемо. Объём — 1800–3000 знаков.",
  };

  function maxTokensFor(format, options) {
    if (format !== "pptx") return MAX_TOKENS;
    const slides = options && options.slides ? options.slides : 6;
    return Math.min(MAX_TOKENS, 3000 + Math.max(0, slides - 6) * 160);
  }

  function schemaHint(format, options) {
    let brief = FORMAT_BRIEF[format] || "";
    if (format === "pptx" && options && options.slides) {
      const slides = options.slides;
      const content = Math.max(1, slides - 2);
      brief =
        "Презентация ровно на " +
        slides +
        " слайдов: обложка с темой, затем " +
        content +
        " содержательных слайдов и финальный слайд с темой. Верни ровно " +
        content +
        " разделов — по одному содержательному слайду на раздел. Таблицы не используй. В каждом разделе один вводный абзац 180–240 знаков и 4–5 пунктов по 10–18 слов: каждый пункт раскрывает мысль, а не обозначает её. Объём — " +
        Math.round(content * 450) +
        "–" +
        Math.round(content * 750) +
        " знаков.";
    }
    return [
      'Формат: {"title": "…", "subtitle": "…", "author": "", "sections": [',
      '{"heading": "…", "paragraphs": ["…"], "bullets": ["…"], "table": {"headers": ["…"], "rows": [["…"]]}}',
      "]}",
      "Поле table необязательно: если таблица не нужна, не добавляй его вовсе.",
      "Формат документа: " + FORMAT_NAMES[format] + ". " + brief,
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
      return parsed && typeof parsed === "object" ? pairSplitSections(parsed) : null;
    } catch (error) {
      try {
        const parsed = JSON.parse(relaxJson(cleaned));
        return parsed && typeof parsed === "object" ? pairSplitSections(parsed) : null;
      } catch (relaxError) {
        return null;
      }
    }
  }

  function asText(value) {
    return Forma.textOf ? Forma.textOf(value) : String(value == null ? "" : value);
  }

  function hasBody(section) {
    if (!section || typeof section !== "object") return false;
    const paragraphs = Array.isArray(section.paragraphs) ? section.paragraphs : [];
    const bullets = Array.isArray(section.bullets) ? section.bullets : [];
    return (
      paragraphs.some(function (part) {
        return asText(part).trim();
      }) ||
      bullets.some(function (part) {
        return asText(part).trim();
      }) ||
      Boolean(section.table && Array.isArray(section.table.headers) && section.table.headers.length)
    );
  }

  // Модель иногда разбивает раздел на {"heading":"…"} и заголовок без секции
  // {"paragraphs":[…],"bullets":[…]}; склеиваем их, чтобы раздел не потерялся.
  function pairSplitSections(value) {
    if (!value || !Array.isArray(value.sections)) return value;
    const sections = [];
    let pending = "";

    function flush() {
      if (!pending) return;
      sections.push({ heading: pending, paragraphs: [], bullets: [], table: null });
      pending = "";
    }

    value.sections.forEach(function (item) {
      if (!item || typeof item !== "object") return;
      const heading = String(item.heading || "").trim();
      if (!heading) {
        if (pending && hasBody(item)) {
          sections.push({
            heading: pending,
            paragraphs: Array.isArray(item.paragraphs) ? item.paragraphs : [],
            bullets: Array.isArray(item.bullets) ? item.bullets : [],
            table: item.table || null,
          });
          pending = "";
        }
        return;
      }
      flush();
      if (hasBody(item)) sections.push(item);
      else pending = heading;
    });
    flush();
    value.sections = sections;
    return value;
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
          return asText(part).trim().length > 40;
        }) ||
        bullets.length >= 3 ||
        Boolean(section.table);
      if (section.heading && hasText) filled += 1;
      paragraphs.concat(bullets).forEach(function (part) {
        chars += asText(part).length;
      });
    });
    return { filled: filled, chars: chars };
  }

  function isThin(doc, format, options) {
    const weight = documentWeight(doc);
    const targetSections =
      format === "pptx" && options && options.slides ? Math.max(1, options.slides - 2) : 4;
    const minChars =
      format === "pptx"
        ? options && options.slides
          ? Math.max(600, Math.min(2600, targetSections * 300))
          : 900
        : 900;
    return weight.filled < targetSections || weight.chars < minChars;
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

  function buildMessages(topic, format, options) {
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
          schemaHint(format, options),
      },
    ];
  }

  function expansionPrompt(format, options) {
    if (format === "pptx" && options && options.slides) {
      const content = Math.max(1, options.slides - 2);
      return (
        "Содержания слишком мало: нужно ровно " +
        content +
        " содержательных разделов — по одному слайду на раздел, без таблиц. " +
        "В каждом разделе вводный абзац 180–240 знаков и 4–5 пунктов по 10–18 слов: пункты должны быть развёрнутыми, с действием и результатом, а не короткими ярлыками. " +
        "Расширь презентацию и верни полный JSON по схеме из первого сообщения, без пояснений."
      );
    }
    return "Содержания слишком мало: нужно 4–6 разделов с конкретикой — действиями, сроками, ответственными и рисками. В каждом разделе 1–3 абзаца по 400–700 знаков. Расширь документ и верни полный JSON по схеме из первого сообщения, без пояснений.";
  }

  function cleanSections(doc) {
    if (doc && Array.isArray(doc.sections)) {
      doc.sections = doc.sections.filter(function (item) {
        return item && String(item.heading || "").trim() && hasBody(item);
      });
    }
    return doc;
  }

  function sectionCount(value) {
    if (!value || !Array.isArray(value.sections)) return 0;
    return value.sections.length;
  }

  function usableSections(value) {
    if (!value || !Array.isArray(value.sections)) return [];
    return value.sections.filter(function (item) {
      return item && typeof item === "object";
    });
  }

  async function topUpSections(doc, options, messages, jsonMode, signal) {
    const target = Math.max(1, (options && options.slides ? options.slides : 6) - 2);
    cleanSections(doc);
    if (sectionCount(doc) >= target) return doc;
    if (!Array.isArray(doc.sections)) doc.sections = [];

    for (let attempt = 0; attempt < 2 && sectionCount(doc) < target; attempt += 1) {
      const missing = target - sectionCount(doc);
      const existing = doc.sections
        .map(function (item) {
          return String((item && item.heading) || "").trim();
        })
        .filter(Boolean);

      const prompt = [
        "Для презентации нужно ровно " +
          target +
          " содержательных разделов, сейчас есть " +
          existing.length +
          ".",
        existing.length ? "Уже заняты темы: «" + existing.join("», «") + "»." : "",
        "Добавь ровно " + missing + " новых разделов на ту же тему, не повторяя существующие.",
        'Верни строго один JSON-объект вида {"sections":[{"heading":"…","paragraphs":["…"],"bullets":["…","…"]}]} — только новые разделы, без таблиц. В каждом разделе вводный абзац 180–240 знаков и 4–5 пунктов по 10–18 слов. Без пояснений, markdown и обёртки из ```.',
      ]
        .filter(Boolean)
        .join("\n\n");

      const answer = await complete(
        messages.concat([{ role: "user", content: prompt }]),
        { json: jsonMode, temperature: 0.5, maxTokens: 1800 },
        signal
      );
      const fresh = cleanSections(parseContent(answer.content));
      if (!fresh || !fresh.sections || !fresh.sections.length) continue;

      const seen = new Set(
        existing.map(function (heading) {
          return heading.toLowerCase();
        })
      );
      fresh.sections.forEach(function (item) {
        if (sectionCount(doc) >= target) return;
        const heading = String(item.heading || "").trim();
        if (!heading || seen.has(heading.toLowerCase())) return;
        seen.add(heading.toLowerCase());
        doc.sections.push(item);
      });
    }

    return doc;
  }

  function headingKey(value) {
    return String(value || "")
      .replace(/[^0-9a-zA-Zа-яёА-ЯЁ]+/gi, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  async function askJson(messages, settings, signal) {
    let answer;
    try {
      answer = await complete(messages, Object.assign({ json: true }, settings), signal);
    } catch (error) {
      if (error.status !== 400 && error.status !== 422) throw error;
      answer = await complete(messages, Object.assign({ json: false }, settings), signal);
    }
    return parseContent(answer.content);
  }

  function headingList(value) {
    return usableSections(value)
      .map(function (item) {
        return String(item.heading || "").trim();
      })
      .filter(Boolean);
  }

  async function buildDeck(topic, options, signal) {
    const target = Math.max(1, options.slides - 2);
    const system = { role: "system", content: SYSTEM_PROMPT };
    const lead = "Тема: «" + topic + "»\n\n";

    const outline = await askJson(
      [
        system,
        {
          role: "user",
          content:
            lead +
            "Нужен план презентации ровно на " +
            options.slides +
            " слайдов: обложка с темой, " +
            target +
            " содержательных слайдов и финальный слайд с темой.\n" +
            'Верни только JSON вида {"title":"…","subtitle":"…","sections":[{"heading":"…"}]} — ровно ' +
            target +
            " объектов в sections, в каждом только heading. Без bullets, paragraphs и table. Заголовки до 8 слов, разные аспекты темы, без повторов, в логическом порядке. Без пояснений и markdown.",
        },
      ],
      { temperature: 0.5, maxTokens: 1200 },
      signal
    );

    const headings = headingList(outline).slice(0, target);

    if (headings.length < target) {
      const missing = target - headings.length;
      const more = await askJson(
        [
          system,
          {
            role: "user",
            content:
              lead +
              (headings.length
                ? "Для плана презентации уже есть разделы: «" + headings.join("», «") + "». "
                : "") +
              "Добавь ровно " +
              missing +
              " новых заголовков до 8 слов на ту же тему, не повторяя существующие.\n" +
              'Верни только JSON {"sections":[{"heading":"…"}]} с новыми заголовками. Без пояснений и markdown.',
          },
        ],
        { temperature: 0.6, maxTokens: 700 },
        signal
      );
      const seen = new Set(headings.map(headingKey));
      headingList(more).forEach(function (heading) {
        if (headings.length >= target) return;
        const key = headingKey(heading);
        if (!key || seen.has(key)) return;
        seen.add(key);
        headings.push(heading);
      });
    }

    if (!headings.length) return null;

    const filled = new Map();

    function storeItem(key, item) {
      if (!key || !item || filled.has(key)) return;
      const paragraphs = (Array.isArray(item.paragraphs) ? item.paragraphs : [])
        .map(function (part) {
          return String(part || "").trim();
        })
        .filter(Boolean)
        .slice(0, 2);
      const bullets = (Array.isArray(item.bullets) ? item.bullets : [])
        .map(function (part) {
          return String(part || "").trim();
        })
        .filter(Boolean)
        .slice(0, 5);
      if (!paragraphs.length && !bullets.length) return;
      filled.set(key, { paragraphs: paragraphs, bullets: bullets });
    }

    async function fillChunk(chunk) {
      const response = await askJson(
        [
          system,
          {
            role: "user",
            content:
              lead +
              "Разделы презентации: «" +
              chunk.join("», «") +
              "».\nДля каждого заголовка напиши вводный абзац 180–240 знаков и 4–5 конкретных пунктов по 10–18 слов; заголовки оставь дословно. Без таблиц.\n" +
              'Верни только JSON {"sections":[{"heading":"…","paragraphs":["…"],"bullets":["…","…"]}]} — по одному объекту на каждый заголовок, в том же порядке. Без пояснений и markdown.',
          },
        ],
        { temperature: 0.6, maxTokens: 2200 },
        signal
      );
      const items = usableSections(response).filter(function (item) {
        return (
          (Array.isArray(item.bullets) && item.bullets.length) ||
          (Array.isArray(item.paragraphs) && item.paragraphs.length)
        );
      });

      const paired = Math.min(items.length, chunk.length);
      for (let index = 0; index < paired; index += 1) {
        storeItem(headingKey(chunk[index]), items[index]);
      }

      items.slice(paired).forEach(function (item) {
        storeItem(headingKey(item.heading), item);
      });
    }

    const chunks = [];
    for (let start = 0; start < headings.length; start += 3) {
      chunks.push(headings.slice(start, start + 3));
    }
    for (let index = 0; index < chunks.length; index += 1) {
      if (filled.size >= headings.length) break;
      await fillChunk(chunks[index]).catch(function () {
        return null;
      });
    }

    for (let round = 0; round < 2; round += 1) {
      const missing = headings.filter(function (heading) {
        return !filled.has(headingKey(heading));
      });
      if (!missing.length) break;
      const step = round === 0 ? 2 : 1;
      for (let start = 0; start < missing.length; start += step) {
        await fillChunk(missing.slice(start, start + step)).catch(function () {
          return null;
        });
      }
    }

    const sections = headings
      .filter(function (heading) {
        return filled.has(headingKey(heading));
      })
      .map(function (heading) {
        const content = filled.get(headingKey(heading));
        return {
          heading: heading,
          paragraphs: content.paragraphs,
          bullets: content.bullets,
          table: null,
        };
      });

    if (!sections.length) return null;

    return {
      title: String((outline && outline.title) || "").trim(),
      subtitle: String((outline && outline.subtitle) || "").trim(),
      author: "",
      date: "",
      sections: sections,
    };
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

  async function requestCompletion(topic, format, options) {
    const controller = new AbortController();
    const timer = setTimeout(function () {
      controller.abort();
    }, TIMEOUT_MS);
    const messages = buildMessages(topic, format, options);
    const budget = maxTokensFor(format, options);

    try {
      if (format === "pptx" && options && options.slides && options.slides >= 12) {
        const deck = await buildDeck(topic, options, controller.signal).catch(function () {
          return null;
        });
        if (deck) {
          const target = options.slides - 2;
          cleanSections(deck);
          if (sectionCount(deck) >= target) return deck;
          const topped = await topUpSections(deck, options, messages, true, controller.signal).catch(
            function () {
              return deck;
            }
          );
          if (sectionCount(topped) >= target) return topped;
        }
      }

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
          { json: jsonMode, temperature: 0.2, maxTokens: MAX_TOKENS },
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
            { json: jsonMode, temperature: 0.2, maxTokens: MAX_TOKENS },
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

      cleanSections(parsed);

      if (isThin(parsed, format, options)) {
        const expanded = await complete(
          messages.concat([{ role: "user", content: expansionPrompt(format, options) }]),
          { json: jsonMode, temperature: 0.4, maxTokens: budget },
          controller.signal
        );
        const fuller = cleanSections(parseContent(expanded.content));
        if (
          fuller &&
          fuller.sections &&
          (sectionCount(fuller) > sectionCount(parsed) ||
            documentWeight(fuller).chars > documentWeight(parsed).chars)
        ) {
          parsed = fuller;
        }
      }

      if (format === "pptx" && options && options.slides) {
        parsed = await topUpSections(parsed, options, messages, jsonMode, controller.signal);
        if (sectionCount(parsed) < options.slides - 2) {
          const error = new Error("ИИ вернул меньше разделов, чем нужно для выбранного числа слайдов.");
          error.code = "short-deck";
          throw error;
        }
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
      if (!data.key) return { ok: false, reason: "На сервере не настроен ключ ИИ." };
      return { ok: true, model: String(data.model || "") };
    } catch (error) {
      return { ok: false, reason: "Не удалось связаться с сервисом ИИ." };
    } finally {
      clearTimeout(timer);
    }
  }

  async function generateContent(topic, format, options) {
    const demoDoc = Forma.normalize(Forma.demo(topic, format, options), format, options);

    try {
      const raw = await requestCompletion(topic, format, options);
      const doc = Forma.normalize(raw, format, options);
      if (raw.author) doc.author = Forma.clean(raw.author, 80);
      if (raw.date) doc.date = Forma.clean(raw.date, 40);
      return { doc: doc, provider: "ИИ", model: "", note: "" };
    } catch (error) {
      let reason;
      if (error && error.name === "AbortError") {
        reason = "ИИ не ответил за 90 секунд.";
      } else if (error && error.name === "TypeError") {
        reason = "Не удалось связаться с сервисом ИИ — проверьте подключение к сети.";
      } else {
        reason = (error && error.message) || "Не удалось получить ответ ИИ.";
      }
      const retry = /попробуйте/i.test(reason) ? "" : " Попробуйте сгенерировать ещё раз.";
      return {
        doc: demoDoc,
        provider: "демо-режим",
        model: "",
        note: reason + " Показан демонстрационный документ." + retry,
      };
    }
  }

  Object.assign(Forma, {
    ai: {
      health: checkHealth,
      generateContent: generateContent,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);
