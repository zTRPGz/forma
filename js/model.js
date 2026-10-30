(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});

  const LIMITS = Object.freeze({
    title: 120,
    heading: 90,
    text: 700,
    bullet: 220,
    cell: 120,
    sections: 7,
    paragraphs: 3,
    bullets: 6,
    rows: 6,
    cols: 4,
  });

  const MONTHS = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
  ];

  const CONTROL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g;
  const SPACES = /\s+/g;
  const LEADING_MARK = /^\s*(?:\{\s*\}|\(\s*\)|\[\s*\])[\s:;.,—–-]*/;
  const LEADING_BULLET = /^\s*[-–—•]\s+/;
  const TEXT_KEYS = [
    "text", "point", "value", "content", "title", "header",
    "label", "name", "description", "desc",
  ];

  function textOf(value) {
    if (value == null) return "";
    if (typeof value === "string") return value;
    if (typeof value === "number" || typeof value === "boolean") return String(value);
    if (Array.isArray(value)) {
      return value.map(textOf).filter(Boolean).join(" ");
    }
    if (typeof value === "object") {
      for (let i = 0; i < TEXT_KEYS.length; i += 1) {
        const part = textOf(value[TEXT_KEYS[i]]);
        if (part) return part;
      }
      const keys = Object.keys(value);
      for (let i = 0; i < keys.length; i += 1) {
        const part = textOf(value[keys[i]]);
        if (part) return part;
      }
    }
    return "";
  }

  function clean(value, limit) {
    let text = String(textOf(value)).replace(CONTROL, " ").replace(SPACES, " ").trim();
    let guard = 0;
    while (guard < 4 && (LEADING_MARK.test(text) || LEADING_BULLET.test(text))) {
      text = text.replace(LEADING_MARK, "").replace(LEADING_BULLET, "").trim();
      guard += 1;
    }
    if (text.length > limit) {
      let cut = text.slice(0, limit);
      const space = cut.lastIndexOf(" ");
      if (space > 0) cut = cut.slice(0, space);
      text = cut.replace(/[ ,.;:\-—]+$/, "") + "…";
    }
    return text;
  }

  function cleanList(values, limit, count) {
    if (!Array.isArray(values)) return [];
    return values
      .slice(0, count)
      .map(function (value) {
        return clean(value, limit);
      })
      .filter(Boolean);
  }

  function todayRu() {
    const now = new Date();
    return now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear();
  }

  function tidy(value) {
    return String(value || "").replace(/\s*:$/, "").trim();
  }

  function normalize(raw, fmt) {
    const source = raw && typeof raw === "object" ? raw : {};
    const sections = [];
    const rawSections = Array.isArray(source.sections) ? source.sections.slice(0, LIMITS.sections) : [];

    rawSections.forEach(function (item) {
      if (!item || typeof item !== "object") return;
      const heading = tidy(clean(item.heading, LIMITS.heading));
      if (!heading) return;

      let table = null;
      if (item.table && Array.isArray(item.table.headers)) {
        const headers = cleanList(item.table.headers, 40, LIMITS.cols);
        if (headers.length) {
          const rows = [];
          const rawRows = Array.isArray(item.table.rows) ? item.table.rows.slice(0, LIMITS.rows) : [];
          rawRows.forEach(function (row) {
            if (!Array.isArray(row)) return;
            const cells = row.slice(0, headers.length).map(function (cell) {
              return clean(cell, LIMITS.cell);
            });
            while (cells.length < headers.length) cells.push("");
            if (cells.some(Boolean)) rows.push(cells);
          });
          if (rows.length) table = { headers: headers, rows: rows };
        }
      }

      const section = {
        heading: heading,
        paragraphs: cleanList(item.paragraphs, LIMITS.text, LIMITS.paragraphs),
        bullets: cleanList(item.bullets, LIMITS.bullet, LIMITS.bullets),
        table: table,
      };
      if (section.paragraphs.length || section.bullets.length || section.table) {
        sections.push(section);
      }
    });

    if (!sections.length) {
      sections.push({
        heading: "Обзор",
        paragraphs: ["Документ пока пуст — попробуйте сформулировать тему подробнее."],
        bullets: [],
        table: null,
      });
    }

    return {
      title: tidy(clean(source.title, LIMITS.title)) || "Без названия",
      subtitle: clean(source.subtitle, 120),
      author: clean(source.author, 80),
      date: clean(source.date, 40) || todayRu(),
      sections: sections,
      format: fmt,
    };
  }

  const SUBTITLES = {
    docx: "Обзор, ключевые идеи и план",
    pdf: "Аналитическая записка",
    pptx: "Ключевые тезисы",
  };

  function demo(topic, fmt) {
    const base = String(topic || "").trim().replace(/[.!?]+$/, "");
    const title = base ? base.charAt(0).toUpperCase() + base.slice(1) : "Новый документ";

    return {
      title: title,
      subtitle: SUBTITLES[fmt] || "",
      author: "",
      date: todayRu(),
      sections: [
        {
          heading: "Контекст",
          paragraphs: [
            "«" + title + "» — тема, которая всё чаще возникает при планировании работы. " +
              "Ниже собраны краткий разбор, ключевые идеи и практический план в одном документе.",
            "Цель — дать целостную картину и вывести читателя к конкретным действиям, " +
              "а не к общим рассуждениям.",
          ],
        },
        {
          heading: "Ключевые идеи",
          bullets: [
            "Начните с ясной цели: что должно измениться после того, как тема «" + title + "» будет проработана.",
            "Разделите задачу на короткие этапы с проверяемым результатом на каждом.",
            "Назначьте ответственного и срок для каждого этапа — без этого план остаётся пожеланием.",
            "Соберите обратную связь до того, как решение станет дорогим в изменении.",
            "Зафиксируйте допущения: они устаревают быстрее всего.",
          ],
        },
        {
          heading: "Три подхода",
          table: {
            headers: ["Подход", "Что даёт", "Ограничение"],
            rows: [
              ["Быстрый старт", "Результат за дни", "Мало глубины"],
              ["Основательная проработка", "Устойчивое решение", "Требует времени"],
              ["Поэтапный подход", "Баланс скорости и качества", "Нужна дисциплина"],
            ],
          },
        },
        {
          heading: "План действий",
          bullets: [
            "Сформулировать цель и критерии успеха.",
            "Собрать исходные данные и ограничения.",
            "Сделать первый ограниченный прототип.",
            "Проверить результат с заинтересованными людьми.",
            "Зафиксировать выводы и решить, что масштабировать.",
          ],
        },
        {
          heading: "Выводы",
          paragraphs: [
            "«" + title + "» проще всего сдвинуть с места небольшим, но завершённым первым шагом. " +
              "Дальше решение уточняется по обратной связи.",
            "Документ стоит воспринимать как рабочую рамку: меняйте формулировки " +
              "и приоритеты под свою ситуацию.",
          ],
        },
      ],
    };
  }

  const TRANSLIT = {
    а: "a", б: "b", в: "v", г: "g", д: "d", е: "e", ё: "e", ж: "zh", з: "z",
    и: "i", й: "y", к: "k", л: "l", м: "m", н: "n", о: "o", п: "p", р: "r",
    с: "s", т: "t", у: "u", ф: "f", х: "h", ц: "ts", ч: "ch", ш: "sh", щ: "sch",
    ъ: "", ы: "y", ь: "", э: "e", ю: "yu", я: "ya",
  };

  function slugify(text) {
    const ascii = Array.from(String(text || "").toLowerCase())
      .map(function (char) {
        return Object.prototype.hasOwnProperty.call(TRANSLIT, char) ? TRANSLIT[char] : char;
      })
      .join("");
    const slug = ascii.replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60).replace(/-+$/, "");
    return slug || "document";
  }

  Object.assign(Forma, {
    LIMITS: LIMITS,
    clean: clean,
    cleanList: cleanList,
    textOf: textOf,
    normalize: normalize,
    demo: demo,
    todayRu: todayRu,
    slugify: slugify,
  });
})(typeof window !== "undefined" ? window : globalThis);
