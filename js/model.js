(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});

  const LIMITS = Object.freeze({
    title: 120,
    heading: 90,
    text: 700,
    bullet: 220,
    cell: 120,
    sections: 20,
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
      let cut = text.slice(0, Math.max(1, limit - 1));
      const space = cut.lastIndexOf(" ");
      if (space > 0) cut = cut.slice(0, space);
      text = cut.replace(/[ ,.;:\-—]+$/, "") + "…";
    }
    return text;
  }

  function cleanList(values, limit, count) {
    if (!Array.isArray(values)) return [];
    return values
      .map(function (value) {
        return clean(value, limit);
      })
      .filter(Boolean)
      .slice(0, count);
  }

  function todayRu() {
    const now = new Date();
    return now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear();
  }

  function tidy(value) {
    return String(value || "").replace(/\s*:$/, "").trim();
  }

  function titleFromTopic(topic) {
    const base = tidy(clean(topic, LIMITS.title)).replace(/[.!?]+$/, "");
    return base ? base.charAt(0).toUpperCase() + base.slice(1) : "";
  }

  function normalize(raw, fmt, options) {
    const source = raw && typeof raw === "object" ? raw : {};
    const settings = options && typeof options === "object" ? options : {};
    const slides =
      fmt === "pptx" && Number.isFinite(settings.slides) && settings.slides > 1
        ? Math.min(20, Math.floor(settings.slides))
        : 0;
    const allowTable = !(fmt === "pptx" && slides);
    const sections = [];
    const maxSections = slides ? Math.max(1, slides - 2) : LIMITS.sections;
    const rawSections = Array.isArray(source.sections) ? source.sections : [];

    for (const item of rawSections) {
      if (sections.length >= maxSections) break;
      if (!item || typeof item !== "object") continue;
      const heading = tidy(clean(item.heading, LIMITS.heading));
      if (!heading) continue;

      let table = null;
      if (allowTable && item.table && Array.isArray(item.table.headers)) {
        const headers = cleanList(item.table.headers, 40, LIMITS.cols);
        if (headers.length) {
          const rows = [];
          const rawRows = Array.isArray(item.table.rows) ? item.table.rows : [];
          for (const row of rawRows) {
            if (rows.length >= LIMITS.rows) break;
            if (!Array.isArray(row)) continue;
            const cells = row.slice(0, headers.length).map(function (cell) {
              return clean(cell, LIMITS.cell);
            });
            while (cells.length < headers.length) cells.push("");
            if (cells.some(Boolean)) rows.push(cells);
          }
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
    }

    if (!sections.length) {
      sections.push({
        heading: "Обзор",
        paragraphs: ["Документ пока пуст — попробуйте сформулировать тему подробнее."],
        bullets: [],
        table: null,
      });
    }

    return {
      // ИИ иногда присылает только разделы: без заголовка из темы первая и
      // последняя страницы выходят «Без названия».
      title: tidy(clean(source.title, LIMITS.title)) ||
        titleFromTopic(settings.topic) ||
        "Без названия",
      subtitle: clean(source.subtitle, 120),
      author: clean(source.author, 80),
      date: clean(source.date, 40) || todayRu(),
      sections: sections,
      format: fmt,
      slides: slides ? Math.min(slides, sections.length + 2) : null,
      theme: settings.theme || "paper",
    };
  }

  const SUBTITLES = {
    docx: "Обзор, ключевые идеи и план",
    pdf: "Аналитическая записка",
    pptx: "Ключевые тезисы",
  };

  function demo(topic, fmt, options) {
    const title = titleFromTopic(topic) || "Новый документ";
    const tableSection =
      fmt === "pptx"
        ? {
            heading: "Три подхода",
            bullets: [
              "Быстрый старт — результат за дни, но меньше глубины.",
              "Основательная проработка — устойчивое решение ценой времени.",
              "Поэтапный подход — баланс скорости и качества при дисциплине.",
            ],
          }
        : {
            heading: "Три подхода",
            table: {
              headers: ["Подход", "Что даёт", "Ограничение"],
              rows: [
                ["Быстрый старт", "Результат за дни", "Мало глубины"],
                ["Основательная проработка", "Устойчивое решение", "Требует времени"],
                ["Поэтапный подход", "Баланс скорости и качества", "Нужна дисциплина"],
              ],
            },
          };

    const sections = [
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
      tableSection,
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
      {
        heading: "Риски",
        bullets: [
          "Расплывчатая цель растягивает сроки и обесценивает результат.",
          "Решения без ответственного остаются намерениями.",
          "Обратная связь, собранная в конце, уже не спасает решение.",
        ],
      },
      {
        heading: "Что измерить",
        bullets: [
          "Сколько этапов закрыто в срок и с ожидаемым качеством.",
          "Где возникли задержки и по какой причине.",
          "Какие допущения не подтвердились на практике.",
        ],
      },
      {
        heading: "Роли и сроки",
        bullets: [
          "Владелец темы отвечает за результат и приоритеты.",
          "Исполнители ведут этапы и сообщают о блокерах заранее.",
          "Руководитель снимает спорные решения и выделяет ресурсы.",
        ],
      },
      {
        heading: "Следующие шаги",
        bullets: [
          "Согласовать цель и критерии успеха с командой.",
          "Запустить первый ограниченный этап.",
          "Назначить дату разбора результатов и обновления плана.",
        ],
      },
      {
        heading: "Ресурсы и бюджет",
        bullets: [
          "Определите, какие люди и время нужны на каждом этапе.",
          "Отделите обязательные расходы от желательных.",
          "Заложите резерв на непредвиденные работы.",
        ],
      },
      {
        heading: "Коммуникация",
        bullets: [
          "Договоритесь, кто и как сообщает о статусе и решениях.",
          "Сведите отчёты к одному формату и одному сроку.",
          "Отмечайте не только результаты, но и снятые блокеры.",
        ],
      },
      {
        heading: "Контроль качества",
        bullets: [
          "Заранее опишите, что считается готовым результатом.",
          "Проверяйте работу на коротких этапах, а не в конце.",
          "Фиксируйте замечания и проверяйте, что они закрыты.",
        ],
      },
      {
        heading: "Обучение команды",
        bullets: [
          "Объясните, зачем меняется процесс и что это даёт каждому.",
          "Дайте короткую инструкцию и пример готовой работы.",
          "Назначьте наставника на первые недели.",
        ],
      },
      {
        heading: "Метрики",
        bullets: [
          "Выберите два-три показателя, а не длинный список.",
          "Сравнивайте результат с исходной точкой, а не с ожиданиями.",
          "Пересматривайте метрики, если меняется цель.",
        ],
      },
      {
        heading: "Ограничения",
        bullets: [
          "Учитывайте доступное время и загрузку ключевых людей.",
          "Проверьте, какие решения требуют согласования.",
          "Не начинайте этап, пока не сняты критические зависимости.",
        ],
      },
      {
        heading: "Первые две недели",
        bullets: [
          "Соберите рабочую группу и распределите роли.",
          "Согласуйте план и критерии успеха первого этапа.",
          "Запустите пилот и назначьте дату разбора.",
        ],
      },
      {
        heading: "Масштабирование",
        bullets: [
          "Расширяйте только то, что подтвердилось на пилоте.",
          "Описывайте процесс так, чтобы его мог повторить другой.",
          "Следите, чтобы рост не разрушал качество и сроки.",
        ],
      },
      {
        heading: "Частые ошибки",
        bullets: [
          "Слишком широкий старт без проверяемого результата.",
          "Отсутствие владельца и даты у ключевых шагов.",
          "Отчётность ради отчётности вместо решения проблем.",
        ],
      },
      {
        heading: "Чек-лист запуска",
        bullets: [
          "Цель, роли и сроки зафиксированы письменно.",
          "Первый этап ограничен по объёму и длительности.",
          "Есть дата промежуточной проверки и разбора итогов.",
        ],
      },
    ];

    if (fmt === "pptx") {
      const slides =
        options && Number.isFinite(options.slides)
          ? Math.min(20, Math.max(2, Math.floor(options.slides)))
          : 6;
      sections.length = Math.min(sections.length, Math.max(1, slides - 2));
    } else {
      sections.length = Math.min(sections.length, 6);
    }

    return {
      title: title,
      subtitle: SUBTITLES[fmt] || "",
      author: "",
      date: todayRu(),
      sections: sections,
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
    clean: clean,
    textOf: textOf,
    normalize: normalize,
    demo: demo,
    slugify: slugify,
  });
})(typeof window !== "undefined" ? window : globalThis);
