(function () {
  "use strict";

  const NAMES = { docx: "Word", pdf: "PDF", pptx: "PowerPoint" };
  const EXTS = { docx: "DOCX", pdf: "PDF", pptx: "PPTX" };
  const MIME = {
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    pdf: "application/pdf",
  };
  const FORMAT_KEY = "forma.format.v1";
  const SLIDES_KEY = "forma.slides.v1";
  const THEME_KEY = "forma.theme.v1";
  const SLIDE_COUNTS = [4, 6, 8, 10, 12, 16, 20];
  const FALLBACK_IDEAS = [
    "онбординг новых сотрудников за первые две недели",
    "план запуска продукта на новом рынке",
    "анализ рынка доставки продуктов",
    "регламент совещаний и рабочих встреч",
    "стратегия удержания клиентов",
    "программа обучения руководителей",
  ];

  const topic = document.getElementById("topic");
  const count = document.getElementById("count");
  const counter = count.parentElement;
  const generateButton = document.getElementById("generate");
  const ctaText = document.getElementById("ctaText");
  const ctaFormat = document.getElementById("ctaFormat");
  const statusLine = document.getElementById("status");
  const result = document.getElementById("result");
  const resultTitle = document.getElementById("result-title");
  const resultMeta = document.getElementById("result-meta");
  const resultNote = document.getElementById("result-note");
  const downloadLink = document.getElementById("download");
  const downloadLabel = document.getElementById("download-label");
  const retryButton = document.getElementById("retry");
  const outlineList = document.getElementById("outline-list");
  const preview = document.getElementById("preview");
  const badge = document.getElementById("llm-badge");
  const formatsEl = document.querySelector(".formats");
  const formatButtons = Array.prototype.slice.call(document.querySelectorAll(".format"));
  const slidesRow = document.getElementById("slides-row");
  const slidesEl = document.getElementById("slides");
  const slideButtons = Array.prototype.slice.call(document.querySelectorAll(".seg-btn"));
  const templateList = document.getElementById("template-list");
  const templateFile = document.getElementById("template-file");
  const templateAttach = document.getElementById("template-attach");
  const attachLabel = document.getElementById("attach-label");
  const templateNote = document.getElementById("template-note");
  const ideaButtons = Array.prototype.slice.call(document.querySelectorAll("#ideas .idea"));
  const IDEAS = ideaButtons
    .map(function (button) {
      return button.dataset.topic;
    })
    .filter(Boolean);
  if (!IDEAS.length) IDEAS.push.apply(IDEAS, FALLBACK_IDEAS);

  let format = "docx";
  let busy = false;
  let lastUrl = null;
  let slides = 6;
  let themeId = "paper";
  let customTheme = null;
  let templateMessage = null;
  let lastIdea = -1;

  function setStatus(message, isError) {
    statusLine.textContent = message;
    statusLine.classList.toggle("is-error", Boolean(isError));
  }

  function setBusy(value) {
    busy = value;
    generateButton.disabled = value;
    generateButton.setAttribute("aria-busy", String(value));
    statusLine.classList.toggle("is-busy", value);
    ctaText.textContent = value ? "Собираю…" : "Сгенерировать";
  }

  function readStoredFormat() {
    try {
      const stored = localStorage.getItem(FORMAT_KEY);
      return stored && NAMES[stored] ? stored : "docx";
    } catch (error) {
      return "docx";
    }
  }

  function storeFormat() {
    try {
      localStorage.setItem(FORMAT_KEY, format);
    } catch (error) {
      /* приватный режим — формат запомнится только до перезагрузки */
    }
  }

  function readStoredSlides() {
    try {
      const stored = Number(localStorage.getItem(SLIDES_KEY));
      return SLIDE_COUNTS.indexOf(stored) >= 0 ? stored : 6;
    } catch (error) {
      return 6;
    }
  }

  function storeSlides() {
    try {
      localStorage.setItem(SLIDES_KEY, String(slides));
    } catch (error) {
      /* не критично */
    }
  }

  function readStoredTheme() {
    try {
      const stored = localStorage.getItem(THEME_KEY);
      if (stored === "custom") return "paper";
      if (stored && Forma.themes && Forma.themes.resolve(stored).id === stored) return stored;
      return "paper";
    } catch (error) {
      return "paper";
    }
  }

  function storeTheme() {
    try {
      localStorage.setItem(THEME_KEY, themeId);
    } catch (error) {
      /* не критично */
    }
  }

  function slidesWord(value) {
    const last = value % 10;
    const tail = value % 100;
    if (last === 1 && tail !== 11) return "слайд";
    if (last >= 2 && last <= 4 && (tail < 12 || tail > 14)) return "слайда";
    return "слайдов";
  }

  function pickIdea() {
    if (IDEAS.length === 1) return IDEAS[0];
    let index = Math.floor(Math.random() * IDEAS.length);
    if (index === lastIdea) index = (index + 1) % IDEAS.length;
    lastIdea = index;
    return IDEAS[index];
  }

  function currentTheme() {
    if (themeId === "custom" && customTheme) return customTheme;
    return Forma.themes ? Forma.themes.resolve(themeId) : themeId;
  }

  function applySlidesUI() {
    slideButtons.forEach(function (button) {
      const active = Number(button.dataset.slides) === slides;
      button.classList.toggle("is-active", active);
      button.setAttribute("aria-checked", String(active));
      button.tabIndex = active ? 0 : -1;
    });
    slidesRow.hidden = format !== "pptx";
  }

  function setTemplateNote(text, isError) {
    templateMessage = text ? { text: text, error: Boolean(isError) } : null;
    updateTemplateNote();
  }

  function updateTemplateNote() {
    if (templateMessage) {
      templateNote.hidden = false;
      templateNote.classList.toggle("is-error", templateMessage.error);
      templateNote.textContent = templateMessage.text;
      return;
    }
    if (themeId === "custom" && customTheme) {
      templateNote.hidden = false;
      templateNote.classList.remove("is-error");
      templateNote.textContent =
        "Шаблон «" + customTheme.name + "» прикреплён: цвета и шрифты взяты из файла.";
      return;
    }
    templateNote.hidden = true;
    templateNote.textContent = "";
    templateNote.classList.remove("is-error");
  }

  function templateChip(theme, isCustom) {
    const key = isCustom ? "custom" : theme.id;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "template" + (themeId === key ? " is-active" : "");
    button.dataset.theme = key;
    button.setAttribute("aria-pressed", String(themeId === key));
    button.title = theme.note || theme.name;

    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.setAttribute(
      "style",
      "background:linear-gradient(90deg,#" +
        theme.navy +
        " 0 36%,#" +
        theme.accent +
        " 36% 66%,#" +
        theme.alt +
        " 66% 100%)"
    );

    const name = document.createElement("span");
    name.className = "template-name";
    name.textContent = theme.name;

    button.append(swatch, name);
    button.addEventListener("click", function () {
      themeId = key;
      templateMessage = null;
      renderTemplates();
      updateTemplateNote();
      storeTheme();
      hideResult();
      setStatus("");
    });
    return button;
  }

  function renderTemplates() {
    if (!templateList || !Forma.themes) return;
    templateList.replaceChildren();
    Forma.themes.presets.forEach(function (theme) {
      templateList.append(templateChip(theme, false));
    });
    if (customTheme) templateList.append(templateChip(customTheme, true));
  }

  async function attachTemplate(file) {
    if (!file || !Forma.themes) return;
    templateAttach.classList.add("is-busy");
    attachLabel.textContent = "Читаю шаблон…";
    templateFile.disabled = true;
    try {
      const theme = await Forma.themes.fromTemplate(file);
      customTheme = theme;
      themeId = "custom";
      templateMessage = null;
      renderTemplates();
      updateTemplateNote();
      storeTheme();
      hideResult();
      setStatus("");
    } catch (error) {
      setTemplateNote(
        "Не удалось прочитать шаблон: " + (error && error.message ? error.message : "неизвестная ошибка."),
        true
      );
    } finally {
      templateAttach.classList.remove("is-busy");
      attachLabel.textContent = "Прикрепить .docx / .pptx";
      templateFile.disabled = false;
      templateFile.value = "";
    }
  }

  function setBadge(state, reason) {
    badge.classList.toggle("badge--live", state === "live");
    badge.classList.toggle("badge--demo", state === "demo");
    if (state === "live") {
      badge.textContent = "ИИ · на связи";
      badge.title = "GigaChat-2 отвечает с сервера FORMA.";
    } else if (state === "demo") {
      badge.textContent = "ИИ · демо-режим";
      badge.title = reason || "Сервис ИИ недоступен — документы собираются в демо-режиме.";
    } else {
      badge.textContent = "ИИ · GigaChat-2";
      badge.title = "Содержание собирает GigaChat на сервере FORMA.";
    }
  }

  function applyFormatUI() {
    formatButtons.forEach(function (button) {
      const active = button.dataset.format === format;
      button.classList.toggle("is-active", active);
      button.setAttribute("aria-checked", String(active));
      button.tabIndex = active ? 0 : -1;
    });
    ctaFormat.textContent = NAMES[format];
    applySlidesUI();
  }

  function selectFormat(next, moveFocus) {
    if (!NAMES[next]) return;
    const changed = format !== next;
    format = next;
    applyFormatUI();
    storeFormat();
    Forma.deps.prefetch(format);

    if (moveFocus) {
      const button = formatButtons.find(function (item) {
        return item.dataset.format === format;
      });
      if (button) button.focus();
    }

    if (changed) {
      hideResult();
      setStatus("");
    }
  }

  function hideResult() {
    result.hidden = true;
    retryButton.hidden = true;
    preview.className = "preview";
    preview.replaceChildren();
    downloadLink.removeAttribute("href");
    if (lastUrl) {
      URL.revokeObjectURL(lastUrl);
      lastUrl = null;
    }
  }

  function updateCounter() {
    const length = topic.value.length;
    count.textContent = String(length);
    counter.classList.toggle("is-near", length >= 270);
  }

  formatButtons.forEach(function (button) {
    button.addEventListener("click", function () {
      selectFormat(button.dataset.format, false);
    });
    button.addEventListener(
      "pointerenter",
      function () {
        Forma.deps.prefetch(button.dataset.format);
      },
      { once: true }
    );
    button.addEventListener(
      "focus",
      function () {
        Forma.deps.prefetch(button.dataset.format);
      },
      { once: true }
    );
  });

  const ideasEl = document.getElementById("ideas");
  if (ideasEl) {
    ideasEl.addEventListener("click", function (event) {
      const button = event.target.closest(".idea");
      if (!button) return;
      topic.value = button.dataset.topic || "";
      updateCounter();
      topic.focus();
      Forma.deps.prefetch(format);
    });
  }

  formatsEl.addEventListener("keydown", function (event) {
    const keys = ["ArrowRight", "ArrowLeft", "ArrowDown", "ArrowUp", "Home", "End"];
    if (keys.indexOf(event.key) < 0) return;
    const index = formatButtons.indexOf(document.activeElement);
    if (index < 0) return;

    event.preventDefault();
    let nextIndex;
    if (event.key === "Home") nextIndex = 0;
    else if (event.key === "End") nextIndex = formatButtons.length - 1;
    else if (event.key === "ArrowRight" || event.key === "ArrowDown") {
      nextIndex = (index + 1) % formatButtons.length;
    } else {
      nextIndex = (index - 1 + formatButtons.length) % formatButtons.length;
    }
    selectFormat(formatButtons[nextIndex].dataset.format, true);
  });

  function selectSlides(next, moveFocus) {
    if (!next || next === slides) return;
    slides = next;
    applySlidesUI();
    storeSlides();
    hideResult();
    setStatus("");
    if (moveFocus) {
      const button = slideButtons.find(function (item) {
        return Number(item.dataset.slides) === slides;
      });
      if (button) button.focus();
    }
  }

  slideButtons.forEach(function (button) {
    button.addEventListener("click", function () {
      selectSlides(Number(button.dataset.slides), false);
    });
  });

  slidesEl.addEventListener("keydown", function (event) {
    const keys = ["ArrowRight", "ArrowLeft", "ArrowDown", "ArrowUp", "Home", "End"];
    if (keys.indexOf(event.key) < 0) return;
    const index = slideButtons.indexOf(document.activeElement);
    if (index < 0) return;

    event.preventDefault();
    let nextIndex;
    if (event.key === "Home") nextIndex = 0;
    else if (event.key === "End") nextIndex = slideButtons.length - 1;
    else if (event.key === "ArrowRight" || event.key === "ArrowDown") {
      nextIndex = (index + 1) % slideButtons.length;
    } else {
      nextIndex = (index - 1 + slideButtons.length) % slideButtons.length;
    }
    selectSlides(Number(slideButtons[nextIndex].dataset.slides), true);
  });

  if (templateFile) {
    templateFile.addEventListener("change", function () {
      const file = templateFile.files && templateFile.files[0];
      if (file) attachTemplate(file);
    });
  }

  topic.addEventListener("input", updateCounter);
  topic.addEventListener(
    "focus",
    function () {
      Forma.deps.prefetch(format);
    },
    { once: true }
  );

  topic.addEventListener("keydown", function (event) {
    if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      generate();
    }
  });

  generateButton.addEventListener("click", generate);
  retryButton.addEventListener("click", generate);

  window.addEventListener("beforeunload", function () {
    if (lastUrl) URL.revokeObjectURL(lastUrl);
  });

  async function generate() {
    if (busy) return;

    let value = topic.value.trim();
    if (!value) {
      value = pickIdea();
      topic.value = value;
      updateCounter();
    }
    if (value.length < 3) {
      setStatus("Опишите тему хотя бы тремя символами.", true);
      topic.focus();
      return;
    }

    setBusy(true);
    hideResult();
    setStatus("GigaChat пишет содержание…");

    const slowHint = setTimeout(function () {
      if (format === "pptx" && slides >= 12) {
        setStatus("Собираю большую презентацию — обычно это занимает до 40 секунд.");
      }
    }, 10000);

    const depsReady = Forma.deps.ensure(format).then(
      function () {
        return null;
      },
      function (error) {
        return error;
      }
    );

    const started = performance.now();

    try {
      const options = {
        slides: format === "pptx" ? slides : null,
        theme: currentTheme(),
      };
      const generated = await Forma.ai.generateContent(value, format, options);

      if (!Forma.deps.ready(format)) setStatus("Загружаю модуль вёрстки…");
      const depsError = await depsReady;
      if (depsError) {
        throw new Error("модуль вёрстки не загрузился — проверьте, что папка vendor загружена полностью, и обновите страницу.");
      }

      clearTimeout(slowHint);
      setStatus("Собираю файл…");
      const blob = await Forma.renderers[format](generated.doc);
      const file = new Blob([blob], { type: MIME[format] });
      const url = URL.createObjectURL(file);
      if (lastUrl) URL.revokeObjectURL(lastUrl);
      lastUrl = url;

      const seconds = (performance.now() - started) / 1000;
      const filename = "forma-" + Forma.slugify(generated.doc.title) + "." + format;

      renderResult(generated, url, filename, seconds);
      if (generated.note) {
        setBadge("demo", generated.note);
        setStatus("Демо-документ готов — можно попробовать ещё раз.");
      } else {
        setBadge("live");
        setStatus("Файл готов.");
      }
    } catch (error) {
      setStatus(
        "Не удалось собрать файл: " + (error && error.message ? error.message : "неизвестная ошибка."),
        true
      );
    } finally {
      clearTimeout(slowHint);
      setBusy(false);
    }
  }

  function renderResult(generated, url, filename, seconds) {
    const doc = generated.doc;

    resultTitle.textContent = doc.title;
    const parts = [NAMES[format]];
    if (format === "pptx" && doc.slides) parts.push(doc.slides + " " + slidesWord(doc.slides));
    if (Forma.themes) parts.push("Шаблон · " + Forma.themes.resolve(doc.theme).name);
    parts.push(generated.provider);
    if (generated.model) parts.push(generated.model);
    parts.push(seconds.toFixed(1).replace(".", ",") + " с");
    resultMeta.replaceChildren();
    parts.forEach(function (part) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = part;
      resultMeta.append(chip);
    });

    resultNote.hidden = !generated.note;
    resultNote.textContent = generated.note || "";
    retryButton.hidden = !generated.note;

    downloadLink.href = url;
    downloadLink.download = filename;
    downloadLabel.textContent = "Скачать " + EXTS[format];
    downloadLink.setAttribute("aria-disabled", "false");

    outlineList.replaceChildren();
    doc.sections.forEach(function (section, index) {
      const item = document.createElement("li");
      const number = document.createElement("span");
      number.className = "n";
      number.textContent = String(index + 1).padStart(2, "0");

      const body = document.createElement("div");
      const heading = document.createElement("span");
      heading.textContent = section.heading;
      const detail = document.createElement("span");
      detail.className = "d";
      const bits = [];
      if (section.paragraphs.length) bits.push(section.paragraphs.length + " абз.");
      if (section.bullets.length) bits.push(section.bullets.length + " п.");
      if (section.table) bits.push("таблица");
      detail.textContent = bits.join(" · ");

      body.append(heading, detail);
      item.append(number, body);
      outlineList.append(item);
    });

    preview.className = "preview preview--" + format;
    preview.replaceChildren();

    const caption = document.createElement("p");
    caption.className = "preview-caption";
    caption.textContent = "Макет " + NAMES[format] + " · точная вёрстка — в скачанном файле";

    const scroll = document.createElement("div");
    scroll.className = "preview-scroll";
    scroll.append(Forma.preview[format](doc));

    preview.append(caption, scroll);

    result.hidden = false;
    result.scrollIntoView({ block: "start" });
  }

  format = readStoredFormat();
  slides = readStoredSlides();
  themeId = readStoredTheme();
  setBadge("checking");
  renderTemplates();
  updateTemplateNote();
  applyFormatUI();
  updateCounter();

  Forma.ai.health().then(function (result) {
    setBadge(result.ok ? "live" : "demo", result.reason);
  });
})();
