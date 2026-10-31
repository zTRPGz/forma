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
  const outlineCount = document.getElementById("outline-count");
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
  const genEl = document.getElementById("gen");
  const genLiveLabel = document.getElementById("gen-live-label");
  const genFill = document.getElementById("gen-fill");
  const genSteps = Array.prototype.slice.call(document.querySelectorAll(".gen-step"));
  const topbar = document.querySelector(".topbar");
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
  let badgeLocked = false;

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

  function showGen() {
    if (!genEl) return;
    genEl.hidden = false;
    genEl.setAttribute("aria-hidden", "false");
    genEl.classList.remove("is-done", "is-error");
    genEl.dataset.stage = "content";
    if (genLiveLabel) genLiveLabel.textContent = "идёт";
    genSteps.forEach(function (step, index) {
      step.classList.toggle("is-active", index === 0);
      step.classList.remove("is-done");
    });
    if (genFill) genFill.style.transform = "scaleX(0.06)";
    setGenStage("content");
    const calm =
      window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    genEl.scrollIntoView({ block: "nearest", behavior: calm ? "auto" : "smooth" });
  }

  function setGenStage(stage) {
    if (!genEl) return;
    const order = ["content", "layout", "file"];
    const index = order.indexOf(stage);
    if (index < 0) return;
    genEl.dataset.stage = stage;
    genSteps.forEach(function (step, position) {
      step.classList.toggle("is-active", position === index);
      step.classList.toggle("is-done", position < index);
    });
    if (genFill) {
      genFill.style.transform = "scaleX(" + [0.16, 0.56, 0.86][index] + ")";
    }
  }

  function finishGen(failed) {
    if (!genEl || genEl.hidden) return;
    genEl.classList.toggle("is-error", Boolean(failed));
    genEl.classList.toggle("is-done", !failed);
    if (genLiveLabel) genLiveLabel.textContent = failed ? "ошибка" : "готово";
    genSteps.forEach(function (step) {
      step.classList.remove("is-active");
      step.classList.toggle("is-done", !failed);
    });
    if (genFill) genFill.style.transform = failed ? "scaleX(0.86)" : "scaleX(1)";
  }

  function hideGen() {
    if (!genEl) return;
    genEl.hidden = true;
    genEl.setAttribute("aria-hidden", "true");
  }

  function readStoredFormat() {
    try {
      const stored = localStorage.getItem(FORMAT_KEY);
      return stored && Object.prototype.hasOwnProperty.call(NAMES, stored) ? stored : "docx";
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
    if (!file || !Forma.themes || busy) return;
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
      badge.title = "Сервис ИИ отвечает с сервера «Тезиса».";
    } else if (state === "demo") {
      badge.textContent = "ИИ · демо-режим";
      badge.title = reason || "Сервис ИИ недоступен — документы собираются в демо-режиме.";
    } else {
      badge.textContent = "ИИ · подключение";
      badge.title = "Проверяю связь с сервисом ИИ.";
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
    if (!NAMES[next] || busy) return;
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
    hideGen();
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
    if (!next || next === slides || busy) return;
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
      if (!file) return;
      if (busy) {
        templateFile.value = "";
        templateMessage = {
          text: "Шаблон можно прикрепить, когда текущий файл соберётся.",
          error: true,
        };
        updateTemplateNote();
        return;
      }
      attachTemplate(file);
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

    const targetFormat = format;
    const targetSlides = slides;
    const targetTheme = currentTheme();

    setBusy(true);
    hideResult();
    showGen();
    setStatus("ИИ собирает содержание…");

    const slowHint = setTimeout(function () {
      if (targetFormat === "pptx" && targetSlides >= 12) {
        setStatus("Собираю большую презентацию — обычно это занимает до 40 секунд.");
      }
    }, 10000);

    const depsReady = Forma.deps.ensure(targetFormat).then(
      function () {
        return null;
      },
      function (error) {
        return error;
      }
    );

    const started = performance.now();
    let failed = false;

    try {
      const options = {
        slides: targetFormat === "pptx" ? targetSlides : null,
        theme: targetTheme,
      };
      const generated = await Forma.ai.generateContent(value, targetFormat, options);

      if (!Forma.deps.ready(targetFormat)) setStatus("Загружаю модуль вёрстки…");
      setGenStage("layout");
      const depsError = await depsReady;
      if (depsError) {
        throw new Error("модуль вёрстки не загрузился — проверьте, что папка vendor загружена полностью, и обновите страницу.");
      }

      clearTimeout(slowHint);
      setStatus("Собираю файл…");
      setGenStage("file");
      const blob = await Forma.renderers[targetFormat](generated.doc);
      const file = new Blob([blob], { type: MIME[targetFormat] });
      const url = URL.createObjectURL(file);
      if (lastUrl) URL.revokeObjectURL(lastUrl);
      lastUrl = url;

      const seconds = (performance.now() - started) / 1000;
      const filename = "tezis-" + Forma.slugify(generated.doc.title) + "." + targetFormat;

      renderResult(generated, url, filename, seconds, targetFormat);
      badgeLocked = true;
      if (generated.note) {
        setBadge("demo", generated.note);
        setStatus("Демо-документ готов — можно попробовать ещё раз.");
      } else {
        setBadge("live");
        setStatus("Файл готов.");
      }
    } catch (error) {
      failed = true;
      setStatus(
        "Не удалось собрать файл: " + (error && error.message ? error.message : "неизвестная ошибка."),
        true
      );
    } finally {
      clearTimeout(slowHint);
      finishGen(failed);
      setBusy(false);
    }
  }

  function renderResult(generated, url, filename, seconds, targetFormat) {
    const doc = generated.doc;

    resultTitle.textContent = doc.title;
    const parts = [NAMES[targetFormat]];
    if (targetFormat === "pptx" && doc.slides) parts.push(doc.slides + " " + slidesWord(doc.slides));
    if (Forma.themes) parts.push("Шаблон · " + Forma.themes.resolve(doc.theme).name);
    parts.push(generated.provider);
    if (generated.model) parts.push(generated.model);
    parts.push(seconds.toFixed(1).replace(".", ",") + " с");
    resultMeta.replaceChildren();
    parts.forEach(function (part, index) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = part;
      chip.style.animationDelay = 80 + index * 45 + "ms";
      resultMeta.append(chip);
    });

    resultNote.hidden = !generated.note;
    resultNote.textContent = generated.note || "";
    retryButton.hidden = !generated.note;

    downloadLink.href = url;
    downloadLink.download = filename;
    downloadLabel.textContent = "Скачать " + EXTS[targetFormat];
    downloadLink.setAttribute("aria-disabled", "false");

    outlineList.replaceChildren();
    if (outlineCount) outlineCount.textContent = String(doc.sections.length);
    doc.sections.forEach(function (section, index) {
      const item = document.createElement("li");
      item.style.animationDelay = 120 + index * 55 + "ms";
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

    preview.className = "preview preview--" + targetFormat;
    preview.replaceChildren();

    const caption = document.createElement("p");
    caption.className = "preview-caption";
    caption.textContent = "Макет " + NAMES[targetFormat] + " · точная вёрстка — в скачанном файле";

    const scroll = document.createElement("div");
    scroll.className = "preview-scroll";
    scroll.tabIndex = 0;
    scroll.setAttribute("role", "region");
    scroll.setAttribute(
      "aria-label",
      "Предпросмотр " + NAMES[targetFormat] + " — прокрутите, чтобы увидеть весь макет"
    );
    scroll.append(Forma.preview[targetFormat](doc));

    preview.append(caption, scroll);

    result.hidden = false;
    result.scrollIntoView({ block: "start" });
  }

  function syncTopbar() {
    if (topbar) topbar.classList.toggle("is-scrolled", window.scrollY > 8);
  }

  window.addEventListener("scroll", syncTopbar, { passive: true });

  format = readStoredFormat();
  slides = readStoredSlides();
  themeId = readStoredTheme();
  setBadge("checking");
  renderTemplates();
  updateTemplateNote();
  applyFormatUI();
  updateCounter();
  syncTopbar();

  Forma.ai.health().then(function (result) {
    if (badgeLocked) return;
    setBadge(result.ok ? "live" : "demo", result.reason);
  });
})();
