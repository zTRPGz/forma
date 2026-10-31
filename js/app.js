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
  const editOpen = document.getElementById("edit-open");
  const editorEl = document.getElementById("editor");
  const editTitle = document.getElementById("edit-title");
  const editSubtitle = document.getElementById("edit-subtitle");
  const editSections = document.getElementById("edit-sections");
  const editAdd = document.getElementById("edit-add");
  const photoEditor = document.getElementById("photo-editor");
  const photoFile = document.getElementById("photo-file");
  const photoPreview = document.getElementById("photo-preview");
  const photoEmpty = document.getElementById("photo-empty");
  const photoNote = document.getElementById("photo-note");
  const photoZoom = document.getElementById("photo-zoom");
  const photoZoomValue = document.getElementById("photo-zoom-value");
  const photoX = document.getElementById("photo-x");
  const photoXValue = document.getElementById("photo-x-value");
  const photoY = document.getElementById("photo-y");
  const photoYValue = document.getElementById("photo-y-value");
  const photoRemove = document.getElementById("photo-remove");
  const editorStatus = document.getElementById("editor-status");
  const editApply = document.getElementById("edit-apply");
  const editCancel = document.getElementById("edit-cancel");
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
  let active = null;
  let editDraft = null;
  let photoDraft = null;
  let photoTimer = null;
  let editorBusy = false;

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
    closeEditor();
    active = null;
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

  editOpen.addEventListener("click", openEditor);
  editApply.addEventListener("click", applyEdits);
  editCancel.addEventListener("click", function () {
    if (editorBusy) return;
    closeEditor();
    editOpen.focus();
  });

  editAdd.addEventListener("click", function () {
    if (editorBusy || !editDraft) return;
    const current = collectSections();
    if (current.length >= sectionLimit()) {
      updateAddState();
      return;
    }
    current.push({ heading: "", paragraphs: [], bullets: [], table: null });
    renderEditSections(current);
    const last = editSections.lastElementChild;
    if (last) {
      last.scrollIntoView({ block: "nearest" });
      const input = last.querySelector(".editor-heading");
      if (input) input.focus();
    }
  });

  editSections.addEventListener("click", function (event) {
    if (editorBusy) return;
    const button = event.target.closest(".editor-remove");
    if (!button) return;
    const card = button.closest(".editor-section");
    if (!card) return;
    const current = collectSections();
    const index = Array.prototype.indexOf.call(editSections.children, card);
    if (index < 0 || current.length <= 1) return;
    current.splice(index, 1);
    renderEditSections(current);
    setEditorMessage("");
  });

  photoFile.addEventListener("change", function () {
    const file = photoFile.files && photoFile.files[0];
    if (file) handlePhotoFile(file);
  });

  photoZoom.addEventListener("input", function () {
    if (!photoDraft) return;
    photoDraft.zoom = Number(photoZoom.value);
    photoZoomValue.textContent = photoZoom.value + " %";
    schedulePhotoPreview();
  });

  photoX.addEventListener("input", function () {
    if (!photoDraft) return;
    photoDraft.fx = Number(photoX.value);
    photoXValue.textContent = photoX.value + " %";
    schedulePhotoPreview();
  });

  photoY.addEventListener("input", function () {
    if (!photoDraft) return;
    photoDraft.fy = Number(photoY.value);
    photoYValue.textContent = photoY.value + " %";
    schedulePhotoPreview();
  });

  photoRemove.addEventListener("click", function () {
    if (editorBusy) return;
    resetPhotoUI();
    setPhotoNote("Фотография убрана — на слайдах вернётся рамка «Фото».");
  });

  document.addEventListener("keydown", function (event) {
    if (event.key !== "Escape" || editorEl.hidden || editorBusy) return;
    closeEditor();
    editOpen.focus();
  });

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
    active = {
      doc: generated.doc,
      url: url,
      filename: filename,
      format: targetFormat,
      provider: generated.provider,
      model: generated.model,
      note: generated.note,
      seconds: seconds,
      edited: false,
      photo: null,
    };
    paintResult(true);
  }

  function paintResult(scrollToResult) {
    if (!active) return;
    const doc = active.doc;
    const targetFormat = active.format;
    const previous = preview.querySelector(".preview-scroll");
    const keepScroll = previous ? previous.scrollTop : 0;

    resultTitle.textContent = doc.title;
    const parts = [NAMES[targetFormat]];
    if (targetFormat === "pptx" && doc.slides) parts.push(doc.slides + " " + slidesWord(doc.slides));
    if (Forma.themes) parts.push("Шаблон · " + Forma.themes.resolve(doc.theme).name);
    parts.push(active.provider);
    if (active.model) parts.push(active.model);
    if (active.edited) parts.push("изменено");
    if (active.photo) parts.push("фото");
    if (!active.edited) parts.push(active.seconds.toFixed(1).replace(".", ",") + " с");
    resultMeta.replaceChildren();
    parts.forEach(function (part, index) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = part;
      chip.style.animationDelay = 80 + index * 45 + "ms";
      resultMeta.append(chip);
    });

    resultNote.hidden = !active.note;
    resultNote.textContent = active.note || "";
    retryButton.hidden = !active.note;

    downloadLink.href = active.url;
    downloadLink.download = active.filename;
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
    if (keepScroll) scroll.scrollTop = keepScroll;

    preview.append(caption, scroll);

    result.hidden = false;
    if (scrollToResult) result.scrollIntoView({ block: "start" });
  }

  function cloneSections(sections) {
    return sections.map(function (section) {
      return {
        heading: section.heading,
        paragraphs: section.paragraphs.slice(),
        bullets: section.bullets.slice(),
        table: section.table
          ? {
              headers: section.table.headers.slice(),
              rows: section.table.rows.map(function (row) {
                return row.slice();
              }),
            }
          : null,
      };
    });
  }

  function splitLines(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map(function (line) {
        return line.replace(/^\s*[-–—•]\s+/, "").trim();
      })
      .filter(Boolean);
  }

  function sectionLimit() {
    return active && active.format === "pptx" ? 18 : 20;
  }

  function setEditorMessage(text, isError) {
    if (!editorStatus) return;
    editorStatus.textContent = text || "";
    editorStatus.classList.toggle("is-error", Boolean(isError));
  }

  function setEditorBusy(value, message) {
    editorBusy = value;
    editorEl.setAttribute("aria-busy", String(value));
    editApply.disabled = value;
    editCancel.disabled = value;
    editAdd.disabled = value || !editDraft || editDraft.sections.length >= sectionLimit();
    if (message) setEditorMessage(message);
  }

  function updateAddState() {
    if (!editAdd || !editDraft) return;
    const limit = sectionLimit();
    const full = editDraft.sections.length >= limit;
    editAdd.disabled = editorBusy || full;
    editAdd.textContent = full ? "Достигнут лимит разделов" : "+ Добавить раздел";
  }

  function sectionNote(text) {
    const note = document.createElement("p");
    note.className = "editor-section-note";
    note.textContent = text;
    return note;
  }

  function editSectionCard(section, index) {
    const card = document.createElement("article");
    card.className = "editor-section";
    card._table = section.table || null;

    const head = document.createElement("div");
    head.className = "editor-section-head";
    const title = document.createElement("p");
    title.className = "editor-section-title";
    title.textContent = "Раздел " + (index + 1);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "editor-remove";
    remove.textContent = "Удалить";
    remove.setAttribute("aria-label", "Удалить раздел " + (index + 1));
    remove.disabled = editDraft.sections.length <= 1;
    head.append(title, remove);

    const headingField = document.createElement("label");
    headingField.className = "editor-field";
    const headingLabel = document.createElement("span");
    headingLabel.textContent = "Заголовок раздела";
    const heading = document.createElement("input");
    heading.type = "text";
    heading.className = "editor-heading";
    heading.maxLength = 90;
    heading.autocomplete = "off";
    heading.placeholder = "Например, «Ключевые идеи»";
    heading.value = section.heading || "";
    headingField.append(headingLabel, heading);

    const grid = document.createElement("div");
    grid.className = "editor-grid-2";

    const paragraphsField = document.createElement("label");
    paragraphsField.className = "editor-field";
    const paragraphsLabel = document.createElement("span");
    paragraphsLabel.textContent = "Абзацы — по одному в строке";
    const paragraphs = document.createElement("textarea");
    paragraphs.rows = 3;
    paragraphs.maxLength = 2100;
    paragraphs.placeholder = "Текст раздела…";
    paragraphs.value = section.paragraphs.join("\n");
    paragraphsField.append(paragraphsLabel, paragraphs);

    const bulletsField = document.createElement("label");
    bulletsField.className = "editor-field";
    const bulletsLabel = document.createElement("span");
    bulletsLabel.textContent = "Пункты — по одному в строке";
    const bullets = document.createElement("textarea");
    bullets.rows = 3;
    bullets.maxLength = 1320;
    bullets.placeholder = "Пункт списка…";
    bullets.value = section.bullets.join("\n");
    bulletsField.append(bulletsLabel, bullets);

    grid.append(paragraphsField, bulletsField);
    card.append(head, headingField, grid);
    if (card._table) {
      card.append(sectionNote("Таблица в этом разделе сохранится без изменений."));
    }
    card.append(sectionNote("До 3 абзацев по 700 символов и до 6 пунктов по 220 символов."));
    return card;
  }

  function renderEditSections(sections) {
    if (sections) editDraft.sections = sections;
    editSections.replaceChildren();
    editDraft.sections.forEach(function (section, index) {
      editSections.append(editSectionCard(section, index));
    });
    updateAddState();
  }

  function collectSections() {
    return Array.prototype.map.call(
      editSections.querySelectorAll(".editor-section"),
      function (card) {
        const areas = card.querySelectorAll("textarea");
        return {
          heading: card.querySelector(".editor-heading").value.trim(),
          paragraphs: splitLines(areas[0].value),
          bullets: splitLines(areas[1].value),
          table: card._table || null,
        };
      }
    );
  }

  function openEditor() {
    if (!active || busy || editorBusy) return;
    editDraft = { sections: cloneSections(active.doc.sections) };
    editTitle.value = active.doc.title;
    editSubtitle.value = active.doc.subtitle || "";
    renderEditSections();
    resetPhotoUI();
    photoEditor.hidden = active.format !== "pptx";
    setEditorMessage("");
    setEditorBusy(false);
    if (active.format === "pptx" && active.photo && active.photo.data) {
      loadPhotoSource(active.photo.data, active.photo.name);
    }
    editorEl.hidden = false;
    editOpen.setAttribute("aria-expanded", "true");
    const calm =
      window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    editorEl.scrollIntoView({ block: "nearest", behavior: calm ? "auto" : "smooth" });
    editTitle.focus({ preventScroll: true });
  }

  function closeEditor() {
    if (!editorEl) return;
    editorEl.hidden = true;
    editOpen.setAttribute("aria-expanded", "false");
    editSections.replaceChildren();
    editDraft = null;
    clearTimeout(photoTimer);
    resetPhotoUI();
    setEditorMessage("");
    setEditorBusy(false);
  }

  function photoSummary(photo) {
    const name = photo.name ? "«" + photo.name + "»" : "Фотография";
    return name + " · " + photo.width + "×" + photo.height + " px";
  }

  function setPhotoNote(text, isError) {
    if (!photoNote) return;
    photoNote.textContent = text;
    photoNote.classList.toggle("is-error", Boolean(isError));
  }

  function resetPhotoUI() {
    photoDraft = null;
    clearTimeout(photoTimer);
    photoFile.disabled = false;
    photoPreview.hidden = true;
    photoPreview.removeAttribute("src");
    photoEmpty.hidden = false;
    photoRemove.hidden = true;
    photoZoom.disabled = true;
    photoX.disabled = true;
    photoY.disabled = true;
    photoZoom.value = "100";
    photoX.value = "50";
    photoY.value = "50";
    photoZoomValue.textContent = "100 %";
    photoXValue.textContent = "50 %";
    photoYValue.textContent = "50 %";
    photoFile.value = "";
    setPhotoNote("JPEG, PNG или WebP до 16 МБ.");
  }

  function cropPhoto(photo, width, height) {
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    const scale = Math.max(width / photo.width, height / photo.height) * (photo.zoom / 100);
    const drawW = photo.width * scale;
    const drawH = photo.height * scale;
    const maxX = Math.max(0, drawW - width);
    const maxY = Math.max(0, drawH - height);
    const offsetX = Math.min(Math.max((photo.fx / 100) * drawW - width / 2, 0), maxX);
    const offsetY = Math.min(Math.max((photo.fy / 100) * drawH - height / 2, 0), maxY);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
    ctx.drawImage(
      photo.img,
      offsetX / scale,
      offsetY / scale,
      width / scale,
      height / scale,
      0,
      0,
      width,
      height
    );
    return canvas.toDataURL("image/jpeg", 0.9);
  }

  function updatePhotoPreview() {
    if (!photoDraft || editorEl.hidden) return;
    const data = cropPhoto(photoDraft, 360, 480);
    if (data) photoPreview.src = data;
  }

  function schedulePhotoPreview() {
    clearTimeout(photoTimer);
    photoTimer = setTimeout(updatePhotoPreview, 60);
  }

  function syncPhotoUI() {
    const hasPhoto = Boolean(photoDraft);
    photoEmpty.hidden = hasPhoto;
    photoPreview.hidden = !hasPhoto;
    photoRemove.hidden = !hasPhoto;
    photoZoom.disabled = !hasPhoto;
    photoX.disabled = !hasPhoto;
    photoY.disabled = !hasPhoto;
    if (!hasPhoto) return;
    photoZoom.value = String(photoDraft.zoom);
    photoX.value = String(photoDraft.fx);
    photoY.value = String(photoDraft.fy);
    photoZoomValue.textContent = photoDraft.zoom + " %";
    photoXValue.textContent = photoDraft.fx + " %";
    photoYValue.textContent = photoDraft.fy + " %";
    setPhotoNote(photoSummary(photoDraft));
    updatePhotoPreview();
  }

  function loadPhotoSource(src, name) {
    const img = new Image();
    editApply.disabled = true;
    img.onload = function () {
      if (editorEl.hidden) return;
      photoDraft = {
        img: img,
        width: img.naturalWidth,
        height: img.naturalHeight,
        zoom: 100,
        fx: 50,
        fy: 50,
        name: name || "фотография",
      };
      photoFile.disabled = false;
      editApply.disabled = editorBusy;
      syncPhotoUI();
    };
    img.onerror = function () {
      photoFile.disabled = false;
      editApply.disabled = editorBusy;
      if (editorEl.hidden) return;
      setPhotoNote("Не удалось открыть фотографию. Попробуйте другой файл.", true);
    };
    img.src = src;
  }

  function handlePhotoFile(file) {
    if (!file) return;
    if (["image/jpeg", "image/png", "image/webp"].indexOf(file.type) < 0) {
      setPhotoNote("Подойдёт JPEG, PNG или WebP.", true);
      photoFile.value = "";
      return;
    }
    if (file.size > 16 * 1024 * 1024) {
      setPhotoNote("Файл больше 16 МБ — выберите снимок меньшего размера.", true);
      photoFile.value = "";
      return;
    }
    setPhotoNote("Читаю фотографию…");
    editApply.disabled = true;
    photoFile.disabled = true;
    const reader = new FileReader();
    reader.onload = function () {
      if (editorEl.hidden) return;
      loadPhotoSource(String(reader.result), file.name);
    };
    reader.onerror = function () {
      photoFile.disabled = false;
      editApply.disabled = editorBusy;
      setPhotoNote("Не удалось прочитать файл.", true);
      photoFile.value = "";
    };
    reader.readAsDataURL(file);
  }

  async function applyEdits() {
    if (!active || editorBusy || busy) return;

    const title = editTitle.value.trim();
    if (!title) {
      setEditorMessage("Укажите заголовок документа.", true);
      editTitle.focus();
      return;
    }

    const sections = collectSections();
    if (!sections.length) {
      setEditorMessage("Оставьте хотя бы один раздел.", true);
      return;
    }
    for (let index = 0; index < sections.length; index += 1) {
      const section = sections[index];
      const card = editSections.children[index];
      if (!section.heading) {
        setEditorMessage("Укажите заголовок раздела " + (index + 1) + ".", true);
        if (card) card.querySelector(".editor-heading").focus();
        return;
      }
      if (!section.paragraphs.length && !section.bullets.length && !section.table) {
        setEditorMessage("Добавьте текст или пункты в раздел " + (index + 1) + ".", true);
        if (card) card.querySelector("textarea").focus();
        return;
      }
    }

    const nextDoc = Forma.normalize(
      {
        title: title,
        subtitle: editSubtitle.value.trim(),
        sections: sections,
        theme: active.doc.theme,
      },
      active.format,
      {
        slides: active.format === "pptx" ? sections.length + 2 : null,
        theme: active.doc.theme,
      }
    );

    let photo = null;
    if (active.format === "pptx" && photoDraft) {
      const data = cropPhoto(photoDraft, 1035, 1380);
      if (data) {
        photo = {
          data: data,
          name: photoDraft.name,
          width: photoDraft.width,
          height: photoDraft.height,
        };
      }
    }
    if (photo) nextDoc.photo = photo;

    setEditorBusy(true, "Собираю обновлённый файл…");
    try {
      await Forma.deps.ensure(active.format);
      const blob = await Forma.renderers[active.format](nextDoc);
      const url = URL.createObjectURL(new Blob([blob], { type: MIME[active.format] }));
      if (lastUrl) URL.revokeObjectURL(lastUrl);
      lastUrl = url;

      active.doc = nextDoc;
      active.url = url;
      active.filename = "tezis-" + Forma.slugify(nextDoc.title) + "." + active.format;
      active.edited = true;
      active.photo = photo;

      closeEditor();
      paintResult(false);
      setStatus("Файл обновлён — можно скачать.");
    } catch (error) {
      setEditorMessage(
        "Не удалось обновить файл: " +
          (error && error.message ? error.message : "неизвестная ошибка."),
        true
      );
    } finally {
      setEditorBusy(false);
    }
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
