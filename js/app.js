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

  let format = "docx";
  let busy = false;
  let lastUrl = null;

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

    const value = topic.value.trim();
    if (value.length < 3) {
      setStatus("Опишите тему хотя бы тремя символами.", true);
      topic.focus();
      return;
    }

    setBusy(true);
    hideResult();
    setStatus("GigaChat пишет содержание…");

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
      const generated = await Forma.ai.generateContent(value, format);

      if (!Forma.deps.ready(format)) setStatus("Загружаю модуль вёрстки…");
      const depsError = await depsReady;
      if (depsError) {
        throw new Error("модуль вёрстки не загрузился — проверьте, что папка vendor загружена полностью, и обновите страницу.");
      }

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
        setStatus("Демо-документ готов — ИИ не ответил.");
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
      setBusy(false);
    }
  }

  function renderResult(generated, url, filename, seconds) {
    const doc = generated.doc;

    resultTitle.textContent = doc.title;
    const parts = [NAMES[format], generated.provider];
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

    if (format === "pdf") {
      const frame = document.createElement("iframe");
      frame.src = url;
      frame.title = "Предпросмотр PDF";
      preview.append(frame);
    } else {
      const caption = document.createElement("p");
      caption.className = "preview-caption";
      caption.textContent =
        "Макет " + NAMES[format] + " · точная вёрстка — в скачанном файле";

      const scroll = document.createElement("div");
      scroll.className = "preview-scroll";
      scroll.append(Forma.preview[format](doc));

      preview.append(caption, scroll);
    }

    result.hidden = false;
    result.scrollIntoView({ block: "start" });
  }

  format = readStoredFormat();
  setBadge("checking");
  applyFormatUI();
  updateCounter();

  Forma.ai.health().then(function (result) {
    setBadge(result.ok ? "live" : "demo", result.reason);
  });
})();
