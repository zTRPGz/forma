(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const doc = global.document;

  let fontsLoaded = false;

  const LIBS = {
    docx: { src: "./vendor/docx.umd.js", ready: function () { return !!global.docx; } },
    pptx: { src: "./vendor/pptxgen.bundle.js", ready: function () { return !!global.PptxGenJS; } },
    pdfmake: { src: "./vendor/pdfmake.min.js", ready: function () { return !!global.pdfMake; } },
    fonts: { src: "./vendor/vfs_fonts.js", ready: function () { return fontsLoaded; } },
  };

  const ORDER = {
    docx: ["docx"],
    pptx: ["pptx"],
    pdf: ["pdfmake", "fonts"],
  };

  const pending = {};
  const warmed = {};

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      const script = doc.createElement("script");
      script.src = src;
      script.async = true;
      script.dataset.formaLib = src;
      script.onload = function () {
        resolve();
      };
      script.onerror = function () {
        script.remove();
        reject(new Error("Не удалось загрузить " + src));
      };
      doc.head.appendChild(script);
    });
  }

  function ensureLib(key) {
    const lib = LIBS[key];
    if (!lib) return Promise.reject(new Error("Неизвестный модуль: " + key));
    if (lib.ready()) return Promise.resolve();

    if (!pending[key]) {
      pending[key] = loadScript(lib.src)
        .then(function () {
          if (key === "fonts") fontsLoaded = true;
          if (!lib.ready()) throw new Error("Модуль " + key + " не инициализировался.");
        })
        .catch(function (error) {
          delete pending[key];
          throw error;
        });
    }
    return pending[key];
  }

  function ensure(format) {
    const keys = ORDER[format];
    if (!keys) return Promise.reject(new Error("Неизвестный формат: " + format));
    return keys.reduce(function (chain, key) {
      return chain.then(function () {
        return ensureLib(key);
      });
    }, Promise.resolve());
  }

  function ready(format) {
    const keys = ORDER[format] || [];
    return keys.every(function (key) {
      return LIBS[key].ready();
    });
  }

  function prefetch(format) {
    if (!ORDER[format] || warmed[format] || ready(format)) return;
    warmed[format] = true;
    ensure(format).catch(function () {
      warmed[format] = false;
    });
  }

  Forma.deps = { ensure: ensure, prefetch: prefetch, ready: ready };
})(typeof window !== "undefined" ? window : globalThis);
