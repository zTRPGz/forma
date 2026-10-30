(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const MAX_TEMPLATE_BYTES = 20 * 1024 * 1024;
  const WHITE = "FFFFFF";

  const PRESETS = [
    {
      id: "paper",
      name: "Формовка",
      note: "Тёплая бумага",
      ink: "1F2A37",
      navy: "16243A",
      accent: "B45309",
      muted: "6B5E50",
      rule: "D9CFC1",
      alt: "F3EDE4",
      light: "EADFD0",
      meta: "9AA7B8",
      sheet: WHITE,
      headFont: "Georgia",
      bodyFont: "Segoe UI",
    },
    {
      id: "graphite",
      name: "Графит",
      note: "Холодный синий",
      ink: "1B2530",
      navy: "26384D",
      accent: "3F6E95",
      muted: "5D6B79",
      rule: "D2DAE2",
      alt: "EDF2F7",
      light: "E4ECF4",
      meta: "8FA1B4",
      sheet: WHITE,
      headFont: "Cambria",
      bodyFont: "Segoe UI",
    },
    {
      id: "forest",
      name: "Лес",
      note: "Глубокий зелёный",
      ink: "1E2B24",
      navy: "1C3A2E",
      accent: "4F7D58",
      muted: "5D6B60",
      rule: "D4DCD3",
      alt: "EDF2EB",
      light: "DDE8DD",
      meta: "93A79A",
      sheet: WHITE,
      headFont: "Georgia",
      bodyFont: "Segoe UI",
    },
    {
      id: "clay",
      name: "Терракота",
      note: "Тёплый кирпич",
      ink: "2B2421",
      navy: "3B2A25",
      accent: "A64B2A",
      muted: "6F615A",
      rule: "E0D0C6",
      alt: "F6EDE7",
      light: "F1DCCF",
      meta: "B29C8D",
      sheet: WHITE,
      headFont: "Palatino Linotype",
      bodyFont: "Segoe UI",
    },
  ];

  const FALLBACK = PRESETS[0];

  function normalizeHex(value) {
    const raw = String(value || "").replace(/[^0-9a-f]/gi, "");
    if (raw.length === 3) {
      return raw
        .split("")
        .map(function (char) {
          return char + char;
        })
        .join("")
        .toUpperCase();
    }
    return raw.length >= 6 ? raw.slice(0, 6).toUpperCase() : "";
  }

  function channels(hex) {
    return [0, 2, 4].map(function (offset) {
      return parseInt(hex.slice(offset, offset + 2), 16);
    });
  }

  function toHex(values) {
    return values
      .map(function (value) {
        const clamped = Math.max(0, Math.min(255, Math.round(value)));
        return clamped.toString(16).padStart(2, "0");
      })
      .join("")
      .toUpperCase();
  }

  function blend(from, to, amount) {
    const a = channels(from);
    const b = channels(to);
    return toHex([
      a[0] + (b[0] - a[0]) * amount,
      a[1] + (b[1] - a[1]) * amount,
      a[2] + (b[2] - a[2]) * amount,
    ]);
  }

  function luminance(hex) {
    const rgb = channels(hex).map(function (value) {
      const channel = value / 255;
      return channel <= 0.03928 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4);
    });
    return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2];
  }

  function ensureDark(candidate, fallback) {
    if (candidate && luminance(candidate) <= 0.36) return candidate;
    if (fallback && luminance(fallback) <= 0.36) return fallback;
    return "1F2A37";
  }

  function ensureAccent(candidates) {
    for (let index = 0; index < candidates.length; index += 1) {
      const value = normalizeHex(candidates[index]);
      if (value) {
        const level = luminance(value);
        if (level >= 0.09 && level <= 0.55) return value;
      }
    }
    return FALLBACK.accent;
  }

  function resolve(theme) {
    if (theme && typeof theme === "object") {
      return Object.assign({}, FALLBACK, theme, {
        id: theme.id || "custom",
        headFont: theme.headFont || FALLBACK.headFont,
        bodyFont: theme.bodyFont || FALLBACK.bodyFont,
      });
    }
    const id = String(theme || "paper");
    for (let index = 0; index < PRESETS.length; index += 1) {
      if (PRESETS[index].id === id) return PRESETS[index];
    }
    return FALLBACK;
  }

  // ---------- чтение theme1.xml из DOCX/PPTX ----------

  function readU16(bytes, offset) {
    return bytes[offset] | (bytes[offset + 1] << 8);
  }

  function readU32(bytes, offset) {
    return (
      (bytes[offset] |
        (bytes[offset + 1] << 8) |
        (bytes[offset + 2] << 16) |
        (bytes[offset + 3] << 24)) >>>
      0
    );
  }

  function findEndOfCentralDirectory(bytes) {
    const stop = Math.max(0, bytes.length - 66000);
    for (let offset = bytes.length - 22; offset >= stop; offset -= 1) {
      if (readU32(bytes, offset) === 0x06054b50) return offset;
    }
    return -1;
  }

  function findEntry(bytes, pattern) {
    const end = findEndOfCentralDirectory(bytes);
    if (end < 0) throw new Error("это не ZIP-архив Office");

    let offset = readU32(bytes, end + 16);
    const total = readU16(bytes, end + 10);
    if (!total || offset >= bytes.length) throw new Error("архив пуст");

    const decoder = new TextDecoder("utf-8");
    for (let index = 0; index < total; index += 1) {
      if (readU32(bytes, offset) !== 0x02014b50) break;
      const method = readU16(bytes, offset + 10);
      const compressed = readU32(bytes, offset + 20);
      const nameLength = readU16(bytes, offset + 28);
      const extraLength = readU16(bytes, offset + 30);
      const commentLength = readU16(bytes, offset + 32);
      const localOffset = readU32(bytes, offset + 42);
      const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));

      if (pattern.test(name)) {
        if (readU32(bytes, localOffset) !== 0x04034b50) throw new Error("повреждённый архив");
        const localName = readU16(bytes, localOffset + 26);
        const localExtra = readU16(bytes, localOffset + 28);
        const start = localOffset + 30 + localName + localExtra;
        return { method: method, bytes: bytes.subarray(start, start + compressed) };
      }

      offset += 46 + nameLength + extraLength + commentLength;
    }

    return null;
  }

  async function inflate(entry) {
    if (entry.method === 0) return entry.bytes;
    if (entry.method !== 8) throw new Error("неизвестный способ сжатия темы");
    if (typeof global.DecompressionStream !== "function") {
      throw new Error("браузер не умеет распаковывать шаблоны — обновите его");
    }
    const stream = new Blob([entry.bytes])
      .stream()
      .pipeThrough(new global.DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function themeXml(bytes) {
    const doc = new DOMParser().parseFromString(new TextDecoder("utf-8").decode(bytes), "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("тема в файле повреждена");
    return doc;
  }

  function firstNode(doc, name) {
    return doc.getElementsByTagNameNS("*", name)[0] || null;
  }

  function schemeColor(doc, name) {
    const node = firstNode(doc, name);
    if (!node) return "";
    const srgb = node.getElementsByTagNameNS("*", "srgbClr")[0];
    const sys = node.getElementsByTagNameNS("*", "sysClr")[0];
    return normalizeHex(srgb ? srgb.getAttribute("val") : sys ? sys.getAttribute("lastClr") : "");
  }

  function typeface(doc, schemeName) {
    const scheme = firstNode(doc, schemeName);
    if (!scheme) return "";
    const latin = scheme.getElementsByTagNameNS("*", "latin")[0];
    const value = latin ? String(latin.getAttribute("typeface") || "").trim() : "";
    return value && value.charAt(0) !== "+" ? value.slice(0, 40) : "";
  }

  function fromThemeDocument(doc, name) {
    const colors = {
      dk1: schemeColor(doc, "dk1"),
      lt1: schemeColor(doc, "lt1"),
      dk2: schemeColor(doc, "dk2"),
      accent1: schemeColor(doc, "accent1"),
      accent2: schemeColor(doc, "accent2"),
      accent3: schemeColor(doc, "accent3"),
      accent4: schemeColor(doc, "accent4"),
    };

    const ink = ensureDark(colors.dk1, colors.dk2);
    const navy = ensureDark(colors.dk2, blend(ink, "000000", 0.18));
    const accent = ensureAccent([colors.accent1, colors.accent2, colors.accent3, colors.accent4]);
    const light = colors.lt1 && luminance(colors.lt1) >= 0.72 ? colors.lt1 : WHITE;
    const sheet = colors.lt1 && luminance(colors.lt1) >= 0.93 ? colors.lt1 : WHITE;

    return {
      id: "custom",
      name: name || "Свой шаблон",
      note: "Цвета и шрифты из файла",
      custom: true,
      ink: ink,
      navy: navy,
      accent: accent,
      muted: blend(ink, WHITE, 0.52),
      rule: blend(ink, WHITE, 0.82),
      alt: blend(accent, WHITE, 0.9),
      light: light,
      meta: blend(light, navy, 0.35),
      sheet: sheet,
      headFont: typeface(doc, "majorFont") || FALLBACK.headFont,
      bodyFont: typeface(doc, "minorFont") || FALLBACK.bodyFont,
    };
  }

  async function fromTemplate(file) {
    if (!file) throw new Error("файл не выбран");
    if (file.size > MAX_TEMPLATE_BYTES) throw new Error("файл больше 20 МБ");
    const bytes = new Uint8Array(await file.arrayBuffer());
    const entry = findEntry(bytes, /^(?:word|ppt)\/theme\/theme1\.xml$/i);
    if (!entry) throw new Error("в файле нет темы оформления — нужен DOCX или PPTX");
    const xml = await inflate(entry);
    const name = String(file.name || "")
      .replace(/\.[^.]+$/, "")
      .trim()
      .slice(0, 60);
    return fromThemeDocument(themeXml(xml), name);
  }

  Object.assign(Forma, {
    themes: {
      presets: PRESETS,
      resolve: resolve,
      fromTemplate: fromTemplate,
      blend: blend,
      luminance: luminance,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);
