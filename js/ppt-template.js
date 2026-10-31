(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const BUILTIN = {
    mfua: "./vendor/mfua-template.pptx",
  };
  const MAX_BYTES = 20 * 1024 * 1024;
  const MAX_ENTRY_BYTES = 10 * 1024 * 1024;
  const RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const REL_TYPES = {
    officeDocument: OFFICE_REL_NS + "/officeDocument",
    slideMaster: OFFICE_REL_NS + "/slideMaster",
    slide: OFFICE_REL_NS + "/slide",
    slideLayout: OFFICE_REL_NS + "/slideLayout",
    notesSlide: OFFICE_REL_NS + "/notesSlide",
    image: OFFICE_REL_NS + "/image",
  };
  const SLIDE_CONTENT_TYPE =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
  const IMAGE_CONTENT_TYPE = {
    jpeg: "image/jpeg",
    png: "image/png",
  };
  const MEDIA_PHOTO_NAME = "forma-photo";
  const SLIDE_REL_PATTERN = /^ppt\/slides\/[^/]+\.xml$/;

  let builtinPromise = null;

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

  function setU16(view, offset, value) {
    view.setUint16(offset, value & 0xffff, true);
  }

  function setU32(view, offset, value) {
    view.setUint32(offset, value >>> 0, true);
  }

  function hasZip64Marker(value) {
    return value === 0xffff || value === 0xffffffff;
  }

  function findEocd(bytes) {
    const stop = Math.max(0, bytes.length - 66000);
    for (let offset = bytes.length - 22; offset >= stop; offset -= 1) {
      if (readU32(bytes, offset) === 0x06054b50) return offset;
    }
    return -1;
  }

  async function inflateEntry(bytes, entry) {
    if (entry.method === 0) return bytes.subarray(entry.start, entry.start + entry.compressed);
    if (entry.method !== 8) throw new Error("неизвестный способ сжатия шаблона");
    if (typeof global.DecompressionStream !== "function") {
      throw new Error("браузер не умеет распаковывать шаблоны — обновите его");
    }
    const stream = new Blob([bytes.subarray(entry.start, entry.start + entry.compressed)])
      .stream()
      .pipeThrough(new global.DecompressionStream("deflate-raw"));
    const reader = stream.getReader();
    const chunks = [];
    let total = 0;
    while (true) {
      const part = await reader.read();
      if (part.done) break;
      total += part.value.byteLength;
      if (total > MAX_ENTRY_BYTES) {
        await reader.cancel();
        throw new Error("элемент шаблона слишком большой");
      }
      chunks.push(part.value);
    }
    const result = new Uint8Array(total);
    let offset = 0;
    chunks.forEach(function (chunk) {
      result.set(chunk, offset);
      offset += chunk.byteLength;
    });
    return result;
  }

  async function readZip(bytes) {
    if (bytes.length > MAX_BYTES) throw new Error("шаблон больше 20 МБ");
    const end = findEocd(bytes);
    if (end < 0) throw new Error("это не PowerPoint-шаблон");
    const total = readU16(bytes, end + 10);
    if (hasZip64Marker(total)) throw new Error("слишком большой PowerPoint-шаблон");
    const map = new Map();
    let offset = readU32(bytes, end + 16);
    const decoder = new TextDecoder("utf-8");

    for (let index = 0; index < total; index += 1) {
      if (readU32(bytes, offset) !== 0x02014b50) break;
      const method = readU16(bytes, offset + 10);
      const compressed = readU32(bytes, offset + 20);
      const uncompressed = readU32(bytes, offset + 24);
      const nameLength = readU16(bytes, offset + 28);
      const extraLength = readU16(bytes, offset + 30);
      const commentLength = readU16(bytes, offset + 32);
      const localOffset = readU32(bytes, offset + 42);
      const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));

      if (hasZip64Marker(compressed) || hasZip64Marker(uncompressed) || hasZip64Marker(localOffset)) {
        throw new Error("PowerPoint-шаблон использует неподдерживаемый ZIP64");
      }
      if (uncompressed > MAX_ENTRY_BYTES) throw new Error("элемент шаблона слишком большой");
      if (readU32(bytes, localOffset) !== 0x04034b50) throw new Error("повреждённый PowerPoint-шаблон");
      const localName = readU16(bytes, localOffset + 26);
      const localExtra = readU16(bytes, localOffset + 28);
      const start = localOffset + 30 + localName + localExtra;

      if (!map.has(name)) {
        map.set(
          name,
          await inflateEntry(bytes, {
            method: method,
            start: start,
            compressed: compressed,
          })
        );
      }

      offset += 46 + nameLength + extraLength + commentLength;
    }

    if (!map.has("[Content_Types].xml")) throw new Error("в шаблоне нет описания содержимого");
    return { map: map };
  }

  let crcTable = null;

  function crc32(bytes) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (let index = 0; index < 256; index += 1) {
        let value = index;
        for (let bit = 0; bit < 8; bit += 1) {
          value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
        }
        crcTable[index] = value >>> 0;
      }
    }
    let crc = 0xffffffff;
    for (let index = 0; index < bytes.length; index += 1) {
      crc = crcTable[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function writeZip(entries) {
    const encoder = new TextEncoder();
    const dosTime = 0;
    const dosDate = 0x5821;
    const files = entries.map(function (entry) {
      const name = encoder.encode(entry[0]);
      const data = entry[1];
      return {
        name: name,
        data: data,
        crc: crc32(data),
        localSize: 30 + name.length + data.length,
        centralSize: 46 + name.length,
      };
    });
    let localOffset = 0;
    let centralOffset = 0;
    files.forEach(function (file) {
      file.offset = localOffset;
      localOffset += file.localSize;
      centralOffset += file.centralSize;
    });

    const bytes = new Uint8Array(localOffset + centralOffset + 22);
    const view = new DataView(bytes.buffer);
    let cursor = 0;

    files.forEach(function (file) {
      setU32(view, cursor, 0x04034b50);
      setU16(view, cursor + 4, 20);
      setU16(view, cursor + 8, 0);
      setU16(view, cursor + 10, dosTime);
      setU16(view, cursor + 12, dosDate);
      setU32(view, cursor + 14, file.crc);
      setU32(view, cursor + 18, file.data.length);
      setU32(view, cursor + 22, file.data.length);
      setU16(view, cursor + 26, file.name.length);
      bytes.set(file.name, cursor + 30);
      bytes.set(file.data, cursor + 30 + file.name.length);
      cursor += file.localSize;
    });

    const centralStart = cursor;
    files.forEach(function (file) {
      setU32(view, cursor, 0x02014b50);
      setU16(view, cursor + 4, 20);
      setU16(view, cursor + 6, 20);
      setU16(view, cursor + 10, 0);
      setU16(view, cursor + 12, dosTime);
      setU16(view, cursor + 14, dosDate);
      setU32(view, cursor + 16, file.crc);
      setU32(view, cursor + 20, file.data.length);
      setU32(view, cursor + 24, file.data.length);
      setU16(view, cursor + 28, file.name.length);
      setU32(view, cursor + 42, file.offset);
      bytes.set(file.name, cursor + 46);
      cursor += file.centralSize;
    });

    setU32(view, cursor, 0x06054b50);
    setU16(view, cursor + 8, files.length);
    setU16(view, cursor + 10, files.length);
    setU32(view, cursor + 12, cursor - centralStart);
    setU32(view, cursor + 16, centralStart);
    return bytes;
  }

  function decodeText(value) {
    return new TextDecoder("utf-8").decode(value);
  }

  function encodeText(value) {
    return new TextEncoder().encode(value);
  }

  function escapeXml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function directoryOf(path) {
    const slash = path.lastIndexOf("/");
    return slash < 0 ? "" : path.slice(0, slash + 1);
  }

  function relationshipPath(partName) {
    const directory = directoryOf(partName);
    const fileName = partName.slice(directory.length);
    return directory + "_rels/" + fileName + ".rels";
  }

  function normalizePath(path) {
    const parts = path.split("/");
    const output = [];
    parts.forEach(function (part) {
      if (!part || part === ".") return;
      if (part === "..") output.pop();
      else output.push(part);
    });
    return output.join("/");
  }

  function resolvePath(sourceDirectory, target) {
    if (/^[a-zA-Z][a-zA-Z\d+\-.]*:/.test(target)) return "";
    return normalizePath(sourceDirectory + target.replace(/\\/g, "/"));
  }

  function attribute(xml, name) {
    const match = new RegExp(name + '\\s*=\\s*"([^"]*)"', "i").exec(xml || "");
    return match ? match[1] : "";
  }

  function parseRelationships(xml) {
    if (!xml) return [];
    const result = [];
    const pattern = /<Relationship\b[\s\S]*?\/>/g;
    let match;
    while ((match = pattern.exec(xml))) {
      const id = attribute(match[0], "Id");
      const type = attribute(match[0], "Type");
      const target = attribute(match[0], "Target");
      const mode = attribute(match[0], "TargetMode");
      if (id && type && target && mode !== "External") result.push({ id: id, type: type, target: target });
    }
    return result;
  }

  function relationshipsFor(entries, partName) {
    const relName = relationshipPath(partName);
    const xml = entries.map.get(relName);
    return xml ? parseRelationships(decodeText(xml)) : [];
  }

  function preservePart(entries, partName, packageMap) {
    if (!partName || packageMap.has(partName)) return;
    const bytes = entries.map.get(partName);
    if (!bytes) return;
    packageMap.set(partName, bytes);
    const relName = relationshipPath(partName);
    const relBytes = entries.map.get(relName);
    const relationships = relBytes ? parseRelationships(decodeText(relBytes)) : [];
    if (relBytes) packageMap.set(relName, relBytes);
    relationships.forEach(function (relationship) {
      if (
        relationship.type === REL_TYPES.slide ||
        relationship.type === REL_TYPES.notesSlide
      )
        return;
      const target = resolvePath(directoryOf(partName), relationship.target);
      if (target.startsWith("ppt/")) preservePart(entries, target, packageMap);
    });
  }

  function presentationRels(entries) {
    const xml = entries.map.get("ppt/_rels/presentation.xml.rels");
    return xml ? parseRelationships(decodeText(xml)) : [];
  }

  function keptPresentationRels(entries) {
    return presentationRels(entries).filter(function (relationship) {
      return (
        relationship.type !== REL_TYPES.slide &&
        relationship.type !== REL_TYPES.notesSlide
      );
    });
  }

  function nextRelationshipId(relationships) {
    let largest = 0;
    relationships.forEach(function (relationship) {
      const match = /^rId(\d+)$/i.exec(relationship.id || "");
      if (match) largest = Math.max(largest, Number(match[1]));
    });
    return largest + 1;
  }

  function placeholderSlots(layoutXml) {
    const slots = { title: "", bodyIdx: "", numberIdx: "" };
    const pattern = /<p:ph\b[^>]*\/?>/g;
    let match;
    while ((match = pattern.exec(layoutXml))) {
      const tag = match[0];
      const type = attribute(tag, "type");
      const idx = attribute(tag, "idx");
      if (type === "title" || type === "ctrTitle") {
        if (!slots.title) slots.title = type;
      } else if (type === "sldNum") {
        if (!slots.numberIdx) slots.numberIdx = idx;
      } else if (!type || type === "body") {
        if (!slots.bodyIdx) slots.bodyIdx = idx;
      }
    }
    if (!slots.title && /<p:ph\b[^>]*type="title"/.test(layoutXml)) slots.title = "title";
    return slots;
  }

  function layoutCandidates(entries) {
    const candidates = [];
    const slideRelationship = presentationRels(entries).find(function (relationship) {
      return relationship.type === REL_TYPES.slide;
    });
    if (slideRelationship) {
      const slide = resolvePath("ppt/", slideRelationship.target);
      const layout = relationshipsFor(entries, slide).find(function (relationship) {
        return relationship.type === REL_TYPES.slideLayout;
      });
      if (layout) candidates.push(resolvePath(directoryOf(slide), layout.target));
    }
    presentationRels(entries).forEach(function (relationship) {
      if (relationship.type !== REL_TYPES.slideMaster) return;
      const master = resolvePath("ppt/", relationship.target);
      relationshipsFor(entries, master).forEach(function (inner) {
        if (inner.type !== REL_TYPES.slideLayout) return;
        const part = resolvePath(directoryOf(master), inner.target);
        if (candidates.indexOf(part) < 0) candidates.push(part);
      });
    });
    return candidates.filter(function (part) {
      return entries.map.has(part);
    });
  }

  function chooseLayout(entries) {
    const candidates = layoutCandidates(entries);
    let best = null;
    candidates.forEach(function (part) {
      const slots = placeholderSlots(decodeText(entries.map.get(part)));
      const score = (slots.bodyIdx ? 2 : 0) + (slots.title ? 1 : 0);
      if (!best || score > best.score) best = { part: part, slots: slots, score: score };
    });
    return best;
  }

  function paragraph(text, level, points) {
    const value = String(text || "").trim();
    if (!value) return "";
    // Макет МФЮА даёт первый уровень со знаком «•»: абзацам маркер отключаем,
    // пункты остаются на втором уровне со знаком «–».
    const properties =
      level > 0 ? '<a:pPr lvl="' + level + '"/>' : "<a:pPr><a:buNone/></a:pPr>";
    return (
      "<a:p>" +
      properties +
      '<a:r><a:rPr lang="ru-RU" sz="' +
      Math.round(points * 100) +
      '" dirty="0"/><a:t>' +
      escapeXml(value) +
      "</a:t></a:r></a:p>"
    );
  }

  function xfrm(frame) {
    if (!frame) return "";
    return (
      "<a:xfrm><a:off x=\"" +
      Math.round(frame.x) +
      "\" y=\"" +
      Math.round(frame.y) +
      '"/><a:ext cx="' +
      Math.round(frame.w) +
      "\" cy=\"" +
      Math.round(frame.h) +
      '"/></a:xfrm>'
    );
  }

  function slideSize(presentationXml) {
    const tag = /<p:sldSz\b[^>]*>/.exec(presentationXml);
    if (!tag) return null;
    const cx = /cx="(\d+)"/.exec(tag[0]);
    const cy = /cy="(\d+)"/.exec(tag[0]);
    if (!cx || !cy) return null;
    const w = Number(cx[1]);
    const h = Number(cy[1]);
    if (!w || !h) return null;
    return { w: w, h: h };
  }

  const MASTER_FONT_FALLBACK = { titlePt: 44, bodyPt: 32, subPt: 28 };

  function styleSize(block) {
    if (!block) return 0;
    const match = /<a:defRPr[^>]*sz="(\d+)"/.exec(block);
    return match ? Number(match[1]) / 100 : 0;
  }

  // Кегль из p:txStyles мастера: нужен, чтобы предпросмотр показывал те же
  // размеры шрифта, что и PowerPoint в исходном шаблоне.
  function masterFontSizes(masterXml) {
    const title = /<p:titleStyle>([\s\S]*?)<\/p:titleStyle>/.exec(masterXml);
    const body = /<p:bodyStyle>([\s\S]*?)<\/p:bodyStyle>/.exec(masterXml);
    let titlePt = 0;
    let bodyPt = 0;
    let subPt = 0;
    if (title) {
      const level = /<a:lvl1pPr[\s\S]*?<\/a:lvl1pPr>/.exec(title[1]);
      titlePt = styleSize(level && level[0]);
    }
    if (body) {
      const inner = body[1];
      const first = /<a:lvl1pPr[\s\S]*?<\/a:lvl1pPr>/.exec(inner);
      const second = /<a:lvl2pPr[\s\S]*?<\/a:lvl2pPr>/.exec(inner);
      bodyPt = styleSize(first && first[0]);
      subPt = styleSize(second && second[0]);
    }
    return {
      titlePt: titlePt || MASTER_FONT_FALLBACK.titlePt,
      bodyPt: bodyPt || MASTER_FONT_FALLBACK.bodyPt,
      subPt: subPt || bodyPt || MASTER_FONT_FALLBACK.subPt,
    };
  }

  // Доли слайда для текстовых рамок: ниже синей шапки МФЮА, выше жёлтого
  // квадрата с номером. Предпросмотр берёт эти же значения из previewAssets.
  const CONTENT_BOXES = {
    title: { top: 0.15, height: 0.14, side: 0.05 },
    body: { top: 0.31, height: 0.54, side: 0.05 },
  };

  // Итоговая страница всегда хранит место под снимок: вертикальный кадр 3:4
  // по центру. Без фотографии остаётся пунктирная рамка с подписью.
  const PHOTO_BOX = { top: 0.30, height: 0.50, ratio: 0.75 };

  function photoFrame(size) {
    if (!size) return null;
    const height = size.h * PHOTO_BOX.height;
    const width = Math.min(height * PHOTO_BOX.ratio, size.w * 0.46);
    return { x: (size.w - width) / 2, y: size.h * PHOTO_BOX.top, w: width, h: height };
  }

  function contentFrames(presentationXml) {
    const size = slideSize(presentationXml);
    if (!size) return null;
    return {
      title: frameOf(CONTENT_BOXES.title, size),
      body: frameOf(CONTENT_BOXES.body, size),
      photo: photoFrame(size),
    };
  }

  function dataUrlBytes(dataUrl) {
    const value = String(dataUrl || "");
    const comma = value.indexOf(",");
    if (comma < 0 || !/;base64/i.test(value.slice(0, comma))) return null;
    const header = value.slice(0, comma);
    let extension = "jpeg";
    if (/image\/png/i.test(header)) extension = "png";
    else if (!/image\/j/i.test(header)) return null;
    const body = value.slice(comma + 1).replace(/\s+/g, "");
    let bytes = null;
    if (typeof global.Buffer === "function") {
      bytes = new Uint8Array(global.Buffer.from(body, "base64"));
    } else if (typeof global.atob === "function") {
      const binary = global.atob(body);
      bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
    }
    if (!bytes || !bytes.length || bytes.length > MAX_ENTRY_BYTES) return null;
    return { extension: extension, bytes: bytes };
  }

  function frameOf(box, size) {
    const x = size.w * box.side;
    return {
      x: x,
      y: size.h * box.top,
      w: size.w - x * 2,
      h: size.h * box.height,
    };
  }

  const TEXT_METRICS = {
    charWidth: 0.5,
    leading: 1.2,
    paraSpacing: 0.2,
    subIndent: 1.4,
    minScale: 0.3,
    step: 0.05,
  };

  function boxPoints(frame) {
    return frame ? { w: frame.w / 12700, h: frame.h / 12700 } : null;
  }

  function blockHeight(lines, box, sizeOf, scale) {
    return lines.reduce(function (sum, line, index) {
      const points = sizeOf(line.level) * scale;
      const width = box.w - (line.level ? points * TEXT_METRICS.subIndent : 0);
      const perLine = Math.max(1, Math.floor(width / (TEXT_METRICS.charWidth * points)));
      const wrapped = Math.max(1, Math.ceil(String(line.text).length / perLine));
      return (
        sum +
        wrapped * TEXT_METRICS.leading * points +
        (index ? TEXT_METRICS.paraSpacing * points : 0)
      );
    }, 0);
  }

  // PowerPoint пересчитывает normAutofit только при правке текста, поэтому
  // кегль уменьшаем сами и записываем его в запуск.
  function fitScale(lines, box, sizeOf) {
    if (!box || !lines.length) return 1;
    let scale = 1;
    while (scale > TEXT_METRICS.minScale && blockHeight(lines, box, sizeOf, scale) > box.h) {
      scale -= TEXT_METRICS.step;
    }
    return Math.max(TEXT_METRICS.minScale, scale);
  }

  function layoutFonts(entries, layoutPart) {
    const masterRel = relationshipsFor(entries, layoutPart).filter(function (relationship) {
      return relationship.type === REL_TYPES.slideMaster;
    })[0];
    if (!masterRel) return masterFontSizes("");
    const masterPart = resolvePath(directoryOf(layoutPart), masterRel.target);
    const masterXml = entries.map.get(masterPart);
    return masterFontSizes(masterXml ? decodeText(masterXml) : "");
  }

  function titlePlaceholder(title, slots, frame, fonts) {
    const box = boxPoints(frame);
    const scale = fitScale([{ level: 0, text: title }], box, function () {
      return fonts.titlePt;
    });
    return (
      "<p:sp>" +
      '<p:nvSpPr><p:cNvPr id="2" name="Заголовок"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="' +
      (slots.title || "title") +
      '"/></p:nvPr></p:nvSpPr><p:spPr>' +
      xfrm(frame) +
      "</p:spPr>" +
      "<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>" +
      paragraph(title, 0, fonts.titlePt * scale) +
      "</p:txBody></p:sp>"
    );
  }

  function bodyPlaceholder(lines, slots, frame, fonts) {
    const filled = lines.filter(function (line) {
      return String(line.text || "").trim();
    });
    if (!filled.length) return "";
    const sizeOf = function (level) {
      return level ? fonts.subPt : fonts.bodyPt;
    };
    const scale = fitScale(filled, boxPoints(frame), sizeOf);
    const content = filled
      .map(function (line) {
        return paragraph(line.text, line.level, sizeOf(line.level) * scale);
      })
      .join("");
    return (
      "<p:sp>" +
      '<p:nvSpPr><p:cNvPr id="3" name="Содержимое"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph idx="' +
      (slots.bodyIdx || "1") +
      '"/></p:nvPr></p:nvSpPr><p:spPr>' +
      xfrm(frame) +
      "</p:spPr>" +
      "<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>" +
      content +
      "</p:txBody></p:sp>"
    );
  }

  function numberPlaceholder(position, slots) {
    if (!slots.numberIdx) return "";
    return (
      "<p:sp>" +
      '<p:nvSpPr><p:cNvPr id="4" name="Номер слайда"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="sldNum" sz="quarter" idx="' +
      slots.numberIdx +
      '"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/>" +
      '<a:p><a:pPr algn="ctr"/><a:fld id="{17793251-9A28-40A8-A22D-0F67443F6B26}" type="slidenum">' +
      '<a:rPr lang="ru-RU" smtClean="0"/>' +
      '<a:pPr algn="ctr"/><a:t>' +
      position +
      '</a:t></a:fld><a:endParaRPr lang="ru-RU"/></a:p>' +
      "</p:txBody></p:sp>"
    );
  }

  function centeredRun(text, points, bold) {
    return (
      '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ru-RU" sz="' +
      Math.round(points * 100) +
      '"' +
      (bold ? ' b="1"' : "") +
      ' dirty="0"><a:solidFill><a:srgbClr val="5A6B7D"/></a:solidFill></a:rPr><a:t>' +
      escapeXml(text) +
      "</a:t></a:r></a:p>"
    );
  }

  // Слот под снимок на итоговой странице: настоящая картинка, если она
  // приложена, иначе пунктирная рамка — место, куда пользователь вставит своё.
  function photoShape(frame, imageExtension) {
    if (!frame) return "";
    const points = frame.h / 12700;
    if (imageExtension) {
      return (
        "<p:pic>" +
        '<p:nvPicPr><p:cNvPr id="5" name="Фотография"/><p:cNvPicPr>' +
        '<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
        '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
        "<p:spPr>" +
        xfrm(frame) +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
      );
    }
    return (
      "<p:sp>" +
      '<p:nvSpPr><p:cNvPr id="5" name="Место для фотографии"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      "<p:spPr>" +
      xfrm(frame) +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln w="12700" cap="flat"><a:solidFill><a:srgbClr val="9AA9BB"/></a:solidFill>' +
      '<a:prstDash val="dash"/><a:round/></a:ln></p:spPr>' +
      '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>' +
      centeredRun("ФОТО", points * 0.075, true) +
      centeredRun("место для изображения", points * 0.045, false) +
      "</p:txBody></p:sp>"
    );
  }

  function slideXml(title, lines, position, slots, frames, fonts, options) {
    const photo = options || {};
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      titlePlaceholder(title, slots, frames && frames.title, fonts) +
      bodyPlaceholder(lines, slots, frames && frames.body, fonts) +
      (photo.slot ? photoShape(frames && frames.photo, photo.extension) : "") +
      numberPlaceholder(position, slots) +
      "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    );
  }

  function slideRelsXml(layoutPart, imageExtension) {
    const target =
      "../" +
      layoutPart
        .split("/")
        .slice(1)
        .join("/");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' +
      RELATIONSHIP_NS +
      '"><Relationship Id="rId1" Type="' +
      REL_TYPES.slideLayout +
      '" Target="' +
      escapeXml(target) +
      '"/>' +
      (imageExtension
        ? '<Relationship Id="rId2" Type="' +
          REL_TYPES.image +
          '" Target="../media/' +
          MEDIA_PHOTO_NAME +
          '.' +
          imageExtension +
          '"/>'
        : "") +
      "</Relationships>"
    );
  }

  function sectionLines(section) {
    const lines = [];
    section.paragraphs.forEach(function (text) {
      lines.push({ level: 0, text: text });
    });
    section.bullets.forEach(function (text) {
      lines.push({ level: 1, text: text });
    });
    if (section.table) {
      section.table.headers.forEach(function (header, index) {
        const value = section.table.rows.map(function (row) {
          return row[index] || "";
        });
        lines.push({
          level: 1,
          text: [header].concat(value).filter(Boolean).join(" — "),
        });
      });
    }
    return lines;
  }

  function documentSlides(doc) {
    const cover = [doc.subtitle, doc.date].filter(Boolean).map(function (text) {
      return { level: 0, text: text };
    });
    const slides = [{ title: doc.title, lines: cover }];
    doc.sections.forEach(function (section) {
      const lines = sectionLines(section);
      if (section.paragraphs.length || section.bullets.length || lines.length) {
        slides.push({ title: section.heading, lines: lines });
      }
    });
    slides.push({ title: doc.title, lines: [], photo: true });
    return slides;
  }

  function buildPresentationXml(original, slideIds) {
    const slideList =
      "<p:sldIdLst>" +
      slideIds
        .map(function (item) {
          return '<p:sldId id="' + item.id + '" r:id="' + item.relId + '"/>';
        })
        .join("") +
      "</p:sldIdLst>";
    if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(original)) {
      return original.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, slideList);
    }
    const size = original.indexOf("<p:sldSz");
    if (size > 0) return original.slice(0, size) + slideList + original.slice(size);
    return original.replace("</p:presentation>", slideList + "</p:presentation>");
  }

  function relationshipsXml(relationships) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' +
      RELATIONSHIP_NS +
      '">' +
      relationships
        .map(function (relationship) {
          return (
            '<Relationship Id="' +
            escapeXml(relationship.id) +
            '" Type="' +
            escapeXml(relationship.type) +
            '" Target="' +
            escapeXml(relationship.target) +
            '"/>'
          );
        })
        .join("") +
      "</Relationships>"
    );
  }

  function buildRootRelsXml(original, packageMap) {
    const kept = parseRelationships(original).filter(function (relationship) {
      if (relationship.type === REL_TYPES.officeDocument) return true;
      return packageMap.has(resolvePath("", relationship.target));
    });
    return relationshipsXml(kept);
  }

  function ensureDefaultContentTypes(xml, extraDefaults) {
    const defaults = [
      ["xml", "application/xml"],
      ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ].concat(extraDefaults || []);
    const opening = /<Types\b[^>]*>/;
    defaults.forEach(function (item) {
      const pattern = new RegExp('<Default[^>]*Extension="' + item[0] + '"', "i");
      if (pattern.test(xml)) return;
      xml = xml.replace(
        opening,
        function (match) {
          return match + '<Default Extension="' + item[0] + '" ContentType="' + item[1] + '"/>';
        }
      );
    });
    return xml;
  }

  function buildContentTypesXml(original, packageMap, slideCount, imageExtension) {
    let xml = original.replace(/<Override\b[\s\S]*?\/>/g, function (tag) {
      const name = attribute(tag, "PartName").replace(/^\//, "");
      if (SLIDE_REL_PATTERN.test(name)) return "";
      return packageMap.has(name) ? tag : "";
    });
    const overrides = Array.from({ length: slideCount }, function (_, index) {
      return (
        '<Override PartName="/ppt/slides/slide' +
        (index + 1) +
        '.xml" ContentType="' +
        SLIDE_CONTENT_TYPE +
        '"/>'
      );
    }).join("");
    xml = xml.replace("</Types>", overrides + "</Types>");
    const image = imageExtension && IMAGE_CONTENT_TYPE[imageExtension];
    return ensureDefaultContentTypes(xml, image ? [[imageExtension, image]] : []);
  }

  // Разбор ZIP дорогой (распаковка всех частей + base64 фона), а шаблон один и
  // тот же на все генерации, поэтому держим последний распактованный пакет.
  let zipCache = null;

  async function readTemplateZip(source) {
    if (zipCache && zipCache.source === source) return zipCache.value;
    const entries = await readZip(source);
    zipCache = { source: source, value: entries };
    return entries;
  }

  async function render(templateBytes, doc) {
    if (!templateBytes) throw new Error("PowerPoint-шаблон не загружен");
    const source =
      templateBytes instanceof Uint8Array
        ? templateBytes
        : new Uint8Array(templateBytes);
    const entries = await readTemplateZip(source);
    const presentationXml = entries.map.get("ppt/presentation.xml");
    const contentTypesXml = entries.map.get("[Content_Types].xml");
    const presentationRelsXml = entries.map.get("ppt/_rels/presentation.xml.rels");
    const rootRelsXml = entries.map.get("_rels/.rels");
    if (!presentationXml || !contentTypesXml || !presentationRelsXml || !rootRelsXml) {
      throw new Error("PowerPoint-шаблон неполный");
    }

    const layout = chooseLayout(entries);
    if (!layout || !layout.slots.bodyIdx) {
      throw new Error("в шаблоне нет содержательного макета слайда");
    }
    const layoutPart = layout.part;
    const slots = layout.slots;
    const frames = contentFrames(decodeText(presentationXml));
    const fonts = layoutFonts(entries, layoutPart);

    const packageMap = new Map();
    const carriedRels = keptPresentationRels(entries);
    carriedRels.forEach(function (relationship) {
      const target = resolvePath("ppt/", relationship.target);
      if (target.startsWith("ppt/")) preservePart(entries, target, packageMap);
    });
    preservePart(entries, layoutPart, packageMap);

    const firstRelId = nextRelationshipId(carriedRels);
    const slides = documentSlides(doc);
    const photo = frames && frames.photo ? dataUrlBytes(doc.photo && doc.photo.data) : null;
    if (photo) {
      packageMap.set("ppt/media/" + MEDIA_PHOTO_NAME + "." + photo.extension, photo.bytes);
    }
    const slideIds = slides.map(function (slide, index) {
      const name = "ppt/slides/slide" + (index + 1) + ".xml";
      const relId = "rId" + (firstRelId + index);
      const extension = slide.photo && photo ? photo.extension : "";
      packageMap.set(
        name,
        encodeText(
          slideXml(slide.title, slide.lines, index + 1, slots, frames, fonts, {
            slot: Boolean(slide.photo),
            extension: extension,
          })
        )
      );
      packageMap.set(relationshipPath(name), encodeText(slideRelsXml(layoutPart, extension)));
      return { id: 256 + index, relId: relId };
    });

    packageMap.set(
      "ppt/presentation.xml",
      encodeText(buildPresentationXml(decodeText(presentationXml), slideIds))
    );
    packageMap.set(
      "ppt/_rels/presentation.xml.rels",
      encodeText(
        relationshipsXml(
          carriedRels.concat(
            slides.map(function (_, index) {
              return {
                id: "rId" + (firstRelId + index),
                type: REL_TYPES.slide,
                target: "slides/slide" + (index + 1) + ".xml",
              };
            })
          )
        )
      )
    );
    packageMap.set(
      "_rels/.rels",
      encodeText(buildRootRelsXml(decodeText(rootRelsXml), packageMap))
    );

    const partsBeforeSlides = new Map(packageMap);
    packageMap.set(
      "[Content_Types].xml",
      encodeText(
        buildContentTypesXml(
          decodeText(contentTypesXml),
          partsBeforeSlides,
          slides.length,
          photo ? photo.extension : ""
        )
      )
    );

    const ordered = [];
    ordered.push(["[Content_Types].xml", packageMap.get("[Content_Types].xml")]);
    ordered.push(["_rels/.rels", packageMap.get("_rels/.rels")]);
    packageMap.forEach(function (value, name) {
      if (name !== "[Content_Types].xml" && name !== "_rels/.rels") {
        ordered.push([name, value]);
      }
    });

    if (typeof global.Blob === "function") {
      return new global.Blob([writeZip(ordered)], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });
    }
    return writeZip(ordered);
  }

  async function loadBuiltin() {
    if (!builtinPromise) {
      builtinPromise = fetch(BUILTIN.mfua)
        .then(function (response) {
          if (!response.ok) throw new Error("Фирменный шаблон недоступен.");
          return response.arrayBuffer();
        })
        .catch(function (error) {
          builtinPromise = null;
          throw error;
        });
    }
    return builtinPromise;
  }

  async function selectedBytes(format, themeId, customBytes) {
    if (format !== "pptx") return null;
    if (themeId === "custom") return customBytes || null;
    if (themeId === "mfua") return loadBuiltin();
    return null;
  }

  function toDataUrl(bytes, partName) {
    const ext = (String(partName).split(".").pop() || "jpeg").toLowerCase();
    const mime =
      ext === "png"
        ? "image/png"
        : ext === "gif"
          ? "image/gif"
          : ext === "svg"
            ? "image/svg+xml"
            : "image/jpeg";
    if (typeof global.Buffer === "function") {
      return "data:" + mime + ";base64," + global.Buffer.from(bytes).toString("base64");
    }
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return "data:" + mime + ";base64," + global.btoa(binary);
  }

  async function previewAssets(templateBytes) {
    if (!templateBytes) return null;
    const source =
      templateBytes instanceof Uint8Array
        ? templateBytes
        : new Uint8Array(templateBytes);
    let entries;
    try {
      entries = await readTemplateZip(source);
    } catch (error) {
      return null;
    }
    const presentationXml = entries.map.get("ppt/presentation.xml");
    if (!presentationXml) return null;
    const size = slideSize(decodeText(presentationXml));
    const base = {
      aspect: size ? size.w / size.h : 16 / 9,
      widthPt: size ? size.w / 12700 : 960,
      backgroundImage: null,
      boxes: CONTENT_BOXES,
      photoBox: PHOTO_BOX,
      fonts: masterFontSizes(""),
    };
    const masterRel = presentationRels(entries).filter(function (relationship) {
      return relationship.type === REL_TYPES.slideMaster;
    })[0];
    const masterPart = masterRel ? resolvePath("ppt/", masterRel.target) : "";
    const masterXml = masterPart ? entries.map.get(masterPart) : null;
    if (!masterXml) return base;

    base.fonts = masterFontSizes(decodeText(masterXml));
    const bgMatch = /<p:bg>[\s\S]*?<a:blip[^>]*r:embed="([^"]+)"/.exec(decodeText(masterXml));
    if (!bgMatch) return base;
    const imageRel = relationshipsFor(entries, masterPart).filter(function (relationship) {
      return relationship.id === bgMatch[1];
    })[0];
    if (!imageRel) return base;
    const imagePart = resolvePath(directoryOf(masterPart), imageRel.target);
    const imageBytes = entries.map.get(imagePart);
    if (!imageBytes) return base;
    base.backgroundImage = toDataUrl(imageBytes, imagePart);
    return base;
  }

  Object.assign(Forma, {
    pptTemplates: {
      render: render,
      fitScale: fitScale,
      selectedBytes: selectedBytes,
      previewAssets: previewAssets,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);