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
    notesMaster: OFFICE_REL_NS + "/notesMaster",
    handoutMaster: OFFICE_REL_NS + "/handoutMaster",
  };
  const SLIDE_CONTENT_TYPE =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
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
    const entries = [];
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
        const value = await inflateEntry(bytes, {
          method: method,
          start: start,
          compressed: compressed,
        });
        entries.push([name, value]);
        map.set(name, value);
      }

      offset += 46 + nameLength + extraLength + commentLength;
    }

    if (!map.has("[Content_Types].xml")) throw new Error("в шаблоне нет описания содержимого");
    return { entries: entries, map: map };
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

  function paragraph(text, level) {
    const value = String(text || "").trim();
    if (!value) return "";
    const properties = level > 0 ? '<a:pPr lvl="' + level + '"/>' : "";
    return (
      "<a:p>" +
      properties +
      '<a:r><a:rPr lang="ru-RU" dirty="0"/><a:t>' +
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

  function contentFrames(presentationXml) {
    const match = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentationXml);
    if (!match) return null;
    const w = Number(match[1]);
    const h = Number(match[2]);
    if (!w || !h) return null;
    const side = w * 0.05;
    const inner = w - side * 2;
    return {
      title: { x: side, y: h * 0.18, w: inner, h: h * 0.11 },
      body: { x: side, y: h * 0.32, w: inner, h: h * 0.52 },
    };
  }

  function titlePlaceholder(title, slots, frame) {
    return (
      "<p:sp>" +
      '<p:nvSpPr><p:cNvPr id="2" name="Заголовок"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="' +
      (slots.title || "title") +
      '"/></p:nvPr></p:nvSpPr><p:spPr>' +
      xfrm(frame) +
      "</p:spPr>" +
      "<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>" +
      paragraph(title, 0) +
      "</p:txBody></p:sp>"
    );
  }

  function bodyPlaceholder(lines, slots, frame) {
    const content = lines
      .map(function (line) {
        return paragraph(line.text, line.level);
      })
      .filter(Boolean)
      .join("");
    if (!content) return "";
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

  function slideXml(title, lines, position, slots, frames) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      titlePlaceholder(title, slots, frames && frames.title) +
      bodyPlaceholder(lines, slots, frames && frames.body) +
      numberPlaceholder(position, slots) +
      "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    );
  }

  function slideRelsXml(layoutPart) {
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
      '"/></Relationships>'
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
      if (section.table && !lines.length) {
        slides.push({ title: section.heading, lines: sectionLines(section) });
      }
    });
    slides.push({ title: doc.title, lines: [] });
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

  function ensureDefaultContentTypes(xml) {
    const defaults = [
      ["xml", "application/xml"],
      ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ];
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

  function buildContentTypesXml(original, packageMap, slideCount) {
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
    return ensureDefaultContentTypes(xml);
  }

  async function render(templateBytes, doc) {
    if (!templateBytes) throw new Error("PowerPoint-шаблон не загружен");
    const source =
      templateBytes instanceof Uint8Array
        ? templateBytes
        : new Uint8Array(templateBytes);
    const entries = await readZip(source);
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

    const packageMap = new Map();
    const carriedRels = keptPresentationRels(entries);
    carriedRels.forEach(function (relationship) {
      const target = resolvePath("ppt/", relationship.target);
      if (target.startsWith("ppt/")) preservePart(entries, target, packageMap);
    });
    preservePart(entries, layoutPart, packageMap);

    const firstRelId = nextRelationshipId(carriedRels);
    const slides = documentSlides(doc);
    const slideIds = slides.map(function (slide, index) {
      const name = "ppt/slides/slide" + (index + 1) + ".xml";
      const relId = "rId" + (firstRelId + index);
      packageMap.set(name, encodeText(slideXml(slide.title, slide.lines, index + 1, slots, frames)));
      packageMap.set(relationshipPath(name), encodeText(slideRelsXml(layoutPart)));
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
      encodeText(buildContentTypesXml(decodeText(contentTypesXml), partsBeforeSlides, slides.length))
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
      entries = await readZip(source);
    } catch (error) {
      return null;
    }
    const presentationXml = entries.map.get("ppt/presentation.xml");
    if (!presentationXml) return null;
    const sizeMatch = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(decodeText(presentationXml));
    const aspect = sizeMatch ? Number(sizeMatch[1]) / Number(sizeMatch[2]) : 16 / 9;
    const masterRel = presentationRels(entries).filter(function (relationship) {
      return relationship.type === REL_TYPES.slideMaster;
    })[0];
    if (!masterRel) return { aspect: aspect, backgroundImage: null };
    const masterPart = resolvePath("ppt/", masterRel.target);
    const masterXml = entries.map.get(masterPart);
    if (!masterXml) return { aspect: aspect, backgroundImage: null };
    const bgMatch = /<p:bg>[\s\S]*?<a:blip[^>]*r:embed="([^"]+)"/.exec(decodeText(masterXml));
    if (!bgMatch) return { aspect: aspect, backgroundImage: null };
    const imageRel = relationshipsFor(entries, masterPart).filter(function (relationship) {
      return relationship.id === bgMatch[1];
    })[0];
    if (!imageRel) return { aspect: aspect, backgroundImage: null };
    const imagePart = resolvePath(directoryOf(masterPart), imageRel.target);
    const imageBytes = entries.map.get(imagePart);
    if (!imageBytes) return { aspect: aspect, backgroundImage: null };
    return { aspect: aspect, backgroundImage: toDataUrl(imageBytes, imagePart) };
  }

  Object.assign(Forma, {
    pptTemplates: {
      render: render,
      loadBuiltin: loadBuiltin,
      selectedBytes: selectedBytes,
      previewAssets: previewAssets,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);