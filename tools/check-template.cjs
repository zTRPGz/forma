"use strict";

const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const root = path.join(__dirname, "..");
require(path.join(root, "js", "ppt-template.js"));

const templatePath = path.join(root, "vendor", "mfua-template.pptx");
const outPath = path.join(process.env.TEMP || "/tmp", "forma-template-check.pptx");

const doc = {
  title: "План запуска продукта <на новом рынке>",
  subtitle: "Обсуждение команды & сроков",
  date: "31.10.2026",
  sections: [
    {
      heading: "Цель",
      paragraphs: ["Выйти на рынок за один квартал."],
      bullets: ["Охватить три региона", "Обучить партнёров"],
      table: null,
    },
    {
      heading: "Риски",
      paragraphs: [],
      bullets: ["Задержка поставки", "Дефицит кадров"],
      table: {
        headers: ["Риск", "Мера"],
        rows: [["Задержка", "Страховой запас"], ["Кадры", "Найм заранее"]],
      },
    },
  ],
};

function readZipEntries(buffer) {
  const end = buffer.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (end < 0) throw new Error("нет центрального каталога");
  const total = buffer.readUInt16LE(end + 10);
  let offset = buffer.readUInt32LE(end + 16);
  const result = new Map();
  for (let index = 0; index < total; index += 1) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) break;
    const method = buffer.readUInt16LE(offset + 10);
    const compressed = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + nameLength).toString("utf8");
    const localName = buffer.readUInt16LE(localOffset + 26);
    const localExtra = buffer.readUInt16LE(localOffset + 28);
    const start = localOffset + 30 + localName + localExtra;
    const raw = buffer.subarray(start, start + compressed);
    result.set(name, method === 0 ? raw : zlib.inflateRawSync(raw));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return result;
}

async function main() {
  const source = fs.readFileSync(templatePath);
  const blob = await Forma.pptTemplates.render(new Uint8Array(source), doc);
  const bytes = Buffer.from(await blob.arrayBuffer());
  fs.writeFileSync(outPath, bytes);

  const original = readZipEntries(source);
  const output = readZipEntries(bytes);
  const problems = [];

  for (const name of ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout2.xml", "ppt/theme/theme1.xml", "ppt/media/image1.jpeg", "ppt/presProps.xml", "ppt/tableStyles.xml"]) {
    if (!output.has(name)) problems.push("потеряна часть " + name);
    else if (!original.get(name).equals(output.get(name))) problems.push("изменена часть " + name);
  }

  const slides = [...output.keys()].filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  if (slides.length !== doc.sections.length + 2) {
    problems.push("ожидалось " + (doc.sections.length + 2) + " слайдов, получено " + slides.length);
  }

  const types = output.get("[Content_Types].xml").toString("utf8");
  const overrideNames = [...types.matchAll(/PartName="([^"]+)"/g)].map((m) => m[1]);
  const duplicates = overrideNames.filter((n, i) => overrideNames.indexOf(n) !== i);
  if (duplicates.length) problems.push("дубли Override: " + duplicates.join(", "));
  for (const name of overrideNames) {
    if (!output.has(name.replace(/^\//, ""))) problems.push("Override указывает на отсутствующую часть " + name);
  }
  for (const name of output.keys()) {
    if (name === "[Content_Types].xml" || name.endsWith(".rels")) continue;
    if (name.endsWith(".xml") && !overrideNames.includes("/" + name)) {
      problems.push("часть " + name + " без описания содержимого");
    }
  }

  const rels = output.get("ppt/_rels/presentation.xml.rels").toString("utf8");
  const pres = output.get("ppt/presentation.xml").toString("utf8");
  const referencedTypes = /slideMaster|\/slide"|notesMaster|handoutMaster/;
  for (const match of rels.matchAll(/<Relationship Id="([^"]+)" Type="[^"]*\/(\w+)" Target="([^"]+)"/g)) {
    const id = match[1];
    const type = match[2];
    const target = match[3].replace(/^\//, "");
    const full = target.startsWith("ppt/") ? target : "ppt/" + target;
    if (!output.has(full)) problems.push("ссылка " + id + " ведёт на отсутствующую часть " + full);
    if (referencedTypes.test(type) && !pres.includes('r:id="' + id + '"')) {
      problems.push("связь " + id + " (" + type + ") не перечислена в presentation.xml");
    }
  }
  for (const id of [...pres.matchAll(/r:id="([^"]+)"/g)].map((m) => m[1])) {
    if (!rels.includes('Id="' + id + '"')) problems.push("presentation.xml ссылается на отсутствующую связь " + id);
  }

  const cover = output.get("ppt/slides/slide1.xml").toString("utf8");
  if (!cover.includes("<a:t>План запуска продукта &lt;на новом рынке&gt;</a:t>")) {
    problems.push("заголовок документа не подставлен на титульный слайд");
  }
  const content = output.get("ppt/slides/slide2.xml").toString("utf8");
  for (const text of ["Цель", "Выйти на рынок за один квартал.", "Обучить партнёров"]) {
    if (!content.includes("<a:t>" + text + "</a:t>")) {
      problems.push("в слайде 2 нет текста «" + text + "»");
    }
  }
  if (!content.includes('lvl="1"')) problems.push("пункты списка потеряли уровень");
  const layout = output.get("ppt/slideLayouts/slideLayout2.xml").toString("utf8");
  if (!layout.includes('type="obj"')) problems.push("макет слайда изменился");

  console.log("файл: " + outPath + " (" + bytes.length + " байт, частей " + output.size + ")");
  if (problems.length) {
    problems.forEach((p) => console.log("ПРОБЛЕМА: " + p));
    process.exitCode = 1;
  } else {
    console.log("шаблон сохранён без изменений, текст подставлен");
  }
}

main().catch((error) => {
  console.error("СБОЙ: " + (error && error.stack ? error.stack : error));
  process.exitCode = 1;
});
