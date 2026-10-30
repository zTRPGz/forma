(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});

  const DOC = {
    width: 580,
    padX: 40,
    bodyLimit: 660,
    bodyChars: 70,
    headingChars: 44,
    bulletChars: 66,
    line: 20,
    bulletLine: 19,
  };

  // 1 pt = 1/960 ширины слайда (16:9, 13.333 x 7.5 дюйма)
  const PT = 100 / 960;

  function node(tag, className, text) {
    const element = document.createElement(tag);
    if (className) element.className = className;
    if (text != null) element.textContent = text;
    return element;
  }

  function metaLine(doc) {
    return [doc.author, doc.date].filter(Boolean).join(" · ");
  }

  function lineCount(text, perLine) {
    return Math.max(1, Math.ceil(String(text).length / perLine));
  }

  // ---------- DOCX ----------

  function columnWidths(table, innerWidth) {
    const weights = table.headers.map(function (_, index) {
      return index === 0 ? 1.25 : 1;
    });
    const sum = weights.reduce(function (total, weight) {
      return total + weight;
    }, 0);
    return weights.map(function (weight) {
      return (innerWidth * weight) / sum;
    });
  }

  function tableHeight(table, innerWidth) {
    const widths = columnWidths(table, innerWidth);
    const rows = [table.headers].concat(table.rows);
    let height = 12;
    rows.forEach(function (row) {
      let count = 1;
      row.forEach(function (cell, index) {
        const perLine = Math.max(5, Math.floor((widths[index] - 18) / 6.3));
        count = Math.max(count, lineCount(cell, perLine));
      });
      height += Math.max(29, count * 16 + 13);
    });
    return height;
  }

  function setWidths(element, table, innerWidth) {
    const widths = columnWidths(table, innerWidth);
    const sum = widths.reduce(function (total, width) {
      return total + width;
    }, 0);
    const group = node("colgroup");
    widths.forEach(function (width) {
      const col = node("col");
      col.style.width = ((width / sum) * 100).toFixed(2) + "%";
      group.append(col);
    });
    element.append(group);
  }

  function docTable(table, innerWidth) {
    const wrap = node("div", "doc-table-wrap");
    const element = node("table", "doc-table");
    setWidths(element, table, innerWidth);

    const head = node("thead");
    const headRow = node("tr");
    table.headers.forEach(function (header) {
      const cell = node("th", null, header);
      cell.scope = "col";
      headRow.append(cell);
    });
    head.append(headRow);
    element.append(head);

    const body = node("tbody");
    table.rows.forEach(function (row) {
      const tr = node("tr");
      row.forEach(function (value) {
        tr.append(node("td", null, value));
      });
      body.append(tr);
    });
    element.append(body);

    wrap.append(element);
    return wrap;
  }

  function docBlocks(doc, innerWidth) {
    const blocks = [];

    doc.sections.forEach(function (section) {
      blocks.push({
        heading: true,
        height: lineCount(section.heading, DOC.headingChars) * 25 + 36,
        make: function () {
          return node("p", "doc-heading", section.heading);
        },
      });

      section.paragraphs.forEach(function (text) {
        blocks.push({
          height: lineCount(text, DOC.bodyChars) * DOC.line + 12,
          make: function () {
            return node("p", "doc-text", text);
          },
        });
      });

      section.bullets.forEach(function (text) {
        blocks.push({
          height: lineCount(text, DOC.bulletChars) * DOC.bulletLine + 7,
          make: function () {
            return node("p", "doc-bullet", text);
          },
        });
      });

      if (section.table) {
        blocks.push({
          height: tableHeight(section.table, innerWidth),
          make: function () {
            return docTable(section.table, innerWidth);
          },
        });
      }
    });

    return blocks;
  }

  function paginate(blocks, limit, firstLimit) {
    const pages = [];
    let current = [];
    let used = 0;
    let capacity = firstLimit || limit;

    function flush() {
      if (current.length) pages.push(current);
      current = [];
      used = 0;
      capacity = limit;
    }

    blocks.forEach(function (block, index) {
      const next = blocks[index + 1];
      const nextHeight = next ? Math.min(next.height, 120) : 0;
      if (used > 0) {
        const orphansHeading = block.heading && used + block.height + nextHeight > capacity;
        if (orphansHeading || used + block.height > capacity) flush();
      }
      current.push(block);
      used += block.height;
    });

    flush();
    return pages;
  }

  function docPage(pageNumber, first) {
    const page = node("div", "doc-page" + (first ? " doc-page--first" : ""));
    const sheet = node("div", "doc-page-sheet");
    const body = node("div", "doc-page-body");
    page.append(sheet);

    if (first) {
      sheet.append(body);
      return { page: page, body: body };
    }

    const head = node("div", "doc-page-head");
    head.append(node("span", "doc-page-brand", "FORMA"));
    head.append(node("span", "doc-page-label", " · документ"));

    sheet.append(head, body, node("div", "doc-page-foot", String(pageNumber)));
    return { page: page, body: body };
  }

  function docx(doc) {
    const root = node("div", "doc-preview");
    const innerWidth = DOC.width - DOC.padX * 2;
    const meta = metaLine(doc);
    const titleHeight =
      lineCount(doc.title, 26) * 35 + (doc.subtitle ? 28 : 0) + (meta ? 56 : 20);
    const pages = paginate(
      docBlocks(doc, innerWidth),
      DOC.bodyLimit,
      Math.max(240, DOC.bodyLimit - titleHeight - 36)
    );

    const first = docPage(1, true);
    const block = node("div", "doc-title-block");
    block.append(node("p", "doc-title", doc.title));
    if (doc.subtitle) block.append(node("p", "doc-subtitle", doc.subtitle));
    if (meta) block.append(node("p", "doc-meta", meta));
    first.body.append(block);
    (pages[0] || []).forEach(function (item) {
      first.body.append(item.make());
    });
    root.append(first.page);

    pages.slice(1).forEach(function (blocks, index) {
      const page = docPage(index + 2, false);
      blocks.forEach(function (item) {
        page.body.append(item.make());
      });
      root.append(page.page);
    });

    return root;
  }

  // ---------- PDF ----------

  function pdfPage(pageNumber) {
    const page = node("div", "doc-page pdf-page");
    const sheet = node("div", "doc-page-sheet");
    const head = node("div", "doc-page-head");
    head.append(node("span", "doc-page-brand", "FORMA"));
    head.append(node("span", "doc-page-label", " · аналитическая записка"));

    const body = node("div", "doc-page-body");
    sheet.append(head, body, node("div", "doc-page-foot", String(pageNumber)));
    page.append(sheet);
    return { page: page, body: body };
  }

  function pdf(doc) {
    const root = node("div", "doc-preview pdf-preview");
    const innerWidth = DOC.width - DOC.padX * 2;
    const meta = metaLine(doc);
    const titleHeight =
      lineCount(doc.title, 26) * 33 + (doc.subtitle ? 28 : 0) + (meta ? 56 : 20);
    const pages = paginate(
      docBlocks(doc, innerWidth),
      DOC.bodyLimit,
      Math.max(240, DOC.bodyLimit - titleHeight)
    );

    const first = pdfPage(1);
    first.page.classList.add("pdf-page--first");
    const block = node("div", "doc-title-block");
    block.append(node("p", "doc-title", doc.title));
    if (doc.subtitle) block.append(node("p", "doc-subtitle", doc.subtitle));
    if (meta) block.append(node("p", "doc-meta", meta));
    block.append(node("div", "pdf-rule"));
    first.body.append(block);
    (pages[0] || []).forEach(function (item) {
      first.body.append(item.make());
    });
    root.append(first.page);

    pages.slice(1).forEach(function (blocks, index) {
      const page = pdfPage(index + 2);
      blocks.forEach(function (item) {
        page.body.append(item.make());
      });
      root.append(page.page);
    });

    return root;
  }

  // ---------- PPTX ----------

  function slideFooter(page) {
    const rule = node("div", "slide-foot-rule");
    const foot = node("div", "slide-foot");
    foot.append(node("span", "slide-foot-brand", "FORMA"));
    foot.append(node("span", "slide-foot-page", String(page).padStart(2, "0")));
    return [rule, foot];
  }

  function pptxCover(doc) {
    const slide = node("div", "slide slide--cover");
    slide.append(node("div", "slide-bar"));
    slide.append(node("div", "slide-brand", "FORMA"));

    const title = node("p", "slide-cover-title", doc.title);
    const length = doc.title.length;
    title.style.fontSize = (length <= 55 ? 40 : length <= 90 ? 32 : 27) * PT + "cqw";
    slide.append(title);

    slide.append(node("div", "slide-cover-rule"));

    if (doc.subtitle) slide.append(node("p", "slide-cover-subtitle", doc.subtitle));

    const meta = metaLine(doc);
    if (meta) slide.append(node("p", "slide-cover-meta", meta));

    return slide;
  }

  function pptxHeading(slide, text) {
    const heading = node("p", "slide-heading", text);
    const length = text.length;
    heading.style.fontSize = (length <= 42 ? 30 : length <= 70 ? 25 : 22) * PT + "cqw";
    slide.append(heading);
    slide.append(node("div", "slide-rule"));
  }

  function bodySize(total) {
    if (total <= 200) return 20;
    if (total <= 340) return 17;
    if (total <= 520) return 15;
    if (total <= 800) return 13;
    if (total <= 1200) return 12;
    if (total <= 1800) return 11;
    return 10;
  }

  function pptxContentSlide(section, page) {
    const slide = node("div", "slide slide--content");
    pptxHeading(slide, section.heading);

    const body = node("div", "slide-body");
    let total = 0;
    section.paragraphs.forEach(function (text) {
      total += text.length;
    });
    section.bullets.forEach(function (text) {
      total += text.length;
    });
    body.style.fontSize = bodySize(total) * PT + "cqw";

    section.paragraphs.forEach(function (text) {
      body.append(node("p", "slide-p", text));
    });

    if (section.bullets.length) {
      const list = node("ul", "slide-bullets");
      section.bullets.forEach(function (text) {
        list.append(node("li", null, text));
      });
      body.append(list);
    }

    slide.append(body);
    slideFooter(page).forEach(function (element) {
      slide.append(element);
    });
    return slide;
  }

  function pptxTableSlide(section, page) {
    const slide = node("div", "slide slide--content");
    pptxHeading(slide, section.heading);

    const table = section.table;
    let chars = table.headers.join(" ").length;
    table.rows.forEach(function (row) {
      chars += row.join(" ").length;
    });
    const size = chars <= 120 ? 11.5 : chars <= 260 ? 10.5 : chars <= 400 ? 9.5 : 8.5;

    const wrap = node("div", "slide-table-wrap");
    const element = node("table", "slide-table");
    element.style.fontSize = size * PT + "cqw";
    setWidths(element, table, 1);

    const head = node("thead");
    const headRow = node("tr");
    table.headers.forEach(function (header) {
      const cell = node("th", null, header);
      cell.scope = "col";
      headRow.append(cell);
    });
    head.append(headRow);
    element.append(head);

    const body = node("tbody");
    table.rows.forEach(function (row) {
      const tr = node("tr");
      row.forEach(function (value) {
        tr.append(node("td", null, value));
      });
      body.append(tr);
    });
    element.append(body);

    wrap.append(element);
    slide.append(wrap);
    slideFooter(page).forEach(function (item) {
      slide.append(item);
    });
    return slide;
  }

  function pptx(doc) {
    const root = node("div", "slides-preview");
    root.append(pptxCover(doc));

    let page = 2;
    doc.sections.forEach(function (section) {
      if (section.paragraphs.length || section.bullets.length) {
        root.append(pptxContentSlide(section, page));
        page += 1;
      }
      if (section.table) {
        root.append(pptxTableSlide(section, page));
        page += 1;
      }
    });

    return root;
  }

  Object.assign(Forma, {
    preview: {
      docx: docx,
      pdf: pdf,
      pptx: pptx,
    },
  });
})(typeof window !== "undefined" ? window : globalThis);
