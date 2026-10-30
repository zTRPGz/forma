(function (global) {
  "use strict";

  const Forma = (global.Forma = global.Forma || {});
  const isNode =
    typeof process !== "undefined" &&
    !!process.versions &&
    !!process.versions.node &&
    typeof window === "undefined";

  const INK = "1F2A37";
  const NAVY = "16243A";
  const ACCENT = "B45309";
  const MUTED = "6B5E50";
  const RULE = "D9CFC1";
  const WHITE = "FFFFFF";
  const ALT = "F3EDE4";
  const LIGHT_TEXT = "EADFD0";
  const COVER_META = "9AA7B8";
  const HEAD_FONT = "Georgia";
  const BODY_FONT = "Segoe UI";
  const BULLET_REF = "forma-bullets";

  const PPT_W = 13.333;
  const PPT_H = 7.5;
  const PPT_M = 0.9;
  const PPT_CONTENT_W = PPT_W - 2 * PPT_M;
  const PDF_CONTENT_W = 471;
  const PDF_INK = "#1F2A37";
  const PDF_NAVY = "#16243A";
  const PDF_ACCENT = "#B45309";
  const PDF_MUTED = "#6B5E50";
  const PDF_RULE = "#D9CFC1";
  const PDF_WHITE = "#FFFFFF";
  const PDF_ALT = "#F3EDE4";

  function metaLine(doc) {
    return [doc.author, doc.date].filter(Boolean).join(" · ");
  }

  // ---------- DOCX ----------

  function docxStyles(D) {
    return {
      default: {
        document: {
          run: { font: BODY_FONT, size: 22, color: INK },
          paragraph: { spacing: { line: 348, after: 160 } },
        },
      },
      paragraphStyles: [
        {
          id: "FormaTitle",
          name: "Forma Title",
          quickFormat: true,
          run: { font: HEAD_FONT, size: 60, bold: true, color: INK },
          paragraph: { spacing: { before: 960, after: 160, line: 252 }, keepNext: true },
        },
        {
          id: "FormaSubtitle",
          name: "Forma Subtitle",
          quickFormat: true,
          run: { font: BODY_FONT, size: 26, color: ACCENT },
          paragraph: { spacing: { before: 0, after: 560, line: 300 }, keepNext: true },
        },
        {
          id: "FormaMeta",
          name: "Forma Meta",
          quickFormat: true,
          run: { font: BODY_FONT, size: 19, color: MUTED },
          paragraph: { spacing: { before: 0, after: 80, line: 300 } },
        },
        {
          id: "FormaHeading",
          name: "Forma Heading",
          quickFormat: true,
          basedOn: "Normal",
          run: { font: HEAD_FONT, size: 32, bold: true, color: INK },
          paragraph: {
            spacing: { before: 400, after: 200, line: 276 },
            keepNext: true,
            border: {
              bottom: { style: D.BorderStyle.SINGLE, size: 8, space: 4, color: ACCENT },
            },
          },
        },
        {
          id: "FormaBody",
          name: "Forma Body",
          quickFormat: true,
          basedOn: "Normal",
          run: { font: BODY_FONT, size: 22, color: INK },
          paragraph: { spacing: { before: 0, after: 160, line: 348 } },
        },
        {
          id: "FormaBullet",
          name: "Forma Bullet",
          quickFormat: true,
          basedOn: "Normal",
          run: { font: BODY_FONT, size: 22, color: INK },
          paragraph: {
            spacing: { before: 0, after: 80, line: 324 },
            indent: { left: 720, hanging: 360 },
          },
        },
        {
          id: "FormaCell",
          name: "Forma Cell",
          quickFormat: true,
          basedOn: "Normal",
          run: { font: BODY_FONT, size: 19, color: INK },
          paragraph: { spacing: { before: 40, after: 40, line: 276 } },
        },
        {
          id: "FormaCellHead",
          name: "Forma Cell Head",
          quickFormat: true,
          basedOn: "FormaCell",
          run: { font: BODY_FONT, size: 19, bold: true, color: WHITE },
        },
      ],
    };
  }

  function docxNumbering(D) {
    return {
      config: [
        {
          reference: BULLET_REF,
          levels: [
            {
              level: 0,
              format: D.LevelFormat.BULLET,
              text: "—",
              alignment: D.AlignmentType.LEFT,
              style: {
                paragraph: { indent: { left: 720, hanging: 360 } },
                run: { color: ACCENT, font: BODY_FONT },
              },
            },
          ],
        },
      ],
    };
  }

  function docxTable(D, table) {
    const weights = table.headers.map(function (_, index) {
      return index === 0 ? 1.25 : 1;
    });
    const weightSum = weights.reduce(function (sum, value) {
      return sum + value;
    }, 0);
    const available = 9186;
    const widths = weights.map(function (weight) {
      return Math.floor((available * weight) / weightSum);
    });

    const border = { style: D.BorderStyle.SINGLE, size: 4, color: RULE };
    const margins = { top: 60, bottom: 60, left: 100, right: 100 };

    function cell(text, index, style, fill) {
      return new D.TableCell({
        width: { size: widths[index], type: D.WidthType.DXA },
        shading: fill ? { type: D.ShadingType.CLEAR, color: "auto", fill: fill } : undefined,
        verticalAlign: D.VerticalAlign.CENTER,
        margins: margins,
        children: [
          new D.Paragraph({
            style: style,
            children: [new D.TextRun({ text: text })],
          }),
        ],
      });
    }

    const rows = [
      new D.TableRow({
        tableHeader: true,
        children: table.headers.map(function (header, index) {
          return cell(header, index, "FormaCellHead", NAVY);
        }),
      }),
    ];

    table.rows.forEach(function (row, rowIndex) {
      rows.push(
        new D.TableRow({
          children: row.map(function (value, index) {
            return cell(value, index, "FormaCell", rowIndex % 2 === 1 ? ALT : null);
          }),
        })
      );
    });

    return new D.Table({
      width: { size: 100, type: D.WidthType.PERCENTAGE },
      columnWidths: widths,
      layout: D.TableLayoutType.FIXED,
      borders: {
        top: border,
        bottom: border,
        left: border,
        right: border,
        insideHorizontal: border,
        insideVertical: border,
      },
      rows: rows,
    });
  }

  function docxHeaderFooter(D) {
    const header = new D.Header({
      children: [
        new D.Paragraph({
          spacing: { after: 0 },
          border: { bottom: { style: D.BorderStyle.SINGLE, size: 4, space: 4, color: RULE } },
          children: [
            new D.TextRun({ text: "FORMA", size: 16, bold: true, color: ACCENT, font: BODY_FONT }),
            new D.TextRun({ text: "   ·   документ", size: 16, color: MUTED, font: BODY_FONT }),
          ],
        }),
      ],
    });

    const footer = new D.Footer({
      children: [
        new D.Paragraph({
          alignment: D.AlignmentType.CENTER,
          children: [
            new D.TextRun({
              children: [D.PageNumber.CURRENT],
              size: 16,
              color: MUTED,
              font: BODY_FONT,
            }),
          ],
        }),
      ],
    });

    return { header: header, footer: footer };
  }

  async function renderDocx(doc) {
    const D = global.docx;
    if (!D) throw new Error("Библиотека docx не загрузилась.");

    const paragraphs = [
      new D.Paragraph({
        style: "FormaTitle",
        children: [new D.TextRun({ text: doc.title })],
      }),
    ];

    if (doc.subtitle) {
      paragraphs.push(
        new D.Paragraph({
          style: "FormaSubtitle",
          children: [new D.TextRun({ text: doc.subtitle })],
        })
      );
    }

    const meta = metaLine(doc);
    if (meta) {
      paragraphs.push(
        new D.Paragraph({ style: "FormaMeta", children: [new D.TextRun({ text: meta })] })
      );
    }

    paragraphs.push(new D.Paragraph({ children: [new D.PageBreak()] }));

    doc.sections.forEach(function (section) {
      paragraphs.push(
        new D.Paragraph({
          style: "FormaHeading",
          children: [new D.TextRun({ text: section.heading })],
        })
      );

      section.paragraphs.forEach(function (text) {
        paragraphs.push(
          new D.Paragraph({ style: "FormaBody", children: [new D.TextRun({ text: text })] })
        );
      });

      section.bullets.forEach(function (item) {
        paragraphs.push(
          new D.Paragraph({
            style: "FormaBullet",
            numbering: { reference: BULLET_REF, level: 0 },
            children: [new D.TextRun({ text: item })],
          })
        );
      });

      if (section.table) {
        paragraphs.push(docxTable(D, section.table));
        paragraphs.push(new D.Paragraph({ spacing: { after: 120 }, children: [] }));
      }
    });

    const parts = docxHeaderFooter(D);

    const document = new D.Document({
      title: doc.title,
      creator: doc.author || "Forma",
      description: "Документ собран в FORMA",
      styles: docxStyles(D),
      numbering: docxNumbering(D),
      sections: [
        {
          properties: {
            titlePage: true,
            page: {
              size: { width: 11906, height: 16838 },
              margin: { top: 1247, right: 1360, bottom: 1134, left: 1360, header: 624, footer: 567 },
            },
          },
          headers: { default: parts.header, first: new D.Header({ children: [] }) },
          footers: { default: parts.footer, first: new D.Footer({ children: [] }) },
          children: paragraphs,
        },
      ],
    });

    if (isNode) return D.Packer.toBuffer(document);
    return D.Packer.toBlob(document);
  }

  // ---------- PPTX ----------

  function pptxText(slide, x, y, w, h, text, options) {
    const base = {
      x: x,
      y: y,
      w: w,
      h: h,
      margin: 0,
      fontFace: BODY_FONT,
      color: INK,
      valign: "top",
    };
    slide.addText(text, Object.assign(base, options));
  }

  function pptxRect(slide, RECT, x, y, w, h, color) {
    slide.addShape(RECT, {
      x: x,
      y: y,
      w: w,
      h: h,
      fill: { color: color },
      line: { type: "none" },
    });
  }

  function pptxCover(pptx, RECT, doc) {
    const slide = pptx.addSlide();
    slide.background = { color: NAVY };
    pptxRect(slide, RECT, 0, 0, 0.16, PPT_H, ACCENT);
    pptxText(slide, PPT_M, 0.62, 5, 0.4, "FORMA", {
      fontFace: BODY_FONT,
      fontSize: 11,
      bold: true,
      color: ACCENT,
    });

    const title = doc.title;
    const titleSize = title.length <= 55 ? 40 : title.length <= 90 ? 32 : 27;
    pptxText(slide, PPT_M, 2.2, PPT_CONTENT_W, 2.05, title, {
      fontFace: HEAD_FONT,
      fontSize: titleSize,
      bold: true,
      color: WHITE,
      lineSpacingMultiple: 0.98,
      fit: "shrink",
    });

    pptxRect(slide, RECT, PPT_M, 4.55, 1.3, 0.06, ACCENT);

    if (doc.subtitle) {
      pptxText(slide, PPT_M, 4.86, PPT_CONTENT_W, 0.6, doc.subtitle, {
        fontSize: 16,
        color: LIGHT_TEXT,
      });
    }

    const meta = metaLine(doc);
    if (meta) {
      pptxText(slide, PPT_M, 6.6, PPT_CONTENT_W, 0.36, meta, {
        fontSize: 11,
        color: COVER_META,
      });
    }
  }

  function pptxHeading(slide, RECT, text) {
    const size = text.length <= 42 ? 30 : text.length <= 70 ? 25 : 22;
    pptxText(slide, PPT_M, 0.62, PPT_CONTENT_W, 0.95, text, {
      fontFace: HEAD_FONT,
      fontSize: size,
      bold: true,
      color: INK,
      lineSpacingMultiple: 1.0,
      fit: "shrink",
    });
    pptxRect(slide, RECT, PPT_M, 1.62, 0.85, 0.055, ACCENT);
  }

  function pptxFooter(slide, RECT, page) {
    pptxRect(slide, RECT, PPT_M, 6.98, PPT_CONTENT_W, 0.01, RULE);
    pptxText(slide, PPT_M, 7.1, 4, 0.26, "FORMA", {
      fontSize: 8.5,
      bold: true,
      color: MUTED,
    });
    pptxText(slide, PPT_W - PPT_M - 1.0, 7.1, 1.0, 0.26, String(page).padStart(2, "0"), {
      fontSize: 8.5,
      color: MUTED,
      align: "right",
    });
  }

  function pptxContentSlide(pptx, RECT, section, page) {
    const slide = pptx.addSlide();
    pptxHeading(slide, RECT, section.heading);

    let total = 0;
    section.paragraphs.forEach(function (text) {
      total += text.length;
    });
    section.bullets.forEach(function (item) {
      total += item.length;
    });
    const size =
      total <= 200
        ? 20
        : total <= 340
        ? 17
        : total <= 520
        ? 15
        : total <= 800
        ? 13
        : total <= 1200
        ? 12
        : total <= 1800
        ? 11
        : 10;
    const paragraphGap = size >= 15 ? 14 : 9;
    const bulletGap = size >= 15 ? 11 : 7;

    const runs = [];
    section.paragraphs.forEach(function (text) {
      runs.push({
        text: text,
        options: {
          breakLine: true,
          fontSize: size,
          color: INK,
          lineSpacingMultiple: 1.12,
          paraSpaceAfter: paragraphGap,
        },
      });
    });
    section.bullets.forEach(function (item) {
      runs.push({
        text: item,
        options: {
          breakLine: true,
          bullet: { code: "2014" },
          fontSize: size,
          color: INK,
          lineSpacingMultiple: 1.08,
          paraSpaceAfter: bulletGap,
        },
      });
    });

    slide.addText(runs, {
      x: PPT_M,
      y: 1.98,
      w: PPT_CONTENT_W,
      h: 4.8,
      margin: 0,
      valign: "top",
      fontFace: BODY_FONT,
      fit: "shrink",
      breakLine: false,
    });

    pptxFooter(slide, RECT, page);
  }

  function pptxTablePlan(table, colW) {
    const sizes = [11.5, 10.5, 9.5, 8.5];
    const maxHeight = 4.9;
    const rows = [table.headers].concat(table.rows);
    const innerWidths = colW.map(function (width) {
      return Math.max(0.5, width - 0.2);
    });

    function measure(fontSize) {
      const charWidth = (fontSize * 0.56) / 72;
      const lineHeight = (fontSize * 1.24) / 72;
      const minRow = fontSize >= 11 ? 0.4 : fontSize >= 9.5 ? 0.36 : 0.32;
      const heights = rows.map(function (row) {
        let lines = 1;
        row.forEach(function (value, index) {
          const perLine = Math.max(1, Math.floor(innerWidths[index] / charWidth));
          lines = Math.max(lines, Math.ceil(String(value).length / perLine));
        });
        return Math.max(minRow, lines * lineHeight + 0.11);
      });
      return {
        fontSize: fontSize,
        heights: heights,
        total: heights.reduce(function (sum, height) {
          return sum + height;
        }, 0),
      };
    }

    for (let index = 0; index < sizes.length; index += 1) {
      const plan = measure(sizes[index]);
      if (plan.total <= maxHeight) return plan;
    }
    return measure(sizes[sizes.length - 1]);
  }

  function pptxTableSlide(pptx, RECT, heading, table, page) {
    const slide = pptx.addSlide();
    pptxHeading(slide, RECT, heading);

    const weights = table.headers.map(function (_, index) {
      return index === 0 ? 1.25 : 1;
    });
    const weightSum = weights.reduce(function (sum, value) {
      return sum + value;
    }, 0);
    const colW = weights.map(function (weight) {
      return (PPT_CONTENT_W * weight) / weightSum;
    });
    const plan = pptxTablePlan(table, colW);

    const cell = function (text, options) {
      return { text: text, options: options };
    };

    const rows = [
      table.headers.map(function (header) {
        return cell(header, {
          fill: { color: NAVY },
          color: WHITE,
          bold: true,
          fontSize: plan.fontSize,
          valign: "middle",
        });
      }),
    ];

    table.rows.forEach(function (row, rowIndex) {
      rows.push(
        row.map(function (value) {
          return cell(value, {
            fill: { color: rowIndex % 2 === 1 ? ALT : WHITE },
            color: INK,
            fontSize: plan.fontSize,
            valign: "middle",
          });
        })
      );
    });

    slide.addTable(rows, {
      x: PPT_M,
      y: 1.95,
      w: PPT_CONTENT_W,
      colW: colW,
      rowH: plan.heights,
      border: { type: "solid", pt: 0.75, color: RULE },
      fontFace: BODY_FONT,
      margin: [3, 7, 3, 7],
      valign: "middle",
      autoPage: false,
    });

    pptxFooter(slide, RECT, page);
  }

  async function renderPptx(doc) {
    const Pptx = global.PptxGenJS;
    if (!Pptx) throw new Error("Библиотека PptxGenJS не загрузилась.");

    const pptx = new Pptx();
    pptx.defineLayout({ name: "FORMA_WIDE", width: PPT_W, height: PPT_H });
    pptx.layout = "FORMA_WIDE";
    const RECT = pptx.ShapeType.rect;
    pptx.title = doc.title;
    pptx.author = doc.author || "Forma";
    pptx.company = "FORMA";

    pptxCover(pptx, RECT, doc);

    let page = 2;
    doc.sections.forEach(function (section) {
      if (section.paragraphs.length || section.bullets.length) {
        pptxContentSlide(pptx, RECT, section, page);
        page += 1;
      }
      if (section.table) {
        pptxTableSlide(pptx, RECT, section.heading, section.table, page);
        page += 1;
      }
    });

    return pptx.write({ outputType: isNode ? "nodebuffer" : "blob" });
  }

  // ---------- PDF ----------

  function pdfLine(width, color, thickness, margin) {
    return {
      canvas: [
        {
          type: "line",
          x1: 0,
          y1: 0,
          x2: width,
          y2: 0,
          lineWidth: thickness,
          lineColor: color,
        },
      ],
      margin: margin,
    };
  }

  function pdfTable(table) {
    const cols = table.headers.length;
    const weights = table.headers.map(function (_, index) {
      return index === 0 ? 1.25 : 1;
    });
    const weightSum = weights.reduce(function (sum, value) {
      return sum + value;
    }, 0);
    const widths = weights.map(function (weight) {
      return (PDF_CONTENT_W * weight) / weightSum;
    });

    const body = [
      table.headers.map(function (header) {
        return { text: header, bold: true, color: PDF_WHITE, fontSize: 9 };
      }),
    ];
    table.rows.forEach(function (row) {
      body.push(
        row.map(function (value) {
          return { text: value, color: PDF_INK, fontSize: 9 };
        })
      );
    });

    return {
      unbreakable: true,
      table: { headerRows: 1, widths: widths, body: body },
      layout: {
        fillColor: function (rowIndex) {
          if (rowIndex === 0) return PDF_NAVY;
          return rowIndex % 2 === 0 ? PDF_ALT : null;
        },
        hLineColor: function () {
          return PDF_RULE;
        },
        vLineColor: function () {
          return PDF_RULE;
        },
        hLineWidth: function () {
          return 0.5;
        },
        vLineWidth: function () {
          return 0.5;
        },
        paddingLeft: function () {
          return 7;
        },
        paddingRight: function () {
          return 7;
        },
        paddingTop: function () {
          return 5;
        },
        paddingBottom: function () {
          return 5;
        },
      },
      margin: [0, 4, 0, 12],
    };
  }

  async function renderPdf(doc) {
    if (!global.pdfMake) throw new Error("Библиотека pdfmake не загрузилась.");

    const content = [{ text: doc.title, style: "title" }];

    if (doc.subtitle) content.push({ text: doc.subtitle, style: "subtitle" });

    const meta = metaLine(doc);
    if (meta) content.push({ text: meta, style: "meta" });

    content.push(pdfLine(PDF_CONTENT_W, PDF_ACCENT, 1.4, [0, 2, 0, 16]));

    doc.sections.forEach(function (section) {
      const nodes = [];
      nodes.push({ text: section.heading, style: "h1" });
      nodes.push(pdfLine(PDF_CONTENT_W * 0.32, PDF_ACCENT, 1.3, [0, 0, 0, 10]));

      let textLength = 0;
      section.paragraphs.forEach(function (text) {
        textLength += text.length;
        nodes.push({ text: text, style: "body" });
      });

      section.bullets.forEach(function (item) {
        textLength += item.length;
        nodes.push({
          columns: [
            { text: "—", width: 14, color: PDF_ACCENT, fontSize: 10.5 },
            { text: item },
          ],
          columnGap: 4,
          margin: [0, 0, 0, 4],
        });
      });

      if (section.table) nodes.push(pdfTable(section.table));

      if (section.table && textLength <= 600) {
        content.push({ unbreakable: true, stack: nodes });
      } else {
        nodes.forEach(function (node) {
          content.push(node);
        });
      }
    });

    const definition = {
      pageSize: "A4",
      pageMargins: [62, 72, 62, 52],
      info: { title: doc.title, author: doc.author || "Forma", creator: "Forma" },
      defaultStyle: { font: "Roboto", fontSize: 10.5, lineHeight: 1.4, color: PDF_INK },
      styles: {
        title: { fontSize: 26, bold: true, color: PDF_INK, margin: [0, 0, 0, 8] },
        subtitle: { fontSize: 12.5, color: PDF_ACCENT, margin: [0, 0, 0, 12] },
        meta: { fontSize: 8.5, color: PDF_MUTED, margin: [0, 0, 0, 4] },
        h1: { fontSize: 15, bold: true, color: PDF_INK, margin: [0, 16, 0, 4] },
        body: { fontSize: 10.5, lineHeight: 1.4, color: PDF_INK, margin: [0, 0, 0, 8] },
      },
      header: function () {
        return {
          margin: [62, 24, 62, 0],
          stack: [
            {
              columns: [
                { text: "FORMA", bold: true, fontSize: 8, color: PDF_ACCENT },
                { text: "документ", fontSize: 8, color: PDF_MUTED, alignment: "right" },
              ],
            },
            {
              margin: [0, 5, 0, 0],
              canvas: [
                {
                  type: "line",
                  x1: 0,
                  y1: 0,
                  x2: PDF_CONTENT_W,
                  y2: 0,
                  lineWidth: 0.6,
                  lineColor: PDF_RULE,
                },
              ],
            },
          ],
        };
      },
      footer: function (currentPage) {
        return [
          {
            margin: [62, 0, 62, 0],
            canvas: [
              {
                type: "line",
                x1: 0,
                y1: 0,
                x2: PDF_CONTENT_W,
                y2: 0,
                lineWidth: 0.6,
                lineColor: PDF_RULE,
              },
            ],
          },
          {
            text: String(currentPage),
            alignment: "center",
            fontSize: 8,
            color: PDF_MUTED,
            margin: [0, 6, 0, 0],
          },
        ];
      },
      content: content,
    };

    return new Promise(function (resolve, reject) {
      try {
        const pdf = global.pdfMake.createPdf(definition);
        if (isNode) {
          pdf.getBuffer(function (buffer) {
            resolve(buffer);
          });
        } else {
          pdf.getBlob(function (blob) {
            resolve(blob);
          });
        }
      } catch (error) {
        reject(error);
      }
    });
  }

  Object.assign(Forma, {
    renderers: { docx: renderDocx, pptx: renderPptx, pdf: renderPdf },
  });
})(typeof window !== "undefined" ? window : globalThis);
