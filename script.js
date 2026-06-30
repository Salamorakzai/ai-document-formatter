/* =========================================================================
   AI Document Formatter — script.js
   100% client-side. No backend, no build step.

   Sections:
     1. DOM references
     2. Block-level parser  (raw text -> structured blocks)
     3. Inline formatter    (bold / italic within a line)
     4. Live preview renderer (blocks -> HTML inside #previewBody)
     5. DOCX export          (blocks -> real .docx via the "docx" library)
     6. PDF export           (renders #previewSheet via html2pdf)
     7. Wiring / event listeners
   ========================================================================= */

(function () {
  "use strict";

  /* -----------------------------------------------------------------------
     1. DOM references
  ----------------------------------------------------------------------- */
  const els = {
    editor: document.getElementById("editor"),
    wordCount: document.getElementById("wordCount"),
    docTitle: document.getElementById("docTitle"),
    docAuthor: document.getElementById("docAuthor"),
    docCompany: document.getElementById("docCompany"),
    docType: document.getElementById("docType"),
    docTypeBadge: document.getElementById("docTypeBadge"),
    previewCover: document.getElementById("previewCover"),
    previewBody: document.getElementById("previewBody"),
    previewSheet: document.getElementById("previewSheet"),
    detectionSummary: document.getElementById("detectionSummary"),
    exportDocxBtn: document.getElementById("exportDocxBtn"),
    exportPdfBtn: document.getElementById("exportPdfBtn"),
    loadSampleBtn: document.getElementById("loadSampleBtn"),
    toast: document.getElementById("toast"),
  };

  const SAMPLE_TEXT =
`# Standard Operating Procedure: Equipment Calibration

This document describes the **calibration process** for laboratory balances and outlines *quality control* checkpoints.

## Purpose

This SOP ensures all balances are calibrated to manufacturer specification before use.

## Scope

Applies to all QC technicians operating analytical balances.

## Procedure Steps

1. Power on the balance and allow a 30-minute warm-up
2. Level the instrument using the built-in bubble level
3. Run the internal calibration routine
4. Record results in the calibration log

## Required Materials

- Certified calibration weights (1g, 10g, 100g)
- Lint-free cleaning cloth
- Calibration logbook

## Acceptance Criteria

| Weight | Tolerance | Result |
|---|---|---|
| 1g | ±0.001g | Pass |
| 10g | ±0.002g | Pass |
| 100g | ±0.005g | Pass |

\`\`\`
calibration_status = "PASS" if deviation <= tolerance else "FAIL"
\`\`\`

Any failed calibration must be reported immediately to the QC supervisor.`;

  /* -----------------------------------------------------------------------
     2. Block-level parser
     Converts the raw pasted text into an array of typed "blocks". This is
     the single source of truth consumed by BOTH the live preview renderer
     and the DOCX generator, so what the user sees is what they export.
  ----------------------------------------------------------------------- */

  /**
   * @typedef {{type:'heading', level:1|2, text:string}} HeadingBlock
   * @typedef {{type:'paragraph', text:string}} ParagraphBlock
   * @typedef {{type:'bullet-list', items:string[]}} BulletBlock
   * @typedef {{type:'numbered-list', items:string[]}} NumberedBlock
   * @typedef {{type:'table', header:string[], rows:string[][]}} TableBlock
   * @typedef {{type:'code', lines:string[]}} CodeBlock
   */

  function parseDocument(raw) {
    const lines = raw.replace(/\r\n/g, "\n").split("\n");
    const blocks = [];
    let i = 0;
    let paragraphBuffer = [];

    const flushParagraph = () => {
      if (paragraphBuffer.length === 0) return;
      const text = paragraphBuffer.join(" ").trim();
      paragraphBuffer = [];
      if (text) blocks.push({ type: "paragraph", text });
    };

    const isBullet = (l) => /^\s*[-*•]\s+/.test(l);
    const isNumbered = (l) => /^\s*\d+[.)]\s+/.test(l);
    const isTableRow = (l) => /^\s*\|.+\|\s*$/.test(l);
    const isTableSeparator = (l) => /^\s*\|?[\s:|-]+\|[\s:|-]*$/.test(l) && l.includes("-");

    while (i < lines.length) {
      const line = lines[i];
      const trimmed = line.trim();

      // Blank line -> paragraph boundary
      if (trimmed === "") {
        flushParagraph();
        i++;
        continue;
      }

      // Code block ``` ... ```
      if (/^```/.test(trimmed)) {
        flushParagraph();
        const codeLines = [];
        i++;
        while (i < lines.length && !/^```/.test(lines[i].trim())) {
          codeLines.push(lines[i]);
          i++;
        }
        i++; // skip closing fence
        blocks.push({ type: "code", lines: codeLines });
        continue;
      }

      // Headings: # Heading / ## Subheading
      const h2 = trimmed.match(/^##\s+(.*)$/);
      const h1 = trimmed.match(/^#\s+(.*)$/);
      if (h2) {
        flushParagraph();
        blocks.push({ type: "heading", level: 2, text: h2[1].trim() });
        i++;
        continue;
      }
      if (h1) {
        flushParagraph();
        blocks.push({ type: "heading", level: 1, text: h1[1].trim() });
        i++;
        continue;
      }

      // Table: a "| a | b |" row immediately followed by a "---|---" separator
      if (isTableRow(line) && lines[i + 1] && isTableSeparator(lines[i + 1])) {
        flushParagraph();
        const splitRow = (r) =>
          r.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
        const header = splitRow(line);
        i += 2;
        const rows = [];
        while (i < lines.length && isTableRow(lines[i])) {
          rows.push(splitRow(lines[i]));
          i++;
        }
        blocks.push({ type: "table", header, rows });
        continue;
      }

      // Bullet list
      if (isBullet(line)) {
        flushParagraph();
        const items = [];
        while (i < lines.length && isBullet(lines[i])) {
          items.push(lines[i].replace(/^\s*[-*•]\s+/, "").trim());
          i++;
        }
        blocks.push({ type: "bullet-list", items });
        continue;
      }

      // Numbered list
      if (isNumbered(line)) {
        flushParagraph();
        const items = [];
        while (i < lines.length && isNumbered(lines[i])) {
          items.push(lines[i].replace(/^\s*\d+[.)]\s+/, "").trim());
          i++;
        }
        blocks.push({ type: "numbered-list", items });
        continue;
      }

      // Default: accumulate into the current paragraph
      paragraphBuffer.push(trimmed);
      i++;
    }

    flushParagraph();
    return blocks;
  }

  /* -----------------------------------------------------------------------
     3. Inline formatter
     Splits a line into {text, bold, italic} runs based on **bold** and
     *italic* markdown-style markers (and __bold__ / _italic_ variants).
  ----------------------------------------------------------------------- */
  function parseInline(text) {
    const runs = [];
    const pattern = /(\*\*[^*]+\*\*)|(__[^_]+__)|(\*[^*]+\*)|(_[^_]+_)/g;
    let lastIndex = 0;
    let match;

    while ((match = pattern.exec(text)) !== null) {
      if (match.index > lastIndex) {
        runs.push({ text: text.slice(lastIndex, match.index), bold: false, italic: false });
      }
      const token = match[0];
      if (token.startsWith("**") || token.startsWith("__")) {
        runs.push({ text: token.slice(2, -2), bold: true, italic: false });
      } else {
        runs.push({ text: token.slice(1, -1), bold: false, italic: true });
      }
      lastIndex = pattern.lastIndex;
    }
    if (lastIndex < text.length) {
      runs.push({ text: text.slice(lastIndex), bold: false, italic: false });
    }
    return runs.length ? runs : [{ text, bold: false, italic: false }];
  }

  /* -----------------------------------------------------------------------
     4. Live preview renderer
  ----------------------------------------------------------------------- */
  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function inlineToHtml(text) {
    return parseInline(text)
      .map((run) => {
        let safe = escapeHtml(run.text);
        if (run.bold) safe = `<strong>${safe}</strong>`;
        if (run.italic) safe = `<em>${safe}</em>`;
        return safe;
      })
      .join("");
  }

  function renderPreview(blocks) {
    if (blocks.length === 0) {
      els.previewBody.innerHTML = `<p class="doc-empty">Your formatted document will appear here as you type.</p>`;
      return;
    }

    const html = blocks
      .map((block) => {
        switch (block.type) {
          case "heading":
            return block.level === 1
              ? `<h1 class="doc-h1">${inlineToHtml(block.text)}</h1>`
              : `<h2 class="doc-h2">${inlineToHtml(block.text)}</h2>`;
          case "paragraph":
            return `<p class="doc-p">${inlineToHtml(block.text)}</p>`;
          case "bullet-list":
            return `<ul class="doc-ul">${block.items
              .map((it) => `<li>${inlineToHtml(it)}</li>`)
              .join("")}</ul>`;
          case "numbered-list":
            return `<ol class="doc-ol">${block.items
              .map((it) => `<li>${inlineToHtml(it)}</li>`)
              .join("")}</ol>`;
          case "code":
            return `<pre class="doc-code">${escapeHtml(block.lines.join("\n"))}</pre>`;
          case "table": {
            const head = `<tr>${block.header
              .map((c) => `<th>${escapeHtml(c)}</th>`)
              .join("")}</tr>`;
            const body = block.rows
              .map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`)
              .join("");
            return `<table class="doc-table"><thead>${head}</thead><tbody>${body}</tbody></table>`;
          }
          default:
            return "";
        }
      })
      .join("");

    els.previewBody.innerHTML = html;
  }

  function renderCover(meta) {
    const dateStr = new Date().toLocaleDateString(undefined, {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    els.previewCover.innerHTML = `
      <span class="cover-type">${escapeHtml(meta.type)}</span>
      <div class="cover-title">${escapeHtml(meta.title || "Untitled Document")}</div>
      <div class="cover-meta">
        ${meta.company ? escapeHtml(meta.company) + " &middot; " : ""}${meta.author ? escapeHtml(meta.author) + " &middot; " : ""}${dateStr}
      </div>
    `;
  }

  function getMeta() {
    return {
      title: els.docTitle.value.trim(),
      author: els.docAuthor.value.trim(),
      company: els.docCompany.value.trim(),
      type: els.docType.value,
    };
  }

  function updateDetectionSummary(blocks) {
    const counts = {};
    blocks.forEach((b) => (counts[b.type] = (counts[b.type] || 0) + 1));
    if (blocks.length === 0) {
      els.detectionSummary.textContent = "Start typing to detect document structure";
      return;
    }
    const labels = {
      heading: "heading",
      paragraph: "paragraph",
      "bullet-list": "bullet list",
      "numbered-list": "numbered list",
      table: "table",
      code: "code block",
    };
    const parts = Object.entries(counts).map(
      ([type, n]) => `${n} ${labels[type] || type}${n > 1 ? "s" : ""}`
    );
    els.detectionSummary.textContent = "Detected " + parts.join(", ");
  }

  function refresh() {
    const meta = getMeta();
    const blocks = parseDocument(els.editor.value);
    renderCover(meta);
    renderPreview(blocks);
    updateDetectionSummary(blocks);
    els.docTypeBadge.textContent = meta.type;

    const words = els.editor.value.trim();
    els.wordCount.textContent = (words ? words.split(/\s+/).length : 0) + " words";

    return { meta, blocks };
  }

  /* -----------------------------------------------------------------------
     5. DOCX export
     Uses the UMD build of the "docx" library (window.docx) to assemble a
     real, structured Word document: cover page, table of contents,
     styled headings, lists, tables, code blocks, header/footer & page
     numbers. Mirrors the docx-js best practices: explicit US Letter page
     size, native numbering (no manual bullet characters), dual table
     widths in DXA, and HeadingLevel-based styles so the TOC can find them.
  ----------------------------------------------------------------------- */
  async function exportDocx() {
    const { meta, blocks } = refresh();

    if (!els.editor.value.trim()) {
      showToast("Paste some content before exporting.");
      return;
    }

    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      Header, Footer, AlignmentType, LevelFormat, TableOfContents,
      HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
    } = docx;

    const PAGE_WIDTH = 12240; // US Letter, DXA
    const PAGE_HEIGHT = 15840;
    const MARGIN = 1440; // 1 inch
    const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 9360
    const ACCENT = "6D5EF8";
    const FONT = "Calibri";

    // --- Inline run builder shared by headings / paragraphs / lists ---
    function inlineRuns(text, extra) {
      return parseInline(text).map(
        (run) =>
          new TextRun({
            text: run.text,
            bold: !!run.bold,
            italics: !!run.italic,
            ...extra,
          })
      );
    }

    // --- Cover page ---
    const dateStr = new Date().toLocaleDateString(undefined, {
      year: "numeric", month: "long", day: "numeric",
    });
    const coverParagraphs = [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 1600, after: 200 },
        children: [
          new TextRun({ text: meta.type.toUpperCase(), size: 20, bold: true, color: ACCENT, font: FONT }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [
          new TextRun({ text: meta.title || "Untitled Document", bold: true, size: 52, font: FONT, color: "111111" }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: ACCENT, space: 1 } },
        spacing: { after: 500 },
        children: [new TextRun({ text: " ", size: 2 })],
      }),
    ];
    if (meta.company) {
      coverParagraphs.push(new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [new TextRun({ text: meta.company, size: 26, font: FONT })],
      }));
    }
    if (meta.author) {
      coverParagraphs.push(new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [new TextRun({ text: "Prepared by " + meta.author, size: 22, color: "555555", font: FONT })],
      }));
    }
    coverParagraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 80 },
      children: [new TextRun({ text: dateStr, size: 22, color: "777777", font: FONT })],
    }));
    coverParagraphs.push(new Paragraph({ children: [new PageBreak()] }));

    // --- Table of contents ---
    const tocParagraphs = [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "Table of Contents", font: FONT })] }),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),
    ];

    // --- Body: blocks -> docx elements ---
    const bodyElements = [];
    blocks.forEach((block) => {
      switch (block.type) {
        case "heading":
          bodyElements.push(new Paragraph({
            heading: block.level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
            children: inlineRuns(block.text, { font: FONT }),
          }));
          break;
        case "paragraph":
          bodyElements.push(new Paragraph({
            spacing: { after: 200, line: 276 },
            children: inlineRuns(block.text, { font: FONT, size: 22 }),
          }));
          break;
        case "bullet-list":
          block.items.forEach((item) => {
            bodyElements.push(new Paragraph({
              numbering: { reference: "doc-bullets", level: 0 },
              spacing: { after: 80 },
              children: inlineRuns(item, { font: FONT, size: 22 }),
            }));
          });
          break;
        case "numbered-list":
          block.items.forEach((item) => {
            bodyElements.push(new Paragraph({
              numbering: { reference: "doc-numbers", level: 0 },
              spacing: { after: 80 },
              children: inlineRuns(item, { font: FONT, size: 22 }),
            }));
          });
          break;
        case "code":
          block.lines.forEach((line, idx) => {
            bodyElements.push(new Paragraph({
              shading: { type: ShadingType.CLEAR, fill: "F5F5F5" },
              spacing: { after: idx === block.lines.length - 1 ? 200 : 0 },
              border: { left: { style: BorderStyle.SINGLE, size: 12, color: "CCCCCC", space: 4 } },
              children: [new TextRun({ text: line.length ? line : " ", font: "Consolas", size: 19 })],
            }));
          });
          break;
        case "table": {
          const colCount = Math.max(block.header.length, 1);
          const colWidth = Math.floor(CONTENT_WIDTH / colCount);
          const widths = new Array(colCount).fill(colWidth);
          widths[colCount - 1] += CONTENT_WIDTH - colWidth * colCount;
          const border = { style: BorderStyle.SINGLE, size: 2, color: "D0D0D0" };
          const borders = { top: border, bottom: border, left: border, right: border };
          const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

          const headerRow = new TableRow({
            tableHeader: true,
            children: block.header.map((cell, idx) => new TableCell({
              borders, width: { size: widths[idx], type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: ACCENT },
              margins: cellMargins,
              children: [new Paragraph({ children: [new TextRun({ text: cell, bold: true, color: "FFFFFF", font: FONT })] })],
            })),
          });
          const bodyRows = block.rows.map((row, rIdx) => new TableRow({
            children: row.map((cell, idx) => new TableCell({
              borders, width: { size: widths[idx] || colWidth, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: rIdx % 2 === 0 ? "FFFFFF" : "F7F7FB" },
              margins: cellMargins,
              children: [new Paragraph({ children: [new TextRun({ text: cell, font: FONT })] })],
            })),
          }));
          bodyElements.push(new Table({
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            columnWidths: widths,
            rows: [headerRow, ...bodyRows],
          }));
          break;
        }
      }
    });

    // --- Header / footer ---
    const header = new Header({
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: meta.title || "Untitled Document", size: 18, color: "888888", font: FONT })],
      })],
    });
    const footerRuns = [];
    if (meta.company) footerRuns.push(new TextRun({ text: meta.company + "  |  ", size: 18, color: "888888", font: FONT }));
    footerRuns.push(new TextRun({ text: "Page ", size: 18, color: "888888", font: FONT }));
    footerRuns.push(new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888", font: FONT }));
    footerRuns.push(new TextRun({ text: " of ", size: 18, color: "888888", font: FONT }));
    footerRuns.push(new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "888888", font: FONT }));
    const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: footerRuns })] });

    const document = new Document({
      styles: {
        default: { document: { run: { font: FONT, size: 22 } } },
        paragraphStyles: [
          {
            id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 34, bold: true, font: FONT, color: ACCENT },
            paragraph: { spacing: { before: 320, after: 220 }, outlineLevel: 0 },
          },
          {
            id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 27, bold: true, font: FONT, color: "1A1A1A" },
            paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 },
          },
        ],
      },
      numbering: {
        config: [
          {
            reference: "doc-bullets",
            levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
          },
          {
            reference: "doc-numbers",
            levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
          },
        ],
      },
      sections: [{
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        headers: { default: header },
        footers: { default: footer },
        children: [...coverParagraphs, ...tocParagraphs, ...bodyElements],
      }],
    });

    try {
      const blob = await Packer.toBlob(document);
      const filename = (meta.title || "document").trim().replace(/[^a-z0-9-_ ]/gi, "").replace(/\s+/g, "-").toLowerCase() + ".docx";
      saveAs(blob, filename);
      showToast("Word document downloaded ✓");
    } catch (err) {
      console.error(err);
      showToast("Something went wrong generating the .docx file.");
    }
  }

  /* -----------------------------------------------------------------------
     6. PDF export
     Renders the visible preview sheet (cover + body) into a paginated PDF
     using html2pdf.js (jsPDF + html2canvas under the hood).
  ----------------------------------------------------------------------- */
  function exportPdf() {
    if (!els.editor.value.trim()) {
      showToast("Paste some content before exporting.");
      return;
    }
    refresh();
    const meta = getMeta();
    const filename = (meta.title || "document").trim().replace(/[^a-z0-9-_ ]/gi, "").replace(/\s+/g, "-").toLowerCase() + ".pdf";

    const opt = {
      margin: 0,
      filename,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"] },
    };

    showToast("Generating PDF…");
    html2pdf().set(opt).from(els.previewSheet).save().then(() => {
      showToast("PDF downloaded ✓");
    }).catch((err) => {
      console.error(err);
      showToast("Something went wrong generating the PDF.");
    });
  }

  /* -----------------------------------------------------------------------
     7. Wiring
  ----------------------------------------------------------------------- */
  let toastTimer;
  function showToast(message) {
    els.toast.textContent = message;
    els.toast.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove("show"), 2600);
  }

  function init() {
    [els.editor, els.docTitle, els.docAuthor, els.docCompany, els.docType].forEach((el) => {
      el.addEventListener("input", refresh);
      el.addEventListener("change", refresh);
    });

    els.loadSampleBtn.addEventListener("click", () => {
      els.editor.value = SAMPLE_TEXT;
      els.docTitle.value = "Standard Operating Procedure: Equipment Calibration";
      els.docAuthor.value = "Jane Doe";
      els.docCompany.value = "Acme Labs Inc.";
      els.docType.value = "SOP";
      refresh();
    });

    els.exportDocxBtn.addEventListener("click", exportDocx);
    els.exportPdfBtn.addEventListener("click", exportPdf);

    refresh();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
