---
name: pdf-word-skill
description: Build editable Word documents from PDFs by splitting the PDF into pages, recognizing text/images with PDF text extraction or OCR, rebuilding editable DOCX pages, comparing rendered Word pages against the source PDF, and iteratively adjusting text boxes and image positions until the layout closely matches the original. Use for WPS-compatible editable Word reconstruction where direct PDF-to-Word conversion creates garbled text, missing images, or layout drift.
---

# PDF to Editable Word: Recognize, Compare, Adjust

This skill is for rebuilding an editable Word document from a PDF while keeping the layout as close to the original PDF as possible. Direct conversion is not trusted as the final result. The core loop is:

**recognize content -> rebuild Word page -> render Word -> compare with PDF -> adjust text/images -> repeat.**

## Operating Principle

- Process one PDF page at a time.
- Do not move to the next page until the current page has a usable editable Word page and a comparison result.
- Prefer WPS-compatible DOCX output.
- Avoid `pdf2docx` text when it produces garbled characters. Use rebuild-only mode for WPS workflows.
- Use source PDF imagery to recover photos, diagrams, logos, and backgrounds.
- Recreate visible text as editable Word text boxes whenever text can be extracted or OCR-recognized.
- Treat text and image position adjustment as the main repair step.

## Page Workflow

1. Split the input PDF into single-page PDFs.
2. For the current page, render the source PDF page to PNG.
3. Recognize page content:
   - Extract text from the PDF text layer first.
   - If text is missing, corrupt, or inside an image, use OCR when available.
   - Identify photos, diagrams, warning labels, logos, tables, and other visual regions.
4. Rebuild a Word page:
   - Create a DOCX page with the same page size as the PDF.
   - Recover photos/diagrams/background artwork from the source PDF render.
   - Place recognized text as editable Word text boxes.
   - Keep text boxes, images, and page geometry close to the PDF coordinates.
5. Render the rebuilt Word page back to PDF/PNG using LibreOffice.
6. Validate recognized text:
   - Compare the text written into Word against the PDF text layer or OCR source text.
   - Check for missing lines, garbled characters, duplicated text, and obvious OCR substitutions.
   - Record a text similarity score in the page report.
7. Compare the rendered Word page against the source PDF page.
8. Adjust and rerender:
   - Move or resize text boxes.
   - Move, crop, or scale images.
   - Adjust font size, box width/height, and line spacing.
   - Fix missing photos or OCR text.
9. Repeat comparison until the page is acceptable or mark it for manual review.
10. Save the approved single-page DOCX and proceed to the next page.
11. Merge approved pages into the final editable DOCX.
12. Render the final merged DOCX and run a final page-by-page comparison.

## Script

Use `scripts/repair_pdf_to_word.py`.

Recommended WPS-compatible rebuild mode:

```powershell
python "$env:CODEX_HOME\skills\pdf2docx-compare-repair\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_output" `
  --pages 1-5 `
  --rebuild-only
```

Single-page test:

```powershell
python "$env:CODEX_HOME\skills\pdf2docx-compare-repair\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_page_1" `
  --single-page 1 `
  --rebuild-only
```

Optional baseline mode, only when `pdf2docx` text is not garbled:

```powershell
python "$env:CODEX_HOME\skills\pdf2docx-compare-repair\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_repair_output" `
  --pages 1-5
```

## Dependencies

- Microsoft Word on Windows for creating editable DOCX pages with positioned text boxes.
- LibreOffice `soffice` for render-based comparison.
- Python packages: `PyMuPDF`, `Pillow`, `numpy`, `python-docx`, `docxcompose`, `pywin32`.
- OCR engine if the PDF lacks a usable text layer or if image-embedded text must become editable.

## Acceptance Rules

- Text that is reconstructed from PDF text/OCR must be editable in Word/WPS.
- Reconstructed Word text must be validated against the PDF text layer or OCR source text.
- Photos, diagrams, and logos must be visually present and close to the original position.
- Page size, margins, major headings, tables, and image blocks should align with the PDF.
- If image-embedded text remains raster-only, note it in the report unless OCR reconstruction was requested for those regions.
- Do not silently accept pages with missing images, garbled text, severe overlap, or major drift.

## Outputs

The script writes:

- `pdf_pages/`: split single-page PDFs.
- `docx_rebuilt/`: rebuilt editable page DOCX files.
- `docx_selected/`: approved page DOCX files used for merge.
- `rendered_pdf/`: Word-rendered PDFs for QA.
- `png_compare/`: source and rendered page PNGs for visual comparison.
- `<source>_editable_repaired.docx`: merged final Word document.
- `repair_report.json`: page-level metrics, selected page sources, and repair status.
