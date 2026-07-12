---
name: pdf-word-skill
description: Build editable Word documents from PDFs by splitting the PDF into pages, recognizing text/images with PDF text extraction or OCR, rebuilding editable DOCX pages, comparing rendered Word pages against the source PDF, and iteratively adjusting text boxes and image positions until the layout closely matches the original. Use for WPS-compatible editable Word reconstruction where direct PDF-to-Word conversion creates garbled text, missing images, or layout drift.
---

# PDF to Editable Word: Recognize, Compare, Adjust

This skill is for rebuilding an editable Word document from a PDF while keeping the layout as close to the original PDF as possible. Direct conversion is not trusted as the final result. The core loop is:

**recognize content -> rebuild Word page -> compare written text -> render Word -> compare with PDF -> adjust text/images -> repeat.**

## Operating Principle

- Start with a 3-page sample run unless the user explicitly provides pages or asks for the full document. When no page range is passed, the script automatically selects representative complex pages and tries to cover image-heavy, table/diagram-heavy, and text-dense layouts.
- Only run the full document after the sample DOCX and comparison outputs look acceptable. Use `--all-pages` for the full run.
- Do not overwrite an existing output directory unless `--force` is passed. Use `--resume` after an interrupted run; resume means rerun only failed pages and never overwrite pages already confirmed successful.
- Process one PDF page at a time.
- Do not move to the next page until the current page has a usable editable Word page and a comparison result.
- Prefer WPS-compatible DOCX output.
- Avoid `pdf2docx` text when it produces garbled characters. Use rebuild-only mode for WPS workflows.
- Use source PDF imagery to recover photos, diagrams, logos, and backgrounds.
- Recreate visible text as editable Word text boxes whenever text can be extracted or OCR-recognized.
- Use OCR fallback for pages without a usable PDF text layer. If OCR is unavailable or produces no text, mark the page failed with a clear reason instead of silently accepting it.
- Use red/yellow/green page quality status in reports and `review.html`.
- Treat text and image position adjustment as the main repair step.

## Page Workflow

1. Split the input PDF into single-page PDFs.
2. For the current page, render the source PDF page to PNG.
3. Recognize page content:
   - Extract text from the PDF text layer first.
   - If text is missing, corrupt, or inside an image, use OCR when available.
   - Record page analysis: layout class, text lines, image blocks, drawings, table-like rows, and OCR need.
   - Identify photos, diagrams, warning labels, logos, tables, and other visual regions.
4. Rebuild a Word page:
   - Create a DOCX page with the same page size as the PDF.
   - Recover photos/diagrams/background artwork from the source PDF render.
   - Place recognized text as editable Word text boxes.
   - Keep text boxes, images, and page geometry close to the PDF coordinates.
5. Validate written text before trusting layout:
   - Compare the text written into Word against the PDF text layer or OCR source text.
   - Check for missing lines, garbled characters, duplicated text, wrong punctuation, and obvious OCR substitutions.
   - Record a text similarity score in the page report.
6. Render the rebuilt Word page back to PDF/PNG using LibreOffice.
7. Compare the rendered Word page against the source PDF page.
8. Classify page quality:
   - Green: accepted.
   - Yellow: accepted but should be manually reviewed.
   - Red: failed and must be retried or manually repaired.
9. Adjust and rerender:
   - Move or resize text boxes.
   - Move, crop, or scale images.
   - Adjust font size, box width/height, and line spacing.
   - Fix missing photos or OCR text.
   - Re-run text comparison whenever rewritten content changes.
10. Repeat comparison until the page is acceptable or mark it for manual review.
11. Save the approved single-page DOCX and proceed to the next page.
12. Write a partial report after every completed page so interrupted runs can resume.
13. On resume, rerun only failed pages; keep successful pages untouched.
14. Merge approved pages into the final editable DOCX only when every target page succeeded.
15. Finalize the DOCX in exactly two steps: re-save the DOCX, then compress embedded bitmap images only.
16. Render the final merged DOCX and run a final page-by-page comparison. Record both `final_docx_created` and `final_docx_accepted`; a file can be generated but still fail final QA.
17. Generate `review.html` with source and rebuilt page renders side by side.

## Script

Use `scripts/repair_pdf_to_word.py`.

Recommended WPS-compatible sample run. Do this first:

```powershell
python "$env:CODEX_HOME\skills\pdf-word-skill\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_sample" `
  --force `
  --rebuild-only
```

When no `--pages`, `--single-page`, or `--all-pages` option is passed, the script processes three representative sample pages automatically.

Single-page test:

```powershell
python "$env:CODEX_HOME\skills\pdf-word-skill\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_page_1" `
  --single-page 1 `
  --force `
  --rebuild-only
```

Explicit page range after reviewing the sample:

```powershell
python "$env:CODEX_HOME\skills\pdf-word-skill\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_pages_1_5" `
  --pages 1-5 `
  --force `
  --rebuild-only
```

Full-document run only after the sample is accepted:

```powershell
python "$env:CODEX_HOME\skills\pdf-word-skill\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_output" `
  --all-pages `
  --force `
  --rebuild-only
```

Resume an interrupted run:

```powershell
python "$env:CODEX_HOME\skills\pdf-word-skill\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_rebuild_output" `
  --all-pages `
  --resume `
  --rebuild-only
```

`--resume` only retries pages that failed in the checkpoint report. It does not rewrite pages already marked successful.

Optional baseline mode, only when `pdf2docx` text is not garbled:

```powershell
python "$env:CODEX_HOME\skills\pdf-word-skill\scripts\repair_pdf_to_word.py" `
  --input "input.pdf" `
  --output-dir "word_repair_output" `
  --force `
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
- The final written text must match the source text closely enough to avoid omissions, substitutions, duplicated lines, and OCR mistakes.
- A page is failed/red if it has missing text, text coverage below threshold, missing validation where source text exists, missing visual metrics, or severe visual drift.
- A page is yellow if it has moderate visual drift or possible duplicate text but no hard text failure.
- Resume mode retries only red/failed pages. Green and yellow pages are treated as successful and are not overwritten.
- Final QA also uses red/yellow/green visual status after pages are merged. If final merged rendering has red pages, keep the DOCX for review, mark `final_docx_accepted=false`, and write those pages back as failed so the next `--resume` retries only them.
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
- `final_compare/`: final merged DOCX render compared back to source pages.
- `review.html`: side-by-side visual review page for processed pages.
- `<source>_editable_repaired.docx`: merged final Word document.
- `repair_report.json`: page-level metrics, selected page sources, and repair status.
- `repair_report.partial.json`: checkpoint report during an interrupted or running job.

Report fields include:

- `page_analysis`: layout class, text lines, image blocks, drawings, table-like rows, and OCR need.
- `severity`: `green`, `yellow`, or `red`.
- `issues`: quality reasons such as `visual_diff_yellow`, `missing_text`, `ocr_fail`, or `possible_duplicate_text`.
- `retry_count`: how many times the page has been retried through resume mode.
- `final_docx_accepted`: whether the final merged DOCX passed final render comparison.
- `final_docx_compression`: re-save and image-compression audit, including per-image before/after size and pixel dimensions.

Final DOCX compression rule:

- Step 1: re-save the merged DOCX.
- Step 2: compress embedded bitmap images only (`.jpg`, `.jpeg`, `.png`).
- Record every compressed or skipped image in `repair_report.json`.
- Do not delete text boxes.
- Do not rebuild document structure.
- Do not run generic ZIP optimization on the DOCX package.
