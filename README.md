# pdf-word-skill

Codex skill for rebuilding editable Word documents from PDFs while preserving page layout.

## Workflow

- Run a 3-page representative sample by default.
- Rebuild pages one at a time from the PDF text layer or OCR fallback.
- Render rebuilt Word pages and compare them with the source PDF.
- Mark pages with green/yellow/red quality status.
- Write checkpoint reports so interrupted jobs can resume.
- On `--resume`, retry only failed pages and keep successful pages untouched.
- Merge the final DOCX only after all target pages succeed.
- Finalize DOCX output with exactly two safe steps: re-save the DOCX, then compress embedded bitmap images.

## Main command

```powershell
python scripts\repair_pdf_to_word.py `
  --input "input.pdf" `
  --output-dir "word_rebuild_sample" `
  --force `
  --rebuild-only
```

Use `--all-pages` only after the sample output is accepted.

See `SKILL.md` for the full operating rules and acceptance criteria.
