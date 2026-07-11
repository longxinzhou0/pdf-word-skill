from __future__ import annotations

import argparse
import difflib
import html
import json
import math
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import fitz
import numpy as np
from docx import Document
from docxcompose.composer import Composer
from PIL import Image, ImageChops, ImageStat


SOFFICE_CANDIDATES = [
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
]


def run(cmd: list[str], cwd: Path, timeout: int = 180) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, cwd=cwd, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, timeout=timeout)


def find_soffice() -> Path:
    for path in SOFFICE_CANDIDATES:
        if path.exists():
            return path
    found = shutil.which("soffice")
    if found:
        return Path(found)
    raise FileNotFoundError("LibreOffice soffice not found.")


def split_page(src_pdf: Path, page_no: int, out_pdf: Path) -> None:
    doc = fitz.open(src_pdf)
    single = fitz.open()
    single.insert_pdf(doc, from_page=page_no - 1, to_page=page_no - 1)
    single.save(out_pdf)


def render_pdf_page(pdf: Path, png: Path, dpi: int) -> tuple[float, float]:
    doc = fitz.open(pdf)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    pix.save(png)
    return page.rect.width, page.rect.height


def compare_pngs(source_png: Path, candidate_png: Path) -> dict:
    src = Image.open(source_png).convert("RGB")
    cand = Image.open(candidate_png).convert("RGB")
    if cand.size != src.size:
        cand = cand.resize(src.size, Image.Resampling.LANCZOS)
    diff = ImageChops.difference(src, cand)
    stat = ImageStat.Stat(diff)
    arr = np.array(diff)
    return {
        "mean_abs_diff": round(sum(stat.mean) / 3, 3),
        "rms_diff": round(math.sqrt(sum(v * v for v in stat.rms) / 3), 3),
        "pixel_diff_gt_30_pct": round(float(((arr > 30).any(axis=2)).mean() * 100), 3),
    }


def convert_docx_to_pdf(docx: Path, out_dir: Path, cwd: Path, soffice: Path) -> Path | None:
    result = run([str(soffice), "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)], cwd, timeout=180)
    expected = out_dir / f"{docx.stem}.pdf"
    if expected.exists():
        return expected
    if result.returncode != 0:
        return None
    matches = sorted(out_dir.glob(f"{docx.stem}*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0] if matches else None


def convert_with_pdf2docx(page_pdf: Path, out_docx: Path, cwd: Path) -> dict:
    result = run(["pdf2docx", "convert", str(page_pdf), str(out_docx)], cwd, timeout=180)
    return {"ok": result.returncode == 0 and out_docx.exists(), "stdout": result.stdout[-1500:], "stderr": result.stderr[-1500:]}


def extract_text_lines(page_pdf: Path) -> tuple[list[dict], tuple[float, float]]:
    doc = fitz.open(page_pdf)
    page = doc[0]
    lines: list[dict] = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = [s for s in line.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            text = "".join(s["text"] for s in spans).strip()
            if not text:
                continue
            x0 = min(s["bbox"][0] for s in spans)
            y0 = min(s["bbox"][1] for s in spans)
            x1 = max(s["bbox"][2] for s in spans)
            y1 = max(s["bbox"][3] for s in spans)
            size = max(float(s.get("size", 10)) for s in spans)
            bold = any("bold" in s.get("font", "").lower() or "black" in s.get("font", "").lower() for s in spans)
            lines.append({"text": text, "bbox": [x0, y0, x1, y1], "size": size, "bold": bold})
    return lines, (page.rect.width, page.rect.height)


def normalize_text(value: str) -> str:
    value = value.replace("\u00a0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip().lower()


def canonical_match_text(value: str) -> str:
    value = normalize_text(value)
    value = re.sub(r"[^0-9a-z]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def extract_docx_text(docx_path: Path) -> str:
    if not docx_path.exists():
        return ""
    texts: list[str] = []
    with zipfile.ZipFile(docx_path) as archive:
        for name in archive.namelist():
            is_body = name == "word/document.xml"
            is_header_footer = name.startswith("word/header") or name.startswith("word/footer")
            if is_body or is_header_footer:
                data = archive.read(name).decode("utf-8", errors="ignore")
                texts.extend(re.findall(r"<w:t[^>]*>(.*?)</w:t>", data))
    return " ".join(html.unescape(re.sub(r"<[^>]+>", "", t)) for t in texts)


def validate_docx_text_against_pdf(page_pdf: Path, docx_path: Path) -> dict:
    lines, _ = extract_text_lines(page_pdf)
    source_lines = [line["text"] for line in lines if line["text"].strip()]
    source_text = normalize_text(" ".join(source_lines))
    docx_text = normalize_text(extract_docx_text(docx_path))
    docx_canonical = canonical_match_text(docx_text)
    if not source_text:
        return {"status": "no_pdf_text", "similarity": None, "missing_line_count": None}
    similarity = difflib.SequenceMatcher(None, source_text, docx_text).ratio()
    missing = []
    matched = 0
    for line in source_lines:
        norm_line = canonical_match_text(line)
        if len(norm_line) < 8:
            continue
        if norm_line in docx_canonical:
            matched += 1
        else:
            missing.append(line)
    checked = matched + len(missing)
    return {
        "status": "ok",
        "similarity": round(similarity, 4),
        "line_coverage": round(matched / checked, 4) if checked else None,
        "source_char_count": len(source_text),
        "docx_char_count": len(docx_text),
        "missing_line_count": len(missing),
        "missing_line_samples": missing[:8],
    }


def make_redacted_background(page_pdf: Path, lines: list[dict], out_png: Path, dpi: int) -> None:
    doc = fitz.open(page_pdf)
    page = doc[0]
    for item in lines:
        x0, y0, x1, y1 = item["bbox"]
        page.add_redact_annot(fitz.Rect(max(0, x0 - 1.5), max(0, y0 - 1.5), x1 + 1.5, y1 + 1.5), fill=(1, 1, 1))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    pix.save(out_png)


def rebuild_page_with_word(page_pdf: Path, out_docx: Path, work_dir: Path, dpi: int) -> dict:
    try:
        import pythoncom
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "reason": f"pywin32 unavailable: {exc}"}

    lines, (page_w, page_h) = extract_text_lines(page_pdf)
    if not lines:
        return {"ok": False, "reason": "no_pdf_text_layer; OCR fallback not configured"}

    bg_png = work_dir / f"{out_docx.stem}-redacted-bg.png"
    make_redacted_background(page_pdf, lines, bg_png, dpi)
    font_scale = 0.95 if len(lines) > 15 else 0.82

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Add()
        setup = doc.PageSetup
        setup.PageWidth = page_w
        setup.PageHeight = page_h
        setup.TopMargin = setup.BottomMargin = setup.LeftMargin = setup.RightMargin = 0
        inline = doc.InlineShapes.AddPicture(str(bg_png.resolve()), False, True)
        inline.Width = page_w
        inline.Height = page_h
        for item in lines:
            x0, y0, x1, y1 = item["bbox"]
            width = max(18, (x1 - x0) * 1.35 + 24)
            height = max(14, (y1 - y0) * 1.8 + 10)
            shape = doc.Shapes.AddTextbox(1, max(0, x0 - 4), max(0, y0 - 4), width, height)
            shape.Fill.Visible = False
            shape.Line.Visible = False
            shape.TextFrame.MarginLeft = 0
            shape.TextFrame.MarginRight = 0
            shape.TextFrame.MarginTop = 0
            shape.TextFrame.MarginBottom = 0
            rng = shape.TextFrame.TextRange
            rng.Text = item["text"]
            rng.Font.Name = "Arial"
            rng.Font.Size = max(4.5, min(32, item["size"] * font_scale))
            rng.Font.Bold = -1 if item["bold"] else 0
        doc.SaveAs2(str(out_docx.resolve()), FileFormat=16)
        doc.Close(False)
    finally:
        word.Quit()
    return {"ok": out_docx.exists(), "text_lines": len(lines)}


def process_page(src_pdf: Path, page_no: int, dirs: dict, cwd: Path, soffice: Path, dpi: int, wps_compatible: bool, rebuild_only: bool) -> dict:
    stem = f"page-{page_no:03d}"
    page_pdf = dirs["pages"] / f"{stem}.pdf"
    split_page(src_pdf, page_no, page_pdf)

    source_png = dirs["png"] / f"{stem}-source.png"
    render_pdf_page(page_pdf, source_png, dpi)

    candidates: list[dict] = []
    baseline_status = {"ok": False, "skipped": rebuild_only}
    if not rebuild_only:
        baseline_docx = dirs["baseline_docx"] / f"{stem}.docx"
        baseline_status = convert_with_pdf2docx(page_pdf, baseline_docx, cwd)
        if baseline_status["ok"]:
            baseline_pdf = convert_docx_to_pdf(baseline_docx, dirs["rendered_pdf"], cwd, soffice)
            if baseline_pdf:
                baseline_png = dirs["png"] / f"{stem}-baseline.png"
                render_pdf_page(baseline_pdf, baseline_png, dpi)
                candidates.append({"kind": "pdf2docx_baseline", "docx": baseline_docx, "metrics": compare_pngs(source_png, baseline_png)})

    rebuild_docx = dirs["rebuild_docx"] / f"{stem}.docx"
    rebuild_status = rebuild_page_with_word(page_pdf, rebuild_docx, dirs["png"], dpi)
    if rebuild_status["ok"]:
        rebuild_pdf = convert_docx_to_pdf(rebuild_docx, dirs["rendered_pdf"], cwd, soffice)
        if rebuild_pdf:
            rebuild_png = dirs["png"] / f"{stem}-rebuild.png"
            render_pdf_page(rebuild_pdf, rebuild_png, dpi)
            candidates.append({
                "kind": "ocr_rebuild",
                "docx": rebuild_docx,
                "metrics": compare_pngs(source_png, rebuild_png),
                "text_validation": validate_docx_text_against_pdf(page_pdf, rebuild_docx),
                "status": rebuild_status,
            })

    if not candidates:
        return {"page": page_no, "selected": "manual_review", "reason": "no_candidate_rendered", "baseline_status": baseline_status, "rebuild_status": rebuild_status}

    baseline = next((c for c in candidates if c["kind"] == "pdf2docx_baseline"), None)
    rebuild = next((c for c in candidates if c["kind"] == "ocr_rebuild"), None)
    if wps_compatible and rebuild:
        selected = rebuild
    elif baseline and rebuild:
        baseline_diff = baseline["metrics"]["pixel_diff_gt_30_pct"]
        rebuild_diff = rebuild["metrics"]["pixel_diff_gt_30_pct"]
        # Keep the pdf2docx page unless rebuilding clearly fixes a major visual problem.
        # Small metric wins often come from text being too small after reconstruction.
        selected = rebuild if baseline_diff - rebuild_diff >= 5.0 else baseline
    else:
        selected = min(candidates, key=lambda c: (c["metrics"]["pixel_diff_gt_30_pct"], c["metrics"]["mean_abs_diff"]))
    selected_path = dirs["selected_docx"] / f"{stem}.docx"
    shutil.copy2(selected["docx"], selected_path)
    return {"page": page_no, "selected": selected["kind"], "selected_docx": str(selected_path), "metrics": selected["metrics"], "candidates": [{k: v for k, v in c.items() if k != "docx"} for c in candidates]}


def merge_selected_pages(selected_dir: Path, output_docx: Path) -> bool:
    pages = sorted(selected_dir.glob("page-*.docx"))
    if not pages:
        return False
    if len(pages) == 1:
        shutil.copy2(pages[0], output_docx)
        return True
    master = Document(str(pages[0]))
    composer = Composer(master)
    for page_docx in pages[1:]:
        composer.append(Document(str(page_docx)))
    composer.save(output_docx)
    return output_docx.exists()


def parse_pages(spec: str, page_count: int) -> list[int]:
    pages: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start, end = int(start_s), int(end_s)
            if start > end:
                start, end = end, start
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    result = sorted(p for p in pages if 1 <= p <= page_count)
    if not result:
        raise ValueError(f"No valid pages in range {spec!r}; document has {page_count} pages.")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--single-page", type=int)
    parser.add_argument("--pages", help='Page range like "1-5,8,10-12".')
    parser.add_argument("--wps-compatible", action="store_true", help="Prefer rebuilt editable text pages to avoid pdf2docx font-encoding garbage in WPS.")
    parser.add_argument("--rebuild-only", action="store_true", help="Skip pdf2docx and rebuild every page from PDF text/OCR plus recovered imagery.")
    parser.add_argument("--dpi", type=int, default=180)
    args = parser.parse_args()

    src_pdf = args.input.resolve()
    out = args.output_dir.resolve()
    cwd = Path.cwd()
    if out.exists():
        shutil.rmtree(out)
    dirs = {
        "pages": out / "pdf_pages",
        "baseline_docx": out / "docx_pdf2docx",
        "rebuild_docx": out / "docx_rebuilt",
        "selected_docx": out / "docx_selected",
        "rendered_pdf": out / "rendered_pdf",
        "png": out / "png_compare",
    }
    for path in dirs.values():
        path.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    page_count = fitz.open(src_pdf).page_count
    if args.single_page and args.pages:
        raise ValueError("Use either --single-page or --pages, not both.")
    if args.single_page:
        pages = [args.single_page]
    elif args.pages:
        pages = parse_pages(args.pages, page_count)
    else:
        pages = list(range(1, page_count + 1))
    final_docx = out / f"{src_pdf.stem}_editable_repaired.docx"
    report = {"source_pdf": str(src_pdf), "final_docx": str(final_docx), "pages": []}
    for page_no in pages:
        print(f"Processing page {page_no}/{page_count}", flush=True)
        report["pages"].append(process_page(src_pdf, page_no, dirs, cwd, soffice, args.dpi, args.wps_compatible, args.rebuild_only))

    report["final_docx_created"] = merge_selected_pages(dirs["selected_docx"], final_docx)
    (out / "repair_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(out / "repair_report.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
