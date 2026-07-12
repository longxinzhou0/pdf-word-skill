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
from io import BytesIO
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


def render_pdf_page_index(pdf: Path, page_index: int, png: Path, dpi: int) -> tuple[float, float]:
    doc = fitz.open(pdf)
    page = doc[page_index]
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    pix.save(png)
    return page.rect.width, page.rect.height


def render_pdf_page(pdf: Path, png: Path, dpi: int) -> tuple[float, float]:
    return render_pdf_page_index(pdf, 0, png, dpi)


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


def analyze_pdf_page(page_pdf: Path) -> dict:
    doc = fitz.open(page_pdf)
    page = doc[0]
    text_dict = page.get_text("dict")
    text_blocks = 0
    text_lines = 0
    image_blocks = 0
    text_area = 0.0
    image_area = 0.0
    table_like_rows = 0
    for block in text_dict.get("blocks", []):
        bbox = block.get("bbox", [0, 0, 0, 0])
        area = max(0.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
        if block.get("type") == 0:
            text_blocks += 1
            lines = block.get("lines", [])
            text_lines += len(lines)
            text_area += area
            for line in lines:
                spans = [span for span in line.get("spans", []) if span.get("text", "").strip()]
                if len(spans) >= 3:
                    table_like_rows += 1
        elif block.get("type") == 1:
            image_blocks += 1
            image_area += area
    page_area = max(1.0, page.rect.width * page.rect.height)
    drawings = len(page.get_drawings())
    return {
        "text_blocks": text_blocks,
        "text_lines": text_lines,
        "image_blocks": image_blocks,
        "drawings": drawings,
        "text_area_ratio": round(text_area / page_area, 4),
        "image_area_ratio": round(image_area / page_area, 4),
        "table_like_rows": table_like_rows,
        "has_text_layer": text_lines > 0,
        "needs_ocr": text_lines == 0 or (text_lines < 3 and image_area / page_area > 0.35),
        "layout_class": classify_page_layout(text_lines, image_blocks, drawings, table_like_rows, image_area / page_area),
    }


def classify_page_layout(text_lines: int, image_blocks: int, drawings: int, table_like_rows: int, image_area_ratio: float) -> str:
    if table_like_rows >= 8 or drawings >= 80:
        return "table_or_diagram"
    if image_area_ratio > 0.45 or image_blocks >= 6:
        return "image_heavy"
    if text_lines >= 80:
        return "text_dense"
    return "mixed"


def extract_text_lines_with_ocr(page_pdf: Path, dpi: int) -> tuple[list[dict], tuple[float, float], dict]:
    lines, size = extract_text_lines(page_pdf)
    if lines:
        return lines, size, {"used": False, "reason": "pdf_text_layer_available"}
    try:
        doc = fitz.open(page_pdf)
        page = doc[0]
        text_page = page.get_textpage_ocr(flags=0, dpi=dpi, full=True)
        ocr_dict = page.get_text("dict", textpage=text_page)
        ocr_lines: list[dict] = []
        for block in ocr_dict.get("blocks", []):
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
                size_pt = max(float(s.get("size", 10)) for s in spans)
                ocr_lines.append({"text": text, "bbox": [x0, y0, x1, y1], "size": size_pt, "bold": False})
        return ocr_lines, (page.rect.width, page.rect.height), {"used": True, "ok": bool(ocr_lines), "engine": "pymupdf_tesseract", "text_lines": len(ocr_lines)}
    except Exception as exc:  # noqa: BLE001
        return [], size, {"used": True, "ok": False, "engine": "pymupdf_tesseract", "reason": "ocr_unavailable", "error": str(exc)}


def resave_docx_with_word(source_docx: Path, target_docx: Path) -> dict:
    try:
        import pythoncom
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        if source_docx != target_docx:
            shutil.copy2(source_docx, target_docx)
        return {"ok": source_docx.exists(), "used_word": False, "reason": f"pywin32 unavailable: {exc}"}

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(source_docx.resolve()), ReadOnly=False, AddToRecentFiles=False)
        doc.SaveAs2(str(target_docx.resolve()), FileFormat=16)
        doc.Close(False)
    finally:
        word.Quit()
    return {"ok": target_docx.exists(), "used_word": True}


def compress_docx_media(source_docx: Path, target_docx: Path, max_long_edge: int = 2200, jpeg_quality: int = 82) -> dict:
    source_size = source_docx.stat().st_size if source_docx.exists() else 0
    replacements = 0
    compressed_size = 0
    images: list[dict] = []
    skipped: list[dict] = []
    with zipfile.ZipFile(source_docx, "r") as zin, zipfile.ZipFile(target_docx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("word/media/"):
                suffix = Path(item.filename).suffix.lower()
                if suffix not in {".jpg", ".jpeg", ".png"}:
                    skipped.append({"name": item.filename, "reason": "non_bitmap_or_vector", "size": len(data)})
                    zout.writestr(item, data)
                    continue
                try:
                    img = Image.open(BytesIO(data))
                    img.load()
                    width, height = img.size
                    long_edge = max(width, height)
                    original_size = len(data)
                    output_width, output_height = width, height
                    if long_edge > max_long_edge:
                        scale = max_long_edge / long_edge
                        new_size = (max(1, int(round(width * scale))), max(1, int(round(height * scale))))
                        img = img.resize(new_size, Image.Resampling.LANCZOS)
                        output_width, output_height = new_size
                    buffer = BytesIO()
                    if suffix in {".jpg", ".jpeg"}:
                        if img.mode not in ("RGB", "L"):
                            img = img.convert("RGB")
                        img.save(buffer, format="JPEG", quality=jpeg_quality, optimize=True, progressive=True)
                    elif suffix == ".png":
                        img.save(buffer, format="PNG", optimize=True)
                    else:
                        img.save(buffer, format=img.format or "PNG")
                    new_data = buffer.getvalue()
                    if new_data and len(new_data) <= len(data):
                        data = new_data
                        replacements += 1
                        compressed_size += len(data)
                        images.append({
                            "name": item.filename,
                            "original_size": original_size,
                            "compressed_size": len(data),
                            "original_pixels": [width, height],
                            "output_pixels": [output_width, output_height],
                        })
                    else:
                        skipped.append({"name": item.filename, "reason": "not_smaller_after_compression", "size": original_size, "pixels": [width, height]})
                except Exception:
                    skipped.append({"name": item.filename, "reason": "image_decode_failed", "size": len(data)})
            zout.writestr(item, data)
    return {
        "ok": target_docx.exists(),
        "source_size": source_size,
        "compressed_size": target_docx.stat().st_size if target_docx.exists() else 0,
        "replacements": replacements,
        "images": images,
        "skipped": skipped,
        "policy": {
            "steps": ["resave_docx", "compress_embedded_bitmap_images"],
            "max_long_edge": max_long_edge,
            "jpeg_quality": jpeg_quality,
            "forbidden": ["delete_textboxes", "rebuild_structure", "generic_zip_optimization"],
        },
    }


def finalize_docx_output(merged_docx: Path, final_docx: Path, work_dir: Path) -> dict:
    resaved_docx = work_dir / f"{final_docx.stem}.resaved.docx"
    compressed_docx = work_dir / f"{final_docx.stem}.compressed.docx"
    resave_status = resave_docx_with_word(merged_docx, resaved_docx)
    if not resave_status.get("ok"):
        return {"ok": False, "stage": "resave", "resave_status": resave_status}
    compress_status = compress_docx_media(resaved_docx, compressed_docx)
    if not compress_status.get("ok"):
        return {"ok": False, "stage": "compress", "resave_status": resave_status, "compress_status": compress_status}
    shutil.copy2(compressed_docx, final_docx)
    return {
        "ok": final_docx.exists(),
        "resave_status": resave_status,
        "compress_status": compress_status,
        "final_size": final_docx.stat().st_size if final_docx.exists() else 0,
    }


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
    collapsed: list[str] = []
    idx = 0
    while idx < len(texts):
        current = texts[idx]
        collapsed.append(current)
        idx += 1
        while idx < len(texts) and texts[idx] == current:
            idx += 1
    texts = collapsed
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
    char_count_ratio = len(docx_text) / max(1, len(source_text))
    duplicate_ratio = max(0.0, char_count_ratio - 1.0)
    return {
        "status": "ok",
        "sequence_similarity": round(similarity, 4),
        "line_coverage": round(matched / checked, 4) if checked else None,
        "source_char_count": len(source_text),
        "docx_char_count": len(docx_text),
        "char_count_ratio": round(char_count_ratio, 4),
        "duplicate_ratio": round(duplicate_ratio, 4),
        "missing_line_count": len(missing),
        "missing_line_samples": missing[:8],
    }


def classify_quality(metrics: dict, validation: dict | None, page_analysis: dict) -> dict:
    issues: list[str] = []
    visual = metrics.get("pixel_diff_gt_30_pct")
    if isinstance(visual, (int, float)):
        if visual >= 22:
            issues.append("visual_diff_red")
        elif visual >= 12:
            issues.append("visual_diff_yellow")
    else:
        issues.append("visual_metrics_missing")

    if validation and validation.get("status") == "ok":
        coverage = validation.get("line_coverage")
        missing = validation.get("missing_line_count")
        duplicate = validation.get("duplicate_ratio")
        if isinstance(coverage, (int, float)) and coverage < 0.98:
            issues.append("text_coverage_low")
        if isinstance(missing, int) and missing > 0:
            issues.append("missing_text")
        if isinstance(duplicate, (int, float)) and duplicate > 0.35:
            issues.append("possible_duplicate_text")
    elif page_analysis.get("has_text_layer"):
        issues.append("text_validation_missing")

    red_issues = {"visual_diff_red", "text_coverage_low", "missing_text", "visual_metrics_missing", "text_validation_missing"}
    severity = "red" if any(issue in red_issues for issue in issues) else "yellow" if issues else "green"
    outcome = "failed" if severity == "red" else "success"
    return {"severity": severity, "outcome": outcome, "issues": issues}


def classify_visual_quality(metrics: dict) -> dict:
    visual = metrics.get("pixel_diff_gt_30_pct")
    if not isinstance(visual, (int, float)):
        return {"severity": "red", "outcome": "failed", "issues": ["visual_metrics_missing"]}
    if visual >= 22:
        return {"severity": "red", "outcome": "failed", "issues": ["visual_diff_red"]}
    if visual >= 12:
        return {"severity": "yellow", "outcome": "success", "issues": ["visual_diff_yellow"]}
    return {"severity": "green", "outcome": "success", "issues": []}


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

    lines, (page_w, page_h), ocr_status = extract_text_lines_with_ocr(page_pdf, dpi)
    if not lines:
        return {"ok": False, "reason": "no_text_after_pdf_text_and_ocr", "ocr_status": ocr_status}

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
    return {"ok": out_docx.exists(), "text_lines": len(lines), "ocr_status": ocr_status}


def process_page(src_pdf: Path, page_no: int, dirs: dict, cwd: Path, soffice: Path, dpi: int, wps_compatible: bool, rebuild_only: bool) -> dict:
    stem = f"page-{page_no:03d}"
    page_pdf = dirs["pages"] / f"{stem}.pdf"
    split_page(src_pdf, page_no, page_pdf)
    page_analysis = analyze_pdf_page(page_pdf)

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
                candidates.append({"kind": "pdf2docx_baseline", "docx": baseline_docx, "metrics": compare_pngs(source_png, baseline_png), "quality": {"severity": "yellow", "outcome": "success", "issues": ["baseline_text_not_validated"]}})

    rebuild_docx = dirs["rebuild_docx"] / f"{stem}.docx"
    rebuild_status = rebuild_page_with_word(page_pdf, rebuild_docx, dirs["png"], dpi)
    if rebuild_status["ok"]:
        rebuild_pdf = convert_docx_to_pdf(rebuild_docx, dirs["rendered_pdf"], cwd, soffice)
        if rebuild_pdf:
            rebuild_png = dirs["png"] / f"{stem}-rebuild.png"
            render_pdf_page(rebuild_pdf, rebuild_png, dpi)
            candidates.append({
                "kind": "text_layer_rebuild",
                "docx": rebuild_docx,
                "metrics": compare_pngs(source_png, rebuild_png),
                "text_validation": validate_docx_text_against_pdf(page_pdf, rebuild_docx),
                "status": rebuild_status,
            })
            candidates[-1]["quality"] = classify_quality(candidates[-1]["metrics"], candidates[-1]["text_validation"], page_analysis)

    if not candidates:
        return {
            "page": page_no,
            "outcome": "failed",
            "status": "failed",
            "severity": "red",
            "selected": "manual_review",
            "reason": "ocr_fail" if page_analysis.get("needs_ocr") else "no_candidate_rendered",
            "page_analysis": page_analysis,
            "baseline_status": baseline_status,
            "rebuild_status": rebuild_status,
        }

    baseline = next((c for c in candidates if c["kind"] == "pdf2docx_baseline"), None)
    rebuild = next((c for c in candidates if c["kind"] == "text_layer_rebuild"), None)
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
    quality = selected.get("quality", {"severity": "yellow", "outcome": "success", "issues": ["quality_not_classified"]})
    return {
        "page": page_no,
        "status": quality["outcome"],
        "outcome": quality["outcome"],
        "severity": quality["severity"],
        "issues": quality["issues"],
        "selected": selected["kind"],
        "selected_docx": str(selected_path),
        "page_analysis": page_analysis,
        "metrics": selected["metrics"],
        "candidates": [{k: v for k, v in c.items() if k != "docx"} for c in candidates],
    }


def merge_selected_pages(selected_dir: Path, output_docx: Path, pages: list[int]) -> bool:
    page_docx_paths = [selected_dir / f"page-{page_no:03d}.docx" for page_no in pages]
    if not page_docx_paths or any(not path.exists() for path in page_docx_paths):
        return False
    if len(page_docx_paths) == 1:
        shutil.copy2(page_docx_paths[0], output_docx)
        return True
    master = Document(str(page_docx_paths[0]))
    composer = Composer(master)
    for page_docx in page_docx_paths[1:]:
        composer.append(Document(str(page_docx)))
    composer.save(output_docx)
    return output_docx.exists()


def write_report(report: dict, path: Path) -> None:
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def load_existing_report(out: Path) -> dict | None:
    for name in ("repair_report.partial.json", "repair_report.json"):
        path = out / name
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    return None


def successful_pages_from_report(report: dict, selected_dir: Path) -> set[int]:
    completed: set[int] = set()
    for item in report.get("pages", []):
        page_no = item.get("page")
        if not isinstance(page_no, int):
            continue
        selected_docx = Path(item.get("selected_docx", selected_dir / f"page-{page_no:03d}.docx"))
        explicit_success = item.get("outcome") == "success"
        legacy_success = "outcome" not in item and item.get("selected") != "manual_review"
        if (explicit_success or legacy_success) and selected_docx.exists():
            completed.add(page_no)
    return completed


def failure_pages_from_report(report: dict, target_pages: list[int]) -> list[int]:
    indexed = {item.get("page"): item for item in report.get("pages", []) if isinstance(item.get("page"), int)}
    failures: list[int] = []
    for page_no in target_pages:
        item = indexed.get(page_no)
        if not item or item.get("outcome") != "success":
            failures.append(page_no)
    return failures


def previous_retry_count(report: dict, page_no: int) -> int:
    for item in report.get("pages", []):
        if item.get("page") == page_no:
            return int(item.get("retry_count", 0))
    return 0


def apply_final_qa_to_pages(report: dict) -> None:
    final_compare = report.get("final_render_compare", {})
    final_metrics = {item.get("source_page"): item for item in final_compare.get("metrics", [])}
    for item in report.get("pages", []):
        page_no = item.get("page")
        final_item = final_metrics.get(page_no)
        if not final_item:
            continue
        quality = final_item.get("quality", {})
        severity = quality.get("severity")
        if severity == "red":
            item["status"] = "failed"
            item["outcome"] = "failed"
            item["severity"] = "red"
            item["issues"] = sorted(set(item.get("issues", []) + ["final_render_red"]))
            item["final_render_quality"] = quality
        elif severity == "yellow" and item.get("severity") == "green":
            item["severity"] = "yellow"
            item["issues"] = sorted(set(item.get("issues", []) + ["final_render_yellow"]))
            item["final_render_quality"] = quality


def make_review_html(report: dict, out: Path) -> Path:
    rows = []
    for item in report.get("pages", []):
        page_no = item.get("page")
        stem = f"page-{page_no:03d}"
        source = Path("png_compare") / f"{stem}-source.png"
        rebuild = Path("png_compare") / f"{stem}-rebuild.png"
        metrics = item.get("metrics", {})
        validation = {}
        selected_candidate = {}
        for candidate in item.get("candidates", []):
            if candidate.get("kind") == item.get("selected"):
                selected_candidate = candidate
                validation = candidate.get("text_validation", {})
                break
        severity = item.get("severity", "red" if item.get("outcome") != "success" else "green")
        issues = ", ".join(item.get("issues", [])) or "none"
        page_analysis = item.get("page_analysis", {})
        ocr_status = selected_candidate.get("status", {}).get("ocr_status", {})
        rows.append(f"""
        <section class="{html.escape(str(severity))}">
          <h2>Page {page_no} - {html.escape(str(item.get("selected", "")))} <span>{html.escape(str(severity)).upper()}</span></h2>
          <p>layout: {html.escape(str(page_analysis.get("layout_class", "n/a")))} | issues: {html.escape(issues)}</p>
          <p>visual diff gt 30: {metrics.get("pixel_diff_gt_30_pct", "n/a")}% | mean diff: {metrics.get("mean_abs_diff", "n/a")} | line coverage: {validation.get("line_coverage", "n/a")} | duplicate ratio: {validation.get("duplicate_ratio", "n/a")} | OCR: {html.escape(str(ocr_status))}</p>
          <div class="pair">
            <figure><figcaption>Source PDF</figcaption><img src="{source.as_posix()}" /></figure>
            <figure><figcaption>Rebuilt Word render</figcaption><img src="{rebuild.as_posix()}" /></figure>
          </div>
        </section>
        """)
    body = "\n".join(rows)
    doc = f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>PDF to Word Review</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; color: #202124; }}
    h1 {{ font-size: 22px; }}
    h2 {{ font-size: 18px; margin-top: 28px; }}
    h2 span {{ font-size: 12px; padding: 2px 8px; border-radius: 999px; margin-left: 8px; }}
    p {{ color: #444; }}
    section {{ border-left: 6px solid #999; padding-left: 12px; }}
    section.green {{ border-left-color: #188038; }}
    section.yellow {{ border-left-color: #f9ab00; }}
    section.red {{ border-left-color: #d93025; }}
    section.green h2 span {{ background: #e6f4ea; color: #137333; }}
    section.yellow h2 span {{ background: #fef7e0; color: #b06000; }}
    section.red h2 span {{ background: #fce8e6; color: #c5221f; }}
    .pair {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; align-items: start; }}
    figure {{ margin: 0; border: 1px solid #ddd; padding: 8px; }}
    figcaption {{ font-size: 13px; margin-bottom: 8px; color: #555; }}
    img {{ width: 100%; height: auto; display: block; background: white; }}
  </style>
</head>
<body>
  <h1>PDF to Word Review</h1>
  <p>Source: {html.escape(report.get("source_pdf", ""))}</p>
  <p>Processed pages: {html.escape(str(report.get("processed_pages", [])))}</p>
  <p>Final DOCX accepted: {html.escape(str(report.get("final_docx_accepted", "n/a")))} | failed pages: {html.escape(str(report.get("failed_pages", [])))} | final QA failed pages: {html.escape(str(report.get("final_render_compare", {}).get("failed_pages", [])))}</p>
  {body}
</body>
</html>
"""
    path = out / "review.html"
    path.write_text(doc, encoding="utf-8")
    return path


def final_render_compare(report: dict, out: Path, cwd: Path, soffice: Path, dpi: int) -> dict:
    final_docx = Path(report["final_docx"])
    if not final_docx.exists():
        return {"ok": False, "reason": "final_docx_missing"}
    final_pdf = convert_docx_to_pdf(final_docx, out / "rendered_pdf", cwd, soffice)
    if not final_pdf:
        return {"ok": False, "reason": "final_docx_pdf_render_failed"}
    final_compare_dir = out / "final_compare"
    final_compare_dir.mkdir(parents=True, exist_ok=True)
    metrics = []
    for idx, page_no in enumerate(report.get("processed_pages", [])):
        final_png = final_compare_dir / f"merged-page-{idx + 1:03d}-from-source-{page_no:03d}.png"
        source_png = out / "png_compare" / f"page-{page_no:03d}-source.png"
        try:
            render_pdf_page_index(final_pdf, idx, final_png, dpi)
            page_metrics = compare_pngs(source_png, final_png)
            metrics.append({"source_page": page_no, "merged_page": idx + 1, "metrics": page_metrics, "quality": classify_visual_quality(page_metrics)})
        except Exception as exc:  # noqa: BLE001
            metrics.append({"source_page": page_no, "merged_page": idx + 1, "error": str(exc), "quality": {"severity": "red", "outcome": "failed", "issues": ["final_render_exception"]}})
    failed = [item["source_page"] for item in metrics if item.get("quality", {}).get("outcome") == "failed"]
    yellow = [item["source_page"] for item in metrics if item.get("quality", {}).get("severity") == "yellow"]
    return {"ok": True, "accepted": not failed, "pdf": str(final_pdf), "failed_pages": failed, "yellow_pages": yellow, "metrics": metrics}


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


def score_page_complexity(page, page_no: int) -> dict:
    text_blocks = 0
    text_lines = 0
    image_blocks = 0
    area_weight = 0.0
    drawings = len(page.get_drawings())
    table_like_rows = 0
    image_area = 0.0
    for block in page.get_text("dict").get("blocks", []):
        bbox = block.get("bbox", [0, 0, 0, 0])
        block_area = max(0.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
        area_weight += block_area / max(1.0, page.rect.width * page.rect.height)
        if block.get("type") == 0:
            text_blocks += 1
            lines = block.get("lines", [])
            text_lines += len(lines)
            table_like_rows += sum(1 for line in lines if len([span for span in line.get("spans", []) if span.get("text", "").strip()]) >= 3)
        elif block.get("type") == 1:
            image_blocks += 1
            image_area += block_area
    image_area_ratio = image_area / max(1.0, page.rect.width * page.rect.height)
    layout_class = classify_page_layout(text_lines, image_blocks, drawings, table_like_rows, image_area_ratio)
    score = text_lines * 2.0 + text_blocks * 1.5 + image_blocks * 8.0 + drawings * 0.4 + table_like_rows * 3.0 + area_weight * 20.0
    return {
        "page": page_no,
        "score": round(score, 3),
        "text_blocks": text_blocks,
        "text_lines": text_lines,
        "image_blocks": image_blocks,
        "drawings": drawings,
        "table_like_rows": table_like_rows,
        "image_area_ratio": round(image_area_ratio, 4),
        "layout_class": layout_class,
    }


def auto_select_sample_pages(src_pdf: Path, page_count: int, sample_count: int = 3) -> tuple[list[int], list[dict]]:
    doc = fitz.open(src_pdf)
    sample_count = min(sample_count, page_count)
    if sample_count <= 1:
        scores = [score_page_complexity(doc[0], 1)]
        return [1], scores

    segments: list[tuple[int, int]] = []
    for idx in range(sample_count):
        start = int(math.floor(idx * page_count / sample_count)) + 1
        end = int(math.floor((idx + 1) * page_count / sample_count))
        segments.append((start, max(start, end)))

    selected: list[int] = []
    selected_scores: list[dict] = []
    for start, end in segments:
        candidates = [score_page_complexity(doc[page_no - 1], page_no) for page_no in range(start, end + 1)]
        best = max(candidates, key=lambda item: (item["score"], item["page"]))
        selected.append(best["page"])
        selected_scores.append({**best, "segment": [start, end]})

    unique_pages = sorted(set(selected))
    all_scores = [score_page_complexity(doc[page_no - 1], page_no) for page_no in range(1, page_count + 1)]
    desired_classes = ["image_heavy", "table_or_diagram", "text_dense"]
    for layout_class in desired_classes:
        if len(unique_pages) >= sample_count:
            break
        if any(item["layout_class"] == layout_class for item in selected_scores):
            continue
        class_candidates = [item for item in all_scores if item["layout_class"] == layout_class and item["page"] not in unique_pages]
        if class_candidates:
            best = max(class_candidates, key=lambda item: (item["score"], item["page"]))
            unique_pages.append(best["page"])
            selected_scores.append({**best, "segment": "layout_coverage"})
    if len(unique_pages) < sample_count:
        for item in sorted(all_scores, key=lambda value: (value["score"], value["page"]), reverse=True):
            if item["page"] not in unique_pages:
                unique_pages.append(item["page"])
            if len(unique_pages) == sample_count:
                break
    return sorted(unique_pages), selected_scores


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--single-page", type=int)
    parser.add_argument("--pages", help='Page range like "1-5,8,10-12".')
    parser.add_argument("--all-pages", action="store_true", help="Process the full document. Without this or a page range, only three representative sample pages are processed.")
    parser.add_argument("--force", action="store_true", help="Delete an existing output directory before running.")
    parser.add_argument("--resume", action="store_true", help="Resume from an existing output directory and retry only failed pages from the checkpoint report.")
    parser.add_argument("--wps-compatible", action="store_true", help="Prefer rebuilt editable text pages to avoid pdf2docx font-encoding garbage in WPS.")
    parser.add_argument("--rebuild-only", action="store_true", help="Skip pdf2docx and rebuild every page from PDF text/OCR plus recovered imagery.")
    parser.add_argument("--dpi", type=int, default=180)
    args = parser.parse_args()

    src_pdf = args.input.resolve()
    out = args.output_dir.resolve()
    cwd = Path.cwd()
    if args.resume and args.force:
        raise ValueError("Use either --resume or --force, not both.")
    if out.exists():
        if args.resume:
            pass
        elif args.force:
            shutil.rmtree(out)
        else:
            raise FileExistsError(f"Output directory already exists: {out}. Use --resume to continue or --force to replace it.")
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
    existing_report = load_existing_report(out) if args.resume else None
    if args.resume and not existing_report:
        raise FileNotFoundError(f"No checkpoint report found in {out}. Use --force for a fresh run.")
    explicit_modes = [bool(args.single_page), bool(args.pages), bool(args.all_pages)]
    if sum(explicit_modes) > 1:
        raise ValueError("Use only one of --single-page, --pages, or --all-pages.")
    auto_sample = False
    auto_sample_scores = []
    if args.single_page:
        if not 1 <= args.single_page <= page_count:
            raise ValueError(f"--single-page must be between 1 and {page_count}; got {args.single_page}.")
        pages = [args.single_page]
    elif args.pages:
        pages = parse_pages(args.pages, page_count)
    elif args.all_pages:
        pages = list(range(1, page_count + 1))
    elif args.resume and existing_report and existing_report.get("processed_pages"):
        pages = [int(page_no) for page_no in existing_report["processed_pages"]]
        auto_sample = bool(existing_report.get("auto_sample", False))
        auto_sample_scores = existing_report.get("auto_sample_scores", [])
    else:
        auto_sample = True
        pages, auto_sample_scores = auto_select_sample_pages(src_pdf, page_count)
    final_docx = out / f"{src_pdf.stem}_editable_repaired.docx"
    if existing_report:
        report = existing_report
        report.update({
            "source_pdf": str(src_pdf),
            "final_docx": str(final_docx),
            "page_count": page_count,
            "processed_pages": pages,
            "auto_sample": auto_sample,
            "auto_sample_scores": auto_sample_scores or report.get("auto_sample_scores", []),
            "resumed": True,
        })
    else:
        report = {
            "source_pdf": str(src_pdf),
            "final_docx": str(final_docx),
            "page_count": page_count,
            "processed_pages": pages,
            "auto_sample": auto_sample,
            "auto_sample_scores": auto_sample_scores,
            "resumed": False,
            "pages": [],
        }
    partial_report = out / "repair_report.partial.json"
    successful_pages = successful_pages_from_report(report, dirs["selected_docx"]) if args.resume else set()
    pages_to_run = failure_pages_from_report(report, pages) if args.resume else pages
    if args.resume and not pages_to_run:
        print("No failed pages to retry; all target pages are already successful.", flush=True)
    for page_no in pages_to_run:
        if page_no in successful_pages:
            print(f"Skipping completed page {page_no}/{page_count}", flush=True)
            continue
        print(f"Processing page {page_no}/{page_count}", flush=True)
        retry_count = previous_retry_count(report, page_no) + (1 if args.resume else 0)
        try:
            page_report = process_page(src_pdf, page_no, dirs, cwd, soffice, args.dpi, args.wps_compatible, args.rebuild_only)
        except Exception as exc:  # noqa: BLE001
            page_report = {"page": page_no, "status": "failed", "outcome": "failed", "severity": "red", "selected": "manual_review", "reason": "exception", "error": str(exc)}
        page_report["retry_count"] = retry_count
        report["pages"] = [item for item in report.get("pages", []) if item.get("page") != page_no]
        report["pages"].append(page_report)
        report["pages"].sort(key=lambda item: item.get("page", 0))
        write_report(report, partial_report)

    successful_pages = successful_pages_from_report(report, dirs["selected_docx"])
    failed_pages = [page_no for page_no in pages if page_no not in successful_pages]
    report["failed_pages"] = failed_pages
    if failed_pages:
        report["final_docx_created"] = False
        report["final_docx_accepted"] = False
        report["final_docx_reason"] = "failed_pages_remaining"
        report["final_render_compare"] = {"ok": False, "accepted": False, "reason": "failed_pages_remaining", "failed_pages": failed_pages}
        report["final_docx_compression"] = {"ok": False, "reason": "final_docx_not_created"}
    else:
        merged_docx = out / f"{src_pdf.stem}_merged_uncompressed.docx"
        merged_ok = merge_selected_pages(dirs["selected_docx"], merged_docx, pages)
        report["final_docx_created"] = False
        if merged_ok:
            compression_status = finalize_docx_output(merged_docx, final_docx, out)
            report["final_docx_compression"] = compression_status
            report["final_docx_created"] = bool(compression_status.get("ok"))
        else:
            report["final_docx_compression"] = {"ok": False, "reason": "merge_failed"}
        report["final_render_compare"] = final_render_compare(report, out, cwd, soffice, args.dpi) if report["final_docx_created"] else {"ok": False, "accepted": False, "reason": "final_docx_not_created"}
        report["final_docx_accepted"] = bool(report["final_render_compare"].get("accepted", False))
        apply_final_qa_to_pages(report)
        successful_pages = successful_pages_from_report(report, dirs["selected_docx"])
        report["failed_pages"] = [page_no for page_no in pages if page_no not in successful_pages]
    report["review_html"] = str(make_review_html(report, out))
    write_report(report, out / "repair_report.json")
    if partial_report.exists():
        partial_report.unlink()
    print(out / "repair_report.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
