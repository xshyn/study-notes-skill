#!/usr/bin/env python3
"""
render_preview.py — visual QA for the generated files.

Converts .docx / .pptx to PDF with LibreOffice, rasterises pages, and builds contact sheets
(grids of thumbnails) so you can eyeball layout problems: text overflowing boxes, cramped
slides, broken tables, missing figures. Also reports a quick overflow heuristic for decks.

  python render_preview.py notes.docx deck.pptx -o preview/
  # → preview/<name>.pdf, preview/<name>_sheet1.png, ...

Requires LibreOffice (soffice) and pdftoppm (poppler). If they are missing the script says
so and exits 0 — preview is a quality aid, not a hard requirement.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def to_pdf(src: Path, outdir: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[preview] LibreOffice not found — skipping render.")
        return None
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
                   capture_output=True, timeout=300)
    pdf = outdir / (src.stem + ".pdf")
    if not pdf.exists():
        return None
    named = outdir / f"{src.stem}.{src.suffix.lstrip('.').lower()}.pdf"
    pdf.replace(named)
    return named


def rasterise(pdf: Path, outdir: Path, dpi: int) -> list[Path]:
    if not shutil.which("pdftoppm"):
        print("[preview] pdftoppm not found — PDF written but no PNGs.")
        return []
    prefix = outdir / f"{pdf.stem}_p"
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(prefix)], capture_output=True, timeout=300)
    return sorted(outdir.glob(f"{pdf.stem}_p-*.png"))


def contact_sheets(pages: list[Path], outdir: Path, stem: str, cols: int, per_sheet: int) -> list[Path]:
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return []
    sheets = []
    for si in range(0, len(pages), per_sheet):
        imgs = [Image.open(p).convert("RGB") for p in pages[si:si + per_sheet]]
        tw = 520
        th = int(tw * imgs[0].height / imgs[0].width)
        rows = (len(imgs) + cols - 1) // cols
        sheet = Image.new("RGB", (cols * (tw + 16) + 16, rows * (th + 36) + 16), "#d9dce3")
        d = ImageDraw.Draw(sheet)
        for i, im in enumerate(imgs):
            r, c = divmod(i, cols)
            x, y = 16 + c * (tw + 16), 16 + r * (th + 36)
            sheet.paste(im.resize((tw, th)), (x, y + 20))
            d.text((x, y + 4), f"#{si + i + 1}", fill="#222")
        out = outdir / f"{stem}_sheet{si // per_sheet + 1}.png"
        sheet.save(out)
        sheets.append(out)
    return sheets


def pptx_overflow_report(path: Path) -> list[str]:
    """Heuristic: flag text boxes whose estimated text height exceeds the box."""
    try:
        from pptx import Presentation
        from pptx.util import Emu
    except ImportError:
        return []
    import math
    issues = []
    prs = Presentation(str(path))
    for si, slide in enumerate(prs.slides, 1):
        for shp in slide.shapes:
            if not shp.has_text_frame or not shp.text_frame.text.strip():
                continue
            w = Emu(shp.width).inches - 0.1
            h = Emu(shp.height).inches
            total = 0.0
            for p in shp.text_frame.paragraphs:
                size = None
                for r in p.runs:
                    if r.font.size:
                        size = r.font.size.pt
                        break
                size = size or 18
                txt = "".join(r.text for r in p.runs)
                cpl = max(6, w * 72 / (size * 0.5))
                total += max(1, math.ceil(len(txt) / cpl)) * size * 1.35 / 72
            if total > h * 1.15 and total - h > 0.25:
                issues.append(f"slide {si}: text may overflow ({total:.1f}in needed vs {h:.1f}in box): "
                              f"{shp.text_frame.text[:60]!r}")
    return issues


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--outdir", default="preview")
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--cols", type=int, default=4)
    ap.add_argument("--per-sheet", type=int, default=16)
    a = ap.parse_args()
    outdir = Path(a.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    for f in a.files:
        src = Path(f)
        if src.suffix.lower() == ".pptx":
            issues = pptx_overflow_report(src)
            print(f"[preview] {src.name}: {len(issues)} possible overflow(s)")
            for i in issues[:20]:
                print("   ", i)
        pdf = src if src.suffix.lower() == ".pdf" else to_pdf(src, outdir)
        if not pdf:
            continue
        pages = rasterise(pdf, outdir, a.dpi)
        sheets = contact_sheets(pages, outdir, pdf.stem.replace(".", "_"), a.cols, a.per_sheet)
        for p in pages:
            p.unlink(missing_ok=True) if sheets else None
        print(f"[preview] {src.name}: {len(pages)} pages -> {pdf}")
        for s in sheets:
            print(f"    contact sheet: {s}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
