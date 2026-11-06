#!/usr/bin/env python3
"""
extract_source.py — turn a learning resource into clean, structured plain text.

Supported inputs
  * Local files: .pdf .docx .pptx .epub .html/.htm .md .txt .rst .ipynb .srt .vtt
  * URLs: web pages (article extraction), direct PDF links, YouTube videos (transcript)

Output
  A UTF-8 text file with page / slide / section markers ("=== Page 12 ===") so the
  note-taker can cite locations, plus a short stats header printed to stderr.
  With --images DIR, figures embedded in PDFs/DOCX/PPTX are saved for reuse in notes.

Every third-party library is optional: the script tries the best extractor first and
falls back to simpler ones, so it works on a bare Python install for most formats.

Usage
  python extract_source.py SOURCE [-o out.txt] [--images DIR] [--pages 1-40]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from html import unescape
from pathlib import Path

UA = "Mozilla/5.0 (study-notes skill; +https://github.com/xshyn/study-notes-skill)"


# ----------------------------------------------------------------------------- helpers
def _clean(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\u00ad", "")
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)          # de-hyphenate line breaks
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _parse_pages(spec: str | None, n: int) -> list[int]:
    if not spec:
        return list(range(n))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a) - 1, min(int(b), n)))
        elif part:
            out.append(int(part) - 1)
    return [p for p in out if 0 <= p < n]


def _html_to_text(html: str) -> str:
    try:
        from bs4 import BeautifulSoup  # type: ignore
        soup = BeautifulSoup(html, "html.parser")
        for t in soup(["script", "style", "nav", "footer", "header", "aside", "form", "noscript"]):
            t.decompose()
        root = soup.find("article") or soup.find("main") or soup.body or soup
        lines = []
        for el in root.find_all(["h1", "h2", "h3", "h4", "p", "li", "pre", "code", "td", "th", "blockquote", "dt", "dd"]):
            if el.name == "code" and el.find_parent("pre"):
                continue
            txt = el.get_text(" " if el.name != "pre" else "", strip=el.name != "pre")
            if not txt.strip():
                continue
            if el.name in ("h1", "h2", "h3", "h4"):
                lines.append("\n" + "#" * int(el.name[1]) + " " + txt.strip())
            elif el.name == "li":
                lines.append("- " + txt)
            elif el.name == "pre":
                lines.append("```\n" + txt.rstrip() + "\n```")
            else:
                lines.append(txt)
        # de-duplicate consecutive repeats (nested elements)
        dedup = []
        for ln in lines:
            if not dedup or dedup[-1] != ln:
                dedup.append(ln)
        return "\n".join(dedup)
    except ImportError:
        html = re.sub(r"(?is)<(script|style).*?</\1>", "", html)
        html = re.sub(r"(?i)<br\s*/?>|</p>|</h\d>|</li>", "\n", html)
        return unescape(re.sub(r"<[^>]+>", " ", html))


# ----------------------------------------------------------------------------- local formats
def extract_pdf(path: Path, pages: str | None, img_dir: Path | None) -> str:
    chunks: list[str] = []
    try:
        import pymupdf as fitz  # type: ignore
    except ImportError:
        try:
            import fitz  # type: ignore
        except ImportError:
            fitz = None
    if fitz is not None:
        doc = fitz.open(str(path))
        for i in _parse_pages(pages, doc.page_count):
            page = doc[i]
            chunks.append(f"=== Page {i + 1} ===\n{page.get_text('text')}")
            if img_dir:
                for j, img in enumerate(page.get_images(full=True)):
                    try:
                        pix = fitz.Pixmap(doc, img[0])
                        if pix.width < 120 or pix.height < 120:
                            continue
                        if pix.n - pix.alpha >= 4:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        pix.save(str(img_dir / f"page{i + 1:03d}_img{j + 1}.png"))
                    except Exception:
                        pass
        return "\n\n".join(chunks)
    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(str(path)) as pdf:
            for i in _parse_pages(pages, len(pdf.pages)):
                chunks.append(f"=== Page {i + 1} ===\n{pdf.pages[i].extract_text() or ''}")
        return "\n\n".join(chunks)
    except ImportError:
        pass
    try:
        from pypdf import PdfReader  # type: ignore
        r = PdfReader(str(path))
        for i in _parse_pages(pages, len(r.pages)):
            chunks.append(f"=== Page {i + 1} ===\n{r.pages[i].extract_text() or ''}")
        return "\n\n".join(chunks)
    except ImportError:
        pass
    import subprocess
    return subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True, text=True).stdout


def extract_docx(path: Path, img_dir: Path | None) -> str:
    try:
        import docx  # type: ignore
        d = docx.Document(str(path))
        out = []
        for p in d.paragraphs:
            st = (p.style.name or "").lower() if p.style is not None else ""
            t = p.text.strip()
            if not t:
                continue
            m = re.match(r"heading (\d)", st)
            out.append(("#" * int(m.group(1)) + " " + t) if m else (("- " + t) if "list" in st else t))
        for ti, table in enumerate(d.tables):
            out.append(f"\n[Table {ti + 1}]")
            for row in table.rows:
                out.append(" | ".join(c.text.strip() for c in row.cells))
        text = "\n".join(out)
    except ImportError:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf8", "ignore")
        text = unescape(re.sub(r"<[^>]+>", "", xml.replace("</w:p>", "\n")))
    if img_dir:
        _dump_zip_media(path, "word/media/", img_dir)
    return text


def extract_pptx(path: Path, img_dir: Path | None) -> str:
    try:
        from pptx import Presentation  # type: ignore
        prs = Presentation(str(path))
        out = []
        for i, slide in enumerate(prs.slides, 1):
            out.append(f"=== Slide {i} ===")
            for shp in slide.shapes:
                if shp.has_text_frame:
                    for para in shp.text_frame.paragraphs:
                        t = "".join(r.text for r in para.runs).strip()
                        if t:
                            out.append("  " * para.level + "- " + t)
                if getattr(shp, "has_table", False) and shp.has_table:
                    for row in shp.table.rows:
                        out.append(" | ".join(c.text for c in row.cells))
            if slide.has_notes_slide:
                n = slide.notes_slide.notes_text_frame.text.strip()
                if n:
                    out.append("[Speaker notes] " + n)
        text = "\n".join(out)
    except ImportError:
        text = ""
    if img_dir:
        _dump_zip_media(path, "ppt/media/", img_dir)
    return text


def extract_epub(path: Path) -> str:
    out = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.lower().endswith((".xhtml", ".html", ".htm"))]
        for n in sorted(names):
            out.append(f"=== Section {Path(n).stem} ===\n" + _html_to_text(z.read(n).decode("utf8", "ignore")))
    return "\n\n".join(out)


def extract_ipynb(path: Path) -> str:
    nb = json.loads(path.read_text(encoding="utf8"))
    out = []
    for c in nb.get("cells", []):
        src = "".join(c.get("source", []))
        out.append(src if c.get("cell_type") == "markdown" else f"```python\n{src}\n```")
    return "\n\n".join(out)


def extract_subtitles(path: Path) -> str:
    raw = path.read_text(encoding="utf8", errors="ignore")
    lines = [ln for ln in raw.splitlines()
             if ln.strip() and not re.match(r"^(\d+|WEBVTT.*|\d\d:\d\d.*-->.*)$", ln.strip())]
    return " ".join(dict.fromkeys(lines))


def _dump_zip_media(path: Path, prefix: str, img_dir: Path) -> None:
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith(prefix) and n.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")):
                (img_dir / Path(n).name).write_bytes(z.read(n))


# ----------------------------------------------------------------------------- URLs
def _youtube_id(url: str) -> str | None:
    m = re.search(r"(?:v=|youtu\.be/|/shorts/|/embed/)([A-Za-z0-9_-]{11})", url)
    return m.group(1) if m else None


def extract_url(url: str, pages: str | None, img_dir: Path | None) -> str:
    vid = _youtube_id(url) if ("youtube" in url or "youtu.be" in url) else None
    if vid:
        try:
            from youtube_transcript_api import YouTubeTranscriptApi  # type: ignore
            try:
                fetched = YouTubeTranscriptApi().fetch(vid)
                segs = [(s.start, s.text) for s in fetched]
            except AttributeError:  # older API
                segs = [(s["start"], s["text"]) for s in YouTubeTranscriptApi.get_transcript(vid)]
            out, last = [], -60.0
            for start, txt in segs:
                if start - last >= 60:
                    out.append(f"\n=== [{int(start // 60):02d}:{int(start % 60):02d}] ===")
                    last = start
                out.append(txt)
            return " ".join(out)
        except Exception as e:  # noqa: BLE001
            raise SystemExit(f"Could not fetch YouTube transcript ({e}). Ask the user for the transcript "
                             f"or use your own web/video tools.")
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=45) as r:
            data, ctype = r.read(), r.headers.get("Content-Type", "")
    except Exception as e:  # noqa: BLE001
        raise SystemExit(
            f"NETWORK_BLOCKED: could not download {url} ({e}).\n"
            "Fallback: fetch the page with your own web tool (e.g. WebFetch / browser), save the full "
            "text to a .md/.txt file, and run this script on that file — or ask the user to attach it.")
    if "pdf" in ctype or url.lower().split("?")[0].endswith(".pdf") or data[:4] == b"%PDF":
        tmp = Path(os.environ.get("TMPDIR", "/tmp")) / "study_notes_download.pdf"
        tmp.write_bytes(data)
        return extract_pdf(tmp, pages, img_dir)
    html = data.decode("utf8", "ignore")
    try:
        import trafilatura  # type: ignore
        txt = trafilatura.extract(html, include_tables=True, include_formatting=True,
                                  include_links=False, output_format="markdown")
        if txt and len(txt) > 400:
            return txt
    except ImportError:
        pass
    return _html_to_text(html)


# ----------------------------------------------------------------------------- main
def extract(source: str, pages: str | None = None, img_dir: Path | None = None) -> str:
    if re.match(r"^https?://", source):
        return extract_url(source, pages, img_dir)
    p = Path(source).expanduser()
    if not p.exists():
        raise SystemExit(f"File not found: {p}")
    ext = p.suffix.lower()
    if ext == ".pdf":
        return extract_pdf(p, pages, img_dir)
    if ext in (".docx", ".dotx"):
        return extract_docx(p, img_dir)
    if ext == ".pptx":
        return extract_pptx(p, img_dir)
    if ext == ".epub":
        return extract_epub(p)
    if ext == ".ipynb":
        return extract_ipynb(p)
    if ext in (".srt", ".vtt"):
        return extract_subtitles(p)
    if ext in (".html", ".htm"):
        return _html_to_text(p.read_text(encoding="utf8", errors="ignore"))
    return p.read_text(encoding="utf8", errors="ignore")  # md, txt, rst, code …


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", help="file path or URL")
    ap.add_argument("-o", "--output", help="write text here (default: stdout)")
    ap.add_argument("--images", help="directory to save embedded figures")
    ap.add_argument("--pages", help="page range for PDFs, e.g. 1-30,45")
    a = ap.parse_args()
    img_dir = Path(a.images) if a.images else None
    if img_dir:
        img_dir.mkdir(parents=True, exist_ok=True)
    text = _clean(extract(a.source, a.pages, img_dir))
    words = len(text.split())
    markers = len(re.findall(r"^=== ", text, flags=re.M))
    if a.output:
        Path(a.output).write_text(text, encoding="utf8")
    else:
        sys.stdout.write(text + "\n")
    print(f"[extract_source] {words:,} words, {markers} page/slide markers"
          + (f", figures -> {img_dir} ({len(list(img_dir.iterdir()))} files)" if img_dir else ""),
          file=sys.stderr)
    if words < 80:
        print("[extract_source] WARNING: very little text. If this is a scanned PDF, OCR it "
              "(e.g. `ocrmypdf in.pdf out.pdf`) or read the page images directly.", file=sys.stderr)


if __name__ == "__main__":
    main()
