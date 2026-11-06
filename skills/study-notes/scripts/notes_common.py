"""Shared helpers for build_docx.py / build_pptx.py / validate_notes.py."""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

RTL_LANGS = {"fa", "ar", "he", "ur", "ps", "ckb", "yi", "dv", "ug", "sd"}

# Each theme: primary (headings, bars), accent (highlights), soft (callout bg), ink (body text)
THEMES = {
    "indigo":  {"primary": "2B3A8C", "accent": "E8A33D", "soft": "EEF0FA", "ink": "1F2333", "muted": "6B7088"},
    "teal":    {"primary": "0F6E6E", "accent": "F2994A", "soft": "E7F4F3", "ink": "1C2B2B", "muted": "5F7373"},
    "crimson": {"primary": "8E1F2F", "accent": "D9A441", "soft": "F8ECEE", "ink": "2A1A1D", "muted": "7A6266"},
    "slate":   {"primary": "2F3E4E", "accent": "3FA7D6", "soft": "EDF1F5", "ink": "1E252C", "muted": "68737E"},
    "forest":  {"primary": "2E5E3A", "accent": "C9A227", "soft": "EDF4EE", "ink": "1D2A20", "muted": "62705F"},
    "plum":    {"primary": "5B2A6E", "accent": "2BB3A3", "soft": "F3ECF6", "ink": "26182C", "muted": "75677B"},
}

# Semantic callout colours (bg, border, label) — shared by both builders
CALLOUTS = {
    "key_point": ("FFF6E0", "E8A33D", "Key point"),
    "warning":   ("FDECEC", "C0392B", "Common mistake"),
    "tip":       ("E8F6EE", "27AE60", "Exam tip"),
    "mnemonic":  ("F1EAFB", "8E44AD", "Memory aid"),
    "note":      ("EAF2FB", "2E86DE", "Note"),
    "intuition": ("EAF7F7", "16A085", "Intuition"),
}

LABELS = {
    "en": {"contents": "Contents", "objectives": "Learning objectives", "prereq": "Prerequisites",
           "overview": "Overview", "in_brief": "In brief", "takeaways": "Key takeaways",
           "mistakes": "Common mistakes", "exam_tips": "Exam tips", "cheat": "Cheat sheet",
           "glossary": "Glossary", "questions": "Review questions", "answers": "Answer key",
           "flashcards": "Flashcards", "further": "Further reading", "source": "Source",
           "example": "Worked example", "solution": "Solution", "problem": "Problem",
           "where": "where", "agenda": "Roadmap", "check": "Check yourself", "summary": "Summary",
           "definition": "Definition", "front": "Prompt", "back": "Answer", "page": "Page",
           "self_test": "Self-test plan", "difficulty": "Difficulty", "answer": "Answer",
           "true_false": "True or false", "continued": "cont.", "term": "Term", "formula": "Formula", "kicker": "STUDY NOTES", "code": "Code", "key_terms": "Key terms", "remember": "Remember",
           "key_point": "Key point", "warning": "Common mistake", "tip": "Exam tip",
           "mnemonic": "Memory aid", "note": "Note", "intuition": "Intuition"},
    "fa": {"contents": "فهرست مطالب", "objectives": "اهداف یادگیری", "prereq": "پیش‌نیازها",
           "overview": "مرور کلی", "in_brief": "به طور خلاصه", "takeaways": "نکات کلیدی",
           "mistakes": "اشتباهات رایج", "exam_tips": "نکات امتحانی", "cheat": "برگه خلاصه",
           "glossary": "واژه‌نامه", "questions": "سؤالات مرور", "answers": "پاسخنامه",
           "flashcards": "فلش‌کارت‌ها", "further": "منابع بیشتر", "source": "منبع",
           "example": "مثال حل‌شده", "solution": "راه‌حل", "problem": "مسئله",
           "where": "که در آن", "agenda": "نقشه راه", "check": "خودآزمایی", "summary": "جمع‌بندی",
           "definition": "تعریف", "front": "پرسش", "back": "پاسخ", "page": "صفحه",
           "self_test": "برنامه خودآزمایی", "difficulty": "سطح دشواری", "answer": "پاسخ",
           "true_false": "درست یا نادرست", "continued": "ادامه", "term": "اصطلاح", "formula": "فرمول", "kicker": "جزوه مطالعه", "code": "کد", "key_terms": "اصطلاحات کلیدی", "remember": "به خاطر بسپارید",
           "key_point": "نکته کلیدی", "warning": "اشتباه رایج", "tip": "نکته امتحانی",
           "mnemonic": "ترفند حفظ کردن", "note": "یادداشت", "intuition": "شهود"},
    "ar": {"contents": "المحتويات", "objectives": "أهداف التعلم", "prereq": "المتطلبات السابقة",
           "overview": "نظرة عامة", "in_brief": "باختصار", "takeaways": "النقاط الرئيسية",
           "mistakes": "أخطاء شائعة", "exam_tips": "نصائح الامتحان", "cheat": "ورقة المراجعة",
           "glossary": "المسرد", "questions": "أسئلة المراجعة", "answers": "مفتاح الإجابات",
           "flashcards": "بطاقات تعليمية", "further": "قراءات إضافية", "source": "المصدر",
           "example": "مثال محلول", "solution": "الحل", "problem": "المسألة", "where": "حيث",
           "agenda": "خريطة الطريق", "check": "اختبر نفسك", "summary": "الخلاصة",
           "definition": "تعريف", "front": "سؤال", "back": "إجابة", "page": "صفحة",
           "self_test": "خطة الاختبار الذاتي", "difficulty": "الصعوبة", "answer": "الإجابة",
           "true_false": "صح أم خطأ", "continued": "تابع", "term": "المصطلح", "formula": "الصيغة", "kicker": "ملاحظات دراسية", "code": "الكود", "key_terms": "المصطلحات الأساسية", "remember": "تذكّر",
           "key_point": "نقطة أساسية", "warning": "خطأ شائع", "tip": "نصيحة امتحان",
           "mnemonic": "وسيلة للتذكر", "note": "ملاحظة", "intuition": "الحدس"},
}


def load_notes(path: str | Path) -> dict:
    with open(path, encoding="utf8") as f:
        data = json.load(f)
    meta = data.setdefault("meta", {})
    lang = (meta.get("language") or "en").split("-")[0].lower()
    meta["language"] = lang
    meta.setdefault("direction", "rtl" if lang in RTL_LANGS else "ltr")
    meta.setdefault("theme", "indigo")
    # sections may be inline objects or paths to per-section JSON files (handy for long sources)
    base = Path(path).resolve().parent
    secs = []
    for s in data.get("sections", []):
        if isinstance(s, str):
            with open(base / s, encoding="utf8") as f:
                loaded = json.load(f)
            secs.extend(loaded if isinstance(loaded, list) else [loaded])
        else:
            secs.append(s)
    data["sections"] = secs
    return data


def labels(meta: dict) -> dict:
    base = dict(LABELS["en"])
    base.update(LABELS.get(meta.get("language", "en"), {}))
    base.update(meta.get("labels", {}))  # caller overrides for any other language
    return base


def theme(meta: dict) -> dict:
    t = dict(THEMES.get(meta.get("theme", "indigo"), THEMES["indigo"]))
    t.update({k: v.lstrip("#") for k, v in (meta.get("colors") or {}).items()})
    return t


INLINE_RE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|(?<![\w*])\*(?!\s)[^*]+?\*(?![\w*])|==.+?==)")


def inline_runs(text: str) -> list[tuple[str, dict]]:
    """Split light markdown (**bold**, *italic*, `code`, ==highlight==) into (text, fmt) runs."""
    out: list[tuple[str, dict]] = []
    for part in INLINE_RE.split(str(text)):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            out.append((part[2:-2], {"bold": True}))
        elif part.startswith("==") and part.endswith("==") and len(part) > 4:
            out.append((part[2:-2], {"bold": True, "highlight": True}))
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            out.append((part[1:-1], {"code": True}))
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            out.append((part[1:-1], {"italic": True}))
        else:
            out.append((part, {}))
    return out


def plain(text: str) -> str:
    return "".join(t for t, _ in inline_runs(text))


def item_text(item) -> str:
    return item.get("text", "") if isinstance(item, dict) else str(item)


def item_children(item) -> list:
    return item.get("children", []) if isinstance(item, dict) else []


def walk_sections(sections, depth=1, prefix=""):
    """Yield (number, depth, section) for every section/subsection in reading order."""
    for i, s in enumerate(sections, 1):
        num = f"{prefix}{i}"
        yield num, depth, s
        yield from walk_sections(s.get("subsections", []), depth + 1, num + ".")


def render_mermaid(code: str, out_png: Path, width: int = 1600) -> bool:
    """Render Mermaid to PNG with mmdc if available. Returns True on success."""
    mmdc = shutil.which("mmdc")
    if not mmdc:
        return False
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "d.mmd"
        src.write_text(code, encoding="utf8")
        cfg = Path(td) / "p.json"
        cfg.write_text(json.dumps({"args": ["--no-sandbox"]}))
        try:
            subprocess.run([mmdc, "-i", str(src), "-o", str(out_png), "-w", str(width),
                            "-b", "white", "-s", "2", "-p", str(cfg)],
                           check=True, capture_output=True, timeout=120)
            return out_png.exists()
        except Exception:
            return False


MATHTEXT_FIXES = [  # LaTeX macros matplotlib's mathtext lacks -> supported equivalents
    (r"\\iff\b", r"\\Leftrightarrow"), (r"\\implies\b", r"\\Rightarrow"), (r"\\impliedby\b", r"\\Leftarrow"),
    (r"\\[dt]frac\b", r"\\frac"), (r"\\text\{", r"\\mathrm{"), (r"\\textbf\{", r"\\mathbf{"),
    (r"\\operatorname\{", r"\\mathrm{"), (r"\\lvert|\\rvert", "|"), (r"\\lVert|\\rVert", r"\\Vert"),
    (r"\\mid\b", "|"), (r"\\displaystyle\s*", ""), (r"\\limits\b", ""), (r"\\,|\\;|\\!", " "),
    (r"\\quad\b|\\qquad\b", r"\\ \\ "), (r"\\le\b", r"\\leq"), (r"\\ge\b", r"\\geq"),
    (r"\\bmod\b", r"\\mathrm{mod}"), (r"\\tag\{[^}]*\}", ""), (r"\\begin\{[^}]*\}|\\end\{[^}]*\}", ""),
    (r"&", ""), (r"\\\\", r"\\ "),
]


def render_latex(expr: str, out_png: Path, fontsize: int = 22, color: str = "1F2333") -> bool:
    """Render a LaTeX math expression to a tight PNG via matplotlib mathtext."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return False
    expr = expr.strip().strip("$")
    for a, b in MATHTEXT_FIXES:
        expr = re.sub(a, b, expr)
    if not (expr.startswith("$") and expr.endswith("$")):
        expr = f"${expr}$"
    try:
        fig = plt.figure(figsize=(0.01, 0.01))
        fig.text(0, 0, expr, fontsize=fontsize, color="#" + color)
        fig.savefig(str(out_png), dpi=220, bbox_inches="tight", pad_inches=0.06, transparent=True)
        plt.close(fig)
        return True
    except Exception:
        plt.close("all")
        return False


def image_size(path: Path) -> tuple[int, int]:
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except Exception:
        return (1600, 900)
