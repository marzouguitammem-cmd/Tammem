# -*- coding: utf-8 -*-
"""المنطق المشترك بين نسخة ويندوز (main.py) ونسخة الأندرويد (android/src/main.py).
ما فيه حتى استعمال لواجهة رسومية: ملفات، بحث، تعمير قوالب Word، واستخراج نص PDF."""

import json
import os
import re
import unicodedata

try:  # مكتبة تعمير القوالب (python-docx)
    import docx
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:  # pragma: no cover
    docx = None

try:  # استخراج نص PDF (مولّد المذكرات)
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    try:
        import fitz
    except ImportError:
        fitz = None

TEMPLATE_NAME = "template.docx"
INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
IGNORED_FILES = {"thumbs.db", "desktop.ini"}
_ARABIC_MAP = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ة": "ه", "ى": "ي", "ؤ": "و", "ئ": "ي"})
_ARABIC_MARKS = re.compile("[ً-ٰٟـ]")


def normalize(text):
    """تطبيع النص للبحث: حروف صغيرة، بدون تشكيل، وتوحيد الألف والتاء المربوطة."""
    return _ARABIC_MARKS.sub("", text.casefold()).translate(_ARABIC_MAP)


def natural_key(name):
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", normalize(name))]


def is_hidden(name):
    return name.startswith(".") or name.startswith("~$") or name.lower() in IGNORED_FILES


def list_entries(path):
    """يرجع (المجلدات, الملفات) مرتبة، بدون الملفات المخفية."""
    dirs, files = [], []
    try:
        with os.scandir(path) as it:
            for e in it:
                if is_hidden(e.name):
                    continue
                try:
                    (dirs if e.is_dir() else files).append(e.name)
                except OSError:
                    continue
    except OSError:
        pass
    dirs.sort(key=natural_key)
    files.sort(key=natural_key)
    return dirs, files


def subfolders(path):
    return list_entries(path)[0]


def validate_name(name):
    """يرجع رسالة خطأ بالعربية أو None إذا الاسم صالح."""
    if not name:
        return "الاسم فارغ."
    if INVALID_CHARS.search(name):
        return 'الاسم فيه رموز ممنوعة: < > : " / \\ | ? *'
    if name != name.rstrip(" ."):
        return "الاسم ما ينجمش ينتهي بنقطة أو فراغ."
    if name.split(".")[0].upper() in RESERVED_NAMES:
        return "هذا الاسم محجوز من النظام."
    return None


def search_names(root, query, limit=1000):
    """يبحث بالاسم (مع تجاهل الهمزات والتشكيل) في كل المجلدات.
    يرجّع (قائمة (مسار، مجلد؟)، النتائج مقطوعة؟)."""
    query = normalize(query.strip())
    results = []
    if not query or not root:
        return results, False
    for folder, dirs, files in os.walk(root):
        dirs[:] = sorted((d for d in dirs if not is_hidden(d)), key=natural_key)
        names = [(d, True) for d in dirs] + [(f, False) for f in sorted(files, key=natural_key) if not is_hidden(f)]
        for name, is_dir in names:
            if query in normalize(name):
                if len(results) >= limit:
                    return results, True
                results.append((os.path.join(folder, name), is_dir))
    return results, False


# ---------- تعمير النموذج (الكلمات المعلّمة) ----------
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
_INVISIBLE = re.compile("[​-‏‪-‮⁦-⁩﻿]")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def placeholder_key(text):
    return normalize(_INVISIBLE.sub("", text).strip())


def fill_paragraph(p_el, mapping):
    """يبدّل {{...}} في فقرة. الكلمة المعلّمة ممكن تكون مقسومة على عدة runs:
    القيمة الجديدة تتكتب في أول run (فتبقى تنسيقاتو) والبقية تتفرّغ من الجزء المعلّم."""
    nodes = list(p_el.iter(qn("w:t")))
    originals = [t.text or "" for t in nodes]
    full = "".join(originals)
    if "{{" not in full:
        return
    matches = [m for m in PLACEHOLDER_RE.finditer(full) if placeholder_key(m.group(1)) in mapping]
    if not matches:
        return
    starts, pos = [], 0
    for text in originals:
        starts.append(pos)
        pos += len(text)
    texts = list(originals)
    for m in reversed(matches):  # من الآخر باش الإزاحات ما تتلخبطش
        start, end = m.span()
        value = mapping[placeholder_key(m.group(1))]
        first = True
        for i, original in enumerate(originals):
            a, b = starts[i], starts[i] + len(original)
            if not original or b <= start or a >= end:
                continue
            lo, hi = max(start, a) - a, min(end, b) - a
            texts[i] = texts[i][:lo] + (value if first else "") + texts[i][hi:]
            first = False
    for node, new, old in zip(nodes, texts, originals):
        if new != old:
            set_node_text(node, new)


def set_node_text(node, text):
    """يكتب النص في w:t، وكل سطر جديد (\\n) يتحوّل لفاصل سطر w:br داخل نفس الـ run."""
    first, *rest = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    node.text = first
    node.set(XML_SPACE, "preserve")
    prev = node
    for part in rest:
        br, t = OxmlElement("w:br"), OxmlElement("w:t")
        t.text = part
        t.set(XML_SPACE, "preserve")
        prev.addnext(br)
        br.addnext(t)
        prev = t


def fill_docx(src, dst, values):
    """ينسخ النموذج src إلى dst ويعمّر الكلمات المعلّمة في النص والجداول والـ header والـ footer."""
    if docx is None:
        raise RuntimeError("مكتبة python-docx غير مثبتة. نفّذ: pip install python-docx")
    mapping = {placeholder_key(k): v for k, v in values.items()}
    document = docx.Document(src)
    for p_el in iter_paragraph_elements(document):  # الفقرات والجداول (حتى المتداخلة) وصناديق النص والـ header والـ footer
        fill_paragraph(p_el, mapping)
    document.save(dst)


def auto_file_name(subject, lesson, week):
    parts = [subject.strip(), lesson.strip(), f"أسبوع {week}"]
    name = " - ".join(p for p in parts if p)
    name = re.sub(r"\s+", " ", INVALID_CHARS.sub(" ", name))
    return name.strip(" .")


# ---------- مولّد المذكرات (بلا إنترنت): استخراج النص من PDF وتقسيمه ----------
GUIDE_FILE, BOOK_FILE, KEYWORDS_FILE = "guide.pdf", "book.pdf", "keywords.json"
BOOK_KEY = "الكتاب"
STANDARD_KEYS = {"المادة", "الدرس", "الأسبوع", "التاريخ", BOOK_KEY}
DEFAULT_KEYWORDS = {
    "الأهداف": ["الأهداف", "الأهداف التعلمية", "الكفاءة", "الكفاءات", "الهدف"],
    "الوضعية الانطلاقية": ["الوضعية الانطلاقية", "الوضعية المشكلة", "وضعية الانطلاق", "التمهيد"],
    "المراحل": ["المراحل", "مراحل الدرس", "سير الحصة", "سير الدرس", "سير الأنشطة"],
    "التقييم": ["التقييم", "التقويم", "تقييم", "تقويم"],
}
# كلمات عربية شائعة تُستعمل لكشف الحروف المعكوسة داخل الكلمة
COMMON_WORDS = {
    normalize(w) for w in (
        "في من على إلى الى عن أن ان التي الذي هذا هذه ذلك كل ثم أو او مع بين عند بعد قبل "
        "ما لا هو هي هم كان يكون تم حول خلال أي إذا اذا الدرس التلميذ المتعلم النشاط نشاط"
    ).split()
}
_ARABIC_CHAR = re.compile("[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")
_ARABIC_LETTERS = re.compile("[ء-ْ]+")
_CONTROL_CHARS = re.compile("[​-‏‪-‮⁦-⁩﻿]")
_LEADING_JUNK = re.compile(r"^[\W\d_]+")
_LIST_ITEM = re.compile(r"^(?:[-•·*–—▪●◦]|\d+\s*[-.)]|[ء-ي]\s*[-.)]|\(\s*\w{1,3}\s*\))\s*")
TEXT_MODES = [
    ("تلقائي", "auto"),
    ("ترتيب PDF الأصلي (بدون إعادة ترتيب)", "raw"),
    ("عكس حروف الكلمات", "letters"),
    ("عكس ترتيب الكلمات في السطر", "words"),
]


class MemoError(Exception):
    """خطأ بمعنى واضح للمستخدم (الرسالة بالعربية)."""


def load_keywords(subject_dir):
    """يقرأ keywords.json للمادة، وإذا ما كانش موجود يكوّنو بالقيم الافتراضية."""
    path = os.path.join(subject_dir, KEYWORDS_FILE)
    if not os.path.isfile(path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_KEYWORDS, f, ensure_ascii=False, indent=2)
        except OSError as e:
            raise MemoError(f"ما نجمتش نكوّن ملف {KEYWORDS_FILE} في مجلد المادة:\n{e}")
        return {k: list(v) for k, v in DEFAULT_KEYWORDS.items()}
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise MemoError(f"ملف {KEYWORDS_FILE} فيه غلطة في الكتابة (سطر {e.lineno}، عمود {e.colno}): {e.msg}\nصلّحو وعاود حاول.")
    except OSError as e:
        raise MemoError(f"ما نجمتش نقرا {KEYWORDS_FILE}:\n{e}")
    return validate_keywords(data)


def validate_keywords(data):
    """يتثبّت من محتوى keywords.json ويرجّع قاموس القسم ← قائمة كلمات."""
    if not isinstance(data, dict) or not data:
        raise MemoError(f"ملف {KEYWORDS_FILE} لازم يكون قاموس: اسم القسم ← قائمة كلمات مفتاحية.")
    result = {}
    for section, words in data.items():
        if isinstance(words, str):
            words = [words]
        if not isinstance(words, list) or not all(isinstance(w, str) for w in words):
            raise MemoError(f"في {KEYWORDS_FILE}: قيمة القسم «{section}» لازم تكون قائمة نصوص.")
        if section in STANDARD_KEYS:
            raise MemoError(f"في {KEYWORDS_FILE}: الاسم «{section}» محجوز، بدّلو (المحجوزة: {'، '.join(sorted(STANDARD_KEYS))}).")
        result[section] = [w for w in words if w.strip()]
    return result


def open_pdf(path):
    if fitz is None:
        raise MemoError("مكتبة PyMuPDF غير مثبتة. نفّذ: pip install PyMuPDF")
    if not os.path.isfile(path):
        raise MemoError(f"الملف غير موجود: {os.path.basename(path)}")
    try:
        doc = fitz.open(path)
    except Exception as e:  # noqa: BLE001
        raise MemoError(f"ما نجمتش نفتح الملف {os.path.basename(path)} (ربما تالف):\n{e}")
    if doc.needs_pass:
        doc.close()
        raise MemoError(f"الملف {os.path.basename(path)} محمي بكلمة سر. شيل الحماية وعاود حاول.")
    return doc


def pdf_page_count(path):
    doc = open_pdf(path)
    try:
        return doc.page_count
    finally:
        doc.close()


def _has_arabic(text):
    return bool(_ARABIC_CHAR.search(text))


def page_lines(page, mode):
    """أسطر الصفحة بترتيب القراءة. نعتمد على مواقع الكلمات في الصفحة (وليس ترتيب التخزين في PDF)
    فنتفادو الكلمات المعكوسة الترتيب، ونلصقو الكلمات المقطّعة (حروف منفصلة عن بعضها)."""
    if mode == "raw":
        return page.get_text("text", sort=True).splitlines()
    words = [w for w in page.get_text("words") if w[4].strip()]
    rows = []  # كل سطر: [مركز y، ارتفاع، كلمات]
    for w in sorted(words, key=lambda w: (w[1] + w[3]) / 2):
        yc, h = (w[1] + w[3]) / 2, max(w[3] - w[1], 1)
        if rows and abs(yc - rows[-1][0]) <= 0.5 * max(h, rows[-1][1]):
            rows[-1][2].append(w)
        else:
            rows.append([yc, h, [w]])
    lines = []
    for _yc, h, row in rows:
        arabic = sum(len(_ARABIC_CHAR.findall(w[4])) for w in row)
        latin = sum(len(re.findall(r"[A-Za-z]", w[4])) for w in row)
        if arabic > latin:
            tokens = _rtl_tokens(sorted(row, key=lambda w: -w[2]), h)
        else:
            tokens = [w[4] for w in sorted(row, key=lambda w: w[0])]
        lines.append(" ".join(tokens))
    return lines


def _rtl_tokens(row, height):
    """row مرتبة من اليمين لليسار. نلصقو الأجزاء المتلاصقة، ونرجّعو عبارات اللاتينية لاتجاهها."""
    tokens, prev = [], None
    for w in row:
        text = unicodedata.normalize("NFKC", w[4])
        if prev is not None and prev[0] - w[2] < 0.1 * height and _has_arabic(text) and _has_arabic(tokens[-1]):
            tokens[-1] += text
        else:
            tokens.append(text)
        prev = w[0:1]
    # علامة ترقيم في أول كلمة عربية (مثل «:الأهداف») مكانها الصحيح في آخرها
    tokens = [re.sub(r"^([:：،؛.!؟]+)(.*[\u0621-\u064a])$", r"\2\1", t) if _has_arabic(t) else t for t in tokens]
    # عبارة لاتينية متتالية (كلمات فيها حروف A-Z) تتكتب من اليسار لليمين
    out, i = [], 0
    while i < len(tokens):
        if re.search(r"[A-Za-z]", tokens[i]):
            j = i
            while j < len(tokens) and re.search(r"[A-Za-z]", tokens[j]):
                j += 1
            out.extend(reversed(tokens[i:j]))
            i = j
        else:
            out.append(tokens[i])
            i += 1
    return out


def looks_letter_reversed(lines, vocab):
    forward = backward = 0
    for line in lines:
        for token in re.findall(r"[ء-ْ]+", line):
            key = normalize(token)
            if key in vocab:
                forward += 3
            if key[::-1] in vocab:
                backward += 3
            if len(key) > 3 and key.startswith("ال"):
                forward += 1
            if len(key) > 3 and key.endswith("لا"):
                backward += 1
    return backward > forward * 1.5 and backward >= 3


def repair_lines(lines, mode, vocab):
    lines = [_CONTROL_CHARS.sub("", unicodedata.normalize("NFKC", line)).strip() for line in lines]
    lines = [line for line in lines if line]
    if mode == "letters" or (mode == "auto" and looks_letter_reversed(lines, vocab)):
        lines = [_ARABIC_LETTERS.sub(lambda m: m.group(0)[::-1], line) for line in lines]
    if mode == "words":
        lines = [" ".join(reversed(line.split())) for line in lines]
    return lines


def extract_lines(pdf_path, first, last, mode, vocab):
    """أسطر النص للصفحات first..last (ترقيم من 1)."""
    doc = open_pdf(pdf_path)
    try:
        if not 1 <= first <= last <= doc.page_count:
            raise MemoError(f"مجال الصفحات غير صحيح: {first}-{last} (الملف {os.path.basename(pdf_path)} فيه {doc.page_count} صفحة).")
        raw = []
        for number in range(first - 1, last):
            raw.extend(page_lines(doc[number], mode))
    finally:
        doc.close()
    lines = repair_lines(raw, mode, vocab)
    if not lines:
        raise MemoError(
            f"ما لقيتش نص في الصفحات {first}-{last} من {os.path.basename(pdf_path)}.\n"
            "ربما الصفحات صور ممسوحة ضوئيا (Scan). التطبيق يقرا النص الحقيقي برك، وما يعملش OCR."
        )
    return lines


def reflow(lines):
    """يلصق الأسطر اللي كانت مقطوعة بسبب عرض الصفحة، ويخلي القوائم والعناوين كل وحدة في سطر."""
    paragraphs, last_len = [], 0
    for line in lines:
        line = line.strip()
        if not line:
            continue
        joinable = (
            paragraphs and last_len >= 45 and not _LIST_ITEM.match(line)
            and not re.search(r"[.:؟?!؛]$", paragraphs[-1])
        )
        if joinable:
            paragraphs[-1] += " " + line
        else:
            paragraphs.append(line)
        last_len = len(line)
    return "\n".join(paragraphs)


def _normalized_with_map(text):
    chars, index = [], []
    for i, ch in enumerate(text):
        for nc in normalize(ch):
            chars.append(nc)
            index.append(i)
    return "".join(chars), index


def match_heading(line, keyword_index):
    """إذا السطر عنوان قسم يرجّع (القسم، بقية النص بعد العنوان) وإلا None."""
    stripped = _LEADING_JUNK.sub("", line.strip())
    norm, index = _normalized_with_map(stripped)
    for keyword, section in keyword_index:
        if not norm.startswith(keyword):
            continue
        if len(norm) > len(keyword) and "ء" <= norm[len(keyword)] <= "ي":
            continue  # الكلمة المفتاحية لازم تكون كلمة كاملة
        rest = stripped[index[len(keyword) - 1] + 1:]
        colon = re.match(r"^[^:：]{0,40}[:：]\s*(.*)$", rest)
        if colon:
            return section, colon.group(1).strip()
        if len(rest.strip()) <= 30 and not re.search(r"[.؟?!]", rest):
            return section, ""
    return None


def segment_guide(lines, keywords):
    """يقسّم أسطر الدليل حسب العناوين. يرجّع (قاموس القسم ← نص، أسطر ما قبل أول عنوان)."""
    keyword_index = sorted(
        ((normalize(kw.strip()), section) for section, kws in keywords.items() for kw in kws if kw.strip()),
        key=lambda item: -len(item[0]),
    )
    found = {section: [] for section in keywords}
    current, unclassified = None, []
    for line in lines:
        heading = match_heading(line, keyword_index)
        if heading:
            current = heading[0]
            if heading[1]:
                found[current].append(heading[1])
        elif current is None:
            unclassified.append(line)
        else:
            found[current].append(line)
    return {section: reflow(body) for section, body in found.items()}, unclassified


def iter_paragraph_elements(document):
    """كل فقرات المستند: الجسم (وفيه الجداول وصناديق النص) والـ header والـ footer."""
    roots = [document.element.body]
    for section in document.sections:
        for part in (
            section.header, section.first_page_header, section.even_page_header,
            section.footer, section.first_page_footer, section.even_page_footer,
        ):
            if not part.is_linked_to_previous:
                roots.append(part._element)
    for root in roots:
        yield from list(root.iter(qn("w:p")))


def template_placeholders(template_path):
    """أسماء الكلمات المعلّمة الموجودة في النموذج (مطبّعة)."""
    if docx is None:
        raise MemoError("مكتبة python-docx غير مثبتة. نفّذ: pip install python-docx")
    try:
        document = docx.Document(template_path)
    except Exception as e:  # noqa: BLE001
        raise MemoError(f"ما نجمتش نفتح النموذج {os.path.basename(template_path)} (ربما تالف):\n{e}")
    found = {}
    for p_el in iter_paragraph_elements(document):
        text = "".join(t.text or "" for t in p_el.iter(qn("w:t")))
        for m in PLACEHOLDER_RE.finditer(text):
            found[placeholder_key(m.group(1))] = m.group(1).strip()
    return found


def fmt_date(d):
    return f"{d.day():02d}/{d.month():02d}/{d.year()}"
