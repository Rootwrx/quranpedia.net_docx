#!/usr/bin/env python3
"""Build DOCX files from Quranpedia tafsir dumps (tafsir-book-<id>.json[.gz]).

Examples:
  python build_tafsir_docx.py --data . --ids 1-4,7-9,12 --mode complete
  python build_tafsir_docx.py --data . --ids 3 --mode surah --surahs 1,2,18-20
  python build_tafsir_docx.py --data . --ids 3 --mode both -o out

Requires: pip install python-docx
"""
import argparse, gzip, json, re, sys
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

SURAHS = """الفاتحة البقرة|آل عمران|النساء|المائدة|الأنعام|الأعراف|الأنفال|التوبة|يونس|هود|يوسف|الرعد|إبراهيم|الحجر|النحل|الإسراء|الكهف|مريم|طه|الأنبياء|الحج|المؤمنون|النور|الفرقان|الشعراء|النمل|القصص|العنكبوت|الروم|لقمان|السجدة|الأحزاب|سبأ|فاطر|يس|الصافات|ص|الزمر|غافر|فصلت|الشورى|الزخرف|الدخان|الجاثية|الأحقاف|محمد|الفتح|الحجرات|ق|الذاريات|الطور|النجم|القمر|الرحمن|الواقعة|الحديد|المجادلة|الحشر|الممتحنة|الصف|الجمعة|المنافقون|التغابن|الطلاق|التحريم|الملك|القلم|الحاقة|المعارج|نوح|الجن|المزمل|المدثر|القيامة|الإنسان|المرسلات|النبأ|النازعات|عبس|التكوير|الانفطار|المطففين|الانشقاق|البروج|الطارق|الأعلى|الغاشية|الفجر|البلد|الشمس|الليل|الضحى|الشرح|التين|العلق|القدر|البينة|الزلزلة|العاديات|القارعة|التكاثر|العصر|الهمزة|الفيل|قريش|الماعون|الكوثر|الكافرون|النصر|المسد|الإخلاص|الفلق|الناس"""
SURAHS = SURAHS.replace("الفاتحة البقرة", "الفاتحة|البقرة").split("|")
assert len(SURAHS) == 114, len(SURAHS)

FONT = "KFGQPC Uthman Taha Naskh"   # body text
AYAH_FONT = "Kitab"                 # Quran text
BODY_PT, FOOT_PT, AYAH_PT = 16, 12, 16
AYAH_COLOR = RGBColor(0x1F, 0x5C, 0x3A)
FOOT_COLOR = RGBColor(0x55, 0x55, 0x55)
PAGE_MARK = re.compile(r"\s*-\[[0-9٠-٩]+\]-\s*")



# ---------- honorifics ----------
HON_COLOR = RGBColor(0x9C, 0x1C, 0x1C)
HONORIFICS_ON = True
_HON_PHRASES = [
    # prophetic
    "صلى الله عليه وسلم", "صلى الله عليه وآله وسلم", "صلى الله عليه وعلى آله وسلم",
    "صلى الله عليه وعلى آله وصحبه وسلم", "صلى الله وسلم عليه", "صلى الله عليهم وسلم",
    "صلوات الله وسلامه عليه", "صلوات الله وسلامه عليهم", "عليه الصلاة والسلام",
    "عليهم الصلاة والسلام", "عليهما الصلاة والسلام", "عليه أفضل الصلاة والسلام",
    "عليه السلام", "عليها السلام", "عليهما السلام", "عليهم السلام", "عليهن السلام",
    "سلام الله عليه", "سلام الله عليهم", "صلى الله عليه", "صلى الله عليهم",
    # companions
    "رضي الله عنه", "رضي الله عنها", "رضي الله عنهما", "رضي الله عنهم", "رضي الله عنهن",
    "رضي الله تعالى عنه", "رضي الله تعالى عنها", "رضي الله تعالى عنهما", "رضي الله تعالى عنهم",
    "رضوان الله عليه", "رضوان الله عليها", "رضوان الله عليهم", "رضوان الله عليهم أجمعين",
    "رضوان الله تعالى عليهم", "كرم الله وجهه",
    # scholars
    "رحمه الله تعالى", "رحمها الله تعالى", "رحمهم الله تعالى", "رحمهما الله تعالى",
    "رحمه الله", "رحمها الله", "رحمهم الله", "رحمهما الله", "رحمهن الله",
    "عليه رحمة الله", "عليهم رحمة الله", "قدس الله روحه", "نور الله ضريحه",
    # divine
    "سبحانه وتعالى", "تبارك وتعالى", "تبارك وتقدس", "جل وعلا", "جل جلاله", "عز وجل",
    "سبحانه", "تعالى", "تبارك اسمه", "جل شأنه",
    # more prophetic
    "عليه الصلاة والتسليم", "عليه أزكى الصلاة وأتم التسليم", "عليه أفضل الصلاة وأتم التسليم",
    "عليها الصلاة والسلام", "عليهن الصلاة والسلام", "عليه وعلى آله الصلاة والسلام",
    "عليه وعلى نبينا الصلاة والسلام", "على نبينا وعليه الصلاة والسلام",
    "صلى الله على نبينا محمد وعلى آله وسلم", "صلى الله على محمد وسلم", "صلى الله على محمد",
    "صلوات الله عليه", "صلوات الله عليهم", "صلوات ربي وسلامه عليه", "صلى الله عليه وسلم تسليما",
    # more companions / scholars
    "رضي الله عنهم أجمعين", "رضي الله عنهما وأرضاهما", "رضي الله عنه وأرضاه", "رحمه الله ورضي عنه",
    "رحمة الله عليه", "رحمة الله عليهم", "غفر الله له", "غفر الله لهم", "عفا الله عنه",
    "قدس الله سره", "قدس الله أرواحهم", "نضر الله وجهه", "حفظه الله", "أيده الله", "وفقه الله",
    "تغمده الله برحمته", "طيب الله ثراه",
    # more divine
    "سبحانه وبحمده", "تقدست أسماؤه", "تعالى شأنه", "عز شأنه", "جل ثناؤه", "جل في علاه",
    "عز اسمه", "تبارك الله", "عز من قائل", "جل من قائل", "تعالى ذكره", "تقدس وتعالى",
    "عز سلطانه", "جلت قدرته", "جل اسمه", "تعالى جده", "سبحانه عما يشركون",
]
_LETTER_MAP = {"ا": "[اأإآ]", "أ": "[اأإآ]", "إ": "[اأإآ]", "آ": "[اأإآ]", "ى": "[ىي]", "ي": "[ىي]"}
_MARKS = "[\u064B-\u065F\u0670\u0640]*"


def _phrase_re(ph):
    out = []
    for ch in ph:
        if ch == " ":
            out.append(r"\s+")
        else:
            out.append(_LETTER_MAP.get(ch, re.escape(ch)) + _MARKS)
    return "".join(out)


_alts = [_phrase_re(p) for p in sorted(set(_HON_PHRASES), key=len, reverse=True)]
HON_RE = re.compile(
    "(?:(?<![\u0621-\u064A])|(?<=(?<![\u0621-\u064A])[\u0648\u0641]))(" + "|".join(_alts) + ")(?![\u0621-\u064A])"
    "|([\uFDFA\uFDFB\uFD40-\uFD4F])"
)


LIGATURES_ON = True
_MARK_RE = re.compile("[\u064B-\u065F\u0670\u0640]")


def _norm(t):
    t = _MARK_RE.sub("", t)
    t = re.sub(r"\s+", " ", t).strip()
    return re.sub("[أإآ]", "ا", t).replace("ى", "ي")


_SALAWAT = {_norm(p) for p in [
    "صلى الله عليه وسلم", "صلى الله عليه وآله وسلم", "صلى الله عليه وعلى آله وسلم",
    "صلى الله عليه وعلى آله وصحبه وسلم", "صلى الله وسلم عليه",
]}
_JALLA = {_norm("جل جلاله")}


def _ligature(chunk):
    if not LIGATURES_ON:
        return chunk
    n = _norm(chunk)
    if n in _SALAWAT:
        return "\uFDFA"  # ﷺ
    if n in _JALLA:
        return "\uFDFB"  # ﷻ
    return chunk


def split_honorifics(text):
    """-> [(chunk, is_honorific)]"""
    if not HONORIFICS_ON:
        return [(text, False)]
    res, pos = [], 0
    for m in HON_RE.finditer(text):
        if m.start() > pos:
            res.append((text[pos:m.start()], False))
        res.append((_ligature(m.group(0)), True))
        pos = m.end()
    if pos < len(text):
        res.append((text[pos:], False))
    return res


# ---------- selection parsing ----------
def parse_ranges(spec, lo=None, hi=None):
    out = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    seen, res = set(), []
    for n in out:
        if n not in seen:
            seen.add(n)
            res.append(n)
    return res


# ---------- loading ----------
def find_book_file(data_dir: Path, bid: int):
    n = f"tafsir-book-{bid}.json"
    for c in (data_dir / (n + ".gz"), data_dir / n / n, data_dir / n):
        if c.is_file():
            return c
    return None


def load_json(path: Path):
    if path.suffix == ".gz":
        with gzip.open(path, "rt", encoding="utf-8") as f:
            return json.load(f)
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------- HTML -> paragraphs ----------
class TafsirHTML(HTMLParser):
    """Produces paragraphs: {'kind': 'body'|'foot', 'runs': [(text, is_ayah)]}"""

    def __init__(self, keep_page_marks=False):
        super().__init__()
        self.paras, self.cur = [], []
        self.ayah = False
        self.foot = False
        self.keep = keep_page_marks

    def flush(self):
        runs = [(t, a) for t, a in self.cur if t]
        self.cur = []
        if not runs or not "".join(t for t, _ in runs).strip():
            return
        runs[0] = (runs[0][0].lstrip(), runs[0][1])
        runs[-1] = (runs[-1][0].rstrip(), runs[-1][1])
        self.paras.append({"kind": "foot" if self.foot else "body", "runs": runs})

    def handle_starttag(self, tag, attrs):
        cls = dict(attrs).get("class", "") or ""
        if tag == "br":
            self.flush()
        elif tag == "span" and "book-ayah" in cls:
            self.ayah = True
        elif tag == "div" and "foot-notes" in cls:
            self.flush()
            self.foot = True

    def handle_endtag(self, tag):
        if tag == "span":
            self.ayah = False
        elif tag == "div":
            self.flush()

    def handle_data(self, data):
        data = re.sub(r"[\r\n]+", " ", data)
        if not self.keep:
            data = PAGE_MARK.sub(" ", data)
        self.cur.append((data, self.ayah))


def html_to_paras(html, keep_page_marks=False):
    p = TafsirHTML(keep_page_marks)
    p.feed(html or "")
    p.close()
    p.flush()
    return p.paras


# ---------- docx helpers ----------
def _rtl_par(p):
    p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))  # before spacing/jc
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _style_run(run, size, color=None, bold=False, font=None):
    font = font or FONT
    run.font.name = font
    rpr = run._r.get_or_add_rPr()
    rpr.rFonts.set(qn("w:cs"), font)
    rpr.rFonts.set(qn("w:eastAsia"), font)
    if bold:
        run.bold = True
        b = rpr.find(qn("w:b"))
        bcs = OxmlElement("w:bCs")
        b.addnext(bcs)
    if color is not None:
        run.font.color.rgb = color
    run.font.size = Pt(size)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(size * 2)))
    rpr.find(qn("w:sz")).addnext(szcs)
    rpr.append(OxmlElement("w:rtl"))


def add_par(doc, runs, size, color_for_plain=None, ayah_color=AYAH_COLOR, space_after=6, style=None):
    p = doc.add_paragraph(style=style)
    _rtl_par(p)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.1
    for text, is_ayah in runs:
        if is_ayah:
            r = p.add_run(text)
            _style_run(r, size, ayah_color, font=AYAH_FONT)
            continue
        for chunk, is_hon in split_honorifics(text):
            r = p.add_run(chunk)
            _style_run(r, size, HON_COLOR if is_hon else color_for_plain, font=FONT)
    return p


def add_heading(doc, text, size, center=True):
    p = doc.add_paragraph(style="Heading 1")
    _rtl_par(p)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(text)
    _style_run(r, size, RGBColor(0, 0, 0), bold=True)
    return p


def add_toc(doc):
    p = doc.add_paragraph()
    _rtl_par(p)
    r = p.add_run()
    for kind, txt in (("begin", None), ("instr", 'TOC \\o "1-1" \\h \\z \\u'), ("separate", None), ("text", "اضغط F9 لتحديث الفهرس"), ("end", None)):
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
            r._r.append(el)
        elif kind == "text":
            t = OxmlElement("w:t")
            t.text = txt
            r._r.append(t)
        else:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            r._r.append(fc)


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def new_document(with_toc_update=False):
    doc = Document()
    for s in doc.sections:
        s.right_margin = s.left_margin = Cm(2.2)
        s.top_margin = s.bottom_margin = Cm(2.2)
        s.footer.is_linked_to_previous = False
    if with_toc_update:
        upd = OxmlElement("w:updateFields")
        upd.set(qn("w:val"), "true")
        doc.settings.element.append(upd)
    return doc


def title_block(doc, book):
    b = book.get("book", {})
    add_par(doc, [(b.get("name") or b.get("short_name") or "", False)], 30, RGBColor(0, 0, 0), space_after=18).alignment = WD_ALIGN_PARAGRAPH.CENTER
    if b.get("short_name") and b.get("short_name") != b.get("name"):
        add_par(doc, [(b["short_name"], False)], 20, space_after=12).alignment = WD_ALIGN_PARAGRAPH.CENTER
    lines = []
    au = b.get("author") or {}
    if au.get("full_name") or au.get("ar_name"):
        lines.append(au.get("full_name") or au.get("ar_name"))
    if b.get("mohaqeq"):
        lines.append("تحقيق: " + b["mohaqeq"])
    if b.get("nasher"):
        lines.append("الناشر: " + b["nasher"])
    if b.get("edition"):
        lines.append("الطبعة: " + b["edition"])
    for ln in lines:
        add_par(doc, [(ln, False)], 16, space_after=6).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_par(doc, [("المصدر: الموسوعة القرآنية quranpedia.net", False)], 12, FOOT_COLOR).alignment = WD_ALIGN_PARAGRAPH.CENTER


# ---------- core ----------
def group_by_surah(book, wanted_surahs):
    by = {}
    for a in sorted(book.get("ayahs", []), key=lambda x: (x.get("surah", 0), x.get("ayah", 0))):
        s = a.get("surah")
        if wanted_surahs and s not in wanted_surahs:
            continue
        by.setdefault(s, []).append(a)
    return by


def is_heading_spillover(html):
    """Short page-boundary leftovers like 'تفسير سورة البقرة / وهي مدنية'."""
    plain = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html or "")).strip()
    return plain.startswith("تفسير سورة") and len(plain) < 80


def surah_paras(ayahs, seen, keep_marks):
    """Dedupe repeated entries (a tafsir block is repeated for every ayah it covers)."""
    out = []
    for a in ayahs:
        for c in a.get("content", []):
            key = (c.get("part"), c.get("page"), c.get("text"))
            if key in seen:
                continue
            seen.add(key)
            if is_heading_spillover(c.get("text", "")):
                continue
            out.extend(html_to_paras(c.get("text", ""), keep_marks))
            out.append({"kind": "sep"})
    return out


def write_paras(doc, paras):
    for p in paras:
        if p["kind"] == "sep":
            continue
        if p["kind"] == "foot":
            add_par(doc, p["runs"], FOOT_PT, FOOT_COLOR, ayah_color=FOOT_COLOR, space_after=2)
        else:
            add_par(doc, p["runs"], BODY_PT)


def sname(n):
    return f"سورة {SURAHS[n-1]}" if 1 <= n <= 114 else f"سورة {n}"


def safe(s):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("_")


def build_complete(book, by, out_path, keep_marks):
    doc = new_document(with_toc_update=True)
    title_block(doc, book)
    page_break(doc)
    add_heading(doc, "الفهرس", 22)
    add_toc(doc)
    seen = set()
    for s, ayahs in by.items():
        paras = surah_paras(ayahs, seen, keep_marks)
        if not paras:
            continue
        page_break(doc)
        add_heading(doc, sname(s), 24)
        write_paras(doc, paras)
    doc.save(out_path)


def build_surah(book, s, ayahs, out_path, keep_marks):
    paras = surah_paras(ayahs, set(), keep_marks)
    if not paras:
        return False
    doc = new_document()
    add_heading(doc, sname(s), 24)
    b = book.get("book", {})
    add_par(doc, [(b.get("short_name") or b.get("name") or "", False)], 14, FOOT_COLOR).alignment = WD_ALIGN_PARAGRAPH.CENTER
    write_paras(doc, paras)
    doc.save(out_path)
    return True


def list_books(data_dir, out_file):
    """Scan data_dir for tafsir-book-*.json[.gz] and write 'id<TAB>short_name<TAB>name<TAB>author'."""
    ids = {}
    for p in data_dir.glob("tafsir-book-*.json*"):
        m = re.match(r"tafsir-book-(\d+)\.json(\.gz)?$", p.name)
        if m:
            ids[int(m.group(1))] = int(m.group(1))
    rows = []
    dec = json.JSONDecoder()
    for bid in sorted(ids):
        f = find_book_file(data_dir, bid)
        if not f:
            continue
        try:
            op = gzip.open if f.suffix == ".gz" else open
            with op(f, "rt", encoding="utf-8") as fh:
                head = fh.read(200000)
            i = head.index('"book"')
            obj, _ = dec.raw_decode(head, head.index("{", i))
            au = obj.get("author") or {}
            rows.append((bid, obj.get("short_name") or "", obj.get("name") or "", au.get("ar_name") or au.get("full_name") or ""))
        except Exception as e:
            rows.append((bid, "?", f"unreadable: {e}", ""))
    with open(out_file, "w", encoding="utf-8") as o:
        o.write("id\tshort_name\tname\tauthor\n")
        for r in rows:
            o.write("\t".join(str(x) for x in r) + "\n")
    print(f"{len(rows)} tafsirs -> {out_file}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=".", help="folder containing tafsir-book-*.json[.gz]")
    ap.add_argument("--list", metavar="FILE", help="write id/name/author of every tafsir in --data to FILE and exit")
    ap.add_argument("--ids", help="e.g. 1-4,7-9,12")
    ap.add_argument("--mode", choices=["complete", "surah", "both"], default="complete")
    ap.add_argument("--surahs", default="", help="optional filter, e.g. 1,2,18-20 (default: all)")
    ap.add_argument("-o", "--out", default="out")
    ap.add_argument("--no-ligatures", action="store_true", help="keep full salawat text instead of ﷺ")
    ap.add_argument("--no-honorifics", action="store_true", help="don't colour honorific phrases")
    ap.add_argument("--keep-page-marks", action="store_true", help="keep -[٤٠]- page markers")
    args = ap.parse_args()

    if args.list:
        list_books(Path(args.data), args.list)
        return
    if not args.ids:
        ap.error("--ids is required (or use --list FILE)")
    global HONORIFICS_ON, LIGATURES_ON
    LIGATURES_ON = not args.no_ligatures
    HONORIFICS_ON = not args.no_honorifics
    data, out = Path(args.data), Path(args.out)
    wanted = set(parse_ranges(args.surahs)) if args.surahs else None
    for bid in parse_ranges(args.ids):
        f = find_book_file(data, bid)
        if not f:
            print(f"[{bid}] not found, skipped", file=sys.stderr)
            continue
        print(f"[{bid}] loading {f}")
        book = load_json(f)
        b = book.get("book", {})
        label = f"{bid}-{safe(b.get('short_name') or b.get('name') or 'book')}"
        by = group_by_surah(book, wanted)
        if args.mode in ("complete", "both"):
            out.mkdir(parents=True, exist_ok=True)
            p = out / f"{label}.docx"
            build_complete(book, by, p, args.keep_page_marks)
            print(f"  -> {p}")
        if args.mode in ("surah", "both"):
            d = out / label
            d.mkdir(parents=True, exist_ok=True)
            n = 0
            for s, ayahs in by.items():
                nm = SURAHS[s - 1] if 1 <= s <= 114 else str(s)
                if build_surah(book, s, ayahs, d / f"{s:03d}-{safe(nm)}.docx", args.keep_page_marks):
                    n += 1
            print(f"  -> {d}/ ({n} files)")


if __name__ == "__main__":
    main()
