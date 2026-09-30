"""Word .docx 읽기 — 표준 라이브러리(zipfile + xml)만 씁니다.

.docx 는 XML 파일 묶음(zip)입니다. word/document.xml 에서 문단·표·그림·수식을,
word/styles.xml 에서 제목 스타일을 읽어 절 구조를 만듭니다.
구형 .doc 는 읽지 않습니다 — Word 에서 .docx 로 저장한 뒤 넣습니다(기획서 7장, 가정).
"""

import re
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
DC = "http://purl.org/dc/elements/1.1/"
NS = {"w": W, "m": M}


def _q(tag):
    pre, name = tag.split(":")
    return "{%s}%s" % ({"w": W, "m": M, "dc": DC}[pre], name)


HEADING_STYLE_RE = re.compile(r"^(?:heading|제목)\s*(\d)$", re.I)
NUM_TITLE_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)\.?\s+(\S.*)$")


class DocxError(Exception):
    pass


def _para_text(p):
    parts = []
    for el in p.iter():
        if el.tag == _q("w:t"):
            parts.append(el.text or "")
        elif el.tag in (_q("w:tab"),):
            parts.append(" ")
        elif el.tag in (_q("w:br"), _q("w:cr")):
            parts.append(" ")
    return re.sub(r"[ \t ]+", " ", "".join(parts)).strip()


def _read_styles(z):
    styles = {}
    if "word/styles.xml" not in z.namelist():
        return styles
    root = ET.fromstring(z.read("word/styles.xml"))
    for st in root.findall("w:style", NS):
        sid = st.get(_q("w:styleId"))
        name_el = st.find("w:name", NS)
        name = name_el.get(_q("w:val")) if name_el is not None else ""
        lvl = None
        m = HEADING_STYLE_RE.match(name or "")
        if m:
            lvl = int(m.group(1))
        ol = st.find("w:pPr/w:outlineLvl", NS)
        if lvl is None and ol is not None:
            v = int(ol.get(_q("w:val"), "9"))
            if v < 9:
                lvl = v + 1
        based = st.find("w:basedOn", NS)
        styles[sid] = {
            "name": name or "",
            "level": lvl,
            "based": based.get(_q("w:val")) if based is not None else None,
        }
    # basedOn 을 따라 제목 수준 물려받기
    for sid, info in styles.items():
        seen = set()
        cur = info
        while cur["level"] is None and cur["based"] and cur["based"] in styles and cur["based"] not in seen:
            seen.add(cur["based"])
            cur = styles[cur["based"]]
        if info["level"] is None:
            info["level"] = cur["level"]
    return styles


def _body_children(body):
    """w:sdt(콘텐츠 컨트롤) 안쪽까지 펼쳐 문단·표를 순서대로 돌려줍니다."""
    for el in list(body):
        if el.tag == _q("w:sdt"):
            content = el.find("w:sdtContent", NS)
            if content is not None:
                yield from _body_children(content)
        elif el.tag in (_q("w:p"), _q("w:tbl")):
            yield el


def read_docx(path, number_titles=True):
    """docx 를 읽어 블록 목록과 제목을 돌려줍니다 (구조화는 structure.build)."""
    try:
        z = zipfile.ZipFile(path)
    except (FileNotFoundError, IsADirectoryError):
        raise DocxError("파일을 찾을 수 없습니다: %s" % path)
    except zipfile.BadZipFile as e:
        raise DocxError("docx 파일을 열 수 없습니다: %s (%s). 구형 .doc 이면 Word 에서 .docx 로 저장해 주세요." % (path, e))
    with z:
        header_lines = []
        for name in sorted(n for n in z.namelist() if re.match(r"word/(header|footer)\d*\.xml$", n)):
            try:
                hroot = ET.fromstring(z.read(name))
            except ET.ParseError:
                continue
            for p in hroot.iter(_q("w:p")):
                t = _para_text(p)
                if t and t not in header_lines:
                    header_lines.append(t)
        if "word/document.xml" not in z.namelist():
            raise DocxError("word/document.xml 이 없습니다. Word(.docx) 파일이 맞는지 확인해 주세요: %s" % path)
        styles = _read_styles(z)
        root = ET.fromstring(z.read("word/document.xml"))
        core_title = ""
        if "docProps/core.xml" in z.namelist():
            core = ET.fromstring(z.read("docProps/core.xml"))
            t = core.find("{%s}title" % DC)
            if t is not None and t.text:
                core_title = t.text.strip()
    body = root.find("w:body", NS)
    if body is None:
        raise DocxError("문서 본문이 비어 있습니다: %s" % path)
    blocks = []
    for el in _body_children(body):
        if el.tag == _q("w:p"):
            ppr = el.find("w:pPr", NS)
            sid = None
            outline = None
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None:
                    sid = ps.get(_q("w:val"))
                ol = ppr.find("w:outlineLvl", NS)
                if ol is not None:
                    v = int(ol.get(_q("w:val"), "9"))
                    outline = v + 1 if v < 9 else None
            st = styles.get(sid, {}) if sid else {}
            level = outline or st.get("level")
            has_fig = any(True for _ in el.iter(_q("w:drawing"))) or any(True for _ in el.iter(_q("w:pict"))) or any(True for _ in el.iter(_q("w:object")))
            blocks.append({
                "kind": "p",
                "text": _para_text(el),
                "style": (st.get("name") or sid or ""),
                "style_level": level,
                "figure": has_fig,
                "omath": sum(1 for _ in el.iter(_q("m:oMath"))),
            })
        else:
            rows = []
            for tr in el.findall("w:tr", NS):
                cells = []
                for tc in tr.findall("w:tc", NS):
                    cells.append(" ".join(_para_text(p) for p in tc.iter(_q("w:p")) if _para_text(p)))
                rows.append(cells)
            blocks.append({
                "kind": "tbl",
                "rows": rows,
                "text": " / ".join(" | ".join(r) for r in rows),
                "omath": sum(1 for _ in el.iter(_q("m:oMath"))),
                "figure": False,
            })
    if not core_title and header_lines:
        from .pdf_reader import header_title
        core_title = header_title(header_lines)
    return build(blocks, core_title=core_title, number_titles=number_titles, source="docx", header=header_lines)


# 개정 이력 머리말 — 이 문단 뒤는 본문 절이 아니라 개정 이력으로 봅니다.
REV_RE = re.compile(r"(제\s*[/·]?\s*개정\s*(이력|기록)|개정\s*(이력|기록)|Revision\s*(History|Log|Record)|Change\s*Log)", re.I)
REV_ROW_RE = re.compile(r"^\s*(?:Rev\.?\s*)?(\d{1,2}|[0-9xX]{2})\s+(\d{4}\s*[.\-/]\s*\d{1,2}\s*[.\-/]\s*\d{1,2})\s*(.*)$")
DATE_RE = re.compile(r"(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})")


def _rev_rows(blocks):
    rows = []
    for b in blocks:
        if b["kind"] == "tbl":
            for r in b["rows"]:
                m = REV_ROW_RE.match(" ".join(c for c in r if c))
                if m:
                    rows.append({"rev": m.group(1), "date": re.sub(r"\s", "", m.group(2)), "rest": m.group(3).strip()})
        elif b.get("text"):
            m = REV_ROW_RE.match(b["text"])
            if m:
                rows.append({"rev": m.group(1), "date": re.sub(r"\s", "", m.group(2)), "rest": m.group(3).strip()})
    return rows


def _looks_like_title(text, limit=40):
    t = text.strip()
    if len(t) > limit:
        return False
    if t.endswith((".", "다", "다.", "함", "음", ":", "：")):
        return False
    return True


def build(blocks, core_title="", number_titles=True, source="docx", header=None, title_max=40):
    """블록에 절 번호를 매기고 제목 목록을 만듭니다. (docx 없이도 시험할 수 있게 분리)"""
    headings = []
    counters = []
    current = ""
    top_last = 0
    title = core_title
    in_rev = False
    rev_blocks = []
    for i, b in enumerate(blocks):
        b["idx"] = i
        b["heading"] = None
        if not in_rev and b["kind"] == "p" and b.get("text") and len(b["text"]) <= 40 and REV_RE.search(b["text"]):
            in_rev = True
        if in_rev:
            b["rev"] = True
            b["sec"] = ""
            rev_blocks.append(b)
            continue
        if b["kind"] == "p" and b["text"]:
            style_name = (b.get("style") or "").lower()
            if not title and (style_name in ("title", "제목") or (not headings and i == 0 and not NUM_TITLE_RE.match(b["text"]))):
                title = b["text"]
                b["sec"] = ""
                b["is_title"] = True
                continue
            lvl = b.get("style_level")
            m = NUM_TITLE_RE.match(b["text"])
            num = None
            name = b["text"]
            if lvl:
                if m:
                    num, name = m.group(1), m.group(2).strip()
                else:
                    counters = (counters + [0] * lvl)[:lvl]
                    counters[lvl - 1] += 1
                    num = ".".join(str(c) for c in counters)
            elif number_titles and m and b.get("title_ok", True) and _looks_like_title(b["text"], title_max):
                cand = m.group(1)
                parts = cand.split(".")
                top = int(parts[0])
                if len(parts) == 1:
                    if top_last < top <= top_last + 3:
                        num, name = cand, m.group(2).strip()
                else:
                    if current and current.split(".")[0] == parts[0]:
                        num, name = cand, m.group(2).strip()
            if num:
                parts = [int(x) for x in num.split(".")]
                counters = parts[:]
                if len(parts) == 1:
                    top_last = parts[0]
                current = num
                h = {"num": num, "level": len(parts), "title": name, "idx": i}
                headings.append(h)
                b["heading"] = h
        b["sec"] = current
    doc = _finish(blocks, headings, title)
    doc["source"] = source
    doc["header"] = list(header or [])
    doc["revision"] = {"found": bool(rev_blocks), "rows": _rev_rows(rev_blocks)}
    return doc


def _finish(blocks, headings, title):
    return {"blocks": blocks, "headings": headings, "title": title}


def read_document(path, number_titles=True):
    """확장자를 보고 .docx 또는 .pdf 를 읽습니다."""
    suffix = str(path).lower().rsplit(".", 1)[-1]
    if suffix == "pdf":
        from .pdf_reader import read_pdf
        return read_pdf(path, number_titles=number_titles)
    return read_docx(path, number_titles=number_titles)


# ---------- 절 도우미 ----------

def top_of(sec):
    try:
        return int(sec.split(".")[0]) if sec else 0
    except ValueError:
        return 0


def in_section(sec, num):
    """sec 가 num 절(또는 그 하위)에 속하는지."""
    return bool(sec) and (sec == num or sec.startswith(num + "."))


def parse_target(target):
    """「1,4+」 → 판정 함수. 「N+」는 N항 이후 전체."""
    tokens = [t.strip() for t in str(target or "").replace(";", ",").split(",") if t.strip()]

    def match(sec):
        for t in tokens:
            if t.endswith("+"):
                try:
                    if top_of(sec) >= int(t[:-1]):
                        return True
                except ValueError:
                    pass
            elif in_section(sec, t):
                return True
        return False

    return tokens, match


def section_blocks(doc, target):
    _, match = parse_target(target)
    return [b for b in doc["blocks"] if match(b.get("sec", "")) and not b.get("heading")]


def section_text(doc, target, limit=None):
    """대상 절의 제목과 본문을 이어 붙인 글(프롬프트용)."""
    _, match = parse_target(target)
    lines = []
    for b in doc["blocks"]:
        if not match(b.get("sec", "")):
            continue
        if b.get("heading"):
            h = b["heading"]
            lines.append("")
            lines.append("%s %s" % (h["num"], h["title"]))
        elif b["kind"] == "tbl":
            for r in b["rows"]:
                lines.append("| " + " | ".join(r) + " |")
        elif b["text"]:
            lines.append(b["text"])
        elif b.get("figure"):
            lines.append("[그림]")
    text = "\n".join(lines).strip()
    if limit and len(text) > limit:
        text = text[:limit] + "\n…(이하 생략 — 설정 「프롬프트_본문_최대글자」를 넘었습니다)"
    return text


def has_section(doc, num):
    return any(h["num"] == str(num) for h in doc["headings"])
