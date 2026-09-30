"""학습 프로필 — 사용자가 넣은 Template·승인 표준·기존 회의록에서 규칙을 뽑아 내 PC 에만 저장합니다.

「학습」은 AI 모델 학습이 아니라, 파일의 구조와 표기 습관을 세어 규칙으로 바꾸는 일입니다.

- Template 학습   : 상위·하위 항목 목록과 순서, 번호 방식, 예시 문구·자리표시(xx, 0000), 개정 이력 칸 유무
- 기존 표준 학습  : 용어정의 목록, 인용 문서번호의 모양(글자·숫자 자릿수), 표·그림 캡션 방식, 영문 표기 습관
- 회의록 학습     : 지적 범주(구분) 이름, 범주별 건수, 문장 끝맺음(…바랍니다 / …필요 / …요망 등), 머리 기호, 열 이름

프로필(JSON)은 기본으로 data/profile.json 에 저장되며 .gitignore 로 저장소에 올라가지 않습니다.
프로필에는 사내 문서의 용어·지적 문구가 담기므로 PC 밖으로 옮기지 않습니다.
"""

import copy
import datetime
import json
import re
from collections import Counter
from pathlib import Path

from .docx_reader import in_section

PROFILE_VERSION = 1

# ---------- 공통 ----------


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def empty_profile():
    return {
        "version": PROFILE_VERSION,
        "note": "이 파일은 사내 문서에서 뽑은 규칙입니다. 이 PC 밖으로 옮기거나 저장소에 올리지 마세요.",
        "updated": "",
        "templates": {},     # 표준 종류 → Template 학습 결과
        "standards": {},     # 파일 이름 → 기존 표준 학습 결과
        "minutes": {},       # 파일 이름 → 회의록 학습 결과
    }


class ProfileError(Exception):
    pass


def load_profile(path):
    p = Path(path)
    if not p.exists():
        return empty_profile()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (ValueError, OSError) as e:
        raise ProfileError("학습 프로필을 읽을 수 없습니다: %s (%s)" % (p, e))
    base = empty_profile()
    base.update(data)
    return base


def save_profile(profile, path):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    profile["updated"] = now()
    p.write_text(json.dumps(profile, ensure_ascii=False, indent=1), encoding="utf-8")


def is_empty(profile):
    return not (profile.get("templates") or profile.get("standards") or profile.get("minutes"))


# ---------- Template 학습 ----------

# 자리표시 — Template 에 채워 넣으라고 비워 둔 칸(xx, XX, 0000, ####, YYYY.MM.DD, ○○○)
PLACEHOLDER_RES = [
    r"(?<![A-Za-z])[xX]{2,}(?![A-Za-z])",
    r"(?<![\d,.])0{3,4}(?![\d,.])",
    r"\d{4}\s*\.\s*00\s*\.\s*00|0000\s*\.\s*\d{2}\s*\.\s*\d{2}",
    r"Rev\.?\s*0[xX]",
    r"[○◯OΟ]{3}|○{2,}|△{2,}",
    r"#{3,}",
    r"YYYY|(?<![A-Za-z])MM\s*\.\s*DD(?![A-Za-z])",
]
PLACEHOLDER_RE = re.compile("|".join("(?:%s)" % p for p in PLACEHOLDER_RES))
# 「(예시)」처럼 괄호로 감싼 표기는 가상 예시 문서의 머리 표시로도 쓰이므로 넣지 않습니다.
EXAMPLE_MARK_RE = re.compile(r"^\s*(?:예시\s*\)|예\s*\)|\(\s*참고\s*\)|작성\s*예시|작성\s*요령|※\s*작성)")
FLOW_WORDS_RE = re.compile(r"flow\s*chart|flowchart|흐름도|순서도|플로우", re.I)
OPTIONAL_ITEM_RE = re.compile(r"기타|appendix|첨부|부록", re.I)


def _body_blocks(doc):
    return [b for b in doc["blocks"] if not b.get("rev") and not b.get("heading") and b.get("text")]


def _example_sub_nums(doc):
    """상위 항목 바로 아래에 「예시)」 표시가 있고 그 뒤에 나오는 하위 제목(예: 3항의 예시 용어)은
    Template 의 필수 하위 항목이 아니라 예시로 봅니다."""
    out, marked_top = set(), None
    for b in doc["blocks"]:
        if b.get("rev"):
            break
        h = b.get("heading")
        if h:
            if h["level"] == 1:
                marked_top = None
            elif h["level"] == 2 and marked_top and h["num"].split(".")[0] == marked_top:
                out.add(h["num"])
            continue
        sec = str(b.get("sec") or "")
        if sec and "." not in sec and EXAMPLE_MARK_RE.match(b.get("text") or ""):
            marked_top = sec
    return out


def learn_template(doc, std_type, source):
    tops = [h for h in doc["headings"] if h["level"] == 1]
    ex_subs = _example_sub_nums(doc)
    items = []
    for order, h in enumerate(tops, 1):
        subs = [{"num": s["num"], "name": s["title"]} for s in doc["headings"]
                if s["level"] == 2 and s["num"].split(".")[0] == h["num"] and s["num"] not in ex_subs]
        items.append({"order": order, "num": h["num"], "name": h["title"],
                      "required": not OPTIONAL_ITEM_RE.search(h["title"]), "subs": subs})
    # 예시 문구: 「예시)」「(참고)」 표시 줄과, 그 뒤 같은 절 안에서 이어지는 예시 본문
    example_lines, marker_count = [], 0
    last_marker_sec = None
    for b in doc["blocks"]:
        if b.get("rev"):
            break
        if b.get("heading"):
            # 예시 하위 제목(예: 3항 예시 용어) 아래 글도 예시 문구로 모읍니다.
            last_marker_sec = b["heading"]["num"] if b["heading"]["num"] in ex_subs else None
            continue
        t = (b.get("text") or "").strip()
        if not t:
            continue
        if EXAMPLE_MARK_RE.match(t):
            marker_count += 1
            last_marker_sec = b.get("sec")
            rest = EXAMPLE_MARK_RE.sub("", t).strip()
            if len(rest) >= 6:
                example_lines.append(rest)
            continue
        if last_marker_sec is not None and b.get("sec") == last_marker_sec and len(t) >= 6:
            example_lines.append(t)
    example_heads = [h["title"] for h in doc["headings"] if h["num"] in ex_subs]
    header_ph = [t for t in doc.get("header", []) if PLACEHOLDER_RE.search(t)]
    flow = [s["num"] for it in items for s in it["subs"] if FLOW_WORDS_RE.search(s["name"])]
    return {
        "type": std_type, "source": source, "learned": now(),
        "items": items,
        "title": doc.get("title", ""),
        "example_lines": sorted(set(example_lines)),
        "example_terms": example_heads,
        "example_markers": marker_count,
        "header_lines": len(doc.get("header", [])),
        "header_placeholders": header_ph,
        "has_revision": bool(doc.get("revision", {}).get("found")),
        "flow_sections": flow,
    }


# ---------- 기존 표준 학습 ----------

CITE_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9])[A-Z][A-Z0-9]{0,7}(?:-[A-Za-z0-9]{1,8}){1,4}(?![A-Za-z0-9])")
CAPTION_STYLE_RE = re.compile(r"^\s*([\[<〈(]?)\s*(표|그림|Table|Tab|Figure|Fig)\s*(\.?)\s*(\d+)", re.I)
LATIN_TOKEN_RE = re.compile(r"(?<![A-Za-z])[A-Za-z]{2,}(?:-[A-Za-z]{2,})*(?![A-Za-z])")


def shape_of(token):
    """문서번호 모양 — 실제 번호 대신 글자·숫자 자릿수만 남긴 정규식."""
    out = []
    for run in re.findall(r"[A-Z]+|[a-z]+|\d+|-", token):
        if run == "-":
            out.append("-")
        elif run.isdigit():
            n = len(run)
            out.append(r"\d{%d,%d}" % (max(1, n - 1), n + 1))
        else:
            n = len(run)
            out.append("[A-Za-z]{1,%d}" % (n + 2) if run.islower() else "[A-Z]{1,%d}" % (n + 2))
    return "".join(out)


def is_cite_token(tok):
    return "-" in tok and any(ch.isdigit() for ch in tok) and sum(ch.isalpha() for ch in tok) >= 2


def caption_style(text):
    m = CAPTION_STYLE_RE.match(text or "")
    if not m:
        return None
    word = m.group(2)
    kind = "표" if word.lower() in ("표", "table", "tab") else "그림"
    return {"kind": kind, "word": word, "style": "%s%s%s n%s" % (m.group(1), word, m.group(3), {"[": "]", "<": ">", "〈": "〉", "(": ")"}.get(m.group(1), ""))}


def learn_standard(doc, source):
    from .checks import citation_lines, term_entries
    terms = []
    for e in term_entries(doc):
        m = re.search(r"[(（]\s*([^)）]+?)\s*[)）]", e["raw"])
        terms.append({"term": e["term"], "alt": m.group(1).strip() if m else ""})
    shapes = Counter()
    for t, _ in citation_lines(doc):
        for tok in CITE_TOKEN_RE.findall(t):
            if is_cite_token(tok):
                shapes[shape_of(tok)] += 1
    for t in doc.get("header", []):
        for tok in CITE_TOKEN_RE.findall(t):
            if is_cite_token(tok) and not PLACEHOLDER_RE.search(tok):
                shapes[shape_of(tok)] += 1
    styles, words = Counter(), {"표": Counter(), "그림": Counter()}
    for b in doc["blocks"]:
        if b.get("heading") or b.get("rev") or not b.get("text"):
            continue
        cs = caption_style(b["text"])
        if cs and (b.get("style") == "caption" or len(b["text"]) <= 80):
            styles[cs["style"]] += 1
            words[cs["kind"]][cs["word"]] += 1
    forms = {}
    for b in _body_blocks(doc):
        if in_section(b.get("sec", ""), "2"):
            continue
        for tok in LATIN_TOKEN_RE.findall(b["text"]):
            if len(tok) < 3:
                continue
            key = tok.lower().replace("-", "")
            forms.setdefault(key, Counter())[tok] += 1
    return {
        "source": source, "learned": now(),
        "terms": terms,
        "cite_shapes": dict(shapes),
        "caption_styles": dict(styles),
        "caption_words": {k: dict(v) for k, v in words.items()},
        "latin_forms": {k: dict(v) for k, v in forms.items()},
        "counts": {"절": len(doc["headings"]), "용어": len(terms), "인용번호모양": len(shapes), "캡션": sum(styles.values())},
    }


# ---------- 회의록 학습 ----------

# 지적 범주 — 낱말로 나눕니다. 코드 → 기준표 No (없으면 기타)
CATEGORY_RULES = [
    ("LEFTOVER", r"추후|예정|미완|미작성|작성\s*필요|TBD|자리표시|예시\s*문구|내용\s*없음"),
    ("REVISION", r"개정|이력|Rev\.?|머리글|표지|작성일|발행일"),
    ("FLOWCHART", r"흐름도|flow\s*chart|flowchart|순서도|플로우"),
    ("CAPTIONS", r"그림\s*/\s*표|표\s*/\s*그림|그림\s*번호|표\s*번호|캡션|Tab\.?|Fig\.?|그림.*번호|표.*번호|번호.*(그림|표)"),
    ("CITATION", r"인용|관련\s*표준|참고\s*문헌|문서\s*번호|표준\s*번호|법규|References"),
    ("TERM", r"용어|표기|약어|영문|한글|통일|정의"),
    ("PURPOSE", r"목적|범위|Scope"),
    ("QUANT", r"정량|수치|기준값|근거|단위|정성|구체적\s*값"),
    ("VERIFY", r"검증|시험|평가|Test"),
    ("SAFETY", r"안전|환경"),
    ("TEMPLATE", r"Template|템플릿|양식|항목|목차|순서|구성"),
    ("TYPO", r"오타|맞춤법|띄어쓰기|오기|문법"),
]
CODE_TO_NO = {"TEMPLATE": "공통-1", "LEFTOVER": "공통-4", "REVISION": "공통-3", "PURPOSE": "1-1", "CITATION": "2-1",
              "TERM": "3-2", "FLOWCHART": "4-2", "CAPTIONS": "4-4", "QUANT": "5-2", "VERIFY": "5-1", "SAFETY": "7-1"}
NO_TO_CODE = {"공통-1": "TEMPLATE", "공통-3": "REVISION", "공통-4": "LEFTOVER", "1-1": "PURPOSE", "2-1": "CITATION",
              "4-6": "CITATION", "3-1": "TERM", "3-2": "TERM", "4-1": "TEMPLATE", "4-2": "FLOWCHART", "4-3": "QUANT",
              "4-4": "CAPTIONS", "4-5": "PURPOSE", "5-1": "VERIFY", "5-2": "QUANT", "6-1": "TEMPLATE", "7-1": "SAFETY",
              "8-1": "TEMPLATE"}
CODE_LABELS = {"TEMPLATE": "양식·항목", "LEFTOVER": "미완성·잔재", "REVISION": "머리글·개정 이력", "PURPOSE": "목적·범위",
               "CITATION": "인용 표준", "TERM": "용어", "FLOWCHART": "흐름도", "CAPTIONS": "표·그림 번호", "QUANT": "정량 기준",
               "VERIFY": "검증 방법", "SAFETY": "안전", "TYPO": "오탈자", "ETC": "기타"}


def classify(text, category=""):
    """지적 범주 코드. 회의록의 「구분」 값이 있으면 그것을 먼저 보고, 모르면 지적 문장으로 판단합니다."""
    for part in (category, text):
        for code, rx in CATEGORY_RULES:
            if part and re.search(rx, part, re.I):
                return code
    return "ETC"


# 문장 끝맺음 — 회의록 문체를 이 묶음으로 셉니다.
ENDINGS = [
    ("바랍니다", r"바랍니다\.?$|바람\.?$|주십시오\.?$|주세요\.?$"),
    ("요청", r"요청\s*(드립니다|합니다|함|드림)?\.?$"),
    ("요망", r"요망\.?$"),
    ("필요", r"필요\s*(함|합니다|해\s*보임|)\.?$"),
    ("할 것", r"(할|될|하실|하도록\s*할)\s*것\.?$"),
    ("검토", r"검토\s*(바람|바랍니다|필요|요망|요청)?\.?$"),
    ("함", r"[가-힣]\s*함\.?$|[가-힣]음\.?$"),
]


def ending_family(text):
    t = re.sub(r"[\s)）\]」』]+$", "", str(text or ""))
    for fam, rx in ENDINGS:
        if re.search(rx, t):
            return fam
    return ""


LEAD_RE = re.compile(r"^\s*([-•·▶▷○●◦※*]|\d{1,2}[.)]|[①-⑳]|[가-하][.)])\s*")
LOC_LEAD_RE = re.compile(r"^\s*(?:\(\s*(\d+(?:\.\d+)*)\s*(?:항|절)?\s*\)|\[\s*(\d+(?:\.\d+)*)\s*(?:항|절)?\s*\]|(\d+(?:\.\d+)*)\s*(항|절)|p\.?\s*\d+|\d+\s*(?:쪽|page))", re.I)

ROLE_WORDS = [
    ("seq", r"^(순번|번호|No\.?|NO|#)$"),
    ("category", r"구분|분류|유형|범주|카테고리|category|항목"),
    ("location", r"위치|페이지|쪽|해당\s*(절|항)|조항|page|절"),
    ("action", r"조치|답변|반영|회신|처리|수정\s*여부|결과"),
    ("comment", r"검토\s*의견|의견|지적|코멘트|comment|내용|요청|개선"),
]


def _role_of(header):
    h = str(header or "").strip()
    for role, rx in ROLE_WORDS:
        if re.search(rx, h, re.I):
            return role
    return ""


def _find_header_row(rows):
    for i, r in enumerate(rows[:15]):
        cells = [("" if c is None else str(c).strip()) for c in r]
        filled = [c for c in cells if c]
        if len(filled) >= 2 and all(len(c) <= 20 for c in filled) and any(_role_of(c) in ("comment", "category") for c in filled):
            return i
    return None


def _minutes_from_xlsx(path):
    from .config import _openpyxl
    openpyxl = _openpyxl()
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        raise ProfileError("회의록 엑셀을 열 수 없습니다: %s (%s)" % (path, e))
    items, headers, roles = [], [], {}
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        hi = _find_header_row(rows)
        if hi is None:
            continue
        head = [("" if c is None else str(c).strip()) for c in rows[hi]]
        rmap = {}
        for j, h in enumerate(head):
            role = _role_of(h)
            if role and role not in rmap:
                rmap[role] = j
        data = [[("" if c is None else str(c).strip()) for c in r] for r in rows[hi + 1:]]
        data = [r for r in data if any(r)]
        if "comment" not in rmap and data:
            # 열 이름으로 못 찾으면 글이 가장 긴 열을 의견 열로 봅니다.
            width = max(len(r) for r in data)
            avg = [sum(len(r[j]) if j < len(r) else 0 for r in data) / len(data) for j in range(width)]
            rmap["comment"] = max(range(width), key=lambda j: avg[j])
        if not headers:
            headers, roles = head, {k: head[v] for k, v in rmap.items() if v < len(head)}
        cat_last = ""
        for r in data:
            def g(role):
                j = rmap.get(role)
                return r[j] if j is not None and j < len(r) else ""
            text = g("comment")
            cat = g("category") or cat_last      # 병합 셀(구분)이면 위 값을 이어 씁니다.
            cat_last = cat
            if len(text) < 4:
                continue
            items.append({"text": text, "category": cat, "location": g("location")})
    return items, headers, roles


REQUEST_END_RE = re.compile(r"(바랍니다|바람|요청\S*|요망|필요\S*|할\s*것|검토\s*\S*|수정\s*\S*|주십시오|주세요)\.?\s*$")


def _minutes_from_text(path):
    from .docx_reader import read_document
    doc = read_document(path)
    items, cat = [], ""
    for b in doc["blocks"]:
        if b.get("heading"):
            cat = b["heading"]["title"]
            continue
        rows = b["rows"] if b["kind"] == "tbl" else [[b.get("text") or ""]]
        for r in rows:
            t = " ".join(c for c in r if c).strip()
            if not t:
                continue
            if len(t) <= 20 and not REQUEST_END_RE.search(t):
                cat = t.strip("[]【】<> :：")
                continue
            if REQUEST_END_RE.search(t) or LEAD_RE.match(t):
                m = LOC_LEAD_RE.match(LEAD_RE.sub("", t))
                items.append({"text": t, "category": cat, "location": m.group(0).strip() if m else ""})
    return items, [], {}


def read_minutes(path):
    """기존 회의록 → ([{text, category, location}], 열 이름 목록, {역할: 열 이름})."""
    suffix = str(path).lower().rsplit(".", 1)[-1]
    if suffix in ("xlsx", "xlsm"):
        return _minutes_from_xlsx(path)
    if suffix in ("docx", "pdf"):
        return _minutes_from_text(path)
    raise ProfileError("회의록은 .xlsx, .docx, .pdf 를 읽습니다: %s" % path)


def learn_minutes(items, headers, roles, source):
    cats, codes, endings, leads, locs = Counter(), Counter(), Counter(), Counter(), Counter()
    code_labels, examples = {}, {}
    tmpl = 0
    for it in items:
        t = it["text"].strip()
        code = classify(t, it.get("category", ""))
        codes[code] += 1
        if it.get("category"):
            cats[it["category"]] += 1
            code_labels.setdefault(code, Counter())[it["category"]] += 1
        fam = ending_family(t)
        if fam:
            endings[fam] += 1
        m = LEAD_RE.match(t)
        leads[m.group(1) if m else ""] += 1
        body = LEAD_RE.sub("", t)
        lm = LOC_LEAD_RE.match(body)
        if lm:
            if lm.group(1):
                locs["(n)"] += 1
            elif lm.group(2):
                locs["[n]"] += 1
            elif lm.group(3):
                locs["n" + (lm.group(4) or "")] += 1
            else:
                locs["쪽"] += 1
        elif it.get("location"):
            locs["열"] += 1
        else:
            locs[""] += 1
        if re.search(r"Template|템플릿|양식", t, re.I):
            tmpl += 1
        no = CODE_TO_NO.get(code)
        if no:
            lst = examples.setdefault(no, [])
            if len(lst) < 10 and t not in lst:
                lst.append(t)
    return {
        "source": source, "learned": now(), "count": len(items),
        "headers": headers, "roles": roles,
        "categories": dict(cats), "codes": dict(codes),
        "code_labels": {k: dict(v) for k, v in code_labels.items()},
        "endings": dict(endings), "leads": dict(leads), "locations": dict(locs),
        "template_mention": tmpl,
        "examples": examples,
    }


# ---------- 합치기와 적용 ----------


def _sum_counters(dicts):
    c = Counter()
    for d in dicts:
        c.update(d or {})
    return c


def minutes_style(profile):
    """여러 회의록 학습 결과를 합친 문체. 회의록을 학습하지 않았으면 None."""
    ms = list(profile.get("minutes", {}).values())
    if not ms:
        return None
    total = sum(m["count"] for m in ms) or 1
    endings = _sum_counters(m["endings"] for m in ms)
    leads = _sum_counters(m["leads"] for m in ms)
    locs = _sum_counters(m["locations"] for m in ms)
    labels = {}
    for m in ms:
        for code, d in m.get("code_labels", {}).items():
            labels.setdefault(code, Counter()).update(d)
    lead = leads.most_common(1)[0][0] if leads else ""
    loc = locs.most_common(1)[0][0] if locs else ""
    roles = {}
    for m in ms:
        for k, v in (m.get("roles") or {}).items():
            roles.setdefault(k, v)
    return {
        "ending": endings.most_common(1)[0][0] if endings else "",
        "lead": lead if leads.get(lead, 0) * 2 >= total else "",
        "location": loc if loc in ("(n)", "[n]", "n항", "n절") and locs.get(loc, 0) * 2 >= total else "",
        "template_prefix": sum(m["template_mention"] for m in ms) * 3 >= total,
        "labels": {k: v.most_common(1)[0][0] for k, v in labels.items()},
        "roles": roles,
    }


def minutes_examples(profile):
    out = {}
    for m in profile.get("minutes", {}).values():
        for no, lst in m.get("examples", {}).items():
            cur = out.setdefault(no, [])
            for t in lst:
                if t not in cur:
                    cur.append(t)
    return out


def _kw(name):
    name = re.split(r"↔", str(name or ""))[0]
    stop = {"사항", "관련", "및", "항목", "방법", "이하", "전체", "내용"}
    words = [w for w in re.split(r"[\s/,·()（）&|:：\-]+", name.lower()) if len(w) >= 2 and w not in stop]
    return words


def _overlap(a, b):
    return any(x in y or y in x for x in a for y in b)


def retarget(criteria, items, std_type):
    """기준표의 「N항 이름」이 학습한 Template 의 N항과 이름이 안 맞으면, 같은 이름의 항목으로 대상을 옮기거나
    그 Template 에 없는 항목이면 「해당 없음」 표시를 붙입니다. (표준 종류마다 항목 번호가 다르기 때문)"""
    by_num = {str(it["num"]): it for it in items}
    notes = []
    for c in criteria:
        m = re.match(r"^\s*(\d+)항\s*(.*)$", str(c.get("item", "")))
        if not m:
            continue
        num, name = m.group(1), m.group(2)
        kws = _kw(name)
        if not kws:
            continue
        cur = by_num.get(num)
        if cur and _overlap(kws, _kw(cur["name"])):
            continue
        found = next((it for it in items if _overlap(kws, _kw(it["name"]))), None)
        if found:
            new = str(found["num"])
            toks = []
            for t in str(c.get("target", "")).split(","):
                t = t.strip()
                if t == num:
                    t = new
                elif t.startswith(num + "."):
                    t = new + t[len(num):]
                if t:
                    toks.append(t)
            c["target"] = ",".join(toks)
            c["retarget"] = "학습한 %s Template 에서 %s항 → %s항(%s)" % (std_type, num, new, found["name"])
            notes.append("%s: %s" % (c["no"], c["retarget"]))
        else:
            c["na"] = "학습한 %s Template 에 「%s」 항목이 없어 이 기준은 해당 없음으로 둡니다." % (std_type, name.strip())
            notes.append("%s: 해당 없음" % c["no"])
    return notes


def apply_profile(cfg, profile, std_type=None):
    """기준표 설정에 학습 결과를 덧입힌 새 설정과, 무엇을 바꿨는지 적은 목록을 돌려줍니다(원본은 그대로)."""
    cfg = copy.deepcopy(cfg)
    notes = []
    if not profile or is_empty(profile):
        return cfg, notes
    tpl = (profile.get("templates") or {}).get(std_type) if std_type else None
    if tpl:
        items = [{"type": std_type, "order": it["order"], "num": it["num"], "name": it["name"],
                  "required": it.get("required", True)} for it in tpl["items"]]
        cfg["items"] = [it for it in cfg.get("items", []) if it.get("type") != std_type] + items
        cfg["template"] = tpl
        notes.append("Template 항목 %d개(하위 %d개)를 학습 결과로 적용" % (len(items), sum(len(it["subs"]) for it in tpl["items"])))
        notes += retarget(cfg.get("criteria", []), tpl["items"], std_type)
    stds = list((profile.get("standards") or {}).values())
    if stds:
        shapes = _sum_counters(s.get("cite_shapes") for s in stds)
        known = {p.get("regex") for p in cfg.get("std_patterns", [])}
        added = 0
        for shape, n in shapes.most_common():
            rx = r"(?<![A-Za-z0-9])%s(?![A-Za-z0-9])" % shape
            if rx not in known:
                cfg.setdefault("std_patterns", []).append({"name": "학습한 번호 모양 (%d건)" % n, "regex": rx})
                known.add(rx)
                added += 1
        if added:
            notes.append("인용 문서번호 모양 %d개 추가" % added)
        s = cfg.setdefault("settings", {})
        for kind, key in (("표", "표_캡션_머리말"), ("그림", "그림_캡션_머리말")):
            words = _sum_counters((st.get("caption_words") or {}).get(kind) for st in stds)
            have = [w.strip() for w in str(s.get(key, "")).split(",") if w.strip()]
            extra = [w for w, _ in words.most_common() if w not in have]
            if extra:
                s[key] = ",".join(have + extra)
                notes.append("%s 캡션 머리말 추가: %d개" % (kind, len(extra)))
        forms = {}
        for st in stds:
            for key, d in (st.get("latin_forms") or {}).items():
                forms.setdefault(key, Counter()).update(d)
        cfg["latin_pref"] = {k: v.most_common(1)[0][0] for k, v in forms.items() if len(v) >= 1}
        gl = []
        for st in stds:
            for t in st.get("terms", []):
                if t["term"] and t not in gl:
                    gl.append(t)
        cfg["glossary"] = gl
        notes.append("승인 표준 %d건의 용어 %d개·영문 표기 %d개 참고" % (len(stds), len(gl), len(cfg["latin_pref"])))
    style = minutes_style(profile)
    if style:
        cfg["minutes_style"] = style
        roles = style.get("roles") or {}
        mapping = {"seq": roles.get("seq"), "item": roles.get("category"), "location": roles.get("location"),
                   "fix": roles.get("comment")}
        used = {c["header"] for c in cfg.get("columns", [])}
        for col in cfg.get("columns", []):
            new = mapping.get(col["key"])
            if new and new != col["header"] and new not in used:
                used.discard(col["header"])
                col["header"] = new
                used.add(new)
        notes.append("회의록 문체 적용(끝맺음 「%s」)" % (style["ending"] or "기본"))
    return cfg, notes


def summary(profile):
    """학습 현황 — 건수만 보여 줍니다(내용은 보여 주지 않음)."""
    lines = []
    for t, tpl in (profile.get("templates") or {}).items():
        lines.append("Template[%s] %s: 상위 항목 %d개, 하위 항목 %d개, 예시 문구 %d줄, 머리글 자리표시 %d줄, 개정 이력 %s, 흐름도 절 %d개" % (
            t, tpl["source"], len(tpl["items"]), sum(len(i["subs"]) for i in tpl["items"]), len(tpl["example_lines"]),
            len(tpl["header_placeholders"]), "있음" if tpl["has_revision"] else "없음", len(tpl["flow_sections"])))
    for name, st in (profile.get("standards") or {}).items():
        c = st["counts"]
        lines.append("기존 표준 %s: 절 %d개, 용어 %d개, 인용 번호 모양 %d개, 캡션 %d개(방식 %d종), 영문 표기 %d종" % (
            name, c["절"], c["용어"], c["인용번호모양"], c["캡션"], len(st["caption_styles"]), len(st["latin_forms"])))
    for name, m in (profile.get("minutes") or {}).items():
        lines.append("회의록 %s: 지적 %d건, 범주 %d종, 끝맺음 %d종, 열 %d개" % (
            name, m["count"], len(m["categories"]), len(m["endings"]), len(m["headers"])))
    return lines


# ---------- 회의록 문체로 고쳐 쓰기 ----------

STEM_RE = re.compile(r"\s*주시기\s*바랍니다\.?\s*$")


def restyle(hint, style, location=""):
    """도구의 기본 수정 요청 문구(「…해 주시기 바랍니다.」)를 학습한 회의록 끝맺음으로 바꿉니다.
    바꿀 수 없는 문장은 그대로 둡니다(억지로 바꾸어 뜻이 흐려지지 않게)."""
    text = str(hint or "").strip()
    if not style or not text:
        return text
    fam = style.get("ending") or ""
    m = STEM_RE.search(text)
    if m and fam and fam != "바랍니다":
        stem = text[:m.start()].rstrip()
        noun = stem[:-1].rstrip() if stem.endswith("해") else None
        if noun:
            # 「X를 추가」 → 「X 추가」: 명사형 끝맺음 앞의 목적격 조사를 뗍니다.
            noun = re.sub(r"(\S)(을|를)\s+(\S+)$", r"\1 \3", noun)
        if fam == "요청":
            text = stem + " 주실 것을 요청드립니다."
        elif fam in ("필요", "요망") and noun:
            text = "%s %s" % (noun, fam)
        elif fam == "할 것" and noun:
            text = noun + "할 것"
        elif fam == "검토" and noun:
            text = noun + " 검토 바랍니다."
        elif fam in ("필요", "요망", "할 것", "함", "검토"):
            text = stem + "야 함"
    lead = style.get("lead") or ""
    loc_style = style.get("location") or ""
    sec = re.match(r"^\s*(\d+(?:\.\d+)*)", str(location or ""))
    if loc_style and sec:
        n = sec.group(1)
        pre = {"(n)": "(%s) " % n, "[n]": "[%s] " % n, "n항": "%s항 " % n, "n절": "%s절 " % n}.get(loc_style, "")
        text = pre + text
    if lead:
        text = (lead if lead in "-•·▶▷○●◦※*" else lead) + " " + text
    return text
