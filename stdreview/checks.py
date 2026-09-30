"""규칙 점검 — 기획서 5.1 의 「자동」「보조」 항목.

모든 함수는 순수 함수입니다(파일·DB 를 건드리지 않음). doc 은 docx_reader.build 결과,
cfg 는 기준표 설정(defaults.default_config 와 같은 모양)입니다.
결과 한 건 = {"no", "result", "detail", "location", "hint"}.
hint 는 수정 요청 문구 틀의 {내용} 자리에 들어갑니다.
"""

import re
from collections import Counter

from .docx_reader import in_section, top_of, has_section, section_blocks

# ---------- 공통 도우미 ----------


def setting(cfg, key, default=""):
    return str(cfg.get("settings", {}).get(key, default) or default)


def setting_num(cfg, key, default):
    try:
        return float(setting(cfg, key, str(default)))
    except ValueError:
        return float(default)


def split_list(s):
    return [x.strip() for x in str(s or "").split(",") if x.strip()]


def norm(s):
    return re.sub(r"[\s·/()（）\[\]\-_.,:：]", "", str(s or "")).lower()


def loc(b, prefix=""):
    sec = b.get("sec") or "머리말"
    snip = (b.get("text") or "").strip()
    if len(snip) > 28:
        snip = snip[:28] + "…"
    where = ("%s절" % sec) if sec != "머리말" else sec
    if prefix:
        where += " " + prefix
    return "%s 「%s」" % (where, snip) if snip else where


def _has_batchim(word):
    w = str(word or "").rstrip(" 」』)）]\"'")
    if not w:
        return False
    ch = w[-1]
    if "가" <= ch <= "힣":
        return (ord(ch) - 0xAC00) % 28 != 0
    if ch.isdigit():
        return ch in "013678"
    if ch.isalpha():
        return ch.lower() in "lmn"
    return False


def j(word, pair):
    """조사 붙이기: j("표", "이가") → "표가", j("그림", "이가") → "그림이". pair: 이가/을를/은는/과와/으로."""
    word = str(word)
    if pair == "으로":
        w = word.rstrip(" 」』)）]")
        last = w[-1] if w else ""
        rieul = "가" <= last <= "힣" and (ord(last) - 0xAC00) % 28 == 8
        return word + ("로" if (not _has_batchim(word) or rieul) else "으로")
    a, b = pair[0], pair[1]
    return word + (a if _has_batchim(word) else b)


def finding(no, result, detail, location="", hint=""):
    return {"no": no, "result": result, "detail": detail, "location": location, "hint": hint}


def _caption_re(prefixes):
    alts = sorted((re.escape(p) for p in prefixes), key=len, reverse=True)
    return re.compile(r"^\s*[\[<〈(]?\s*(?:%s)\s*\.?\s*(\d+(?:[-.]\d+)*)" % "|".join(alts), re.I)


def _ref_re(prefixes):
    alts = sorted((re.escape(p) for p in prefixes), key=len, reverse=True)
    return re.compile(r"(?:%s)\s*\.?\s*(\d+(?:-\d+)*)" % "|".join(alts), re.I)


def is_caption_like(b):
    """캡션 스타일이거나, 짧고 문장(「…다.」)으로 끝나지 않는 문단만 캡션 후보로 봅니다.
    「표 3의 입력값으로 다시 계산한다.」 같은 본문 문장을 캡션으로 오인하지 않기 위함입니다."""
    style = (b.get("style") or "").lower()
    if style in ("caption", "캡션") or "caption" in style:
        return True
    t = b.get("text", "").strip()
    return len(t) <= 80 and not re.search(r"(다|요|음|함)\.?$", t)


def captions(doc, cfg):
    """표·그림 캡션을 찾고 표·그림 객체와 짝짓습니다."""
    tprefix = split_list(setting(cfg, "표_캡션_머리말", "표,Table"))
    fprefix = split_list(setting(cfg, "그림_캡션_머리말", "그림,Fig.,Fig,Figure"))
    tre, fre = _caption_re(tprefix), _caption_re(fprefix)
    blocks = doc["blocks"]
    caps = []
    for b in blocks:
        if b["kind"] != "p" or not b["text"] or b.get("heading") or b.get("rev"):
            continue
        if not is_caption_like(b):
            continue
        m = tre.match(b["text"])
        if m:
            caps.append({"kind": "표", "num": m.group(1).rstrip("."), "idx": b["idx"], "sec": b.get("sec", ""), "text": b["text"], "used": False})
            continue
        m = fre.match(b["text"])
        if m:
            caps.append({"kind": "그림", "num": m.group(1).rstrip("."), "idx": b["idx"], "sec": b.get("sec", ""), "text": b["text"], "used": False})
    by_idx = {c["idx"]: c for c in caps}

    def neighbor(i, step):
        j = i + step
        while 0 <= j < len(blocks):
            nb = blocks[j]
            if nb["kind"] == "p" and not nb["text"] and not nb.get("figure"):
                j += step
                continue
            return j
        return None

    objects = []
    for b in blocks:
        if b.get("rev"):
            continue                      # 개정 이력 표는 번호 캡션 대상이 아닙니다.
        if b["kind"] == "tbl":
            kind = "표"
        elif b.get("figure"):
            kind = "그림"
        else:
            continue
        cap = None
        if kind == "그림" and b["idx"] in by_idx and by_idx[b["idx"]]["kind"] == "그림":
            cap = by_idx[b["idx"]]
        order = (-1, 1) if kind == "표" else (1, -1)
        for step in order:
            if cap:
                break
            j = neighbor(b["idx"], step)
            if j is not None and j in by_idx and by_idx[j]["kind"] == kind and not by_idx[j]["used"]:
                cap = by_idx[j]
        if cap:
            cap["used"] = True
        objects.append({"kind": kind, "block": b, "caption": cap})
    return caps, objects, (tprefix, fprefix)


def _is_int(s):
    return s.isdigit()


# ---------- 기준별 점검 ----------


def check_template_items(doc, cfg, std_type, no):
    out = []
    items = sorted([r for r in cfg.get("items", []) if r.get("type") == std_type], key=lambda r: int(r.get("order") or 0))
    tops = [h for h in doc["headings"] if h["level"] == 1]
    top_nums = [h["num"] for h in tops]
    if not items:
        return [finding(no, "판단 필요", "표준 종류 「%s」의 Template 항목이 기준표 「항목」 시트에 없습니다." % std_type, "", "")]
    for it in items:
        num = str(it.get("num"))
        label = "%s항%s" % (num, (" " + it["name"].split("|")[0]) if it.get("name") else "")
        h = next((h for h in tops if h["num"] == num), None)
        if not h:
            if it.get("required"):
                out.append(finding(no, "부적합", "Template %s 없습니다." % j(label, "이가"), "문서 전체",
                                   "%s 추가해 주시기 바랍니다." % j(label, "을를")))
            else:
                out.append(finding(no, "권고", "선택 항목 %s 없습니다." % j(label, "이가"), "문서 전체",
                                   "필요하면 %s 추가해 주시기 바랍니다." % j(label, "을를")))
            continue
        names = [n for n in str(it.get("name") or "").split("|") if n.strip()]
        if names:
            hn = norm(h["title"])
            if not any(norm(n) in hn or hn in norm(n) for n in names if norm(n)):
                out.append(finding(no, "권고", "%s항 제목 %s Template 항목명 %s 다릅니다." % (num, j("「%s」" % h["title"], "이가"), j("「%s」" % " / ".join(names), "과와")),
                                   "%s절 「%s」" % (num, h["title"]),
                                   "%s항 제목을 %s 맞춰 주시기 바랍니다." % (num, j("「%s」" % names[0], "으로"))))
    # 순서
    order = {str(it["num"]): int(it.get("order") or 0) for it in items}
    seq = [order[n] for n in top_nums if n in order]
    if seq != sorted(seq):
        out.append(finding(no, "부적합", "상위 항목 순서가 Template 과 다릅니다: %s" % " → ".join(top_nums), "문서 전체",
                           "항목 순서를 Template 순서(%s)로 바로잡아 주시기 바랍니다." % " → ".join(str(it["num"]) for it in items)))
    dup = sorted({n for n in top_nums if top_nums.count(n) > 1})
    if dup:
        out.append(finding(no, "부적합", "같은 항목 번호가 두 번 이상 있습니다: %s" % ", ".join(dup), "문서 전체",
                           "항목 번호가 겹치지 않게 고쳐 주시기 바랍니다."))
    extra = [h for h in tops if h["num"] not in order]
    for h in extra:
        out.append(finding(no, "권고", "Template 에 없는 상위 항목 %s 있습니다." % j("「%s %s」" % (h["num"], h["title"]), "이가"),
                           "%s절" % h["num"], "Template 항목 안으로 옮기거나 8항 Appendix 로 옮겨 주시기 바랍니다."))
    out += check_subitems(doc, cfg)
    if not out:
        out.append(finding(no, "적합", "Template 항목 %d개가 순서대로 모두 있습니다." % len(items), "", ""))
    return out


def check_subitems(doc, cfg):
    """학습한 Template 의 하위 항목(2.1, 4.2 …)이 있는지, 이름이 같은지(학습 프로필이 있을 때만)."""
    tpl = cfg.get("template")
    if not tpl:
        return []
    no = "공통-1"
    out = []
    have = {h["num"]: h for h in doc["headings"]}
    for it in tpl["items"]:
        if str(it["num"]) not in have:
            continue                      # 상위 항목이 없으면 위에서 이미 지적
        for sub in it.get("subs", []):
            h = have.get(sub["num"])
            if not h:
                out.append(finding(no, "부적합", "Template 하위 항목 %s 없습니다." % j("%s %s" % (sub["num"], sub["name"]), "이가"),
                                   "%s절" % it["num"], "%s 추가해 주시기 바랍니다." % j("%s %s" % (sub["num"], sub["name"]), "을를")))
                continue
            a, b = norm(h["title"]), norm(sub["name"])
            if b and not (a in b or b in a):
                out.append(finding(no, "권고", "%s절 제목 %s Template 하위 항목명 %s 다릅니다." % (sub["num"], j("「%s」" % h["title"], "이가"), j("「%s」" % sub["name"], "과와")),
                                   "%s절 「%s」" % (sub["num"], h["title"]),
                                   "%s절 제목을 Template 대로 %s 맞추거나, 내용이 다르면 Template 하위 항목 아래로 옮겨 주시기 바랍니다." % (sub["num"], j("「%s」" % sub["name"], "으로"))))
    return out


def title_keywords(title):
    stop = {"예시", "표준", "기술", "기술표준", "설계표준", "시험검증표준", "가상검증표준", "에", "대한", "관한", "의", "및", "등"}
    words = re.split(r"[\s·,()（）/\-]+", title or "")
    out = []
    for w in words:
        w = re.sub(r"(에|의|을|를|과|와|에서|으로|로)$", "", w)
        if len(w) >= 2 and w not in stop and w not in out:
            out.append(w)
    return out


def check_purpose(doc, cfg, crit):
    no = crit["no"]
    if not has_section(doc, "1"):
        return [finding(no, "부적합", "1항을 찾지 못했습니다.", "문서 전체", "1항 목적 및 정의를 작성해 주시기 바랍니다.")]
    body = " ".join(b["text"] for b in section_blocks(doc, "1") if b["text"])
    kws = title_keywords(doc.get("title", ""))
    found = [k for k in kws if k in body]
    miss = [k for k in kws if k not in body]
    detail = "1항 본문 %d자. 제목 핵심어 중 1항에 나온 것: %s / 안 나온 것: %s. 목적·정의의 구체성은 판단 프롬프트로 확인합니다." % (
        len(body), ", ".join(found) or "없음", ", ".join(miss) or "없음")
    return [finding(no, "판단 필요", detail, "1절", "")]


def find_std_numbers(text, cfg):
    """표준번호 형식 정규식으로 번호를 찾습니다. 겹치면 긴 쪽을 남깁니다."""
    spans = []
    for p in cfg.get("std_patterns", []):
        try:
            rx = re.compile(p["regex"])
        except re.error:
            continue
        for m in rx.finditer(text or ""):
            spans.append((m.start(), m.end(), m.group(0)))
    spans.sort(key=lambda s: (s[0], -(s[1] - s[0])))
    out, end = [], -1
    for s, e, g in spans:
        if s >= end:
            out.append((s, e, g))
            end = e
    return out


def std_key(num):
    return re.sub(r"\s+", "", num).upper()


def _base_key(k):
    """연도(「:2011」)를 뗀 비교 키."""
    return re.sub(r":\d{4}$", "", k)


def is_listed(k, listed):
    """본문 인용 번호가 2항 목록에 있는지 — 연도 유무, 국가 채택 접두(KS M ISO …)의 차이는 같은 문서로 봅니다."""
    b = _base_key(k)
    for x in listed:
        y = _base_key(x)
        if b == y or (len(b) >= 6 and len(y) >= 6 and (y.endswith(b) or b.endswith(y))):
            return True
    return False


NO_NUMBER_SECTION_RE = re.compile(r"참고\s*문헌|참고\s*자료|참고\s*도서|Bibliography|Reading", re.I)


def citation_lines(doc):
    """2항의 목록 줄(문단 + 표 행). 「참고문헌」처럼 문서번호가 없는 하위 절은 뺍니다."""
    lines = []
    skip = [h["num"] for h in doc["headings"] if h["level"] >= 2 and h["num"].split(".")[0] == "2" and NO_NUMBER_SECTION_RE.search(h["title"])]
    blocks = [b for b in section_blocks(doc, "2") if not any(in_section(b.get("sec", ""), n) for n in skip)]
    has_table = any(b["kind"] == "tbl" for b in blocks)
    for b in blocks:
        if b["kind"] == "p":
            t = b.get("text") or ""
            # 표가 있으면 표 행만 목록으로 보고, 캡션·안내 문장(「…다.」)은 목록에서 뺍니다.
            if has_table or not t or re.match(r"^\s*(표|그림|Table|Fig)", t, re.I) or t.rstrip().endswith("다."):
                continue
        if b["kind"] == "tbl":
            for ri, r in enumerate(b["rows"]):
                t = " ".join(c for c in r if c).strip()
                if not t:
                    continue
                if ri == 0 and ("번호" in t or "제목" in t) and not re.search(r"\d", t):
                    continue
                lines.append((t, b))
        elif b["text"]:
            lines.append((b["text"], b))
    return lines


def check_citation_list(doc, cfg, crit):
    no = crit["no"]
    if not has_section(doc, "2"):
        return [finding(no, "부적합", "2항을 찾지 못했습니다.", "문서 전체", "2항 관련 표준 및 인용을 작성해 주시기 바랍니다.")]
    lines = citation_lines(doc)
    if not lines:
        return [finding(no, "부적합", "2항에 관련 표준 목록이 없습니다.", "2절", "관련 표준의 문서번호와 제목을 적어 주시기 바랍니다.")]
    out = []
    ok = 0
    for t, b in lines:
        nums = find_std_numbers(t, cfg)
        if not nums:
            out.append(finding(no, "부적합", "문서번호 형식에 맞는 번호가 없습니다: 「%s」" % t[:60], loc(b),
                               "「%s」에 문서번호를 정해진 형식으로 적어 주시기 바랍니다." % t[:30]))
            continue
        rest = t
        for _, _, g in nums:
            rest = rest.replace(g, " ")
        rest = re.sub(r"[\s\-–—:：|,.()（）\[\]0-9]+", "", rest)
        if len(rest) < 2:
            out.append(finding(no, "부적합", "문서번호 %s 에 제목이 없습니다." % ", ".join(g for _, _, g in nums), loc(b),
                               "%s 의 제목을 함께 적어 주시기 바랍니다." % nums[0][2]))
            continue
        ok += 1
    if not out:
        out.append(finding(no, "적합", "2항 목록 %d줄 모두 문서번호와 제목이 짝으로 있습니다." % ok, "2절", ""))
    return out


def term_entries(doc):
    """3항 용어정의 — 표(용어 | 설명), 「용어: 설명」 문단, 또는 「3.1 용어」 하위 제목 + 설명 문단."""
    entries = []
    subs = [h for h in doc["headings"] if h["level"] == 2 and h["num"].split(".")[0] == "3"]
    if subs:
        for h in subs:
            desc = " ".join(b["text"] for b in doc["blocks"] if b.get("sec") == h["num"] and not b.get("heading") and b.get("text"))
            entries.append({"raw": h["title"], "term": term_key(h["title"]), "desc": desc.strip(), "block": doc["blocks"][h["idx"]]})
    for b in section_blocks(doc, "3"):
        if subs and b["kind"] != "tbl":
            continue
        if b["kind"] == "tbl":
            for ri, r in enumerate(b["rows"]):
                cells = [c.strip() for c in r]
                if len(cells) < 2:
                    continue
                if ri == 0 and ("용어" in cells[0] and len(cells[0]) <= 6):
                    continue
                if not cells[0]:
                    continue
                entries.append({"raw": cells[0], "term": term_key(cells[0]), "desc": " ".join(cells[1:]).strip(), "block": b})
        elif b["text"]:
            m = re.match(r"^\s*([^:：]{1,40}?)\s*[:：]\s*(.*)$", b["text"])
            if m:
                entries.append({"raw": m.group(1), "term": term_key(m.group(1)), "desc": m.group(2).strip(), "block": b})
    return entries


TERM_LEAD_RE = re.compile(r"^\s*(?:[•⚫●▪■□◦○◆◇▶▷※\-–·*]\s*)?(?:\d{1,2}\)|\(\d{1,2}\)|[①-⑳])?\s*")


def term_key(raw):
    """「ECU (Electronic Control Unit)」 → 「ECU」. 앞의 글머리 기호·「1)」 번호는 뗍니다."""
    raw = TERM_LEAD_RE.sub("", str(raw or ""))
    t = re.split(r"[(（]", raw)[0].strip()
    return t or raw.strip()


def check_term_def(doc, cfg, crit):
    no = crit["no"]
    if not has_section(doc, "3"):
        return [finding(no, "부적합", "3항을 찾지 못했습니다.", "문서 전체", "3항 용어정의를 작성해 주시기 바랍니다.")]
    entries = term_entries(doc)
    if not entries:
        return [finding(no, "부적합", "3항에서 용어정의(표 또는 「용어: 설명」)를 찾지 못했습니다.", "3절",
                        "용어와 설명을 표로 정리해 주시기 바랍니다.")]
    out = []
    short = int(setting_num(cfg, "짧은_설명_글자수", 10))
    empty = [e for e in entries if not e["desc"]]
    for e in empty:
        out.append(finding(no, "부적합", "용어 「%s」의 설명이 비어 있습니다." % e["term"], loc(e["block"]),
                           "용어 「%s」의 설명을 이 표준에 적용되는 뜻으로 적어 주시기 바랍니다." % e["term"]))
    shorts = [e["term"] for e in entries if e["desc"] and len(e["desc"]) < short]
    detail = "용어 %d개. 설명이 %d자 미만으로 짧은 용어: %s. 일반 상식 수준인지는 판단 프롬프트로 확인합니다." % (
        len(entries), short, ", ".join(shorts) or "없음")
    out.append(finding(no, "판단 필요", detail, "3절", ""))
    return out


def _variants_for(cfg, term):
    for v in cfg.get("variants", []):
        if str(v.get("term", "")).strip().lower() == term.lower():
            return [x for x in v.get("variants", []) if x]
    return []


def check_term_consistency(doc, cfg, crit):
    no = crit["no"]
    entries = term_entries(doc)
    terms = []
    for e in entries:
        if e["term"] and e["term"] not in terms:
            terms.append(e["term"])
    # 용어정의 표에 없어도 사전에 있는 용어는 점검
    for v in cfg.get("variants", []):
        t = str(v.get("term", "")).strip()
        if t and t not in terms:
            terms.append(t)
    # 인용 목록 줄(2항, 글머리 기호 + 문서번호)은 원 문서 제목 표기를 그대로 옮기므로 표기 통일 대상에서 뺍니다.
    body = [b for b in doc["blocks"] if not in_section(b.get("sec", ""), "3") and not b.get("is_title") and b.get("text") and not b.get("rev")
            and not in_section(b.get("sec", ""), "2") and not (re.match(r"^\s*[•⚫●▪■□◦○◆◇▶▷\-–·*]", b["text"]) and find_std_numbers(b["text"], cfg))]
    out = []
    reported = set()
    for t in terms:
        used = 0
        wrong = {}
        ascii_term = bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 \-/]*", t))
        # 뒤에 붙은 각주 번호(「NBR3)」)는 같은 용어로 봅니다.
        end = r"(?![A-Za-z])(?!\d+(?![\d)]))"
        pat_exact = re.compile(r"(?<![A-Za-z0-9])%s%s" % (re.escape(t), end)) if ascii_term else re.compile(re.escape(t))
        pat_ci = re.compile(r"(?<![A-Za-z0-9])%s%s" % (re.escape(t), end), re.I) if ascii_term else None
        # 한글 용어는 띄어쓰기만 다른 표기(「사이드 링」/「사이드링」)도 찾습니다.
        pat_sp = None
        if not ascii_term and len(t.replace(" ", "")) >= 2:
            pat_sp = re.compile(r"\s*".join(re.escape(ch) for ch in t.replace(" ", "")))
        vars_ = _variants_for(cfg, t)
        spacing = {}
        for b in body:
            text = b["text"]
            used += len(pat_exact.findall(text))
            if pat_sp:
                for m in pat_sp.finditer(text):
                    if m.group(0) != t:
                        spacing.setdefault(m.group(0), []).append(b)
            if pat_ci:
                for m in pat_ci.finditer(text):
                    if m.group(0) != t:
                        wrong.setdefault(m.group(0), []).append(b)
            for v in vars_:
                if v in text:
                    wrong.setdefault(v, []).append(b)
        for w, bl in spacing.items():
            locs = "; ".join(loc(b) for b in bl[:5]) + (" 외 %d곳" % (len(bl) - 5) if len(bl) > 5 else "")
            out.append(finding(no, "권고", "용어 %s 띄어쓰기가 다른 %s %d곳에 쓰였습니다." % (j("「%s」" % t, "이가"), j("「%s」" % w, "으로"), len(bl)), locs,
                               "「%s」 표기를 용어정의대로 %s 통일해 주시기 바랍니다." % (w, j("「%s」" % t, "으로"))))
        for w, bl in wrong.items():
            locs = "; ".join(loc(b) for b in bl[:5]) + (" 외 %d곳" % (len(bl) - 5) if len(bl) > 5 else "")
            out.append(finding(no, "부적합", "용어 %s %s %d곳에 다르게 쓰였습니다." % (j("「%s」" % t, "이가"), j("「%s」" % w, "으로"), len(bl)), locs,
                               "「%s」 표기를 용어정의대로 %s 통일해 주시기 바랍니다." % (w, j("「%s」" % t, "으로"))))
        if t in [e["term"] for e in entries] and used == 0 and not wrong and not spacing:
            out.append(finding(no, "권고", "용어정의에 있는 %s 본문에 쓰이지 않았습니다." % j("「%s」" % t, "이가"), "3절",
                               "쓰이지 않는 용어 「%s」는 용어정의에서 빼거나 본문에 맞게 고쳐 주시기 바랍니다." % t))
        for w in wrong:
            reported.add((t.lower().replace("-", ""), w))
    defined = {t.lower().replace("-", "").replace(" ", ""): t for t in terms if re.fullmatch(r"[A-Za-z][A-Za-z\-]*", t)}
    out += check_latin_variants(doc, cfg, crit, defined, reported)
    if not out:
        out.append(finding(no, "적합", "용어 %d개가 본문에서 같은 표기로 쓰였습니다." % len(terms), "", ""))
    return out


def check_section_sequence(doc, cfg, crit):
    no = crit["no"]
    hs = [h for h in doc["headings"] if top_of(h["num"]) >= 4]
    if not hs:
        return [finding(no, "부적합", "4항 이하 절을 찾지 못했습니다.", "문서 전체", "4항 설계절차/시험절차를 항목 순서대로 작성해 주시기 바랍니다.")]
    children = {}
    for h in hs:
        parts = h["num"].split(".")
        if len(parts) >= 2:
            children.setdefault(".".join(parts[:-1]), []).append(int(parts[-1]))
    out = []
    for parent, nums in children.items():
        seen = set()
        for n in nums:
            if n in seen:
                out.append(finding(no, "부적합", "%s.%d 절 번호가 두 번 있습니다." % (parent, n), "%s.%d절" % (parent, n),
                                   "%s 아래 절 번호가 겹치지 않게 고쳐 주시기 바랍니다." % parent))
            seen.add(n)
        expect = list(range(1, max(nums) + 1))
        missing = [x for x in expect if x not in seen]
        if missing:
            out.append(finding(no, "부적합", "%s 아래 절 번호가 빠졌습니다: %s" % (parent, ", ".join("%s.%d" % (parent, m) for m in missing)),
                               "%s절" % parent, "%s 아래 절 번호를 빠짐없이 차례대로 매겨 주시기 바랍니다." % parent))
        if nums != sorted(nums):
            out.append(finding(no, "부적합", "%s 아래 절 순서가 뒤바뀌었습니다: %s" % (parent, ", ".join("%s.%d" % (parent, n) for n in nums)),
                               "%s절" % parent, "%s 아래 절을 번호 순서대로 배치해 주시기 바랍니다." % parent))
    outline = " / ".join("%s %s" % (h["num"], h["title"]) for h in hs if h["level"] <= 3)
    if not out:
        out.append(finding(no, "적합", "4항 이하 절 번호가 빠짐·중복 없이 이어집니다. 절 목록: %s" % outline, "4절 이하", ""))
    else:
        out.append(finding(no, "판단 필요", "절 목록(순서 확인용): %s" % outline, "4절 이하", ""))
    return out


def check_flowchart(doc, cfg, crit):
    no = crit["no"]
    words = [w.lower() for w in split_list(setting(cfg, "흐름도_캡션_낱말"))]
    _, objects, _ = captions(doc, cfg)
    secs = ["4.1", "4.2"] + [s for s in (cfg.get("template") or {}).get("flow_sections", []) if s not in ("4.1", "4.2")]
    figs = [o for o in objects if o["kind"] == "그림" and any(in_section(o["block"].get("sec", ""), s) for s in secs)]
    flow = [o for o in figs if o["caption"] and any(w in o["caption"]["text"].lower() for w in words)]
    if flow:
        o = flow[0]
        return [finding(no, "적합", "%s절에 흐름도 그림이 있습니다: 「%s」. ISO 기호 준수 여부는 검토자가 그림을 보고 확인합니다." % (o["block"]["sec"], o["caption"]["text"][:40]),
                        loc(o["caption"] and {"sec": o["block"]["sec"], "text": o["caption"]["text"]}), "")]
    if figs:
        return [finding(no, "권고", "%s절에 그림이 %d개 있으나 캡션에 흐름도 낱말(%s)이 없습니다." % ("·".join(secs), len(figs), ", ".join(split_list(setting(cfg, "흐름도_캡션_낱말")))),
                        "·".join(secs) + "절", "절차 흐름도라면 캡션에 흐름도(FlowChart)임을 밝혀 주시기 바랍니다.")]
    return [finding(no, "부적합", "%s절에 흐름도(FlowChart) 그림이 없습니다." % "·".join(secs), "·".join(secs) + "절",
                    "4.1 또는 4.2절에 설계·시험 절차를 ISO 기준 흐름도(FlowChart)로 제시해 주시기 바랍니다.")]


def check_equations(doc, cfg, crit):
    no = crit["no"]
    counts = {}
    for b in doc["blocks"]:
        if top_of(b.get("sec", "")) >= 4 and b.get("omath"):
            counts[b["sec"]] = counts.get(b["sec"], 0) + b["omath"]
    total4 = sum(v for k, v in counts.items() if top_of(k) == 4)
    if doc.get("source") == "pdf":
        return [finding(no, "판단 필요", "PDF 에서는 Word 수식 개체를 셀 수 없습니다. 수식 제시 여부는 검토자가 확인합니다(원본 .docx 를 넣으면 자동으로 셉니다).", "4절 이하", "")]
    detail = "절별 Word 수식 개수: %s" % (", ".join("%s절 %d개" % kv for kv in sorted(counts.items())) or "없음")
    if total4 == 0:
        return [finding(no, "권고", "4항 설계·시험 절차에 Word 수식이 없습니다. " + detail, "4절",
                        "가능하면 절차의 계산·판정 기준을 수식으로 제시해 주시기 바랍니다.")]
    return [finding(no, "적합", detail, "4절 이하", "")]


def check_captions(doc, cfg, crit):
    no = crit["no"]
    caps, objects, (tprefix, fprefix) = captions(doc, cfg)
    out = []
    for kind in ("표", "그림"):
        styles = {}
        for c in caps:
            if c["kind"] == kind:
                m = re.match(r"^\s*([\[<〈(]?)\s*([^\s\d.]+)\s*(\.?)", c["text"])
                if m:
                    close = {"[": "]", "<": ">", "〈": "〉", "(": ")"}.get(m.group(1), "")
                    styles.setdefault("%s%s%s n%s" % (m.group(1), m.group(2), m.group(3), close), []).append(c)
        if len(styles) > 1:
            main = max(styles, key=lambda k: len(styles[k]))
            odd = [c for k, v in styles.items() if k != main for c in v]
            out.append(finding(no, "권고", "%s 캡션 형식이 %d가지로 섞여 있습니다(가장 많은 형식 「%s」과 다른 캡션 %d개)." % (kind, len(styles), main, len(odd)),
                               "; ".join(loc({"sec": c["sec"], "text": c["text"]}) for c in odd[:3]),
                               "%s 캡션을 「%s」 형식 하나로 맞춰 주시기 바랍니다." % (kind, main)))
    for o in objects:
        if not o["caption"]:
            b = o["block"]
            what = "표" if o["kind"] == "표" else "그림"
            snip = (b["rows"][0][0] if b["kind"] == "tbl" and b["rows"] and b["rows"][0] else "")[:20]
            where = "%s절 %s%s" % (b.get("sec") or "머리말", what, (" 「%s…」" % snip) if snip else "")
            out.append(finding(no, "부적합", "번호 캡션이 없는 %s 있습니다." % j(what, "이가"), where,
                               "이 %s에 「%s n」 형식의 번호 캡션을 달아 주시기 바랍니다." % (what, "표" if what == "표" else "그림(Fig)")))
    for kind in ("표", "그림"):
        nums = [c["num"] for c in caps if c["kind"] == kind]
        seen = []
        for c in [c for c in caps if c["kind"] == kind]:
            if c["num"] in seen:
                out.append(finding(no, "부적합", "%s %s 번호가 두 번 쓰였습니다." % (kind, c["num"]), loc({"sec": c["sec"], "text": c["text"]}),
                                   "%s 번호가 겹치지 않게 다시 매겨 주시기 바랍니다." % kind))
            seen.append(c["num"])
        ints = [int(n) for n in nums if _is_int(n)]
        if ints and len(ints) == len(nums):
            uniq = sorted(set(ints))
            missing = [x for x in range(1, max(uniq) + 1) if x not in uniq]
            if missing:
                out.append(finding(no, "부적합", "%s 번호가 빠졌습니다: %s" % (kind, ", ".join("%s %d" % (kind, m) for m in missing)), "문서 전체",
                                   "%s 번호를 1부터 빠짐없이 차례대로 매겨 주시기 바랍니다." % kind))
            if ints != sorted(ints):
                out.append(finding(no, "부적합", "%s 번호가 나오는 순서와 다릅니다: %s" % (kind, ", ".join(nums)), "문서 전체",
                                   "%s 번호를 문서에 나오는 순서대로 매겨 주시기 바랍니다." % kind))
    # 본문 참조
    known = {"표": {c["num"] for c in caps if c["kind"] == "표"}, "그림": {c["num"] for c in caps if c["kind"] == "그림"}}
    cap_idx = {c["idx"] for c in caps}
    for kind, pre in (("표", tprefix), ("그림", fprefix)):
        rx = _ref_re(pre)
        for b in doc["blocks"]:
            if b["kind"] != "p" or b["idx"] in cap_idx or not b.get("text") or b.get("heading"):
                continue
            for m in rx.finditer(b["text"]):
                n = m.group(1)
                if n not in known[kind]:
                    out.append(finding(no, "부적합", "본문이 참조하는 %s 문서에 없습니다." % j("「%s」" % m.group(0).strip(), "이가"), loc(b),
                                       "「%s」 참조를 실제 %s 번호에 맞게 고쳐 주시기 바랍니다." % (m.group(0).strip(), kind)))
    if not out and doc.get("source") == "pdf":
        out.append(finding(no, "적합", "캡션 표 %d개, 그림 %d개의 번호·참조가 맞습니다. PDF 에서는 캡션이 아예 없는 표·그림은 찾지 못하므로 검토자가 확인합니다." % (
            sum(1 for c in caps if c["kind"] == "표"), sum(1 for c in caps if c["kind"] == "그림")), "", ""))
    if not out:
        out.append(finding(no, "적합", "표 %d개, 그림 %d개 모두 번호 캡션이 있고 번호·참조가 맞습니다." % (
            sum(1 for o in objects if o["kind"] == "표"), sum(1 for o in objects if o["kind"] == "그림")), "", ""))
    return out


def sentences(text):
    parts = re.split(r"(?<=[.!?다])\s+|\n", text or "")
    return [p.strip() for p in parts if p and p.strip()]


def check_citation_ref(doc, cfg, crit):
    no = crit["no"]
    listed = set()
    for t, _ in citation_lines(doc):
        for _, _, g in find_std_numbers(t, cfg):
            listed.add(std_key(g))
    out = []
    reported = set()
    for b in doc["blocks"]:
        if in_section(b.get("sec", ""), "2") or not b.get("text") or b.get("is_title") or b.get("heading"):
            continue
        for _, _, g in find_std_numbers(b["text"], cfg):
            k = std_key(g)
            if not is_listed(k, listed) and k not in reported:
                reported.add(k)
                out.append(finding(no, "부적합", "본문에 인용한 %s 2항 관련 표준 목록에 없습니다." % j(g, "이가"), loc(b),
                                   "%s 2항 관련 표준 목록에 문서번호·제목과 함께 추가하거나, 인용 번호를 바로잡아 주시기 바랍니다." % j(g, "을를")))
    unit = setting(cfg, "수치_단위_정규식")
    try:
        urx = re.compile(unit)
    except re.error:
        urx = None
    no_basis = []
    if urx:
        for b in doc["blocks"]:
            if top_of(b.get("sec", "")) < 4 or not b.get("text") or b["kind"] != "p" or b.get("heading"):
                continue
            for s in sentences(b["text"]):
                if urx.search(s) and not find_std_numbers(s, cfg):
                    no_basis.append((s, b))
    detail = "4항 이하에서 수치 기준이 있으나 근거 인용(표준번호)이 없는 문장 %d개%s. 정량화·근거 적정성은 판단 프롬프트로 확인합니다." % (
        len(no_basis), (": " + " / ".join("「%s」" % s[:40] for s, _ in no_basis[:5])) if no_basis else "")
    out.append(finding(no, "판단 필요", detail, "; ".join(loc(b) for _, b in no_basis[:3]) or "4절 이하", ""))
    return out


def check_section_judge(doc, cfg, crit):
    no = crit["no"]
    tokens = [t for t in str(crit.get("target", "")).split(",") if t.strip() and not t.strip().endswith("+")]
    main = tokens[-1].strip() if tokens else ""
    if main and not has_section(doc, main):
        return [finding(no, "부적합", "%s항을 찾지 못했습니다." % main, "문서 전체", "%s항을 작성해 주시기 바랍니다." % main)]
    body = " ".join(b["text"] for b in section_blocks(doc, main) if b.get("text")) if main else ""
    return [finding(no, "판단 필요", "%s항 본문 %d자. 내용 판단은 판단 프롬프트로 확인합니다." % (main, len(body)), "%s절" % main, "")]


def check_quant(doc, cfg, crit):
    no = crit["no"]
    if not has_section(doc, "5"):
        return [finding(no, "부적합", "5항을 찾지 못했습니다.", "문서 전체", "5항 설계/시험 검증을 작성해 주시기 바랍니다.")]
    unit = setting(cfg, "수치_단위_정규식")
    try:
        urx = re.compile(unit)
    except re.error:
        urx = None
    words = [w for w in cfg.get("qualitative", []) if w]
    quant, qual = [], []
    for b in section_blocks(doc, "5"):
        if b["kind"] != "p" or not b.get("text"):
            continue
        for s in sentences(b["text"]):
            if urx and urx.search(s):
                quant.append(s)
            elif any(w in s for w in words):
                qual.append(s)
    out = []
    if qual:
        out.append(finding(no, "부적합", "5항에 수치 없이 정성 표현만 있는 문장 %d개: %s" % (len(qual), " / ".join("「%s」" % s[:40] for s in qual[:5])),
                           "5절", "정성 표현(%s)을 수치·단위·허용 범위가 있는 정량 기준으로 바꿔 주시기 바랍니다." % ", ".join(sorted({w for w in words if any(w in s for s in qual)}))))
    out.append(finding(no, "판단 필요", "5항 문장 중 수치+단위가 있는 문장 %d개, 정성 표현만 있는 문장 %d개." % (len(quant), len(qual)), "5절", ""))
    return out


def check_exists(doc, cfg, crit, std_type):
    no = crit["no"]
    num = str(crit.get("target", "")).split(",")[0].strip().rstrip("+")
    if has_section(doc, num):
        h = next(h for h in doc["headings"] if h["num"] == num)
        return [finding(no, "적합", "%s항 %s 있습니다." % (num, j("「%s」" % h["title"], "이가")), "%s절" % num, "")]
    item = next((it for it in cfg.get("items", []) if it.get("type") == std_type and str(it.get("num")) == num), None)
    if item is not None and not item.get("required"):
        return [finding(no, "권고", "선택 항목 %s항이 없습니다." % num, "문서 전체", "필요하면 %s항을 추가해 주시기 바랍니다." % num)]
    return [finding(no, "부적합", "%s항이 없습니다." % num, "문서 전체", "%s항을 작성해 주시기 바랍니다." % num)]


# ---------- 머리글·개정 이력(공통-3), 미완성·양식 잔재(공통-4) ----------

UNFINISHED_RE = re.compile(r"추후.{0,40}?(작성|보완|반영|추가|정리|검토).{0,6}?예정|작성\s*예정|\bTBD\b|\bTBA\b|\(미정\)|미작성")


def _placeholder_re():
    from .profile import PLACEHOLDER_RE
    return PLACEHOLDER_RE


def check_revision(doc, cfg, crit):
    """머리글(표준번호·날짜·Rev)과 개정 이력 — 자리표시가 남았는지, 서로 맞는지."""
    no = crit["no"]
    ph = _placeholder_re()
    tpl = cfg.get("template") or {}
    out = []
    header = doc.get("header", [])
    for t in header:
        if ph.search(t) and not re.search(r"copyright|ⓒ|©", t, re.I):
            out.append(finding(no, "부적합", "머리글에 Template 자리표시가 그대로 있습니다: 「%s」" % t[:40], "머리글",
                               "머리글의 표준번호·날짜·Rev 를 실제 값으로 채워 주시기 바랍니다."))
    if tpl.get("title") and doc.get("title") and norm(doc["title"]) == norm(tpl["title"]):
        out.append(finding(no, "부적합", "표준 제목이 Template 자리 문구 그대로입니다.", "머리글", "표준 제목을 적어 주시기 바랍니다."))
    rev = doc.get("revision") or {}
    need = tpl.get("has_revision", False) or setting(cfg, "개정이력_필수", "아니오").strip() in ("예", "Y", "y")
    if not rev.get("found"):
        if need:
            out.append(finding(no, "부적합", "개정 이력이 없습니다.", "문서 끝",
                               "Template 의 개정 이력 표를 붙이고 Rev·작성일자·작성자·개정 사유를 적어 주시기 바랍니다."))
    else:
        rows = rev.get("rows", [])
        if not rows and need:
            out.append(finding(no, "부적합", "개정 이력 표에 기록 줄이 없습니다.", "개정 이력",
                               "Rev·작성일자·작성자·개정 사유를 적어 주시기 바랍니다."))
        for r in rows:
            if ph.search(r["date"]) or re.search(r"\.00(\.|$)", r["date"]):
                out.append(finding(no, "부적합", "개정 이력 Rev.%s 작성일자가 자리표시(%s)입니다." % (r["rev"], r["date"]), "개정 이력",
                                   "Rev.%s 작성일자를 실제 날짜로 적어 주시기 바랍니다." % r["rev"]))
        real = [r for r in rows if not (ph.search(r["date"]) or re.search(r"\.00(\.|$)", r["date"]))]
        hdate = next((m.group(0) for t in header for m in [re.search(r"\d{4}\s*\.\s*\d{1,2}\s*\.\s*\d{1,2}", t)] if m), "")
        hrev = next((m.group(1) for t in header for m in [re.search(r"Rev\.?\s*(\d{1,2})", t)] if m), "")
        if real and hdate and not ph.search(hdate):
            last = real[-1]
            if re.sub(r"\s", "", hdate) != last["date"]:
                out.append(finding(no, "권고", "머리글 날짜(%s)와 마지막 개정 이력 날짜(%s)가 다릅니다." % (re.sub(r"\s", "", hdate), last["date"]), "머리글·개정 이력",
                                   "머리글 날짜를 마지막 개정일과 맞춰 주시기 바랍니다."))
        if rows and hrev and hrev.isdigit() and rows[-1]["rev"].isdigit() and int(hrev) != int(rows[-1]["rev"]):
            out.append(finding(no, "권고", "머리글 Rev.%s 와 마지막 개정 이력 Rev.%s 가 다릅니다." % (hrev, rows[-1]["rev"]), "머리글·개정 이력",
                               "머리글 Rev 를 마지막 개정 이력과 맞춰 주시기 바랍니다."))
    if not out:
        out.append(finding(no, "적합", "머리글과 개정 이력에 자리표시가 없고 서로 맞습니다." if rev.get("found") else "개정 이력 점검 대상이 아닙니다(Template 에 없음).", "", ""))
    return out


def check_leftover(doc, cfg, crit):
    """Template 예시 문구·자리표시와 「추후 보완 예정」 같은 미완성 표기가 본문에 남았는지."""
    no = crit["no"]
    ph = _placeholder_re()
    tpl = cfg.get("template") or {}
    ex_lines = {norm(t) for t in tpl.get("example_lines", []) if len(norm(t)) >= 6}
    ex_terms = {norm(t) for t in tpl.get("example_terms", []) if norm(t)}
    out = []
    marks, examples, phs, unfinished = [], [], [], []
    from .profile import EXAMPLE_MARK_RE
    for b in doc["blocks"]:
        if b.get("rev") or b.get("is_title"):
            continue
        t = (b.get("text") or "").strip()
        if not t:
            continue
        if b.get("heading"):
            if ex_terms and norm(b["heading"]["title"]) in ex_terms:
                examples.append(b)
            continue
        if EXAMPLE_MARK_RE.match(t):
            marks.append(b)
        elif ex_lines and norm(t) in ex_lines:
            examples.append(b)
        if ph.search(t) and not find_std_numbers(t, cfg):
            phs.append(b)
        if UNFINISHED_RE.search(t):
            unfinished.append(b)
    if marks:
        out.append(finding(no, "부적합", "Template 의 「예시)」「(참고)」 표시가 %d곳 남아 있습니다." % len(marks), "; ".join(loc(b) for b in marks[:3]),
                           "예시 표시와 예시 문구를 지우고 이 표준의 내용으로 바꿔 주시기 바랍니다."))
    if examples:
        out.append(finding(no, "부적합", "Template 예시 문구가 그대로 %d곳 남아 있습니다." % len(examples), "; ".join(loc(b) for b in examples[:3]),
                           "Template 예시 문구를 이 표준의 내용으로 바꿔 주시기 바랍니다."))
    if phs:
        out.append(finding(no, "부적합", "본문에 자리표시(xx, 0000, ○○○ 등)가 %d곳 남아 있습니다." % len(phs), "; ".join(loc(b) for b in phs[:3]),
                           "자리표시를 실제 값으로 채워 주시기 바랍니다."))
    for b in unfinished:
        m = UNFINISHED_RE.search(b["text"])
        out.append(finding(no, "권고", "미완성 표기가 있습니다: 「%s」" % m.group(0)[:40], loc(b),
                           "「%s」 부분을 이번 표준에 반영하거나, 별도 표준으로 둘 경우 관련 표준 번호를 2항에 적어 주시기 바랍니다." % m.group(0)[:30]))
    if not out:
        out.append(finding(no, "적합", "Template 예시 문구·자리표시·미완성 표기가 없습니다.", "", ""))
    return out


def check_latin_variants(doc, cfg, crit, defined, reported):
    """같은 영문 낱말이 대소문자·하이픈만 다르게 섞여 쓰였는지(예: Set-up / Setup). 인용 줄과 2항은 뺍니다."""
    no = crit["no"]
    pref = cfg.get("latin_pref") or {}
    forms, where = {}, {}
    for b in doc["blocks"]:
        if b.get("rev") or b.get("is_title") or not b.get("text") or in_section(b.get("sec", ""), "2") or in_section(b.get("sec", ""), "3"):
            continue
        if find_std_numbers(b["text"], cfg):
            continue
        for tok in re.findall(r"(?<![A-Za-z])[A-Za-z]{2,}(?:-[A-Za-z]{2,})*(?![A-Za-z])", b["text"]):
            if len(tok) < 3:
                continue
            key = tok.lower().replace("-", "")
            forms.setdefault(key, Counter())[tok] += 1
            where.setdefault((key, tok), []).append(b)
    out = []
    for key, c in forms.items():
        if len(c) < 2:
            continue
        if key in defined:
            best = defined[key]
        elif pref.get(key) in c:
            best = pref[key]
        else:
            best = c.most_common(1)[0][0]
        for form, n in c.items():
            if form == best or (key, form) in reported:
                continue
            # 용어정의에 없는 일반 낱말은 대소문자만 다른 경우(문장 첫머리·강조)는 넘어가고 하이픈 차이만 봅니다.
            if key not in defined and "-" not in form + best:
                continue
            bl = where[(key, form)]
            out.append(finding(no, "부적합" if key in defined else "권고",
                               "영문 표기 %s %s와 섞여 %d곳에 쓰였습니다." % (j("「%s」" % form, "이가"), "「%s」" % best, n),
                               "; ".join(loc(b) for b in bl[:3]) + (" 외 %d곳" % (len(bl) - 3) if len(bl) > 3 else ""),
                               "「%s」 표기를 %s 통일해 주시기 바랍니다." % (form, j("「%s」" % best, "으로"))))
    return out


def check_not_applicable(doc, cfg, crit):
    return [finding(crit["no"], "해당 없음", crit.get("na", ""), "", "")]


JUDGE_ONLY_TEXT = "이 기준은 의미 판단이 필요해 판단 프롬프트로 확인합니다."


def check_judge_only(doc, cfg, crit):
    return [finding(crit["no"], "판단 필요", JUDGE_ONLY_TEXT, str(crit.get("target") or "문서 전체"), "")]


CHECKS = {
    "PURPOSE": check_purpose,
    "CITATION_LIST": check_citation_list,
    "TERM_DEF": check_term_def,
    "TERM_CONSISTENCY": check_term_consistency,
    "SECTION_SEQUENCE": check_section_sequence,
    "FLOWCHART": check_flowchart,
    "EQUATIONS": check_equations,
    "CAPTIONS": check_captions,
    "CITATION_REF": check_citation_ref,
    "SECTION_JUDGE": check_section_judge,
    "QUANT": check_quant,
    "REVISION": check_revision,
    "LEFTOVER": check_leftover,
}
KNOWN_CODES = set(CHECKS) | {"TEMPLATE_ITEMS", "FIX_TEMPLATE", "EXISTS"}


def is_judgment(crit):
    return "판단" in str(crit.get("method", ""))


def run_checks(doc, cfg, std_type):
    """사용(Y) 기준 전체를 점검해 결과 목록을 돌려줍니다."""
    results = []
    for crit in cfg.get("criteria", []):
        if not crit.get("active"):
            continue
        code = (crit.get("code") or "").strip()
        if code == "FIX_TEMPLATE":
            continue
        if crit.get("na"):
            res = check_not_applicable(doc, cfg, crit)
        elif code == "TEMPLATE_ITEMS":
            res = check_template_items(doc, cfg, std_type, crit["no"])
        elif code == "EXISTS":
            res = check_exists(doc, cfg, crit, std_type)
        elif code in CHECKS:
            res = CHECKS[code](doc, cfg, crit)
        elif is_judgment(crit) or not code:
            res = check_judge_only(doc, cfg, crit)
        else:
            res = [finding(crit["no"], "판단 필요", "알 수 없는 점검코드 「%s」입니다. 기준표를 확인해 주세요." % code, "", "")]
        results.extend(res)
    return results


def fix_template_default(cfg):
    for c in cfg.get("criteria", []):
        if c.get("code") == "FIX_TEMPLATE" and c.get("fix"):
            return c["fix"]
    return "Template {항목} 기준에 따라 {내용}"


def assemble(results, cfg):
    """점검 결과에 기준표의 항목명·평가 기준·수정 요청 문구를 붙여 회의록 행을 만듭니다."""
    crit_by_no = {c["no"]: c for c in cfg.get("criteria", [])}
    default_fix = fix_template_default(cfg)
    style = cfg.get("minutes_style")
    rows = []
    for i, r in enumerate(results, 1):
        c = crit_by_no.get(r["no"], {})
        fix = ""
        item = c.get("item", "")
        if r["result"] in ("부적합", "권고") and r.get("hint"):
            if style:
                # 학습한 회의록 문체: 「Template … 기준에 따라」 머리를 회의록에서 쓸 때만 붙이고 끝맺음을 맞춥니다.
                from .profile import restyle
                body = restyle(r["hint"], style, r.get("location", ""))
                fix = ("Template %s 기준에 따라 " % item + body) if style.get("template_prefix") and not style.get("lead") else body
            else:
                tpl = c.get("fix") or default_fix
                try:
                    fix = tpl.replace("{항목}", item).replace("{내용}", r["hint"])
                except Exception:
                    fix = r["hint"]
        if style:
            from .profile import NO_TO_CODE
            label = (style.get("labels") or {}).get(NO_TO_CODE.get(r["no"], ""))
            if label:
                item = label
        rows.append({
            "seq": i, "no": r["no"], "item": item, "criterion": c.get("text", ""),
            "result": r["result"], "detail": r["detail"], "location": r.get("location", ""),
            "fix": fix, "verdict": "", "comment": "",
        })
    return rows


def summarize(rows):
    out = {k: 0 for k in ("적합", "부적합", "권고", "판단 필요")}
    if any(r["result"] == "해당 없음" for r in rows):
        out["해당 없음"] = 0
    for r in rows:
        out[r["result"]] = out.get(r["result"], 0) + 1
    return out
