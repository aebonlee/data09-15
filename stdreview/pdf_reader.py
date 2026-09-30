"""PDF 읽기 — pypdf(순수 Python, 오프라인)로 글자를 뽑아 docx 와 같은 블록 구조로 바꿉니다.

PDF 는 Word 처럼 「제목 스타일」「표」「수식」 정보가 없습니다. 그래서
- 절 제목은 「4.1 …」처럼 번호로 시작하는 짧은 줄로 찾고,
- 페이지마다 되풀이되는 머리글·바닥글(문서 종류, 표준번호, 날짜, 저작권 안내, 쪽 번호)은 본문에서 빼서 따로 모으고,
- 그림은 그림 캡션이 있는 자리에 있다고 보고, 표 안의 칸 구분과 수식 개수는 알 수 없는 것으로 둡니다.
표·수식까지 정확히 보려면 원본 .docx 를 넣는 것이 좋습니다.

pypdf 는 vendor/wheels 에 들어 있어 설치 없이도 불러옵니다(README 「폐쇄망 설치」).
"""

import re
from collections import Counter

from .docx_reader import DocxError, NUM_TITLE_RE, REV_RE, _looks_like_title, build

BULLET_RE = re.compile(r"^\s*(?:[•⚫●▪■□◦○◆◇▶▷※\-–·*]|\d{1,2}\)|\(\d{1,2}\)|[가-하]\.|[①-⑳])\s*")
CAPTION_START_RE = re.compile(r"[\[<〈(]?(?:표|그림|Table|Tab|Figure|Fig)\s*\.?\s*\d+", re.I)
PAGE_NO_RE = re.compile(r"^\s*(?:-\s*)?\d{1,3}(?:\s*-)?(?:\s*/\s*\d{1,3})?\s*$")
SENT_END_RE = re.compile(r"(다|음|함|것|요)\s*\.?\s*$|[.!?:：]\s*$")


class PdfError(DocxError):
    pass


def _pypdf():
    from .vendor import import_or_vendor
    try:
        return import_or_vendor("pypdf", ["typing_extensions", "pypdf"])
    except ImportError:
        raise PdfError("PDF 를 읽으려면 pypdf 가 필요합니다. 저장소의 vendor/wheels 폴더가 있는지 확인하거나, "
                       "README 「폐쇄망 설치」대로 python -m pip install --no-index --find-links vendor/wheels pypdf 로 설치해 주세요.")


def _key(line):
    """머리글 판정용 비교 키 — 공백을 없애고 숫자는 # 로 바꿉니다(쪽 번호·날짜가 달라도 같은 줄로 봄)."""
    return re.sub(r"\d+", "#", re.sub(r"\s+", "", line))


def page_lines(path):
    """[(쪽 번호, [줄…])], 그림 수 — pypdf 로 쪽별 글자를 뽑습니다."""
    pypdf = _pypdf()
    try:
        reader = pypdf.PdfReader(str(path))
    except FileNotFoundError:
        raise PdfError("파일을 찾을 수 없습니다: %s" % path)
    except Exception as e:  # pypdf 의 여러 오류를 한 가지 안내로 모읍니다.
        raise PdfError("PDF 를 열 수 없습니다: %s (%s)" % (path, e))
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception:
            raise PdfError("암호가 걸린 PDF 입니다. 암호를 푼 PDF 나 원본 .docx 를 넣어 주세요: %s" % path)
    pages = []
    images = []
    for n, page in enumerate(reader.pages, 1):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            raise PdfError("%d쪽 글자를 읽지 못했습니다: %s (%s)" % (n, path, e))
        pages.append((n, [ln.rstrip() for ln in text.splitlines()]))
        try:
            images.append(len(page.images))
        except Exception:
            images.append(0)
    return pages, images


def repeated_keys(pages):
    """여러 쪽에 되풀이되는 줄(머리글·바닥글)의 키 집합."""
    if len(pages) < 2:
        return set()
    cnt = Counter()
    for _, lines in pages:
        seen = {_key(ln) for ln in lines if ln.strip()}
        cnt.update(seen)
    need = max(2, int(len(pages) * 0.6 + 0.5))
    return {k for k, c in cnt.items() if c >= need and k}


def _split_captions(line):
    """한 줄에 캡션이 둘 이상 붙어 나온 경우(「그림 8. …   그림 9. …」) 나눕니다."""
    starts = [m.start() for m in CAPTION_START_RE.finditer(line)]
    if len(starts) < 2:
        return [line]
    cuts = [starts[0]]
    for s in starts[1:]:
        before = line[:s]
        if re.search(r"(\s{2,}|[>\]〉)]\s*)$", before):
            cuts.append(s)
    if len(cuts) < 2:
        return [line]
    parts = [line[:cuts[0]]] + [line[a:b] for a, b in zip(cuts, cuts[1:] + [len(line)])]
    return [p for p in parts if p.strip()]


PDF_TITLE_MAX = 70


def _top_ok(t):
    """PDF 에서 한 단계 번호(「5 …」)는 표 칸 번호일 때가 많아, 상위 절은 「5. …」처럼 점이 있을 때만 제목으로 봅니다."""
    m = re.match(r"^\s*(\d+(?:\.\d+)*)(\.?)\s", t)
    if not m:
        return False
    return "." in m.group(1) or m.group(2) == "."


def _is_caption(t):
    m = CAPTION_START_RE.match(t)
    return bool(m) and len(t) <= 80 and not re.search(r"(다|요|음|함)\.?\s*$", t)


def lines_to_blocks(pages, repeated=None):
    """쪽별 줄 → (블록 목록, 머리글 줄 목록). 파일 없이 시험할 수 있게 분리했습니다."""
    repeated = repeated if repeated is not None else repeated_keys(pages)
    header = []
    blocks = []
    cur = None

    def flush():
        nonlocal cur
        if cur is not None and cur["text"].strip():
            cur["text"] = re.sub(r"\s+", " ", cur["text"]).strip()
            blocks.append(cur)
        cur = None

    def new_block(text, **kw):
        nonlocal cur
        flush()
        cur = {"kind": "p", "text": text, "style": kw.get("style", ""), "style_level": None,
               "figure": kw.get("figure", False), "omath": 0, "page": kw.get("page"),
               "title_ok": kw.get("title_ok", False)}
        if kw.get("closed"):
            flush()

    for pno, lines in pages:
        for raw in lines:
            if not raw.strip():
                flush()
                continue
            if _key(raw) in repeated:
                t = re.sub(r"\s+", " ", raw).strip()
                if pno == pages[0][0] and t not in header and not PAGE_NO_RE.match(t):
                    header.append(t)
                continue
            if PAGE_NO_RE.match(raw):
                continue
            for piece in _split_captions(raw):
                t = re.sub(r"\s+", " ", piece).strip()
                # PDF 글자 뽑기에서 번호 가운데 끼는 공백(「3691 -1」)을 붙입니다.
                t = re.sub(r"(\d) -(\d)", r"\1-\2", t)
                if not t:
                    continue
                if _is_caption(t):
                    is_fig = bool(re.match(r"[\[<〈(]?\s*(?:그림|Figure|Fig)", t, re.I))
                    # 그림 캡션 자리에 그림이 있다고 봅니다(PDF 에서는 그림 위치를 알 수 없음).
                    new_block(t, style="caption", figure=is_fig, page=pno, closed=True)
                elif len(t) <= 40 and REV_RE.search(t):
                    new_block(t, page=pno, closed=True)
                elif NUM_TITLE_RE.match(t) and _is_caption(NUM_TITLE_RE.match(t).group(2)):
                    # 「1.1 [표 1] …」 처럼 번호 뒤에 캡션이 오는 줄은 캡션으로 봅니다.
                    c = NUM_TITLE_RE.match(t).group(2)
                    new_block(c, style="caption", figure=bool(re.match(r"[\[<〈(]?\s*(?:그림|Figure|Fig)", c, re.I)), page=pno, closed=True)
                elif NUM_TITLE_RE.match(t) and _looks_like_title(t, PDF_TITLE_MAX) and _top_ok(t):
                    new_block(t, page=pno, closed=True, title_ok=True)
                elif BULLET_RE.match(t) and not NUM_TITLE_RE.match(t):
                    new_block(t, page=pno)
                elif re.match(r"^\s*(?:Rev\.?\s*)?\d{1,2}\s+\d{4}\s*[.\-/]", t):
                    new_block(t, page=pno, closed=True)          # 개정 이력 한 줄
                elif cur is None:
                    new_block(t, page=pno)
                else:
                    cur["text"] += " " + t
                    if SENT_END_RE.search(t) and not BULLET_RE.match(cur["text"]):
                        flush()
        flush()
    return blocks, header


def header_title(header):
    """머리글 줄에서 표준 제목을 고릅니다 — 「제목 :」 접두를 떼고, 문서 종류·번호·날짜·저작권 줄은 뺍니다."""
    cands = []
    for t in header:
        s = re.sub(r"^\s*(제목|Title)\s*[:：]\s*", "", t).strip()
        if not s or re.search(r"copyright|all rights|ⓒ|©|문서는|승인|permission|confidential|prohibited", s, re.I):
            continue
        if re.search(r"\d{4}\s*\.\s*\d{1,2}\s*\.\s*\d{1,2}|Rev\.", s) or re.match(r"^\(.*\)$", s):
            continue
        if re.fullmatch(r"[A-Za-z]{2,}[-A-Za-z0-9]*", s.replace(" ", "")) and "-" in s:
            continue
        cands.append(s)
    if not cands:
        return ""
    # 문서 종류(첫 줄, 짧음)보다 긴 줄을 제목으로 봅니다.
    return max(cands, key=len)


def read_pdf(path, number_titles=True):
    pages, images = page_lines(path)
    if not any(ln.strip() for _, lines in pages for ln in lines):
        raise PdfError("PDF 에서 글자를 찾지 못했습니다(스캔 이미지 PDF 일 수 있습니다). 원본 .docx 나 글자가 있는 PDF 를 넣어 주세요: %s" % path)
    blocks, header = lines_to_blocks(pages)
    doc = build(blocks, core_title=header_title(header), number_titles=number_titles, source="pdf", header=header,
                title_max=PDF_TITLE_MAX)
    doc["pages"] = len(pages)
    doc["pdf_images"] = sum(images)
    return doc
