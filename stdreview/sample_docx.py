"""예시(가상) 표준 .docx 만들기 — 표준 라이브러리만 씁니다.

samples/ 의 예시 파일과 테스트용 문서를 만듭니다. 내용은 모두 가상이며 「(예시)」를 붙였습니다.
인용 표준번호(EXS-…)도 실제 문서가 아닌 예시 번호입니다.
"""

import struct
import zipfile
import zlib
from xml.sax.saxutils import escape

NS_DECL = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
           'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')


def _png(w=160, h=80):
    raw = b"".join(b"\x00" + bytes([200, 210, 225]) * w for _ in range(h))

    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _run(text):
    return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % escape(text)


def _p(text, style=None):
    ppr = '<w:pPr><w:pStyle w:val="%s"/></w:pPr>' % style if style else ""
    return "<w:p>%s%s</w:p>" % (ppr, _run(text) if text else "")


def _tbl(rows):
    out = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
           '<w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>']
    for r in rows:
        out.append("<w:tr>" + "".join("<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>%s</w:tc>" % _p(c) for c in r) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def _fig(n):
    cx, cy = 3048000, 1524000
    return ('<w:p><w:r><w:drawing><wp:inline><wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="그림 %d"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="%d" name="img%d.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>') % (cx, cy, n, n, n, n, cx, cy)


def _math(text):
    return '<w:p><m:oMathPara><m:oMath><m:r><m:t>%s</m:t></m:r></m:oMath></m:oMathPara></w:p>' % escape(text)


STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/><w:sz w:val="20"/><w:lang w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
</w:styles>"""

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/figure.png"/>
</Relationships>"""


def build_docx(path, title, blocks, header=None):
    """blocks: ("h1", 글) ("h2", 글) ("p", 글) ("cap", 글) ("tbl", [[..]]) ("fig",) ("math", 글).
    header: 머리글 줄 목록(있으면 word/header1.xml 로 넣습니다)."""
    body = [_p(title, "Title")] if title else []
    fig_n = 0
    for b in blocks:
        k = b[0]
        if k == "h1":
            body.append(_p(b[1], "Heading1"))
        elif k == "h2":
            body.append(_p(b[1], "Heading2"))
        elif k == "p":
            body.append(_p(b[1]))
        elif k == "cap":
            body.append(_p(b[1], "Caption"))
        elif k == "tbl":
            body.append(_tbl(b[1]))
        elif k == "fig":
            fig_n += 1
            body.append(_fig(fig_n))
        elif k == "math":
            body.append(_math(b[1]))
    href = '<w:headerReference w:type="default" r:id="rIdHdr"/>' if header else ""
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr>' + href + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
           '</w:body></w:document>') % (NS_DECL, "".join(body))
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>%s</dc:title></cp:coreProperties>') % escape(title)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        ct = CT if not header else CT.replace("</Types>", '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        rels = DOC_RELS
        if header:
            rels = rels.replace("</Relationships>", '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
            z.writestr("word/header1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr %s>%s</w:hdr>' % (NS_DECL, "".join(_p(t) for t in header)))
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/media/figure.png", _png())
        z.writestr("docProps/core.xml", core)


TITLE = "(예시) 전장 제어기(ECU) 방열 설계표준"


def good_blocks():
    """모든 기준을 지킨 가상 표준 — 도구가 부적합을 내지 않아야 합니다."""
    return [
        ("h1", "1. 목적 및 정의"),
        ("p", "이 표준은 (예시) 건설기계 전장 제어기(ECU)의 방열 설계 절차와 판정 기준을 정한다."),
        ("p", "적용 대상은 ECU 하우징과 방열판이며, 방열 설계란 접합부 온도를 허용값 아래로 유지하도록 방열판을 정하는 일을 말한다."),
        ("h1", "2. 관련 표준 및 인용"),
        ("cap", "표 1. 관련 표준"),
        ("tbl", [["문서번호", "제목"], ["EXS-1001", "(예시) 전장품 환경 시험 기준"], ["EXS-2001", "(예시) 방열판 재료 선정 기준"]]),
        ("h1", "3. 용어정의"),
        ("cap", "표 2. 용어정의"),
        ("tbl", [["용어", "설명"],
                 ["ECU (Electronic Control Unit)", "이 표준에서 건설기계 엔진과 유압을 제어하는 전장 제어기 본체를 말한다."],
                 ["방열판", "ECU 하우징 뒷면에 붙여 열을 공기로 내보내는 알루미늄 판을 말한다."],
                 ["접합부 온도", "ECU 안 반도체 소자의 내부 온도로, 제조사 자료의 최대값과 비교하는 값을 말한다."]]),
        ("h1", "4. 설계절차"),
        ("h2", "4.1 설계 흐름"),
        ("p", "방열 설계는 그림 1의 순서로 진행한다."),
        ("fig",),
        ("cap", "그림 1. 방열 설계 흐름도(FlowChart)"),
        ("h2", "4.2 접합부 온도 계산"),
        ("p", "ECU 접합부 온도는 아래 식으로 계산하고, 입력값은 표 3에 정리한다."),
        ("math", "Tj = Ta + P × Rth"),
        ("p", "주변 온도 Ta 는 85 ℃ 로 둔다(EXS-1001 근거)."),
        ("cap", "표 3. 설계 입력값"),
        ("tbl", [["기호", "뜻"], ["Ta", "주변 온도"], ["P", "소비 전력"], ["Rth", "열 저항"]]),
        ("h2", "4.3 방열판 선정"),
        ("p", "방열판 재료는 EXS-2001 에 따라 선정한다."),
        ("h1", "5. 설계 검증"),
        ("p", "접합부 온도 계산값이 허용값보다 10 ℃ 이상 낮은지 확인한다(EXS-1001 근거)."),
        ("p", "표 3의 입력값으로 다시 계산해 같은 결과가 나오는지 확인한다."),
        ("h1", "6. (예시) 기록 관리"),
        ("p", "설계 계산서는 심의 자료와 함께 보관한다."),
        ("h1", "7. 안전 관련 사항"),
        ("p", "방열판 모서리는 작업자 손 베임을 막도록 모따기 한다."),
        ("h1", "8. 기타(Appendix)"),
        ("p", "부록 A. (예시) ECU 접합부 온도 계산 예"),
    ]


def bad_blocks():
    """일부러 기준을 어긴 가상 표준 — 도구가 각 결함을 잡아야 합니다."""
    return [
        ("h1", "1. 목적 및 정의"),
        ("p", "이 표준은 품질 향상을 위해 만든다."),
        ("h1", "2. 관련 표준 및 인용"),
        ("p", "EXS-1001 (예시) 전장품 환경 시험 기준"),
        ("p", "(예시) 방열판 재료 선정 기준"),                      # 번호 없음 → 2-1
        ("p", "EXS-3001"),                                            # 제목 없음 → 2-1
        ("h1", "3. 용어정의"),
        ("cap", "표 1. 용어정의"),
        ("tbl", [["용어", "설명"],
                 ["ECU", "전장 제어기"],
                 ["방열판", ""],                                      # 설명 비어 있음 → 3-1
                 ["냉각팬", "ECU 옆에 붙이는 팬을 말한다."]]),         # 본문 미사용 → 3-2 권고
        ("h1", "4. 설계절차"),
        ("h2", "4.1 설계 구성"),
        ("p", "이씨유 하우징에 방열판을 붙인다."),                    # 이형 표기 → 3-2
        ("fig",),
        ("cap", "그림 1. 설계 구성도"),                               # 흐름도 낱말 없음 → 4-2 권고
        ("h2", "4.3 방열판 선정"),                                    # 4.2 빠짐 → 4-1
        ("p", "Ecu 방열판은 EXS-9001 에 따라 선정하며 두께는 3 mm 로 한다."),  # 대소문자 → 3-2, 목록에 없는 인용 → 4-6
        ("tbl", [["재료", "두께"], ["알루미늄", "3 mm"]]),             # 캡션 없음 → 4-4
        ("p", "상세 치수는 표 5를 따른다."),                          # 없는 표 참조 → 4-4
        ("fig",),
        ("cap", "그림 3. 방열판 치수"),                               # 그림 2 빠짐 → 4-4
        ("h1", "5. 설계 검증"),
        ("p", "방열 성능이 양호할 것."),                              # 정성 표현 → 5-2
        ("h1", "7. 안전 관련 사항"),                                  # 6항 없음 → 공통-1, 6-1
        ("p", "작업 시 주의한다."),
    ]                                                                 # 8항 없음 → 권고


def make_samples(folder):
    from pathlib import Path
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    a = folder / "예시데이터_설계표준_정상.docx"
    b = folder / "예시데이터_설계표준_오류포함.docx"
    build_docx(a, TITLE, good_blocks())
    build_docx(b, TITLE, bad_blocks())
    return [a, b]
