"""학습 기능 테스트 — python3 -m unittest discover -s test

가상 자료(stdreview/sample_learning.py)로 Template·승인 표준·기존 회의록을 학습하고,
결함을 심은 가상 초안을 점검해 기대한 지적이 나오는지 봅니다. 실제 사내 자료는 쓰지 않습니다.
"""

import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from stdreview import defaults, profile as prof  # noqa: E402
from stdreview.checks import assemble, run_checks  # noqa: E402
from stdreview.docx_reader import read_document  # noqa: E402
from stdreview.pdf_reader import lines_to_blocks  # noqa: E402
from stdreview.docx_reader import build  # noqa: E402

try:
    import openpyxl  # noqa: F401
    HAS_XL = True
except ImportError:
    HAS_XL = False


def tiny_pdf(path, pages):
    """글자만 있는 아주 작은 PDF(Helvetica, 영문)를 만듭니다 — PDF 읽기 경로 시험용."""
    objs = []

    def add(body):
        objs.append(body)
        return len(objs)

    font = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    pages_id = len(objs) + 1 + 2 * len(pages)       # 페이지·내용 객체 뒤에 Pages 객체
    kids = []
    for lines in pages:
        ops = ["BT /F1 11 Tf 72 760 Td 14 TL"]
        for ln in lines:
            ops.append("(%s) Tj T*" % ln.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)"))
        ops.append("ET")
        stream = "\n".join(ops).encode("latin-1")
        cid = add(b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream")
        pid = add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 %d 0 R >> >> /Contents %d 0 R >>"
                  % (pages_id, font, cid))
        kids.append(pid)
    assert add(b"<< /Type /Pages /Kids [%s] /Count %d >>" % (b" ".join(b"%d 0 R" % k for k in kids), len(kids))) == pages_id
    cat = add(b"<< /Type /Catalog /Pages %d 0 R >>" % pages_id)
    out = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, body in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + body + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, cat, x)
    Path(path).write_bytes(bytes(out))


class PdfTest(unittest.TestCase):
    def test_lines_to_blocks_strips_header_and_splits(self):
        head = ["SAMPLE GUIDE", "QSP-B101", "2026.01.02 (Rev.00)"]
        pages = [
            (1, head + ["1. Purpose", "This guide sets the rules", "for brackets.", "2. References", "2.1 Cited", "• QSP-A120 Bracket strength", "1"]),
            (2, head + ["3. Terms", "Fig. 1 Setup A    Fig. 2 Setup B", "5 Row number in a table", "Revision Log", "00 2026.01.02 Writer Initial", "2"]),
        ]
        blocks, header = lines_to_blocks(pages)
        self.assertEqual(header, head)
        doc = build(blocks, source="pdf", header=header)
        self.assertEqual([h["num"] for h in doc["headings"]], ["1", "2", "2.1", "3"])   # 「5 Row…」 는 제목이 아님
        caps = [b["text"] for b in doc["blocks"] if b.get("style") == "caption"]
        self.assertEqual(caps, ["Fig. 1 Setup A", "Fig. 2 Setup B"])
        self.assertTrue(doc["revision"]["found"])
        self.assertEqual(doc["revision"]["rows"][0]["date"], "2026.01.02")
        joined = [b["text"] for b in doc["blocks"] if b.get("sec") == "1" and not b.get("heading")]
        self.assertEqual(joined, ["This guide sets the rules for brackets."])

    def test_read_real_pdf(self):
        p = Path(tempfile.mkdtemp()) / "t.pdf"
        tiny_pdf(p, [["SAMPLE GUIDE", "1. Purpose", "Rules for brackets.", "2. References", "2.1 Cited", "QSP-A120 Bracket strength"],
                     ["SAMPLE GUIDE", "3. Terms", "Revision Log", "00 2026.01.02 Writer Initial"]])
        doc = read_document(p)
        self.assertEqual(doc["source"], "pdf")
        self.assertEqual(doc["header"], ["SAMPLE GUIDE"])
        self.assertEqual([h["num"] for h in doc["headings"]], ["1", "2", "2.1", "3"])
        self.assertTrue(doc["revision"]["found"])


@unittest.skipUnless(HAS_XL, "openpyxl 없음")
class LearnTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from stdreview.sample_learning import make_learning_samples
        cls.tmp = Path(tempfile.mkdtemp())
        cls.paths = make_learning_samples(cls.tmp)
        cls.profile = prof.empty_profile()
        cls.profile["templates"]["설계표준"] = prof.learn_template(read_document(cls.paths["template"]), "설계표준", "t")
        cls.profile["standards"]["s"] = prof.learn_standard(read_document(cls.paths["standard"]), "s")
        items, heads, roles = prof.read_minutes(cls.paths["minutes"])
        cls.profile["minutes"]["m"] = prof.learn_minutes(items, heads, roles, "m")

    def review(self, use_profile=True):
        cfg = defaults.default_config()
        if use_profile:
            cfg, _ = prof.apply_profile(cfg, self.profile, "설계표준")
        doc = read_document(self.paths["draft"])
        return assemble(run_checks(doc, cfg, "설계표준"), cfg)

    def test_template_learning(self):
        t = self.profile["templates"]["설계표준"]
        self.assertEqual([i["num"] for i in t["items"]], ["1", "2", "3", "4", "5", "6", "7"])
        self.assertEqual([s["num"] for i in t["items"] for s in i["subs"]], ["2.1", "2.2", "4.1", "4.2"])  # 3.1 예시 용어는 빠짐
        self.assertEqual(t["example_terms"], ["샘플 용어"])
        self.assertEqual(len(t["header_placeholders"]), 2)
        self.assertTrue(t["has_revision"])
        self.assertEqual(t["flow_sections"], ["4.2"])
        self.assertFalse(t["items"][-1]["required"])            # 부록은 선택 항목

    def test_standard_learning(self):
        s = self.profile["standards"]["s"]
        self.assertIn(prof.shape_of("QSP-A120"), s["cite_shapes"])
        self.assertEqual(s["caption_words"]["표"], {"표": 1})
        self.assertEqual(s["latin_forms"]["setup"], {"Set-up": 2})
        self.assertEqual([t["term"] for t in s["terms"]], ["브래킷", "방진 고무"])

    def test_minutes_learning(self):
        m = self.profile["minutes"]["m"]
        self.assertEqual(m["count"], 6)
        self.assertEqual(m["endings"], {"필요": 6})
        self.assertEqual(m["roles"]["comment"], "검토의견")
        self.assertEqual(m["roles"]["category"], "구분")
        self.assertEqual(m["codes"]["TEMPLATE"], 2)

    def test_review_with_profile(self):
        rows = self.review()
        got = Counter((r["no"], r["result"]) for r in rows if r["result"] in ("부적합", "권고"))
        expected = Counter({
            ("공통-1", "부적합"): 1,   # 6항 안전 사항 없음
            ("공통-1", "권고"): 1,     # 4.2 하위 항목명이 Template 과 다름
            ("공통-3", "부적합"): 2,   # 머리글 번호 자리표시, 개정 이력 날짜 자리표시
            ("공통-4", "부적합"): 2,   # 「예시)」 표시, Template 예시 문구 그대로
            ("공통-4", "권고"): 1,     # 추후 보완 예정
            ("3-2", "권고"): 1,        # Setup ↔ 학습한 Set-up
            ("4-3", "권고"): 1,        # 수식 없음
            ("4-4", "권고"): 1,        # 캡션 형식 섞임
            ("4-6", "부적합"): 1,      # QSP-T999 가 2항에 없음(학습한 번호 모양으로 찾음)
            ("6-1", "부적합"): 1,      # 6항 없음
            ("7-1", "부적합"): 1,      # 안전 기준이 학습 Template 의 6항으로 옮겨져 6항 없음
        })
        self.assertEqual(got, expected)
        fix = next(r["fix"] for r in rows if r["no"] == "3-2" and r["result"] == "권고")
        self.assertEqual(fix, "「Setup」 표기를 「Set-up」로 통일 필요")          # 학습한 회의록 끝맺음
        self.assertEqual(next(r["item"] for r in rows if r["no"] == "3-2"), "용어")  # 학습한 회의록 구분 이름

    def test_learning_changes_result(self):
        before = Counter((r["no"], r["result"]) for r in self.review(use_profile=False))
        after = Counter((r["no"], r["result"]) for r in self.review())
        self.assertEqual(before[("2-1", "부적합")], 1)    # 학습 전: 사내 번호 모양을 몰라 번호 없음으로 봄
        self.assertEqual(after[("2-1", "부적합")], 0)
        self.assertEqual(before[("4-6", "부적합")], 0)    # 학습 전: 없는 인용 번호를 못 찾음
        self.assertEqual(after[("4-6", "부적합")], 1)

    def test_columns_from_minutes(self):
        cfg, _ = prof.apply_profile(defaults.default_config(), self.profile, "설계표준")
        heads = {c["key"]: c["header"] for c in cfg["columns"]}
        # 회의록의 「No」는 이미 기준 번호 열 이름이라 겹치지 않게 순번은 그대로 둡니다.
        self.assertEqual((heads["seq"], heads["item"], heads["location"], heads["fix"]), ("순번", "구분", "해당 절", "검토의견"))
        self.assertEqual(len(set(heads.values())), len(heads))

    def test_retarget(self):
        cfg, notes = prof.apply_profile(defaults.default_config(), self.profile, "설계표준")
        by = {c["no"]: c for c in cfg["criteria"]}
        self.assertEqual(by["7-1"]["target"], "6")
        self.assertEqual(by["8-1"]["target"], "7")
        self.assertEqual(by["1-1"]["target"], "1")

    def test_compare(self):
        import review_tool
        rows = [r for r in self.review() if r["result"] in ("부적합", "권고")]
        past, _, _ = prof.read_minutes(self.paths["minutes"])
        r = review_tool.compare_counts(rows, past)
        self.assertEqual((r["past"], r["past_hit"]), (6, 6))
        self.assertEqual(r["tool"], len(rows))

    def test_profile_roundtrip(self):
        p = self.tmp / "data" / "profile.json"
        prof.save_profile(self.profile, p)
        back = prof.load_profile(p)
        self.assertEqual(prof.summary(back), prof.summary(self.profile))


class RestyleTest(unittest.TestCase):
    def test_endings(self):
        h = "6항을 추가해 주시기 바랍니다."
        self.assertEqual(prof.restyle(h, {"ending": "필요"}), "6항 추가 필요")
        self.assertEqual(prof.restyle(h, {"ending": "요망"}), "6항 추가 요망")
        self.assertEqual(prof.restyle(h, {"ending": "할 것"}), "6항 추가할 것")
        self.assertEqual(prof.restyle(h, {"ending": "바랍니다"}), h)
        self.assertEqual(prof.restyle("번호를 고쳐 주시기 바랍니다.", {"ending": "필요"}), "번호를 고쳐야 함")
        self.assertEqual(prof.restyle(h, {"ending": "필요", "location": "(n)", "lead": "-"}, "4.2절 「…」"), "- (4.2) 6항 추가 필요")

    def test_classify(self):
        self.assertEqual(prof.classify("표 캡션 형식 통일 필요", "그림/표"), "CAPTIONS")
        self.assertEqual(prof.classify("머리글 번호 수정 필요", ""), "REVISION")
        self.assertEqual(prof.ending_family("표기 통일 요망"), "요망")


if __name__ == "__main__":
    unittest.main()
