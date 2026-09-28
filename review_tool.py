#!/usr/bin/env python3
"""기술표준 검토 도구 (data09-15, 1단계) — 명령줄 실행 파일.

  python review_tool.py demo                                  예시 데이터로 전체 흐름 시연
  python review_tool.py review 표준.docx --type 설계표준      표준 점검 → 회의록 초안·프롬프트·보고서
  python review_tool.py inspect 표준.docx                      도구가 읽은 절·표·그림 구조 확인
  python review_tool.py criteria export 기준표.xlsx            현재 평가 기준을 엑셀로 꺼내기
  python review_tool.py criteria import 기준표.xlsx --reason "사유"   고친 기준표 다시 넣기(변경 이력)
  python review_tool.py criteria history                       기준표 변경 이력
  python review_tool.py template import Template.docx --type 설계표준 --reason "사유"
  python review_tool.py feedback 회의록초안.xlsx               검토자 판정(채택/제외·보완 의견) 되받기
  python review_tool.py improve --html 보완후보.html           보완 후보 기준 보고
  python review_tool.py samples                                예시 표준 .docx 다시 만들기

인터넷·AI 호출 없음. 필요한 것: Python 3.8 이상 + openpyxl(엑셀 입출력).
"""

import argparse
import datetime
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from stdreview import defaults  # noqa: E402
from stdreview.checks import assemble, run_checks, setting, setting_num, summarize  # noqa: E402
from stdreview.config import ConfigError, diff, items_from_template, read_xlsx, validate, write_xlsx  # noqa: E402
from stdreview.docx_reader import DocxError, read_docx  # noqa: E402
from stdreview.improve import candidates  # noqa: E402
from stdreview.minutes import parse_map, read_feedback, write_minutes, write_minutes_csv  # noqa: E402
from stdreview.prompts import build_all, safe_name  # noqa: E402
from stdreview.report import improve_html, review_html  # noqa: E402
from stdreview.sample_docx import make_samples  # noqa: E402
from stdreview.store import Store  # noqa: E402

DEFAULT_DB = HERE / "data" / "stdreview.db"
SAMPLES = HERE / "samples"


def say(*a):
    print(*a)


def is_sample_file(p):
    return Path(p).name.startswith("예시데이터")


def confirm(question, assume_yes):
    if assume_yes:
        return True
    try:
        ans = input(question + " [y/N] ").strip().lower()
    except EOFError:
        return False
    return ans in ("y", "yes", "예", "ㅇ")


def number_titles(cfg):
    return setting(cfg, "번호_제목_인식", "예").strip() in ("예", "Y", "y", "yes", "1")


# ---------- review ----------

def do_review(store, path, std_type, out_dir=None, html=True, quiet=False):
    version, cfg = store.current()
    doc = read_docx(path, number_titles=number_titles(cfg))
    rows = assemble(run_checks(doc, cfg, std_type), cfg)
    sample = is_sample_file(path)
    rid = store.add_review(Path(path).name, std_type, doc.get("title", ""), version, rows, is_sample=sample)
    examples = store.adopted_examples(int(setting_num(cfg, "과거사례_최대건수", 5)), include_samples=sample)
    prompts = build_all(cfg, doc, std_type, rows, examples)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    out = Path(out_dir) if out_dir else HERE / "결과" / ("%s_%s" % (Path(path).stem, stamp))
    out.mkdir(parents=True, exist_ok=True)
    info = {
        "검토ID": rid, "표준 파일": Path(path).name, "표준 제목": doc.get("title", ""), "표준 종류": std_type,
        "점검 일시": stamp, "기준표 버전": version if version else "0 (기본 기준표)",
        "예시 데이터": "예" if sample else "아니오",
    }
    xlsx = out / ("%s_회의록초안.xlsx" % Path(path).stem)
    try:
        write_minutes(xlsx, rows, cfg, info, prompts)
    except ConfigError as e:
        # openpyxl 이 없을 때: 회의록을 CSV 로 남기고 계속합니다(판정 되받기는 엑셀이 필요).
        xlsx = out / ("%s_회의록초안.csv" % Path(path).stem)
        write_minutes_csv(xlsx, rows, cfg, info)
        say("주의: %s 회의록 초안을 CSV 로 저장했습니다." % e)
    pdir = out / "prompts"
    pdir.mkdir(exist_ok=True)
    for no, text in prompts:
        (pdir / ("%s.txt" % safe_name(no))).write_text(text, encoding="utf-8")
    report = None
    if html:
        report = out / ("%s_점검보고서.html" % Path(path).stem)
        report.write_text(review_html(info, rows, summarize(rows), prompts, doc["headings"]), encoding="utf-8")
    if not quiet:
        s = summarize(rows)
        say("점검 완료 (검토ID %d, 기준표 버전 %s)%s" % (rid, info["기준표 버전"], " — 예시 데이터" if sample else ""))
        say("  결과: " + ", ".join("%s %d" % kv for kv in s.items()))
        for r in rows:
            if r["result"] in ("부적합", "권고"):
                say("  [%s] %s %s — %s" % (r["result"], r["no"], r["detail"], r["location"]))
        say("  회의록 초안: %s" % xlsx)
        say("  판단 프롬프트 %d개: %s" % (len(prompts), pdir))
        if report:
            say("  보고서: %s" % report)
    return {"id": rid, "rows": rows, "xlsx": xlsx, "prompts": prompts, "out": out, "report": report}


def cmd_review(a, store):
    if a.type not in defaults.STD_TYPES:
        raise ConfigError("--type 은 %s 중 하나입니다." % ", ".join(defaults.STD_TYPES))
    do_review(store, a.file, a.type, a.out, html=not a.no_html)


def cmd_inspect(a, store):
    _, cfg = store.current()
    doc = read_docx(a.file, number_titles=number_titles(cfg))
    from stdreview.checks import captions, term_entries, citation_lines
    say("제목: %s" % (doc.get("title") or "(찾지 못함)"))
    say("절 %d개:" % len(doc["headings"]))
    for h in doc["headings"]:
        say("  " + "  " * (h["level"] - 1) + "%s %s" % (h["num"], h["title"]))
    caps, objects, _ = captions(doc, cfg)
    say("표 %d개 · 그림 %d개 · 캡션 %d개 · 수식 %d개" % (
        sum(1 for o in objects if o["kind"] == "표"), sum(1 for o in objects if o["kind"] == "그림"), len(caps),
        sum(b.get("omath", 0) for b in doc["blocks"])))
    for o in objects:
        say("  %s (%s절) 캡션: %s" % (o["kind"], o["block"].get("sec") or "머리말", o["caption"]["text"] if o["caption"] else "없음"))
    ents = term_entries(doc)
    say("용어정의 %d개: %s" % (len(ents), ", ".join(e["term"] for e in ents)))
    say("2항 목록 %d줄" % len(citation_lines(doc)))


# ---------- criteria ----------

def cmd_criteria(a, store):
    if a.action == "export":
        version, cfg = store.current()
        write_xlsx(cfg, a.file, version_label="%s (%s)" % (version, "기본 기준표" if version == 0 else "DB 저장본"))
        say("기준표를 내보냈습니다: %s (버전 %s)" % (a.file, version))
        return
    if a.action == "history":
        hist = store.history()
        if not hist:
            say("변경 이력이 없습니다. 기본 기준표(버전 0)를 쓰고 있습니다.")
        for v in hist:
            say("버전 %d · %s · %s%s" % (v["id"], v["created"], v["reason"] or "(사유 없음)", (" · " + v["source"]) if v["source"] else ""))
            for c in v["changes"]:
                say("   [%s] %s · %s: %s → %s" % (c["kind"], c["no"], c["field"], c["before"][:60], c["after"][:60]))
        return
    if not a.file:
        raise ConfigError("criteria import 에는 기준표 엑셀 파일이 필요합니다.")
    new = read_xlsx(a.file)
    apply_new_config(store, new, a.reason or "", "기준표 엑셀: %s" % Path(a.file).name, a.yes)


def apply_new_config(store, new, reason, source, assume_yes):
    errs = validate(new)
    if errs:
        say("기준표에 고칠 곳이 있어 넣지 않았습니다:")
        for e in errs:
            say("  - " + e)
        return None
    version, old = store.current()
    changes = diff(old, new)
    if not changes:
        say("바뀐 내용이 없습니다(현재 버전 %s 유지)." % version)
        return None
    say("현재 버전 %s 과 비교한 변경 %d건:" % (version, len(changes)))
    for c in changes:
        say("  [%s] %s · %s: %s → %s" % (c["kind"], c["no"], c["field"], c["before"][:60] or "-", c["after"][:60] or "-"))
    if not reason:
        say("  (사유를 --reason 으로 적으면 변경 이력에 함께 남습니다)")
    if not confirm("이대로 적용할까요?", assume_yes):
        say("적용하지 않았습니다.")
        return None
    vid = store.save_version(new, changes, reason, source)
    say("기준표 버전 %d 로 저장했습니다. 다음 점검부터 적용됩니다." % vid)
    return vid


def cmd_template(a, store):
    if a.type not in defaults.STD_TYPES:
        raise ConfigError("--type 은 %s 중 하나입니다." % ", ".join(defaults.STD_TYPES))
    version, cfg = store.current()
    doc = read_docx(a.file, number_titles=number_titles(cfg))
    items = items_from_template(doc, a.type)
    if not items:
        raise ConfigError("Template 에서 상위 제목(1., 2., …)을 찾지 못했습니다. inspect 명령으로 구조를 확인해 주세요.")
    import copy
    new = copy.deepcopy(cfg)
    old_req = {str(it["num"]): it.get("required", True) for it in cfg["items"] if it["type"] == a.type}
    for it in items:
        it["required"] = old_req.get(str(it["num"]), True)
    new["items"] = [it for it in cfg["items"] if it["type"] != a.type] + items
    apply_new_config(store, new, a.reason or "Template 반영", "Template: %s (%s)" % (Path(a.file).name, a.type), a.yes)


# ---------- feedback / improve ----------

def cmd_feedback(a, store):
    _, cfg = store.current()
    rid, items, warnings = read_feedback(a.file, cfg, parse_map(a.map))
    if not store.review(rid):
        raise ConfigError("검토ID %d 기록이 이 DB 에 없습니다. 점검한 PC·DB 에서 넣어 주세요(--db 로 지정 가능)." % rid)
    store.upsert_feedback(rid, items)
    n_adopt = sum(1 for i in items if i["verdict"] == "채택")
    n_ex = sum(1 for i in items if i["verdict"] == "제외")
    n_c = sum(1 for i in items if i["comment"])
    say("검토ID %d 판정을 넣었습니다: 행 %d개(채택 %d, 제외 %d, 보완 의견 %d)." % (rid, len(items), n_adopt, n_ex, n_c))
    for w in warnings:
        say("  주의: " + w)


def improve_report(store, include_samples=False, html_path=None, quiet=False):
    _, cfg = store.current()
    mc = int(setting_num(cfg, "보완후보_최소판정건수", 3))
    ratio = setting_num(cfg, "보완후보_제외비율", 0.5)
    stats = candidates(store.all_feedback(include_samples), cfg["criteria"], mc, ratio)
    note = "기준: 판정 %d건 이상 & 제외 비율 %.0f%% 이상, 또는 보완 의견이 있는 기준 (설정 시트에서 바꿉니다)" % (mc, ratio * 100)
    if not quiet:
        say(note)
        if not stats:
            say("아직 되받은 판정이 없습니다. 회의록 초안에 채택/제외를 적어 feedback 명령으로 넣어 주세요.")
        shown = [s for s in stats if s["adopt"] or s["exclude"] or s["comments"]]
        if stats and not shown:
            say("판정(채택/제외)이나 보완 의견이 적힌 행이 아직 없습니다.")
        for s in shown:
            flag = "보완 후보" if s["candidate"] else "-"
            say("  %-7s %-6s 채택 %d · 제외 %d · 미판정 %d  %s" % (flag, s["no"], s["adopt"], s["exclude"], s["blank"], " / ".join(s["reasons"])))
            for c in s["comments"][:3]:
                say("           의견: " + c)
    if html_path:
        Path(html_path).write_text(improve_html(stats, store.history(), note,
                                                "예시 데이터 포함 — 시연용 판정이 섞여 있습니다." if include_samples else ""), encoding="utf-8")
        if not quiet:
            say("보고서: %s" % html_path)
    return stats


def cmd_improve(a, store):
    improve_report(store, a.include_samples, a.html)


# ---------- demo ----------

DEMO_EXCLUDE = {"4-3": "(예시 판정) 수식이 필요 없는 절차도 있으므로 권고에서 빼도 됩니다."}


def cmd_demo(a, store):
    files = make_samples(SAMPLES)
    out_root = HERE / "결과_예시"
    out_root.mkdir(exist_ok=True)
    # 시연은 실제 DB 를 건드리지 않도록 별도 DB 를 매번 새로 만듭니다.
    demo_db = out_root / "예시데이터.db"
    if demo_db.exists():
        demo_db.unlink()
    store = Store(demo_db)
    say("예시 데이터 — 가상 표준 2건으로 전체 흐름을 보여 줍니다. (시연용 DB: %s)\n" % demo_db)
    results = []
    for f in files:
        say("■ %s" % f.name)
        results.append(do_review(store, f, "설계표준", out_root / f.stem))
        say("")
    # 오류 포함 문서에 예시 판정을 적어 되받기까지 시연
    bad = results[1]
    items = []
    for r in bad["rows"]:
        if r["result"] in ("부적합", "권고"):
            if r["no"] in DEMO_EXCLUDE:
                items.append(dict(r, verdict="제외", comment=DEMO_EXCLUDE[r["no"]]))
            else:
                items.append(dict(r, verdict="채택", comment=""))
    store.upsert_feedback(bad["id"], items)
    say("■ 예시 판정 되받기: 부적합·권고 %d건에 채택/제외를 적었다고 가정했습니다." % len(items))
    html = out_root / "보완후보_예시.html"
    improve_report(store, include_samples=True, html_path=html)
    say("\n예시 결과 폴더: %s" % out_root)
    store.close()


def cmd_samples(a, store):
    for f in make_samples(SAMPLES):
        say("만들었습니다: %s" % f)


def main(argv=None):
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    p = argparse.ArgumentParser(prog="review_tool.py", description="기술표준 검토 도구 (폐쇄망 로컬, 1단계)")
    p.add_argument("--db", default=str(DEFAULT_DB), help="DB 파일 경로 (기본: data/stdreview.db)")
    sub = p.add_subparsers(dest="cmd")

    s = sub.add_parser("review", help="표준 .docx 점검")
    s.add_argument("file")
    s.add_argument("--type", required=True, help="설계표준 / 시험검증표준 / 가상검증표준")
    s.add_argument("--out", help="결과 폴더 (기본: 결과/파일명_일시)")
    s.add_argument("--no-html", action="store_true", help="HTML 보고서를 만들지 않음")
    s.set_defaults(fn=cmd_review)

    s = sub.add_parser("inspect", help="도구가 읽은 문서 구조 보기")
    s.add_argument("file")
    s.set_defaults(fn=cmd_inspect)

    s = sub.add_parser("criteria", help="기준표 내보내기·다시 넣기·이력")
    s.add_argument("action", choices=["export", "import", "history"])
    s.add_argument("file", nargs="?")
    s.add_argument("--reason", default="")
    s.add_argument("--yes", action="store_true", help="확인 없이 적용")
    s.set_defaults(fn=cmd_criteria)

    s = sub.add_parser("template", help="Template .docx 의 항목 구성을 기준표에 반영")
    s.add_argument("action", choices=["import"])
    s.add_argument("file")
    s.add_argument("--type", required=True)
    s.add_argument("--reason", default="")
    s.add_argument("--yes", action="store_true")
    s.set_defaults(fn=cmd_template)

    s = sub.add_parser("feedback", help="회의록 초안의 검토자 판정 되받기")
    s.add_argument("file")
    s.add_argument("--map", action="append", help="열 매핑 키=열이름 (예: verdict=판정)")
    s.set_defaults(fn=cmd_feedback)

    s = sub.add_parser("improve", help="보완 후보 기준 보고")
    s.add_argument("--html", help="HTML 보고서 경로")
    s.add_argument("--include-samples", action="store_true", help="예시 데이터 판정도 포함")
    s.set_defaults(fn=cmd_improve)

    s = sub.add_parser("demo", help="예시 데이터로 전체 흐름 시연")
    s.set_defaults(fn=cmd_demo)
    s = sub.add_parser("samples", help="예시 표준 .docx 다시 만들기")
    s.set_defaults(fn=cmd_samples)

    a = p.parse_args(argv)
    if not getattr(a, "fn", None):
        p.print_help()
        return 0
    # demo·samples 는 실제 DB 를 열지 않습니다.
    store = None if a.cmd in ("demo", "samples") else Store(a.db)
    try:
        if a.cmd == "criteria" and a.action == "export" and not a.file:
            raise ConfigError("내보낼 파일 이름을 적어 주세요. 예: criteria export 기준표.xlsx")
        a.fn(a, store)
    except (ConfigError, DocxError) as e:
        say("오류: %s" % e)
        return 1
    finally:
        if store:
            store.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
