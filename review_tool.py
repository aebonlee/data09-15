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

  학습(내 PC 의 data/profile.json 에만 저장, 저장소에 올라가지 않음)
  python review_tool.py learn template Template.pdf --type 설계표준   Template 항목·예시 문구·개정 이력 학습
  python review_tool.py learn standard 승인표준1.pdf 승인표준2.docx   용어·인용번호 모양·캡션 방식 학습
  python review_tool.py learn minutes 기존회의록.xlsx                 지적 범주·문체·열 이름 학습
  python review_tool.py learn show | learn reset                      학습 현황(건수만) / 지우기
  python review_tool.py compare 결과\…\회의록초안.xlsx 기존회의록.xlsx   도구 지적과 기존 회의록 일치율

인터넷·AI 호출 없음. 필요한 것: Python 3.8 이상 + openpyxl(엑셀 입출력). PDF 는 vendor/wheels 의 pypdf 를 씁니다.
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
from stdreview.docx_reader import DocxError, read_document  # noqa: E402
from stdreview import profile as prof  # noqa: E402
from stdreview.improve import candidates  # noqa: E402
from stdreview.minutes import parse_map, read_feedback, write_minutes, write_minutes_csv  # noqa: E402
from stdreview.prompts import build_all, safe_name  # noqa: E402
from stdreview.report import improve_html, review_html  # noqa: E402
from stdreview.sample_docx import make_samples  # noqa: E402
from stdreview.sample_learning import make_learning_samples  # noqa: E402
from stdreview.store import Store  # noqa: E402

DEFAULT_DB = HERE / "data" / "stdreview.db"
DEFAULT_PROFILE = HERE / "data" / "profile.json"
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


def ensure_new_criteria(cfg):
    """예전 버전 기준표(DB)에 새로 생긴 기준(공통-3·공통-4)이 없으면 기본값을 채워 넣습니다(저장은 하지 않음)."""
    have = {c.get("no") for c in cfg.get("criteria", [])}
    codes = {c.get("code") for c in cfg.get("criteria", [])}
    for i, c in enumerate(defaults.CRITERIA):
        if c["no"] not in have and c["code"] in ("REVISION", "LEFTOVER") and c["code"] not in codes:
            pos = next((k + 1 for k, x in enumerate(cfg["criteria"]) if x.get("no") == "공통-2"), len(cfg["criteria"]))
            cfg["criteria"].insert(pos, dict(c))
    for k, v in defaults.SETTINGS.items():
        cfg.setdefault("settings", {}).setdefault(k, v)
    return cfg


def effective_cfg(store, profile_path=None, std_type=None, use_profile=True):
    """(기준표 버전, 적용 설정, 학습 적용 메모). 학습 프로필이 있으면 덧입힙니다."""
    version, cfg = store.current()
    cfg = ensure_new_criteria(cfg)
    notes = []
    profile = None
    if use_profile and profile_path and Path(profile_path).exists():
        profile = prof.load_profile(profile_path)
        cfg, notes = prof.apply_profile(cfg, profile, std_type)
    return version, cfg, notes, profile


# ---------- review ----------

def do_review(store, path, std_type, out_dir=None, html=True, quiet=False, profile_path=None, use_profile=True):
    version, cfg, notes, profile = effective_cfg(store, profile_path, std_type, use_profile)
    doc = read_document(path, number_titles=number_titles(cfg))
    rows = assemble(run_checks(doc, cfg, std_type), cfg)
    sample = is_sample_file(path)
    rid = store.add_review(Path(path).name, std_type, doc.get("title", ""), version, rows, is_sample=sample)
    limit = int(setting_num(cfg, "과거사례_최대건수", 5))
    examples = store.adopted_examples(limit, include_samples=sample)
    if profile:
        for no, lst in prof.minutes_examples(profile).items():
            cur = examples.setdefault(no, [])
            for t in lst:
                if len(cur) < limit and t not in cur:
                    cur.append(t)
    prompts = build_all(cfg, doc, std_type, rows, examples)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    out = Path(out_dir) if out_dir else HERE / "결과" / ("%s_%s" % (Path(path).stem, stamp))
    out.mkdir(parents=True, exist_ok=True)
    info = {
        "검토ID": rid, "표준 파일": Path(path).name, "표준 제목": doc.get("title", ""), "표준 종류": std_type,
        "점검 일시": stamp, "기준표 버전": version if version else "0 (기본 기준표)",
        "예시 데이터": "예" if sample else "아니오",
        "문서 형식": doc.get("source", "docx"),
        "학습 프로필": ("적용 — " + " / ".join(notes)) if notes else "없음(기본 기준표만 사용)",
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
        say("  학습 프로필: %s" % info["학습 프로필"])
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
    do_review(store, a.file, a.type, a.out, html=not a.no_html, profile_path=a.profile, use_profile=not a.no_profile)


def cmd_inspect(a, store):
    _, cfg, _, _ = effective_cfg(store, a.profile, None)
    doc = read_document(a.file, number_titles=number_titles(cfg))
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
    rev = doc.get("revision") or {}
    say("형식 %s · 머리글 %d줄 · 개정 이력 %s(%d줄)" % (doc.get("source"), len(doc.get("header", [])), "있음" if rev.get("found") else "없음", len(rev.get("rows", []))))


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
    doc = read_document(a.file, number_titles=number_titles(cfg))
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
    _, cfg, _, _ = effective_cfg(store, a.profile, None)
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


# ---------- learn / compare ----------

def learn_files(kind, files, profile_path, std_type=None, quiet=False):
    """Template·기존 표준·회의록을 읽어 학습 프로필에 더합니다. 같은 파일 이름을 다시 넣으면 덮어씁니다."""
    profile = prof.load_profile(profile_path)
    for f in files:
        name = Path(f).name
        if kind == "template":
            if std_type not in defaults.STD_TYPES:
                raise ConfigError("learn template 에는 --type %s 중 하나가 필요합니다." % " / ".join(defaults.STD_TYPES))
            doc = read_document(f)
            if not [h for h in doc["headings"] if h["level"] == 1]:
                raise ConfigError("Template 에서 상위 항목(1. 2. …)을 찾지 못했습니다. inspect 로 구조를 확인해 주세요: %s" % name)
            entry = prof.learn_template(doc, std_type, name)
            profile["templates"][std_type] = entry
        elif kind == "standard":
            doc = read_document(f)
            profile["standards"][name] = prof.learn_standard(doc, name)
        elif kind == "minutes":
            items, headers, roles = prof.read_minutes(f)
            if not items:
                raise ConfigError("회의록에서 지적 문장을 찾지 못했습니다: %s — 엑셀이면 「검토의견」「구분」 같은 열 이름이 있는지 확인해 주세요." % name)
            profile["minutes"][name] = prof.learn_minutes(items, headers, roles, name)
        else:
            raise ConfigError("learn 종류는 template / standard / minutes 입니다.")
        if not quiet:
            say("학습했습니다: %s (%s)" % (name, {"template": "Template", "standard": "기존 표준", "minutes": "회의록"}[kind]))
    prof.save_profile(profile, profile_path)
    if not quiet:
        for line in prof.summary(profile):
            say("  " + line)
        say("학습 프로필: %s (이 PC 에만 있습니다. 저장소에 올리지 마세요)" % profile_path)
    return profile


def cmd_learn(a, store):
    path = Path(a.profile)
    if a.kind == "show":
        if not path.exists():
            say("학습 프로필이 없습니다: %s" % path)
            return
        p = prof.load_profile(path)
        say("학습 프로필: %s (갱신 %s)" % (path, p.get("updated", "")))
        for line in prof.summary(p) or ["(비어 있음)"]:
            say("  " + line)
        return
    if a.kind == "reset":
        if path.exists() and confirm("학습 프로필 %s 을 지울까요?" % path, a.yes):
            path.unlink()
            say("지웠습니다.")
        return
    if not a.files:
        raise ConfigError("학습할 파일을 적어 주세요. 예: learn %s 파일.pdf" % a.kind)
    learn_files(a.kind, a.files, path, a.type)


def tool_findings(xlsx, cfg):
    """도구가 만든 회의록 초안에서 (No, 결과, 위치) 목록을 읽습니다."""
    from stdreview.config import _openpyxl
    openpyxl = _openpyxl()
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb["회의록 초안"] if "회의록 초안" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    heads = {c["key"]: c["header"] for c in cfg["columns"]}
    for i, r in enumerate(rows[:5]):
        names = [("" if v is None else str(v).strip()) for v in r]
        if heads["no"] in names and heads["result"] in names:
            ix = {k: names.index(h) for k, h in heads.items() if h in names}
            out = []
            for rr in rows[i + 1:]:
                def g(k):
                    j = ix.get(k)
                    return "" if j is None or j >= len(rr) or rr[j] is None else str(rr[j]).strip()
                if g("result") in ("부적합", "권고"):
                    out.append({"no": g("no"), "location": g("location"), "text": g("detail")})
            return out
    raise ConfigError("회의록 초안에서 No·결과 열을 찾지 못했습니다: %s" % xlsx)


def _top_sec(text):
    import re
    m = re.search(r"(?<![\d.])(\d{1,2})(?:\.\d+)*\s*(?:항|절)|\(\s*(\d{1,2})(?:\.\d+)*\s*\)|^\s*(\d{1,2})(?:\.\d+)+", text or "")
    return next((g for g in (m.groups() if m else []) if g), "")


def compare_counts(tool, past):
    """도구 지적과 기존 회의록 지적을 범주(+상위 절)로 맞춰 봅니다. 건수만 돌려줍니다."""
    t = [{"code": prof.NO_TO_CODE.get(x["no"], "ETC"), "sec": _top_sec(x["location"])} for x in tool]
    p = [{"code": prof.classify(x["text"], x.get("category", "")), "sec": _top_sec(x.get("location", "") + " " + x["text"])} for x in past]

    def same(a, b):
        return a["code"] == b["code"] and (not a["sec"] or not b["sec"] or a["sec"] == b["sec"])
    hit_past = sum(1 for x in p if any(same(x, y) for y in t))
    hit_tool = sum(1 for y in t if any(same(x, y) for x in p))
    by = {}
    for x in p:
        d = by.setdefault(x["code"], [0, 0, 0])
        d[0] += 1
        d[1] += 1 if any(same(x, y) for y in t) else 0
    for y in t:
        by.setdefault(y["code"], [0, 0, 0])[2] += 1
    return {"past": len(p), "past_hit": hit_past, "tool": len(t), "tool_hit": hit_tool, "by_code": by}


def cmd_compare(a, store):
    _, cfg, _, _ = effective_cfg(store, a.profile, None)
    tool = tool_findings(a.tool_xlsx, cfg)
    past = []
    for f in a.minutes:
        items, _, _ = prof.read_minutes(f)
        past += items
    return compare_counts_print(tool, past)


def compare_counts_print(tool, past):
    r = compare_counts(tool, past)
    say("기존 회의록 지적 %d건 중 도구가 같은 범주(·절)로 잡은 것 %d건 (재현율 %s)" % (
        r["past"], r["past_hit"], "%.0f%%" % (100.0 * r["past_hit"] / r["past"]) if r["past"] else "-"))
    say("도구 지적(부적합·권고) %d건 중 기존 회의록에도 있는 범주 %d건 (일치율 %s)" % (
        r["tool"], r["tool_hit"], "%.0f%%" % (100.0 * r["tool_hit"] / r["tool"]) if r["tool"] else "-"))
    say("범주별 (기존 지적 / 그중 도구 일치 / 도구 지적):")
    for code, (n, h, tn) in sorted(r["by_code"].items()):
        say("  %-10s %3d / %3d / %3d" % (prof.CODE_LABELS.get(code, code), n, h, tn))
    return r


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

    # 학습 시연: 가상 Template·승인 표준·기존 회의록 → 학습 프로필 → 가상 초안 점검 → 기존 회의록과 비교
    say("\n■ 학습 시연 — 가상 Template·승인 표준·기존 회의록을 학습한 뒤 가상 초안을 점검합니다.")
    paths = make_learning_samples(SAMPLES)
    pp = out_root / "예시_학습프로필.json"
    if pp.exists():
        pp.unlink()
    learn_files("template", [paths["template"]], pp, "설계표준", quiet=True)
    learn_files("standard", [paths["standard"]], pp, quiet=True)
    learn_files("minutes", [paths["minutes"]], pp, quiet=False)
    res = do_review(store, paths["draft"], "설계표준", out_root / paths["draft"].stem, profile_path=pp)
    say("\n■ 도구 회의록 초안과 가상 기존 회의록 비교")
    compare_counts_print(tool_findings(res["xlsx"], effective_cfg(store, pp, None)[1]), prof.read_minutes(paths["minutes"])[0])
    say("\n예시 결과 폴더: %s" % out_root)
    store.close()


def cmd_samples(a, store):
    for f in make_samples(SAMPLES):
        say("만들었습니다: %s" % f)
    for f in make_learning_samples(SAMPLES).values():
        say("만들었습니다: %s" % f)


def main(argv=None):
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    p = argparse.ArgumentParser(prog="review_tool.py", description="기술표준 검토 도구 (폐쇄망 로컬, 1단계)")
    p.add_argument("--db", default=str(DEFAULT_DB), help="DB 파일 경로 (기본: data/stdreview.db)")
    p.add_argument("--profile", default=str(DEFAULT_PROFILE), help="학습 프로필 경로 (기본: data/profile.json)")
    sub = p.add_subparsers(dest="cmd")

    s = sub.add_parser("review", help="표준 .docx / .pdf 점검")
    s.add_argument("file")
    s.add_argument("--no-profile", action="store_true", help="학습 프로필을 쓰지 않고 기본 기준표만으로 점검")
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

    s = sub.add_parser("learn", help="Template·기존 표준·회의록 학습 (내 PC 프로필)")
    s.add_argument("kind", choices=["template", "standard", "minutes", "show", "reset"])
    s.add_argument("files", nargs="*")
    s.add_argument("--type", help="learn template 의 표준 종류: 설계표준 / 시험검증표준 / 가상검증표준")
    s.add_argument("--yes", action="store_true")
    s.set_defaults(fn=cmd_learn)

    s = sub.add_parser("compare", help="도구 회의록 초안과 기존 회의록 비교(건수·일치율)")
    s.add_argument("tool_xlsx")
    s.add_argument("minutes", nargs="+")
    s.set_defaults(fn=cmd_compare)

    s = sub.add_parser("demo", help="예시 데이터로 전체 흐름 시연")
    s.set_defaults(fn=cmd_demo)
    s = sub.add_parser("samples", help="예시 표준 .docx 다시 만들기")
    s.set_defaults(fn=cmd_samples)

    a = p.parse_args(argv)
    if not getattr(a, "fn", None):
        p.print_help()
        return 0
    # demo·samples 는 실제 DB 를 열지 않습니다.
    store = None if a.cmd in ("demo", "samples", "learn") else Store(a.db)
    try:
        if a.cmd == "criteria" and a.action == "export" and not a.file:
            raise ConfigError("내보낼 파일 이름을 적어 주세요. 예: criteria export 기준표.xlsx")
        a.fn(a, store)
    except (ConfigError, DocxError, prof.ProfileError) as e:
        say("오류: %s" % e)
        return 1
    finally:
        if store:
            store.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
