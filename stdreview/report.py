"""간단한 로컬 HTML 보고서 — 외부 파일·인터넷 없이 브라우저에서 바로 열립니다."""

from html import escape

CSS = """
:root{--bg:#fff;--fg:#1b2430;--muted:#5a6675;--line:#d9dee5;--head:#1f4e79;--on-head:#fff;
--bad:#fde2e4;--warn:#fff3c4;--judge:#e3eefa;--ok:#e6f2e0;--banner:#fde2e4;--banner-fg:#8a1020}
@media (prefers-color-scheme:dark){:root{--bg:#14181e;--fg:#e6e9ee;--muted:#a3adba;--line:#323a46;--head:#2b5f91;
--bad:#4a2327;--warn:#443a14;--judge:#1e3450;--ok:#213a24;--banner:#4a2327;--banner-fg:#ffd7dc}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font-family:"Malgun Gothic","Apple SD Gothic Neo",sans-serif;
line-height:1.6;word-break:keep-all;overflow-wrap:break-word}
main{max-width:1180px;margin:0 auto;padding:20px 16px 48px}
h1{font-size:1.4rem;margin:0 0 4px}h2{font-size:1.1rem;margin:28px 0 8px}
.muted{color:var(--muted)}
.banner{background:var(--banner);color:var(--banner-fg);padding:10px 14px;border-radius:6px;font-weight:700;margin:12px 0}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin:12px 0}
.card{border:1px solid var(--line);border-radius:6px;padding:10px 12px}.card b{display:block;font-size:1.4rem}
.wrap{overflow-x:auto;border:1px solid var(--line);border-radius:6px}
table{border-collapse:collapse;width:100%;min-width:720px;font-size:.92rem}
th,td{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
th{background:var(--head);color:var(--on-head);position:sticky;top:0}
td.r-na{color:var(--muted)}td.r-부적합{background:var(--bad)}td.r-권고{background:var(--warn)}td.r-판단{background:var(--judge)}td.r-적합{background:var(--ok)}
details{border:1px solid var(--line);border-radius:6px;margin:8px 0;padding:6px 10px}
summary{cursor:pointer;font-weight:700}
pre{white-space:pre-wrap;word-break:keep-all;overflow-wrap:break-word;font-size:.85rem;margin:8px 0 0}
ul{padding-left:20px}
"""


def _page(title, body):
    return ("<!doctype html><html lang=\"ko\"><head><meta charset=\"utf-8\">"
            "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
            "<title>%s</title><style>%s</style></head><body><main>%s</main></body></html>") % (escape(title), CSS, body)


def review_html(info, rows, summary, prompts, headings):
    e = escape
    parts = ["<h1>표준 점검 보고서</h1>",
             "<p class=\"muted\">%s · %s · 점검 %s · 기준표 버전 %s</p>" % (
                 e(info.get("표준 제목", "")), e(info.get("표준 종류", "")), e(info.get("점검 일시", "")), e(str(info.get("기준표 버전", ""))))]
    if info.get("예시 데이터") == "예":
        parts.append("<div class=\"banner\">예시 데이터 — 가상 표준 문서로 만든 시연용 결과입니다. 실제 심의 결과가 아닙니다.</div>")
    parts.append("<div class=\"cards\">" + "".join(
        "<div class=\"card\"><span class=\"muted\">%s</span><b>%d</b></div>" % (e(k), v) for k, v in summary.items()) + "</div>")
    parts.append("<p class=\"muted\">파일: %s · 검토ID %s. 판정은 회의록 초안 엑셀에 적어 되받습니다.</p>" % (e(info.get("표준 파일", "")), e(str(info.get("검토ID", "")))))
    if info.get("학습 프로필"):
        parts.append("<p class=\"muted\">학습 프로필: %s</p>" % e(info["학습 프로필"]))
    parts.append("<h2>점검 결과</h2><div class=\"wrap\"><table><thead><tr><th>순번</th><th>No</th><th>항목</th><th>결과</th><th>지적 내용</th><th>위치</th><th>수정 요청 문구</th></tr></thead><tbody>")
    for r in rows:
        cls = {"판단 필요": "r-판단", "해당 없음": "r-na"}.get(r["result"], "r-" + r["result"])
        parts.append("<tr><td>%d</td><td>%s</td><td>%s</td><td class=\"%s\">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>" % (
            r["seq"], e(r["no"]), e(r["item"]), cls, e(r["result"]), e(r["detail"]), e(r["location"]), e(r["fix"])))
    parts.append("</tbody></table></div>")
    parts.append("<h2>판단 항목 프롬프트</h2><p class=\"muted\">펼쳐서 전체를 복사해 사내 승인 AI 에 붙여 넣습니다. 같은 내용이 prompts 폴더에 파일로도 있습니다.</p>")
    for no, text in prompts:
        parts.append("<details><summary>%s</summary><pre>%s</pre></details>" % (e(no), e(text)))
    parts.append("<h2>도구가 읽은 절 구조</h2><ul>" + "".join(
        "<li>%s%s %s</li>" % ("&nbsp;&nbsp;" * (h["level"] - 1), e(h["num"]), e(h["title"])) for h in headings) + "</ul>")
    return _page("표준 점검 보고서", "".join(parts))


def improve_html(stats, history, settings_note, sample_note=""):
    e = escape
    parts = ["<h1>평가 기준 보완 후보</h1>", "<p class=\"muted\">%s</p>" % e(settings_note)]
    if sample_note:
        parts.append("<div class=\"banner\">%s</div>" % e(sample_note))
    parts.append("<h2>기준별 판정 누적</h2><div class=\"wrap\"><table><thead><tr><th>No</th><th>항목</th><th>채택</th><th>제외</th><th>미판정</th><th>보완 후보 사유</th><th>보완 의견 / 제외된 지적 예</th></tr></thead><tbody>")
    for s in stats:
        notes = s["comments"][:5] + ["(제외) " + x for x in s["excluded_examples"]]
        parts.append("<tr><td>%s</td><td>%s</td><td>%d</td><td>%d</td><td>%d</td><td class=\"%s\">%s</td><td>%s</td></tr>" % (
            e(s["no"]), e(s["item"]), s["adopt"], s["exclude"], s["blank"], "r-권고" if s["candidate"] else "",
            e(" / ".join(s["reasons"]) or "-"), "<br>".join(e(n) for n in notes) or "-"))
    parts.append("</tbody></table></div>")
    parts.append("<h2>기준표 변경 이력</h2>")
    if not history:
        parts.append("<p>아직 바꾼 적이 없습니다(기본 기준표 사용 중).</p>")
    for v in history:
        parts.append("<details><summary>버전 %d · %s · %s</summary><ul>%s</ul></details>" % (
            v["id"], e(v["created"]), e(v["reason"] or "(사유 없음)"),
            "".join("<li>[%s] %s · %s: %s → %s</li>" % (e(c["kind"]), e(c["no"]), e(c["field"]), e(c["before"][:80]), e(c["after"][:80])) for c in v["changes"]) or "<li>변경 없음</li>"))
    return _page("평가 기준 보완 후보", "".join(parts))
