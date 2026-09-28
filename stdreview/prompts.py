"""판단 항목 프롬프트 — 도구는 AI 를 부르지 않습니다.

평가 기준 + 대상 절 본문 + 도구가 찾은 근거 + (있으면) 과거 채택 지적 예 + 응답 형식을
한 덩어리로 묶어 텍스트로 내보냅니다. 검토자가 사내 승인 AI 에 붙여 넣어 씁니다.
"""

from .checks import JUDGE_ONLY_TEXT, is_judgment, setting_num
from .docx_reader import section_text

RESPONSE_FORMAT = """[응답 형식] 아래 네 줄 형식을 그대로 지켜 답해 주세요. 지적이 여러 개면 네 줄 묶음을 반복합니다.
판정: 적합 | 부적합 | 권고 중 하나
근거 위치: (절 번호와 원문 발췌)
지적 내용: (무엇이 기준에 어긋나는지)
수정 요청: (Template 기준에 따라 어떻게 고치라는 문구)"""


def build_prompt(crit, doc, std_type, evidence="", examples=None, limit=12000):
    target = crit.get("target") or ""
    body = section_text(doc, target, limit=limit) if target else section_text(doc, "0+", limit=limit)
    lines = [
        "[역할] 당신은 기술표준 심의 검토자입니다. 아래 평가 기준 한 가지만으로 표준 본문을 검토해 주세요.",
        "본문에 없는 내용을 지어내지 말고, 판단 근거가 되는 원문 위치를 반드시 적어 주세요.",
        "",
        "[표준 정보] 제목: %s / 표준 종류: %s" % (doc.get("title") or "(제목 없음)", std_type),
        "",
        "[평가 기준] No %s · %s" % (crit["no"], crit.get("item", "")),
        crit.get("text", ""),
    ]
    if crit.get("prompt"):
        lines += ["", "[검토 지시]", crit["prompt"]]
    if evidence:
        lines += ["", "[도구가 먼저 찾은 근거]", evidence]
    if examples:
        lines += ["", "[과거 심의에서 채택된 지적 예 — 문구와 수준을 참고]"]
        lines += ["- " + e for e in examples]
    lines += ["", "[검토 대상 본문] (대상 절: %s)" % (target or "전체"), body or "(대상 절 본문을 찾지 못했습니다)", "", RESPONSE_FORMAT]
    return "\n".join(lines)


def build_all(cfg, doc, std_type, rows, examples_by_no=None):
    """판단 방식 기준마다 프롬프트를 만듭니다. 돌려주는 값: [(no, 텍스트)]."""
    limit = int(setting_num(cfg, "프롬프트_본문_최대글자", 12000))
    out = []
    for crit in cfg.get("criteria", []):
        if not crit.get("active") or not is_judgment(crit):
            continue
        ev = [r["detail"] for r in rows if r["no"] == crit["no"] and r["result"] == "판단 필요" and r["detail"] != JUDGE_ONLY_TEXT]
        ex = (examples_by_no or {}).get(crit["no"], [])
        out.append((crit["no"], build_prompt(crit, doc, std_type, evidence="\n".join(ev), examples=ex, limit=limit)))
    return out


def safe_name(no):
    return "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in no)
