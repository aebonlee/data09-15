"""보완 후보 — 누적된 검토자 판정에서 고칠 만한 기준을 골라냅니다(순수 함수)."""


def candidates(feedback, criteria, min_count, ratio):
    """feedback: [{no, verdict, comment, detail, file}] / criteria: 기준표 행 목록.

    돌려주는 값: 기준별 집계 목록. candidate=True 인 행이 보완 후보입니다.
    - 판정(채택+제외) 건수가 min_count 이상이고 제외 비율이 ratio 이상 → 「제외가 많음(오탐 의심)」
    - 보완 의견이 한 건이라도 있음 → 「보완 의견 있음」
    """
    crit = {c["no"]: c for c in criteria}
    stats = {}
    for f in feedback:
        s = stats.setdefault(f["no"], {"no": f["no"], "adopt": 0, "exclude": 0, "blank": 0, "comments": [], "excluded_examples": []})
        if f["verdict"] == "채택":
            s["adopt"] += 1
        elif f["verdict"] == "제외":
            s["exclude"] += 1
            if len(s["excluded_examples"]) < 3:
                s["excluded_examples"].append(f.get("detail", ""))
        else:
            s["blank"] += 1
        if f.get("comment"):
            s["comments"].append(f["comment"])
    out = []
    for no, s in stats.items():
        judged = s["adopt"] + s["exclude"]
        s["ratio"] = (s["exclude"] / judged) if judged else 0.0
        reasons = []
        if judged >= min_count and s["ratio"] >= ratio:
            reasons.append("제외가 많음(오탐 의심) %d/%d" % (s["exclude"], judged))
        if s["comments"]:
            reasons.append("보완 의견 %d건" % len(s["comments"]))
        s["reasons"] = reasons
        s["candidate"] = bool(reasons)
        c = crit.get(no, {})
        s["item"] = c.get("item", "(기준표에 없는 No)")
        s["text"] = c.get("text", "")
        s["active"] = c.get("active", False)
        out.append(s)

    def order(s):
        try:
            parts = [(0, int(x)) if x.isdigit() else (1, x) for x in s["no"].replace("공통", "0").split("-")]
        except Exception:
            parts = [(1, s["no"])]
        return (not s["candidate"], parts)

    return sorted(out, key=order)
