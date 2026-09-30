"""기본 기준표 — 기획서 5.1 체크리스트를 그대로 옮긴 초기값.

이 값은 처음 한 번만 쓰입니다. 검토자가 기준표 엑셀을 고쳐 다시 넣으면
그 내용이 DB 에 새 버전으로 저장되어 이후 점검에 쓰입니다(기획서 5.3).
(가정) 표시가 붙은 값은 실제 Template·사내 규칙을 받으면 바꿔야 하는 임시값입니다.
"""

import copy

STD_TYPES = ["설계표준", "시험검증표준", "가상검증표준"]

RESULTS = ["적합", "부적합", "권고", "판단 필요", "해당 없음"]
VERDICTS = ["채택", "제외"]

# 점검 방식 표기 — 기획서 5.1 과 같게
METHODS = ["자동", "보조", "판단(프롬프트)"]

DEFAULT_FIX = "Template {항목} 기준에 따라 {내용}"

CRITERIA = [
    {
        "no": "공통-1", "item": "전체",
        "text": "기본 제공된 Template에 정리된 기준을 준수해야 함",
        "method": "자동", "code": "TEMPLATE_ITEMS", "target": "",
        "active": True, "fix": "Template 기준에 따라 {내용}", "prompt": "", "note": "",
    },
    {
        "no": "공통-2", "item": "전체",
        "text": "미준수 시 미준수 내용을 기준 Template 기준으로 수정하라고 멘트 제공",
        "method": "자동", "code": "FIX_TEMPLATE", "target": "",
        "active": True, "fix": DEFAULT_FIX, "prompt": "",
        "note": "이 행의 수정 요청 문구가 다른 행의 문구가 비어 있을 때 쓰는 기본 틀입니다. 회의록에 별도 행은 만들지 않습니다.",
    },
    {
        "no": "공통-3", "item": "전체(머리글·개정 이력)",
        "text": "(Template 학습에서 추가, 가정) 머리글의 표준번호·날짜·Rev 와 개정 이력이 Template 형식대로 채워지고 서로 맞아야 함",
        "method": "자동", "code": "REVISION", "target": "",
        "active": True, "fix": DEFAULT_FIX, "prompt": "",
        "note": "Template 을 학습하면 개정 이력 칸이 필수인지 Template 에서 알아냅니다. 학습 전에는 설정 「개정이력_필수」를 따릅니다.",
    },
    {
        "no": "공통-4", "item": "전체(작성 완결성)",
        "text": "(Template 학습에서 추가, 가정) Template 의 예시 문구·자리표시(xx, 0000)와 「추후 보완 예정」 같은 미완성 표기가 남아 있지 않아야 함",
        "method": "자동", "code": "LEFTOVER", "target": "",
        "active": True, "fix": DEFAULT_FIX, "prompt": "",
        "note": "예시 문구 대조는 Template 을 학습한 뒤부터 됩니다. 미완성 표기는 학습 없이도 찾습니다.",
    },
    {
        "no": "1-1", "item": "1항 목적 및 정의",
        "text": "작성된 표준의 목적과 정의를 구체적으로 정리하여, 제목의 내용에 대한 표준 문서의 목적과 정의를 구체적으로 정리해야 함",
        "method": "판단(프롬프트) + 보조", "code": "PURPOSE", "target": "1",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "1항의 목적과 정의가 표준 제목의 내용에 맞게 구체적으로 적혀 있는지 판단해 주세요. 막연한 표현(예: 품질 향상을 위해)만 있는 문장은 지적해 주세요.",
        "note": "",
    },
    {
        "no": "2-1", "item": "2항 관련 표준 및 인용",
        "text": "표준의 내용에 관련된 표준과 검토된 기술지식에 대한 문서번호와 제목이 정리되어야 함",
        "method": "자동", "code": "CITATION_LIST", "target": "2",
        "active": True, "fix": DEFAULT_FIX, "prompt": "", "note": "",
    },
    {
        "no": "3-1", "item": "3항 용어정의",
        "text": "표준에 사용된 용어에 대한 정의로, 일반적인 상식이 아닌 작성된 표준에 적용되는 용어에 대한 정리와 설명이 구체적으로 정리되어야 함",
        "method": "보조 + 판단(프롬프트)", "code": "TERM_DEF", "target": "3",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "용어정의의 각 설명이 일반 상식 수준의 사전 뜻인지, 이 표준에 적용되는 뜻으로 구체적으로 적혀 있는지 용어별로 판단해 주세요.",
        "note": "",
    },
    {
        "no": "3-2", "item": "3항 용어정의 ↔ 본문 전체",
        "text": "정리된 용어정의는 표준 전체에 동일 기준으로 사용되어야 함. 용어가 통일되게 사용되지 않은 것도 지적 대상 (예: ECU는 계속 ECU로 써야 하며, 한글 「이씨유」 표기 불가)",
        "method": "자동", "code": "TERM_CONSISTENCY", "target": "",
        "active": True, "fix": DEFAULT_FIX, "prompt": "", "note": "원문 예시의 「EUC」는 ECU 의 오기로 보았습니다(가정).",
    },
    {
        "no": "4-1", "item": "4항 설계절차/시험절차",
        "text": "설계절차/시험절차는 설계 방법을 항목 순서 기준으로 정리함",
        "method": "자동 + 보조", "code": "SECTION_SEQUENCE", "target": "4+",
        "active": True, "fix": DEFAULT_FIX, "prompt": "", "note": "",
    },
    {
        "no": "4-2", "item": "4항 설계절차/시험절차",
        "text": "4.1 또는 4.2항에 설계절차, 시험절차를 ISO 기준으로 FlowChart로 제시해야 함",
        "method": "자동 + 보조", "code": "FLOWCHART", "target": "4",
        "active": True, "fix": DEFAULT_FIX, "prompt": "",
        "note": "ISO 기호 준수 여부는 그림 판독이 필요해 검토자 확인 항목입니다.",
    },
    {
        "no": "4-3", "item": "4항 이하",
        "text": "설계·시험 관련 절차를 정리하며 가능하면 수식으로 제시해야 함",
        "method": "보조", "code": "EQUATIONS", "target": "4+",
        "active": True, "fix": DEFAULT_FIX, "prompt": "", "note": "「가능하면」이므로 부적합이 아니라 권고로 적습니다.",
    },
    {
        "no": "4-4", "item": "4항 이하",
        "text": "작성되는 내용 중 그림, 표에는 반드시 표1, 그림(Fig)1 등 번호가 기재되어야 함",
        "method": "자동", "code": "CAPTIONS", "target": "",
        "active": True, "fix": DEFAULT_FIX, "prompt": "", "note": "",
    },
    {
        "no": "4-5", "item": "1항 ↔ 4항 이하",
        "text": "1항 목적 및 범위에 정리된 내용과 4항 이하 설계·시험절차에 정리된 내용을 비교하여 논리적으로 MECE 해야 함",
        "method": "판단(프롬프트)", "code": "", "target": "1,4+",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "1항에 적힌 목적·범위의 요소를 목록으로 뽑은 뒤, 4항 이하 절차가 그 요소를 빠짐없이(누락 없음) 다루는지, 서로 겹치지 않는지(중복 없음) 판단해 주세요. 누락과 중복을 나눠 적어 주세요.",
        "note": "",
    },
    {
        "no": "4-6", "item": "4항 이하",
        "text": "구체적 기준에 대해 근거를 명확하게 정량화하여 표현하고, 관련된 표준·법규 등은 반드시 근거를 명확히 제시해야 함. 제시된 표준 No.와 제목이 맞는지 확인 필요",
        "method": "자동 + 보조 + 판단", "code": "CITATION_REF", "target": "4+",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "4항 이하의 기준 문장이 수치와 단위로 정량화되어 있는지, 수치 기준마다 근거(표준·법규·시험 결과)가 제시되어 있는지 판단해 주세요.",
        "note": "원문의 「정수화」는 「정량화」로 보았습니다(가정). 사내 표준 목록과의 번호-제목 대조는 목록 확보 후(확장).",
    },
    {
        "no": "4-7", "item": "4항",
        "text": "(원문 「7)」 이후 내용 없음 — 기준 확인 필요)",
        "method": "판단(프롬프트)", "code": "", "target": "4+",
        "active": False, "fix": DEFAULT_FIX, "prompt": "",
        "note": "기준을 채운 뒤 사용을 Y 로 바꾸면 판단 프롬프트가 만들어집니다.",
    },
    {
        "no": "5-1", "item": "5항 설계/시험 검증",
        "text": "설계 또는 시험 방법에 제시한 사항을 검증하는 방법을 제시해야 함",
        "method": "보조 + 판단(프롬프트)", "code": "SECTION_JUDGE", "target": "4+,5",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "4항 이하의 절차마다 5항에 대응하는 검증 방법이 있는지 절차별로 짝지어 판단해 주세요. 검증 방법이 없는 절차를 적어 주세요.",
        "note": "",
    },
    {
        "no": "5-2", "item": "5항 설계/시험 검증",
        "text": "최대한 정성적이 아닌 정량적 기준으로 관련 근거를 제기해야 함",
        "method": "보조 + 판단(프롬프트)", "code": "QUANT", "target": "5",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "5항의 검증 기준이 정량적(수치·단위·허용 범위)인지 판단해 주세요. 정성 표현만 있는 문장은 정량 기준으로 바꾸는 예를 들어 주세요.",
        "note": "",
    },
    {
        "no": "5-3", "item": "5항",
        "text": "(원문 「3)」 이후 내용 없음 — 기준 확인 필요)",
        "method": "판단(프롬프트)", "code": "", "target": "5",
        "active": False, "fix": DEFAULT_FIX, "prompt": "",
        "note": "기준을 채운 뒤 사용을 Y 로 바꾸면 판단 프롬프트가 만들어집니다.",
    },
    {
        "no": "6-1", "item": "6항",
        "text": "(원문에 항목명과 기준 「1)」 모두 비어 있음) — 항목명이 정해지기 전까지 존재 여부만 점검",
        "method": "자동(존재만)", "code": "EXISTS", "target": "6",
        "active": True, "fix": DEFAULT_FIX, "prompt": "", "note": "",
    },
    {
        "no": "7-1", "item": "7항 안전 관련 사항",
        "text": "표준을 설계/시험하며 지켜져야 할 안전 기준에 대해 정리되어야 함",
        "method": "보조 + 판단(프롬프트)", "code": "SECTION_JUDGE", "target": "7",
        "active": True, "fix": DEFAULT_FIX,
        "prompt": "7항에 설계·시험 중 지켜야 할 안전 기준이 구체적으로(대상, 조건, 한계값, 조치) 정리되어 있는지 판단해 주세요.",
        "note": "",
    },
    {
        "no": "8-1", "item": "8항 기타(Appendix)",
        "text": "표준 내용을 보완하여 설명을 추가",
        "method": "자동(존재만)", "code": "EXISTS", "target": "8",
        "active": True, "fix": DEFAULT_FIX, "prompt": "",
        "note": "필수인지 선택인지 확인 필요 — 항목 시트의 필수 값을 따릅니다(기본 N, 가정).",
    },
]


def _items_for(std_type):
    rows = [
        (1, "1", "목적 및 정의|목적 및 범위", True),
        (2, "2", "관련 표준 및 인용", True),
        (3, "3", "용어정의", True),
        (4, "4", "설계절차|시험절차", True),
        (5, "5", "설계 검증|시험 검증|설계/시험 검증", True),
        (6, "6", "", True),
        (7, "7", "안전 관련 사항", True),
        (8, "8", "기타|Appendix", False),
    ]
    return [
        {"type": std_type, "order": o, "num": n, "name": name, "required": req}
        for (o, n, name, req) in rows
    ]


ITEMS = [row for t in STD_TYPES for row in _items_for(t)]

# 제출 원문에 나온 예시만 넣었습니다. 나머지는 검토자가 기준표에서 채웁니다.
VARIANTS = [
    {"term": "ECU", "variants": ["이씨유", "이시유"]},
]

# 정성 표현 낱말 — 초기 예시(가정). 기준표 「정성표현」 시트에서 고칩니다.
QUALITATIVE = ["양호", "적절", "적정", "충분", "우수", "원활", "문제없", "이상 없", "만족할 것"]

# 표준번호 형식 — 국제·국가 표준의 일반 표기와 사내 표준 예시 형식(가정).
STD_PATTERNS = [
    {"name": "ISO", "regex": r"ISO(?:/IEC)?\s?\d{3,5}(?:-\d+)*(?::\d{4})?"},
    {"name": "IEC", "regex": r"(?<!/)IEC\s?\d{3,5}(?:-\d+)*(?::\d{4})?"},
    {"name": "KS", "regex": r"KS\s?[A-Z]\s?(?:ISO\s?|IEC\s?)?\d{3,5}(?:-\d+)*"},
    {"name": "SAE", "regex": r"SAE\s?J\d{3,5}"},
    {"name": "법령 공포번호(…령 제 n호)", "regex": r"제\s?\d{1,5}\s?호"},
    {"name": "사내 표준(예시 형식, 가정 — 실제 규칙으로 교체)", "regex": r"\b[A-Z]{2,5}-\d{3,6}(?:-\d+)?\b"},
]

SETTINGS = {
    "짧은_설명_글자수": "10",
    "흐름도_캡션_낱말": "흐름도,FlowChart,Flow Chart,Flowchart,플로우차트,순서도",
    "표_캡션_머리말": "표,Table",
    "그림_캡션_머리말": "그림,Fig.,Fig,Figure",
    "번호_제목_인식": "예",
    "프롬프트_본문_최대글자": "12000",
    "보완후보_최소판정건수": "3",
    "보완후보_제외비율": "0.5",
    "과거사례_최대건수": "5",
    "개정이력_필수": "아니오",
    "수치_단위_정규식": r"\d+(?:\.\d+)?\s?(?:%|℃|°C|mm|cm|km|m|kg|g|kN|N·m|Nm|N|MPa|kPa|Pa|bar|mV|V|mA|A|kW|W|kHz|Hz|rpm|ms|s|min|h|dB|mL|L|시간|분|초|회|개|배)(?![A-Za-z가-힣])",
}

SETTING_NOTES = {
    "짧은_설명_글자수": "용어 설명이 이 글자 수보다 짧으면 3-1 근거로 표시합니다(임시 기본값).",
    "흐름도_캡션_낱말": "4-2: 4.1·4.2 절 그림 캡션에 이 낱말 중 하나가 있으면 흐름도로 봅니다. 쉼표로 구분.",
    "표_캡션_머리말": "표 캡션으로 보는 머리말. 쉼표로 구분.",
    "그림_캡션_머리말": "그림 캡션으로 보는 머리말. 쉼표로 구분.",
    "번호_제목_인식": "예: 「4.1 시험 조건」처럼 번호로 시작하는 짧은 문단도 제목으로 봅니다. 아니오: Word 제목 스타일만 제목으로 봅니다.",
    "프롬프트_본문_최대글자": "프롬프트에 넣는 본문 최대 글자 수. 사내 AI 입력 한도에 맞춰 고칩니다(임시 기본값).",
    "보완후보_최소판정건수": "이 건수 이상 판정이 쌓인 기준만 제외 비율을 봅니다(임시 기본값).",
    "보완후보_제외비율": "제외 판정 비율이 이 값 이상이면 보완 후보로 올립니다. 0~1(임시 기본값).",
    "과거사례_최대건수": "판단 프롬프트에 넣는 과거 채택 지적 예의 최대 건수.",
    "개정이력_필수": "예: 개정 이력이 없으면 부적합. Template 을 학습하면 Template 에 개정 이력이 있는지로 정합니다.",
    "수치_단위_정규식": "수치+단위 문장을 찾는 정규식. 사내에서 쓰는 단위를 더합니다.",
}

# 회의록 열 — 기존 회의록 양식을 받으면 「열 이름」만 바꿔 맞춥니다(열 매핑).
COLUMNS = [
    {"key": "seq", "header": "순번"},
    {"key": "no", "header": "No"},
    {"key": "item", "header": "Template 항목"},
    {"key": "criterion", "header": "평가 기준"},
    {"key": "result", "header": "결과"},
    {"key": "detail", "header": "지적 내용"},
    {"key": "location", "header": "위치"},
    {"key": "fix", "header": "수정 요청 문구"},
    {"key": "verdict", "header": "검토자 판정"},
    {"key": "comment", "header": "보완 의견"},
]
COLUMN_KEYS = [c["key"] for c in COLUMNS]
COLUMN_NOTES = {
    "seq": "필수 — 판정을 되받을 때 행을 찾는 번호입니다. 지우지 마세요.",
    "no": "기준 번호(기준표 No).",
    "verdict": "검토자가 채택 또는 제외를 적습니다.",
    "comment": "기준을 어떻게 고치면 좋을지 적습니다. 보완 후보 보고에 모입니다.",
}


def default_config():
    return copy.deepcopy({
        "criteria": CRITERIA,
        "items": ITEMS,
        "variants": VARIANTS,
        "qualitative": QUALITATIVE,
        "std_patterns": STD_PATTERNS,
        "settings": SETTINGS,
        "columns": COLUMNS,
    })
