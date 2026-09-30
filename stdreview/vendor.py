"""vendor/wheels 의 순수 Python 설치 파일(.whl)을 설치 없이 바로 불러오는 도우미.

폐쇄망 PC에서 pip 설치가 막혀 있어도 저장소 폴더만 있으면 PDF 읽기(pypdf)가 되도록 합니다.
이미 설치된 패키지가 있으면 그것을 먼저 씁니다.
"""

import sys
from pathlib import Path

WHEELS = Path(__file__).resolve().parent.parent / "vendor" / "wheels"


def import_or_vendor(name, wheel_prefixes):
    """name 모듈을 import 합니다. 없으면 vendor/wheels 의 .whl 을 sys.path 에 넣고 다시 시도합니다."""
    try:
        return __import__(name)
    except ImportError:
        pass
    added = False
    for prefix in wheel_prefixes:
        for whl in sorted(WHEELS.glob(prefix + "-*.whl")):
            p = str(whl)
            if p not in sys.path:
                sys.path.append(p)
                added = True
    if not added:
        raise ImportError(name)
    return __import__(name)
