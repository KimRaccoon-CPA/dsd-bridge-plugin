#!/usr/bin/env bash
# dsd-bridge 새 PC 세팅 (macOS/Linux). 이 폴더 안에 .venv 를 만들고 의존성을 설치한 뒤 스모크 테스트를 돌린다.
# 사용: bash ./setup.sh
set -e
cd "$(dirname "$0")"
PY=${PYTHON:-python3}
if ! command -v "$PY" >/dev/null 2>&1; then echo "오류: python3 을 찾을 수 없습니다. Python 3.13 이상을 설치하세요."; exit 1; fi
ver=$("$PY" -c 'import sys; print("%d.%d" % sys.version_info[:2])')
echo "Python $ver ($(command -v "$PY"))"
"$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 13) else 1)' || { echo "오류: Python 3.13 이상이 필요합니다."; exit 1; }
VENV="$HOME/.venvs/dsd-bridge"   # 폴더 밖(로컬)에 둔다 — OneDrive 동기화 회피
[ -x "$VENV/bin/python" ] || { echo "venv 생성 중: $VENV"; "$PY" -m venv "$VENV"; }
VPY="$VENV/bin/python"
case "$(uname -s)-$(uname -m)" in
  Darwin-arm64) W=wheels/mac_arm64 ;;
  *) W="" ;;
esac
ok=0
if [ -n "$W" ] && [ -d "$W" ]; then
  echo "오프라인 설치 시도: $W"
  if "$VPY" -m pip install -q --no-index --find-links "$W" -r requirements.txt; then ok=1; else echo "오프라인 설치 실패 → 온라인 설치로 전환"; fi
fi
[ $ok = 1 ] || "$VPY" -m pip install -q -r requirements.txt
"$VPY" -c 'import openpyxl, lxml; print("openpyxl", openpyxl.__version__, "/ lxml", lxml.__version__)'
echo "--- 스모크 테스트"
"$VPY" dsd_bridge.py --help | head -3
"$VPY" dsd_bridge.py --version
# 시험 폴더는 개발 저장소에만 있다(공유용 플러그인에는 없음)
[ -d tests ] && "$VPY" -m pytest -q --basetemp "$VENV/pytest_tmp" tests/test_review_fixes.py 2>&1 | tail -2
echo
echo "준비 완료. 이후 명령은 이 파이썬으로 실행하세요: $VPY"
echo "예) $VPY dsd_bridge.py export '감사보고서.dsd' -o '감사보고서.xlsx'"
[ -d tests ] && echo '픽스처(시험용 .dsd)가 있으면 DSD_BRIDGE_FIXTURE_DIR 로 지정한 뒤 python -m pytest -q 로 전체 테스트를 돌릴 수 있습니다.'
exit 0
