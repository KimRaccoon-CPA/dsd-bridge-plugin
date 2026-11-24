#!/usr/bin/env bash
# DSD브릿지 실행기 (macOS/Linux) - setup.sh 가 만든 venv(~/.venvs/dsd-bridge)의 파이썬으로 이 폴더의 dsd_bridge.py 를 실행한다
# 사용: ./dsd-bridge.sh export 감사보고서.dsd -o 감사보고서.xlsx
exec "$HOME/.venvs/dsd-bridge/bin/python" "$(cd "$(dirname "$0")" && pwd)/dsd_bridge.py" "$@"
