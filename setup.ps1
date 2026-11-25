# DSD브릿지(dsd-bridge) 새 PC 세팅 (Windows PowerShell 5.1+). 폴더 밖(%LOCALAPPDATA%\dsd-bridge\.venv)에 전용 파이썬을 만들고 의존성을 설치한 뒤 동작을 확인한다.
# 사용: powershell -ExecutionPolicy Bypass -File .\setup.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) { $py = Get-Command py -ErrorAction SilentlyContinue }
if (-not $py) { Write-Host '오류: python 을 찾을 수 없습니다. Python 3.13 이상을 설치하세요 (https://www.python.org/downloads/).'; exit 1 }
$ver = & $py.Source -c "import sys; print('%d.%d' % sys.version_info[:2])"
Write-Host "Python $ver ($($py.Source))"
if ([version]$ver -lt [version]'3.13') { Write-Host '오류: Python 3.13 이상이 필요합니다.'; exit 1 }

# venv 는 폴더 밖(로컬 디스크)에 둔다: 플러그인 폴더는 업데이트 때 바뀌고, OneDrive 안이면 동기화 부담이 크다
$venv = Join-Path $env:LOCALAPPDATA 'dsd-bridge\.venv'
$vpy = Join-Path $venv 'Scripts\python.exe'
if (-not (Test-Path $vpy)) { Write-Host '.venv 생성 중...'; & $py.Source -m venv $venv }

# 플랫폼별 오프라인 휠 폴더
$arch = $env:PROCESSOR_ARCHITECTURE
$wheelDir = if ($arch -eq 'ARM64') { Join-Path $root 'wheels\win_arm64' } else { Join-Path $root 'wheels\win_amd64' }
$ok = $false
if (Test-Path $wheelDir) {
    Write-Host "오프라인 설치 시도: $wheelDir"
    & $vpy -m pip install -q --no-index --find-links $wheelDir -r requirements.txt
    if ($LASTEXITCODE -eq 0) { $ok = $true } else { Write-Host '오프라인 설치 실패 → 온라인 설치로 전환' }
}
if (-not $ok) {
    & $vpy -m pip install -q -r requirements.txt
    if ($LASTEXITCODE -ne 0) { Write-Host '오류: 의존성 설치 실패 (인터넷 또는 휠 폴더를 확인하세요)'; exit 1 }
}
& $vpy -c "import openpyxl, lxml; print('openpyxl', openpyxl.__version__, '/ lxml', lxml.__version__)"

# 엑셀 조종(병합 표 행 끼우기, 스킬의 dsd_insert_rows.py)용 pywin32. 없어도 핵심 기능은 된다(0.3.0)
& $vpy -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('win32com') else 1)"
if ($LASTEXITCODE -ne 0) {
    & $vpy -m pip install -q --retries 1 --timeout 15 pywin32
    if ($LASTEXITCODE -eq 0) { Write-Host 'pywin32 설치 완료(엑셀 행 끼우기용)' }
    else { Write-Host '안내: pywin32 설치 실패. 병합 표 행 끼우기(dsd_insert_rows.py)만 못 쓰고 나머지는 정상입니다.' }
}

Write-Host '--- 스모크 테스트'
& $vpy dsd_bridge.py --help | Select-Object -First 3
& $vpy dsd_bridge.py --version
# 시험 폴더는 개발 저장소에만 있다(공유용 플러그인에는 없음)
if (Test-Path (Join-Path $root 'tests')) {
    & $vpy -m pytest -q --basetemp (Join-Path $env:LOCALAPPDATA 'dsd-bridge\pytest_tmp') tests/test_review_fixes.py 2>&1 | Select-Object -Last 2
}
Write-Host ''
Write-Host "준비 완료. 이후 명령은 이 파이썬으로 실행하세요: $vpy"
Write-Host '예) .\dsd-bridge.cmd export "감사보고서.dsd" -o "감사보고서.xlsx"   (dsd-bridge.cmd 가 위 파이썬을 사용합니다)'
Write-Host '팁: 이 폴더에 dsd-bridge.cmd 가 있으면  .\dsd-bridge.cmd export ...  처럼 짧게 쓸 수 있습니다.'
if (Test-Path (Join-Path $root 'tests')) { Write-Host '픽스처(Client .dsd)가 있으면 $env:DSD_BRIDGE_FIXTURE_DIR 로 지정한 뒤 python -m pytest -q 로 전체 테스트를 돌릴 수 있습니다.' }
