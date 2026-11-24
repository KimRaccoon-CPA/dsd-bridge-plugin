# DSD브릿지 (dsd-bridge)

전기 감사보고서 `.dsd`(DART 편집기 파일)와 **검토가 끝난 당기 정산표(WTB)** 로 당기 감사보고서 `.dsd` 를 만드는 Claude Code 플러그인입니다.

- DSD 를 엑셀로 펼쳐 당기를 전기로 넘기고, 정산표 숫자를 계정에 맞춰 채운 뒤, 다시 DSD 로 되돌립니다. 고치지 않은 부분은 원본과 글자 하나까지 같습니다.
- 되돌리기 전에 **정산표 대사**(보고서 금액 ↔ 정산표)와 **합계 검증**(세로 합계·대차·자본변동표)을 하고, DART 파일 규칙(XSD) 검사도 거칩니다.
- 사전재무제표(감사전 재무제표) `.dsd`, 분·반기 검토보고서 `.dsd` 도 같은 방식으로 만듭니다.
- 마지막 검사(F9·파일형식 검사)와 등록·전송파일(DRT) 생성은 **사람이 DART 편집기에서** 합니다.

## 필요한 것
- Windows PC, Claude Code(Claude 구독), DART 편집기(`C:\DART`)
- 파이썬 3.13 이상(없으면 https://www.python.org/downloads/ 또는 `winget install Python.Python.3.13`)
- 공개 저장소라 GitHub 계정 없이 설치할 수 있습니다

## 설치 (한 번)
Claude Code 창에서 차례로 입력합니다.

```
/plugin marketplace add KimRaccoon-CPA/dsd-bridge-plugin
/plugin install dsd-bridge@dsd-bridge
```

Claude Code 를 다시 시작한 뒤 "DSD브릿지 처음 설정해줘" 라고 하면 Claude 가 전용 파이썬(`%LOCALAPPDATA%\dsd-bridge\.venv`)을 만들고 동작을 확인합니다.

## 쓰는 법
예) "전기 감사보고서 `…제10기.dsd` 랑 검토 끝난 정산표 `…11기 정산표.xlsx` 로 11기 감사보고서 dsd 만들어줘"

Claude 가 펼치기 → 이월 → 정산표로 채우기 → 대사·합계 검증 → 되돌리기 → 후처리 순서로 진행하고, 대사 결과 엑셀과 작업 메모를 함께 남깁니다. 끝나면 DART 편집기에서 할 일을 알려 줍니다.

## 고객 메모 (사용자마다 따로 쌓임)
반복 고객은 작업이 끝날 때 Claude 가 "고객 메모에 저장할까요?" 하고 묻습니다. 저장하면 다음 기수에 그 메모(정산표 시트 대응, 보고서 규칙, 지난번 고친 점)를 먼저 읽고 시작합니다.

- 위치: `%USERPROFILE%\.dsd-bridge\clients\<고객 이름>.md` (내 PC 에만 있고 공유되지 않음)
- 다른 폴더를 쓰려면 `%USERPROFILE%\.dsd-bridge\config.json` 에 `{"clients_dir": "D:\\내 고객메모"}`
- 묻지 않고 항상 저장하려면 같은 파일에 `"memo": "auto"`

## 업데이트
새 판이 나오면 Claude Code 에서:

```
/plugin marketplace update dsd-bridge
/plugin update dsd-bridge@dsd-bridge
```

다시 시작하면 적용됩니다. 고객 메모와 전용 파이썬은 업데이트해도 그대로입니다.

## 주의
- 고객 자료를 Claude 에 넣어도 되는지는 소속 회사 정책을 따르세요.
- 결과 `.dsd` 는 반드시 DART 편집기 검사와 사람 검토를 거친 뒤 제출하세요. 정산표 계정 짝 맞추기는 Claude 의 판단이라, 금액이 우연히 같은 잘못된 짝은 대사로도 걸러지지 않을 수 있습니다.
- 고객 파일은 고객 폴더에서만 다루고, 원본은 덮어쓰지 않고 새 이름으로 저장합니다.
- 작업 엑셀(펼친 `.xlsx`)의 숨김 시트에는 원본 `.dsd` 의 전체 경로(PC 사용자 이름·고객 폴더 이름 포함)가 기록됩니다. 작업 엑셀은 밖으로 보내지 마세요.

## 이용 조건
설치해 그대로 쓰는 것은 자유입니다. 고치거나 다시 배포하려면 미리 허락을 받아야 합니다(이 저장소 Issues 로 요청). 자세한 내용은 [LICENSE](LICENSE).
