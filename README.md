# DSD브릿지 (dsd-bridge)

전기 감사보고서 `.dsd`(DART 편집기 파일)와 **검토가 끝난 당기 정산표(WTB)** 로 당기 감사보고서 `.dsd` 를 만드는 Claude Code 플러그인입니다.

- DSD 를 엑셀로 펼쳐 당기를 전기로 넘기고, 정산표 숫자를 계정에 맞춰 채운 뒤, 다시 DSD 로 되돌립니다. 고치지 않은 부분은 원본과 글자 하나까지 같습니다.
- 되돌리기 전에 **정산표 대사**(보고서 금액 ↔ 정산표)와 **합계 검증**(세로 합계·대차·자본변동표)을 하고, DART 파일 규칙(XSD) 검사도 거칩니다.
- 사전재무제표(감사전 재무제표) `.dsd`, 분·반기 검토보고서 `.dsd` 도 같은 방식으로 만듭니다.
- 마지막 검사(F9·파일형식 검사)와 등록·전송파일(DRT) 생성은 **사람이 DART 편집기에서** 합니다.

## 필요한 것
- Windows PC, Claude 유료 구독(Pro 이상), DART 편집기(`C:\DART`)
- Claude Code 를 쓸 수 있는 곳 하나: **Claude 데스크톱 앱의 Code 탭** 또는 터미널의 Claude Code
- 파이썬 3.13 이상(없으면 https://www.python.org/downloads/ 또는 `winget install Python.Python.3.13`)

> **일반 채팅 화면에서는 동작하지 않습니다.** DSD브릿지는 내 PC 의 파일을 열고 PC 에 설치된 파이썬으로 프로그램을 실행합니다. Claude 앱이라면 반드시 **Code 탭**에서, 환경은 **Local**(내 PC)로 쓰세요. Code 탭에서도 클라우드 환경을 고르면 PC 파일을 쓸 수 없습니다.

## 설치 (한 번)

### Claude 데스크톱 앱을 쓰는 경우
아래 둘 중 편한 방법 하나로 설치합니다.

- **방법 1 (메뉴)**: 앱에서 **Customize → Plugins → Add → Add marketplace** 를 누르고 `KimRaccoon-CPA/dsd-bridge-plugin` 입력 → 목록에 나온 **dsd-bridge** 설치
- **방법 2 (Code 탭 입력창)**: **Code** 탭 입력창에 아래 한 줄 입력
  ```
  /plugin install dsd-bridge --marketplace KimRaccoon-CPA/dsd-bridge-plugin
  ```

설치한 뒤에는 Code 탭에서 **+ New** 로 새 세션을 엽니다.

### 터미널의 Claude Code 를 쓰는 경우
Claude Code 창에서 차례로 입력한 뒤 Claude Code 를 다시 시작합니다.

```
/plugin marketplace add KimRaccoon-CPA/dsd-bridge-plugin
/plugin install dsd-bridge@dsd-bridge
```

> 앱과 터미널은 같은 PC 라면 설정을 함께 씁니다. 한쪽에 설치하면 다른 쪽에서도 보입니다.

## 앱이라면 처음 한 번 설정
1. **Code** 탭 → 입력창 아래에서 **Local** 선택 → **Open folder…** 로 아무 작업 폴더나 엽니다
2. "**DSD브릿지 처음 설정해줘**" 라고 입력합니다
3. Claude 가 전용 파이썬(`%LOCALAPPDATA%\dsd-bridge\.venv`)을 만들고, `dsd-bridge 0.x.y` 버전이 나오는지 확인합니다. 명령 실행 허락 창이 뜨면 허용하세요

### 처음 설정이 막힐 때
| 증상 | 해결 |
|---|---|
| "python 을 찾을 수 없습니다" | 파이썬 3.13 이상을 설치합니다(위 '필요한 것'). 설치 때 **Add python.exe to PATH** 를 체크하고, 설치 후 Claude 를 다시 시작합니다 |
| 부품 설치(pip) 중 인터넷·인증서 오류 | 회사망 보안 때문일 수 있습니다. 다른 네트워크에서 처음 설정만 하거나 회사 IT 에 pip 사용 가능 여부를 확인하세요 |
| "XSD 검증: … 건너뜀" 또는 "XSD 폴더를 찾지 못했습니다" 라고 나옴 | DART 편집기가 `C:\DART` 에 없을 때입니다. DART 편집기를 설치하거나, 다른 곳에 있으면 그 위치를 Claude 에게 알려 주세요 |
| 플러그인을 설치했는데 Claude 가 DSD브릿지를 모름 | Code 탭에서 **+ New** 로 새 세션을 열거나 앱·Claude Code 를 다시 시작합니다. "DSD브릿지" 라는 말을 넣어 요청해 보세요 |

## 쓰는 법
1. **Code** 탭 → **Local** → **Open folder…** 로 그 Client의 기수 폴더를 엽니다
2. 예) "전기 감사보고서 `…제10기.dsd` 랑 검토 끝난 정산표 `…11기 정산표.xlsx` 로 11기 감사보고서 dsd 만들어줘"

Claude 가 펼치기 → 이월 → 정산표로 채우기 → 대사·합계 검증 → 되돌리기 → 점검 순서로 진행하고, 대사 결과 엑셀과 작업 메모를 함께 남깁니다. 끝나면 DART 편집기에서 할 일을 알려 줍니다.

검증에서 사람이 볼 것을 따로 모아 줍니다(0.4.0): 정산표에만 있고 보고서에 없는 계정, 금액은 맞지만 짝이 의심스러운 행(같은 금액 여러 개, 당기·전기 뒤바뀜), 전기와 한 칸도 다르지 않은 주석 표, 금액·날짜가 든 채 전기와 같은 주석 문단, 주석 번호와 참조 불일치, 남은 자리표시. Claude 가 이 목록을 하나씩 확인해 작업 메모에 처리 내용을 적습니다.

## Client 메모 (사용자마다 따로 쌓임)
반복 Client는 작업이 끝날 때 Claude 가 "Client 메모에 저장할까요?" 하고 묻습니다. 저장하면 다음 기수에 그 메모(정산표 시트 대응, 보고서 규칙, 지난번 고친 점)를 먼저 읽고 시작합니다.

- 위치: `%USERPROFILE%\.dsd-bridge\clients\<Client 이름>.md` (내 PC 에만 있고 공유되지 않음)
- 다른 폴더를 쓰려면 `%USERPROFILE%\.dsd-bridge\config.json` 에 `{"clients_dir": "D:\\내 Client메모"}`
- 묻지 않고 항상 저장하려면 같은 파일에 `"memo": "auto"`

## 업데이트
새 판이 나오면:

- **앱(메뉴로 설치한 경우)**: Customize → Plugins 에서 dsd-bridge 를 업데이트합니다
- **Code 탭 입력창 또는 터미널**:
  ```
  /plugin marketplace update dsd-bridge
  /plugin update dsd-bridge@dsd-bridge
  ```

새 세션을 열거나 다시 시작하면 적용됩니다. Client 메모와 전용 파이썬은 업데이트해도 그대로입니다.

## 주의
- 결과 `.dsd` 는 반드시 DART 편집기 검사와 사람 검토를 거친 뒤 제출하세요.
- 작업 엑셀(펼친 `.xlsx`)에는 Client 숫자가 그대로 들어 있습니다. 밖으로 보내지 마세요. 숨김 시트에는 원본 `.dsd` 의 파일 이름만 기록됩니다(0.3.0 부터. 그 전 판으로 만든 작업 엑셀에는 PC 사용자 이름·Client 폴더 이름이 든 전체 경로가 남아 있습니다).

## 이용 조건
설치해 그대로 쓰는 것은 자유입니다. 고치거나 다시 배포하려면 미리 허락을 받아야 합니다(이 저장소 Issues 로 요청). 자세한 내용은 [LICENSE](LICENSE).
