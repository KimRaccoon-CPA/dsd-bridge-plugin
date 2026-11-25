"""dsd-bridge 명령줄 진입점 (계획 §4.13, §5 단계 4).

- `dsd-bridge export <입력.dsd> [-o 출력.xlsx] [-v] [--split-notes]`
- `dsd-bridge import <입력.xlsx> --dsd <원본.dsd> -o <출력.dsd> [-v] [--force-hash]
  [--allow-bulk-delete] [--allow-lossy-mixed] [--no-meta-size] [--acode "라벨=코드" …]`
  — 0.3.0: 서식 조각이 섞인 회색 칸·문단은 서식을 두고 글자만 바꾸고, 날짜 칸(기간·커뮤니케이션 일자)의
  글을 고치면 숨은 날짜 속성값도 맞추며, `--acode` 로 새로 끼운 행의 계정코드를 지정한다.
- `dsd-bridge check <입력.dsd|입력.xlsx> [--json 경로] [--tolerance N] [--xref] [--xref-min N]
  [--settlement 정산표.xlsx] [--settlement-map 경로.json] [--report 경로.xlsx] [-v]`
  — 합계검증 (1.2 WP4). 불일치가 있으면 종료 코드 8 이며, 오류가 아니라 **검사 결과**다.
  `--xref` 는 주석 ↔ 재무제표 상호대사(1.2 WP7)로, 참고용이라 종료 코드에 영향을 주지 않는다.
  `--settlement` 은 정산표 상호대사(1.3)로, 금액 불일치가 있으면 종료 코드 8 이다.
- `dsd-bridge describe <입력.xlsx|입력.dsd> [--json 경로] [--sheet 이름]` — 구조 지도 (1.3).
- 0.6.0: `check … [--decisions 판정.json] [--footing 풋팅.xlsx]`, `diff … [--decisions 판정.json]`,
  `accept 판정.json 번호 --reason 사유`(판정 기록), `compare 전 후 [--report 비교.xlsx]`(두 판 비교),
  `proof 파일 [--since 전] [--json 목록]`(맞춤법 점검 도우미).
- `dsd-bridge template <틀.dsd> [--scrub -o 새틀.dsd [--replace 옛=새 …]] [--check 결과 [--old 문자열 …]]`
  — 전기 .dsd 가 없어 틀 문서로 시작할 때 틀에 든 식별값 보기·비우기·남은 내용 검사 (0.5.0). 식별값 잔존이 있으면 종료 코드 8.
- `dsd-bridge diff <전.xlsx> <후.xlsx> [--rollforward] [--json 경로] [-v]` — 통합문서 비교 (1.3).
- 0.7.0: `from-dart <접수번호|원문.xml> -o 틀.dsd`(공시 원문으로 틀 만들기, `--search 회사명`),
  `template 틀.dsd --blank -o 빈틀.dsd`(구조만 남긴 빈 틀), `clone 원본.dsd -o 새.dsd --table ID|--note N`
  (표·주석 복제), `round 금액… --unit 천원`(합계 맞춤 반올림), `replace 작업.xlsx --set 옛=새 [--apply]`
  (기본정보 한꺼번에 바꾸기). import 는 결과를 다시 펼쳐 엑셀 값과 칸마다 맞춰 본다(재확인).

모든 `DsdBridgeError` 는 `message_ko` 를 표준 오류로 출력하고 그 `exit_code` 로 끝난다.
사용자 대면 플래그는 5개: `--force-hash`, `--allow-bulk-delete`, `--allow-lossy-mixed`,
`--no-meta-size`, `--acode`(0.3.0). `--allow-reorder` 는 없다(rev4 — 지킬 수 없는 약속). `--no-merge` 는 테스트 전용이라
`--help` 에서 숨긴다.
"""

from __future__ import annotations

import argparse
import sys

from dsd_bridge import __version__
from dsd_bridge.errors import DsdBridgeError


def build_parser() -> argparse.ArgumentParser:
    from dsd_bridge.check import DEFAULT_TOLERANCE

    parser = argparse.ArgumentParser(
        prog="dsd-bridge",
        description="DSD(DART 제출용 .dsd) 파일과 Excel 통합문서를 왕복 변환합니다.",
        epilog="자세한 사용법과 주의사항은 README.md 를 참고하세요.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"dsd-bridge {__version__}",
        help="버전을 출력하고 종료합니다.",
    )
    sub = parser.add_subparsers(dest="command", metavar="명령")

    p_export = sub.add_parser(
        "export",
        help=".dsd 를 편집용 Excel 통합문서(.xlsx)로 내보냅니다.",
        description=".dsd 를 편집용 Excel 통합문서(.xlsx)로 내보냅니다.",
    )
    p_export.add_argument("input", metavar="입력.dsd", help="원본 .dsd 파일 경로")
    p_export.add_argument(
        "-o",
        "--output",
        metavar="출력.xlsx",
        help="저장할 .xlsx 경로 (생략하면 입력과 같은 폴더에 같은 이름으로 저장)",
    )
    p_export.add_argument(
        "-v", "--verbose", action="store_true", help="진행 상황을 자세히 출력합니다."
    )
    p_export.add_argument(
        "--split-notes",
        action="store_true",
        help="주석 번호 제목(`1. 회사의 개요` 처럼 `번호.` 로 시작하는 줄)마다 시트를 나눕니다.",
    )
    # 테스트 전용 — 도움말에서 숨긴다 (계획 §4.13 각주).
    p_export.add_argument("--no-merge", action="store_true", help=argparse.SUPPRESS)
    p_export.set_defaults(func=cmd_export)

    p_import = sub.add_parser(
        "import",
        help="편집한 .xlsx 를 원본 .dsd 에 되돌려 새 .dsd 를 만듭니다.",
        description="편집한 .xlsx 를 원본 .dsd 에 되돌려 새 .dsd 를 만듭니다.",
    )
    p_import.add_argument("input", metavar="입력.xlsx", help="편집한 .xlsx 파일 경로")
    p_import.add_argument(
        "--dsd", metavar="원본.dsd", required=True, help="기준이 되는 원본 .dsd 경로 (export 에 썼던 파일)"
    )
    p_import.add_argument(
        "-o", "--output", metavar="출력.dsd", help="저장할 .dsd 경로 (생략하면 <입력명>.out.dsd)"
    )
    p_import.add_argument(
        "-v", "--verbose", action="store_true", help="진행 상황을 자세히 출력합니다."
    )
    p_import.add_argument(
        "--force-hash",
        action="store_true",
        help="원본 해시가 달라도 크기·요소 수·태그 순서가 같으면 진행합니다 (경고 출력).",
    )
    p_import.add_argument(
        "--allow-bulk-delete",
        action="store_true",
        help="한 섹션의 행이 50%% 넘게 사라져도 진행합니다.",
    )
    p_import.add_argument(
        "--allow-lossy-mixed",
        action="store_true",
        help="서식 조각(SPAN 등)을 포함한 문단·셀의 편집을 허용합니다 (서식은 사라집니다).",
    )
    p_import.add_argument(
        "--no-meta-size",
        action="store_true",
        help="meta.xml 의 FILE-INFO size 를 다시 계산하지 않습니다.",
    )
    p_import.add_argument(
        "--acode",
        metavar="라벨=코드",
        action="append",
        default=[],
        help=("새로 끼운 표 행의 계정코드(ACODE)를 지정합니다. 복제한 행은 원본 행의 코드를 물려받습니다. "
              "여러 번 쓸 수 있습니다. 예: --acode \"장기금융상품=11211100010000\" "
              "(표준계정코드표에 없는 신설 계정은 99999999999999)"),
    )
    p_import.set_defaults(func=cmd_import)

    p_check = sub.add_parser(
        "check",
        help="재무제표 표의 합계(세로 합계·총계 등식·대차평형·자본변동표)를 검증합니다.",
        description=(
            ".dsd 또는 내보낸 .xlsx 의 재무제표 표에서 DART 속성(ALEVEL/ADELIM/ACODE)을 읽어\n"
            "세로 합계·총계 등식·대차평형을 검증하고, 자본변동표는 가로(구성요소 = 총계)와\n"
            "세로(기초 + 변동 = 기말)·이월(전기말 = 당기초)을 봅니다.\n"
            "불일치가 있으면 종료 코드 8 입니다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_check.add_argument("input", metavar="입력.dsd|입력.xlsx", help="검사할 .dsd 또는 .xlsx 경로")
    p_check.add_argument(
        "--json", metavar="경로", dest="json_path", help="검사 결과를 JSON 으로 저장할 경로"
    )
    p_check.add_argument(
        "--tolerance",
        metavar="N",
        type=int,
        default=DEFAULT_TOLERANCE,
        help=f"단수차로 넘길 최대 차이 (기본 {DEFAULT_TOLERANCE}, 0 이면 정확히 일치해야 합니다)",
    )
    p_check.add_argument(
        "--xref",
        action="store_true",
        help="재무제표 금액을 주석 표에서 찾아 주는 상호대사를 함께 합니다 (.dsd 만, 참고용).",
    )
    p_check.add_argument(
        "--xref-min",
        metavar="N",
        type=int,
        default=0,
        dest="xref_min",
        help="상호대사에서 건너뛸 금액의 하한 (기본 0 — 0원만 건너뜁니다)",
    )
    p_check.add_argument(
        "--settlement",
        metavar="정산표.xlsx",
        help="정산표와 재무제표·주석 금액을 상호대사합니다 (금액 불일치가 있으면 종료 코드 8).",
    )
    p_check.add_argument(
        "--settlement-map",
        metavar="경로.json",
        dest="settlement_map",
        help="정산표 시트·열 대응을 적은 JSON (생략하면 공시용/BS/IS/CE/CF/FN 기본 대응).",
    )
    p_check.add_argument(
        "--report",
        metavar="경로.xlsx",
        dest="report_path",
        help="상호대사 결과를 색으로 표시한 .xlsx 로 저장합니다.",
    )
    p_check.add_argument(
        "--decisions", metavar="판정.json",
        help="판정 기록 파일. 기록된 경보는 '확인 처리' 로 따로 모으고, 나머지 경보에 번호 [#…] 를 붙입니다.",
    )
    p_check.add_argument(
        "--footing", metavar="풋팅.xlsx",
        help="보고서 모양 그대로 판정 색·틱마크 메모를 단 검토용 풋팅 사본과 한 장 요약을 씁니다.",
    )
    p_check.add_argument(
        "-v", "--verbose", action="store_true", help="OK 인 검사와 제외한 표까지 모두 출력합니다."
    )
    p_check.set_defaults(func=cmd_check)

    p_accept = sub.add_parser(
        "accept",
        help="검사 경보를 판정 기록에 '확인함' 또는 '대상 아님' 으로 적습니다 (번호는 출력의 [#…]).",
    )
    p_accept.add_argument("decisions", metavar="판정.json", help="판정 기록 파일 (없으면 새로 만듭니다)")
    p_accept.add_argument("ids", metavar="번호", nargs="*", help="출력의 [#…] 안 8자리 번호 (여러 개 가능)")
    p_accept.add_argument("--reason", default="", help="확인한 사유 (필수)")
    p_accept.add_argument("--not-applicable", action="store_true", dest="not_applicable",
                          help="상태를 '대상 아님' 으로 적습니다 (기본 '확인함').")
    p_accept.add_argument("--what", default="", help="무엇에 대한 판정인지 메모 (선택)")
    p_accept.add_argument("--remove", action="store_true", help="그 번호의 기록을 지웁니다.")
    p_accept.add_argument("--list", action="store_true", dest="list_", help="기록을 모두 보여 줍니다.")
    p_accept.set_defaults(func=cmd_accept)

    p_compare = sub.add_parser(
        "compare",
        help="같은 보고서의 두 판(.dsd/.xlsx)을 견줍니다 — 문단 문구 추가·삭제·수정과 표 변경.",
    )
    p_compare.add_argument("old", metavar="전판", help="이전 판 .dsd 또는 .xlsx")
    p_compare.add_argument("new", metavar="후판", help="이후 판 .dsd 또는 .xlsx")
    p_compare.add_argument("--report", metavar="비교.xlsx", dest="report_path",
                           help="취소선·밑줄로 표시한 비교 보고서 엑셀을 씁니다.")
    p_compare.add_argument("--json", metavar="경로", dest="json_path", help="결과를 JSON 으로 저장")
    p_compare.add_argument("-v", "--verbose", action="store_true", help="줄인 목록을 모두 출력합니다.")
    p_compare.set_defaults(func=cmd_compare)

    p_proof = sub.add_parser(
        "proof",
        help="맞춤법·띄어쓰기 점검 도우미 — 기계 점검 후보와 점검할 글 목록(--json).",
    )
    p_proof.add_argument("input", metavar="파일", help=".dsd 또는 .xlsx")
    p_proof.add_argument("--since", metavar="전판", help="이 판과 달라진 글만 모읍니다.")
    p_proof.add_argument("--json", metavar="경로", dest="json_path", help="점검할 글 목록을 JSON 으로 저장")
    p_proof.add_argument("-v", "--verbose", action="store_true", help="점검할 글도 모두 출력합니다.")
    p_proof.set_defaults(func=cmd_proof)

    p_describe = sub.add_parser(
        "describe",
        help="통합문서·.dsd 의 구조 지도(시트·표·열 역할·행·기간 표기 셀)를 출력합니다.",
        description=(
            "AI 가 통합문서를 안전하게 편집하도록 구조를 알려 줍니다.\n"
            "시트마다 표를 나열하고, 표준 재무제표 표는 열 역할(과목·당기·전기)을 Excel 열\n"
            "문자로 보여 주며, 롤포워드 때 갱신할 '기간 표기 셀' 을 따로 모아 줍니다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_describe.add_argument("input", metavar="입력.xlsx|입력.dsd", help="읽을 .xlsx 또는 .dsd 경로")
    p_describe.add_argument(
        "--json", metavar="경로", dest="json_path", help="같은 내용을 JSON 으로 저장할 경로"
    )
    p_describe.add_argument("--sheet", metavar="이름", help="이 시트만 출력합니다.")
    p_describe.set_defaults(func=cmd_describe)

    p_template = sub.add_parser(
        "template",
        help="틀 문서로 시작할 때: 틀에 든 식별값 보기 / 숨은 식별값 비우기(--scrub) / 남은 내용 검사(--check).",
        description=(
            "전기 .dsd 가 없어 틀 문서로 시작할 때의 안전장치입니다. 가장 안전한 틀은 DART 편집기 서식으로 만든 빈 문서이고,\n"
            "이전 내용이 들어 있는 틀이라면 그 내용이 결과에 남지 않았는지 반드시 검사합니다.\n"
            "  template 틀.dsd                      틀에 든 이름·숨은 고유번호·요약값·날짜가 어디 있는지 봅니다.\n"
            "  template 틀.dsd --scrub -o 새틀.dsd   숨은 식별값을 비운 새 틀을 만듭니다(--replace 옛=새 로 글자 치환).\n"
            "  template 틀.dsd --blank -o 빈틀.dsd   구조만 남기고 금액·주석 문단·이름을 비운 빈 틀을 만듭니다.\n"
            "  template 틀.dsd --check 결과.dsd      결과에 틀의 이전 내용이 남았는지 봅니다(식별값 잔존이 있으면 종료 코드 8)."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_template.add_argument("input", metavar="틀.dsd", help="틀로 쓰는 .dsd")
    p_template.add_argument("--scrub", action="store_true",
                            help="숨은 식별값(회사 이름·고유번호·법인등록번호·요약값)을 비운 새 틀을 -o 에 씁니다.")
    p_template.add_argument("--blank", action="store_true",
                            help="구조만 남긴 빈 틀을 -o 에 씁니다: 식별값 비우기 + 표 금액 지우기 + 주석 문단 비우기"
                                 "(번호 제목 줄은 남김) + 이름을 ○회사명○ 로.")
    p_template.add_argument("-o", "--output", metavar="새틀.dsd", help="--scrub·--blank 의 출력 경로")
    p_template.add_argument("--replace", metavar="옛=새", action="append", default=[],
                            help="--scrub 때 본문·표지 전체에서 글자를 바꿉니다(여러 번 가능). 예: --replace \"옛회사=새회사\"")
    p_template.add_argument("--check", metavar="결과.dsd|작업.xlsx", dest="check_path",
                            help="이 파일에 틀의 식별값·금액·글이 남았는지 검사합니다.")
    p_template.add_argument("--old", metavar="문자열", action="append", default=[],
                            help="--check 때 더 찾을 틀의 이전 문자열(대표이사 이름, 주소 등, 여러 번 가능)")
    p_template.add_argument("--json", metavar="경로", dest="json_path", help="--check 결과를 JSON 으로 저장할 경로")
    p_template.add_argument("-v", "--verbose", action="store_true", help="틀과 같은 일반 문단까지 모두 출력합니다.")
    p_template.set_defaults(func=cmd_template)

    p_fromdart = sub.add_parser(
        "from-dart",
        help="전기 .dsd 가 없을 때: DART 에 공시된 감사보고서 원문(접수번호)으로 틀 .dsd 를 만듭니다.",
        description=(
            "OpenDART 공시서류원본파일(접수번호 14자리)이나 내려받은 원문 .xml/.zip 으로 .dsd 틀을 만듭니다.\n"
            "공시 원문에는 문단 안 줄바꿈이 없어 복원되지 않고, 회사 이름·요약값은 비어 있습니다.\n"
            "인증키: --key → 환경변수 DART_API_KEY → ~/.dsd-bridge/config.json 의 dart_api_key.\n"
            "  from-dart --search 회사명               최근 외부감사 관련 공시의 접수번호를 찾습니다.\n"
            "  from-dart 20260325001416 -o 전기틀.dsd  그 공시 원문으로 틀을 만듭니다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_fromdart.add_argument("source", metavar="접수번호|원문", nargs="?", help="접수번호 14자리 또는 원문 .xml/.zip")
    p_fromdart.add_argument("-o", "--output", metavar="틀.dsd", help="만들 .dsd 경로")
    p_fromdart.add_argument("--search", metavar="회사명", help="회사 이름(일부)으로 접수번호를 찾습니다.")
    p_fromdart.add_argument("--years", type=int, default=3, help="--search 로 볼 기간(년, 기본 3)")
    p_fromdart.add_argument("--doc", metavar="이름", default="", help="원문 묶음에 파일이 여럿일 때 고를 파일 이름 일부")
    p_fromdart.add_argument("--key", metavar="인증키", help="OpenDART 인증키 (화면에 출력하지 않습니다)")
    p_fromdart.set_defaults(func=cmd_from_dart)

    p_clone = sub.add_parser(
        "clone",
        help="표나 주석 하나를 그대로 복제해 바로 뒤에 붙입니다(새 표·주석을 만들 때).",
        description=(
            "DSD브릿지는 표·주석을 새로 만들지 못하므로, 비슷한 것을 바이트 그대로 복제해 서식을 물려받습니다.\n"
            "  clone 원본.dsd -o 새.dsd --table E00946.46fe [--count 2] [--with-prev 1]\n"
            "  clone 원본.dsd -o 새.dsd --sheet 주석 --index 3        시트의 3번째 표\n"
            "  clone 원본.dsd -o 새.dsd --note 4 [--as 5]             주석 4 를 복제해 주석 5 로\n"
            "복제한 뒤 export 해 내용을 바꾸고, 주석을 늘렸으면 뒤 번호와 (주석N) 참조를 고칩니다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_clone.add_argument("input", metavar="원본.dsd", help="복제할 표·주석이 든 .dsd")
    p_clone.add_argument("-o", "--output", metavar="새.dsd", required=True, help="결과 .dsd (원본과 달라야 함)")
    p_clone.add_argument("--table", metavar="표ID", default="", help="describe 에 나온 표 ID")
    p_clone.add_argument("--sheet", metavar="시트", default="", help="--index 와 함께: 시트 이름")
    p_clone.add_argument("--index", type=int, default=0, help="--sheet 와 함께: 그 시트의 몇 번째 표(1부터)")
    p_clone.add_argument("--count", type=int, default=1, help="표를 몇 개 복제할지 (기본 1)")
    p_clone.add_argument("--with-prev", type=int, default=0, metavar="K",
                         help="표 앞의 형제 요소 K개(제목 문단·단위 표)를 함께 복제")
    p_clone.add_argument("--note", type=int, help="복제할 주석 번호")
    p_clone.add_argument("--as", dest="as_no", type=int, help="복제본 주석 번호 (기본 N+1)")
    p_clone.set_defaults(func=cmd_clone)

    p_round = sub.add_parser(
        "round",
        help="원 단위 금액을 천원·백만원으로 반올림하되 합계가 맞도록 고칠 칸을 골라 줍니다.",
    )
    p_round.add_argument("amounts", metavar="금액", nargs="*", help="원 단위 금액들 (쉼표 가능)")
    p_round.add_argument("--unit", default="천원", choices=["원", "천원", "백만원"], help="바꿀 단위 (기본 천원)")
    p_round.add_argument("--total", metavar="합계", help="맞출 합계(원 단위). 없으면 금액들의 합")
    p_round.add_argument("--from", dest="from_range", metavar="파일.xlsx!시트!C5:C12",
                         help="금액을 엑셀 범위에서 읽습니다 (예: 정산표.xlsx!BS!D5:D20)")
    p_round.set_defaults(func=cmd_round)

    p_replace = sub.add_parser(
        "replace",
        help="작업 엑셀의 보이는 칸에서 회사 이름·대표이사·주소 같은 글자를 한꺼번에 바꿉니다(미리보기 기본).",
    )
    p_replace.add_argument("input", metavar="작업.xlsx", help="export 한 작업 엑셀")
    p_replace.add_argument("--set", dest="pairs", metavar="옛=새", action="append", default=[],
                           help="바꿀 글자 (여러 번 가능). 예: --set \"홍길동=김철수\"")
    p_replace.add_argument("--apply", action="store_true", help="미리보기 대신 실제로 저장합니다.")
    p_replace.set_defaults(func=cmd_replace)

    p_diff = sub.add_parser(
        "diff",
        help="통합문서 둘을 견줍니다 (같은 문서의 편집 전후, 또는 전기·당기 보고서).",
        description=(
            "두 통합문서(또는 .dsd)의 표 행을 짝지어 바뀐 셀을 보여 줍니다.\n"
            "같은 원본에서 내보낸 둘이면 행 _ID 로, 다른 문서 둘이면 정규화한 계정과목으로\n"
            "짝짓고 짝이 없는 행은 삽입·삭제로 적습니다.\n"
            "--rollforward 를 주면 새 문서의 전기 금액이 옛 문서의 당기 금액과 같은지\n"
            "(이월) 검증하고, 당기 미기입·기간 표기 미갱신도 함께 알립니다.\n"
            "참고용이므로 종료 코드는 언제나 0 입니다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_diff.add_argument("old", metavar="전.xlsx", help="이전(전기) 통합문서 또는 .dsd")
    p_diff.add_argument("new", metavar="후.xlsx", help="이후(당기) 통합문서 또는 .dsd")
    p_diff.add_argument(
        "--rollforward", action="store_true", help="이월(전기 = 옛 당기) 검증을 함께 합니다."
    )
    p_diff.add_argument(
        "--json", metavar="경로", dest="json_path", help="비교 결과를 JSON 으로 저장할 경로"
    )
    p_diff.add_argument(
        "--decisions", metavar="판정.json", help="판정 기록 파일 (전기와 같은 표·문단 확인 처리)",
    )
    p_diff.add_argument(
        "-v", "--verbose", action="store_true", help="줄인 목록과 이월 OK 까지 모두 출력합니다."
    )
    p_diff.set_defaults(func=cmd_diff)

    return parser


def _log(verbose: bool):
    if verbose:
        return lambda *parts: print(*parts)
    return lambda *parts: None


def cmd_export(args: argparse.Namespace) -> int:
    from dsd_bridge.export_xlsx import export

    report = export(args.input, args.output, merge=not args.no_merge,
                    split_notes=args.split_notes, log=_log(args.verbose))
    print(f"내보내기 완료: {report.summary_ko()}")
    for w in report.warnings:
        print(w if w.startswith("경고") else f"경고: {w}")
    for n in report.notices:
        print(n if n.startswith("안내") else f"안내: {n}")
    return 0


def _parse_acodes(items) -> dict:
    from dsd_bridge.errors import InputError

    out: dict = {}
    for item in items or []:
        label, sep, code = str(item).rpartition("=")
        if not sep or not label.strip() or not code.strip():
            raise InputError(f"오류: --acode 는 \"라벨=코드\" 형식이어야 합니다: {item}")
        out[label.strip()] = code.strip()
    return out


def cmd_import(args: argparse.Namespace) -> int:
    from dsd_bridge.import_xlsx import import_

    report = import_(
        args.input,
        args.dsd,
        args.output,
        force_hash=args.force_hash,
        allow_bulk_delete=args.allow_bulk_delete,
        allow_lossy_mixed=args.allow_lossy_mixed,
        no_meta_size=args.no_meta_size,
        log=_log(args.verbose),
        acodes=_parse_acodes(args.acode),
    )
    print(report.summary_ko())
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    from dsd_bridge.check import check
    from dsd_bridge.errors import InputError

    settlement = getattr(args, "settlement", None)
    if getattr(args, "settlement_map", None) and not settlement:
        raise InputError("오류: --settlement-map 은 --settlement 과 함께 써야 합니다.")
    if getattr(args, "report_path", None) and not settlement:
        raise InputError("오류: --report 는 --settlement 과 함께 써야 합니다.")

    decisions = _decisions(getattr(args, "decisions", None))
    report = check(args.input, tolerance=args.tolerance,
                   xref=args.xref, xref_min=args.xref_min, decisions=decisions)
    if settlement:
        from dsd_bridge.settlement import check_settlement

        report.settlement = check_settlement(args.input, settlement, args.settlement_map, decisions)
        if args.report_path:
            report.settlement.write_xlsx(args.report_path)
    if args.json_path:
        report.write_json(args.json_path)
    for line in report.lines_ko(verbose=args.verbose):
        print(line)
    if settlement and args.report_path:
        print(f"안내: 상호대사 결과를 저장했습니다 — {args.report_path}")
    if getattr(args, "footing", None):
        from dsd_bridge.footing import write_footing

        write_footing(args.input, args.footing, report)
        print(f"안내: 풋팅 사본(검토용)을 저장했습니다 — {args.footing}")
    return report.exit_code


def _decisions(path):
    if not path:
        return None
    from dsd_bridge.decisions import load

    return load(path)


def cmd_accept(args: argparse.Namespace) -> int:
    from dsd_bridge.decisions import CONFIRMED, NOT_APPLICABLE, load
    from dsd_bridge.errors import InputError

    d = load(args.decisions)
    if args.list_:
        for i, v in sorted(d.items.items()):
            print(f"#{i} {v.get('status')} | {v.get('reason')} | {v.get('what')} | {v.get('date')}")
        print(f"판정 기록 {len(d.items)}건 — {d.path}")
        return 0
    if not args.ids:
        raise InputError("오류: 번호를 하나 이상 적어 주세요 (검사 출력의 [#…] 안 8자리).")
    for i in args.ids:
        i = i.lstrip("#").strip("[]")
        if args.remove:
            print(f"{'지움' if d.remove(i) else '없음'}: #{i}")
        else:
            d.add(i, args.reason, NOT_APPLICABLE if args.not_applicable else CONFIRMED, args.what)
            print(f"기록: #{i} {d.reason_ko(i)}")
    d.save()
    print(f"판정 기록 {len(d.items)}건 — {d.path}")
    return 0


def cmd_compare(args: argparse.Namespace) -> int:
    from dsd_bridge.compare import compare

    report = compare(args.old, args.new)
    if args.report_path:
        report.write_xlsx(args.report_path)
    if args.json_path:
        report.write_json(args.json_path)
    for line in report.lines_ko(verbose=args.verbose):
        print(line)
    if args.report_path:
        print(f"안내: 비교 보고서를 저장했습니다 — {args.report_path}")
    return report.exit_code


def cmd_proof(args: argparse.Namespace) -> int:
    from dsd_bridge.proof import proof

    report = proof(args.input, args.since)
    if args.json_path:
        report.write_json(args.json_path)
    for line in report.lines_ko(verbose=args.verbose):
        print(line)
    return report.exit_code


def cmd_template(args: argparse.Namespace) -> int:
    from dsd_bridge.errors import InputError
    from dsd_bridge.template import check_against, identity, scrub

    if sum(bool(x) for x in (args.scrub, args.check_path, getattr(args, "blank", False))) > 1:
        raise InputError("오류: --scrub, --blank, --check 는 따로 실행하세요.")
    if getattr(args, "blank", False):
        from dsd_bridge.template import blank

        if args.replace:
            raise InputError("오류: --blank 는 이름을 ○회사명○ 로 바꿉니다. 다른 글자를 바꾸려면 --scrub --replace 를 쓰세요.")
        for line in blank(args.input, args.output).lines_ko():
            print(line)
        return 0
    if args.scrub:
        pairs = []
        for item in args.replace:
            old, sep, new = str(item).partition("=")
            if not sep or not old:
                raise InputError(f"오류: --replace 는 \"옛=새\" 형식이어야 합니다: {item}")
            pairs.append((old, new))
        for line in scrub(args.input, args.output, pairs).lines_ko():
            print(line)
        return 0
    if args.replace or args.output:
        raise InputError("오류: --replace 와 -o 는 --scrub(또는 -o 는 --blank)과 함께 씁니다.")
    if args.check_path:
        report = check_against(args.input, args.check_path, args.old)
        if args.json_path:
            report.write_json(args.json_path)
        for line in report.lines_ko(verbose=args.verbose):
            print(line)
        return report.exit_code
    for line in identity(args.input).lines_ko():
        print(line)
    return 0


def _pairs(items, flag: str) -> list:
    from dsd_bridge.errors import InputError

    out = []
    for item in items:
        old, sep, new = str(item).partition("=")
        if not sep or not old:
            raise InputError(f"오류: {flag} 는 \"옛=새\" 형식이어야 합니다: {item}")
        out.append((old, new))
    return out


def cmd_from_dart(args: argparse.Namespace) -> int:
    from dsd_bridge.errors import InputError
    from dsd_bridge.fromdart import from_dart, search

    if args.search:
        rows = search(args.search, key=args.key, years=args.years)
        if not rows:
            print(f"'{args.search}' 이(가) 이름에 든 회사의 외부감사 관련 공시를 찾지 못했습니다.")
            return 0
        print("회사 | 접수번호 | 보고서 | 접수일")
        for corp, no, name, day in rows:
            print(f"{corp} | {no} | {name} | {day}")
        print("다음: dsd-bridge from-dart <접수번호> -o 전기틀.dsd")
        return 0
    if not args.source or not args.output:
        raise InputError("오류: 접수번호(또는 원문 파일)와 -o 출력 경로가 필요합니다. 접수번호는 --search 로 찾습니다.")
    for line in from_dart(args.source, args.output, key=args.key, doc=args.doc).lines_ko():
        print(line)
    return 0


def cmd_clone(args: argparse.Namespace) -> int:
    from dsd_bridge.clone import clone_note, clone_table
    from dsd_bridge.errors import InputError

    if args.note is not None:
        if args.table or args.sheet:
            raise InputError("오류: --note 와 --table/--sheet 는 함께 쓰지 않습니다.")
        report = clone_note(args.input, args.output, args.note, args.as_no)
    elif args.table or (args.sheet and args.index):
        report = clone_table(args.input, args.output, table=args.table, sheet=args.sheet, index_no=args.index,
                             count=args.count, with_prev=args.with_prev)
    else:
        raise InputError("오류: --table <표ID>, --sheet <시트> --index <n>, --note <번호> 중 하나를 주세요.")
    for line in report.lines_ko():
        print(line)
    return 0


def cmd_round(args: argparse.Namespace) -> int:
    from dsd_bridge.assist import read_range, round_to_unit
    from dsd_bridge.errors import InputError

    amounts = list(args.amounts)
    if args.from_range:
        path, sep, ref = args.from_range.partition("!")
        if not sep:
            raise InputError("오류: --from 은 '파일.xlsx!시트!C5:C12' 꼴로 적어 주세요.")
        amounts += read_range(path, ref)
    for line in round_to_unit(amounts, args.unit, args.total).lines_ko():
        print(line)
    return 0


def cmd_replace(args: argparse.Namespace) -> int:
    from dsd_bridge.assist import replace_text

    report = replace_text(args.input, _pairs(args.pairs, "--set"), apply=args.apply)
    for line in report.lines_ko():
        print(line)
    return 0


def cmd_describe(args: argparse.Namespace) -> int:
    from dsd_bridge.describe import describe

    report = describe(args.input, sheet=args.sheet or "")
    if args.json_path:
        report.write_json(args.json_path)
    for line in report.lines_ko():
        print(line)
    return 0


def cmd_diff(args: argparse.Namespace) -> int:
    from dsd_bridge.diff import diff

    report = diff(args.old, args.new, rollforward=args.rollforward,
                  decisions=_decisions(getattr(args, "decisions", None)))
    if args.json_path:
        report.write_json(args.json_path)
    for line in report.lines_ko(verbose=args.verbose):
        print(line)
    return report.exit_code


def _utf8_console() -> None:
    """한국어 Windows 에서 출력을 파일/파이프로 돌리면 stdout 이 cp949 가 되어 `—` 같은 글자에서
    UnicodeEncodeError 가 난다. 출력은 항상 UTF-8 로, 인코딩 불가 문자는 대체한다."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):   # 캡처된 스트림 등
            pass


def main(argv: list[str] | None = None) -> int:
    _utf8_console()
    parser = build_parser()
    args = parser.parse_args(argv)
    if getattr(args, "func", None) is None:
        parser.print_help()
        return 1
    try:
        return args.func(args)
    except DsdBridgeError as e:
        msg = e.message_ko
        if not msg.startswith("오류"):
            msg = "오류: " + msg
        print(msg, file=sys.stderr)
        return e.exit_code
    except ModuleNotFoundError as e:  # 단계 5 이전: import 미구현
        if "import_xlsx" in str(e):
            print("import: 아직 구현되지 않았습니다.", file=sys.stderr)
            return 1
        raise
