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
- `dsd-bridge template <틀.dsd> [--scrub -o 새틀.dsd [--replace 옛=새 …]] [--check 결과 [--old 문자열 …]]`
  — 다른 회사 .dsd 를 틀로 쓸 때의 식별값 보기·비우기·잔존 검사 (0.5.0). 식별값 잔존이 있으면 종료 코드 8.
- `dsd-bridge diff <전.xlsx> <후.xlsx> [--rollforward] [--json 경로] [-v]` — 통합문서 비교 (1.3).

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
        "-v", "--verbose", action="store_true", help="OK 인 검사와 제외한 표까지 모두 출력합니다."
    )
    p_check.set_defaults(func=cmd_check)

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
        help="다른 회사 .dsd 를 틀로 쓸 때: 식별값 보기 / 숨은 식별값 비우기(--scrub) / 틀 잔존 검사(--check).",
        description=(
            "전기 .dsd 가 없어 다른 회사의 .dsd 를 틀로 쓸 때의 안전장치입니다.\n"
            "  template 틀.dsd                      틀의 회사 이름·숨은 고유번호·요약값·날짜가 어디 있는지 봅니다.\n"
            "  template 틀.dsd --scrub -o 새틀.dsd   숨은 식별값을 비운 새 틀을 만듭니다(--replace 옛=새 로 글자 치환).\n"
            "  template 틀.dsd --check 결과.dsd      결과에 틀 회사 것이 남았는지 봅니다(식별값 잔존이 있으면 종료 코드 8)."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_template.add_argument("input", metavar="틀.dsd", help="틀로 쓰는 다른 회사의 .dsd")
    p_template.add_argument("--scrub", action="store_true",
                            help="숨은 식별값(회사 이름·고유번호·법인등록번호·요약값)을 비운 새 틀을 -o 에 씁니다.")
    p_template.add_argument("-o", "--output", metavar="새틀.dsd", help="--scrub 의 출력 경로")
    p_template.add_argument("--replace", metavar="옛=새", action="append", default=[],
                            help="--scrub 때 본문·표지 전체에서 글자를 바꿉니다(여러 번 가능). 예: --replace \"옛회사=새회사\"")
    p_template.add_argument("--check", metavar="결과.dsd|작업.xlsx", dest="check_path",
                            help="이 파일에 틀의 식별값·금액·글이 남았는지 검사합니다.")
    p_template.add_argument("--old", metavar="문자열", action="append", default=[],
                            help="--check 때 더 찾을 틀 회사의 문자열(대표이사 이름, 주소 등, 여러 번 가능)")
    p_template.add_argument("--json", metavar="경로", dest="json_path", help="--check 결과를 JSON 으로 저장할 경로")
    p_template.add_argument("-v", "--verbose", action="store_true", help="틀과 같은 일반 문단까지 모두 출력합니다.")
    p_template.set_defaults(func=cmd_template)

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

    report = check(args.input, tolerance=args.tolerance,
                   xref=args.xref, xref_min=args.xref_min)
    if settlement:
        from dsd_bridge.settlement import check_settlement

        report.settlement = check_settlement(args.input, settlement, args.settlement_map)
        if args.report_path:
            report.settlement.write_xlsx(args.report_path)
    if args.json_path:
        report.write_json(args.json_path)
    for line in report.lines_ko(verbose=args.verbose):
        print(line)
    if settlement and args.report_path:
        print(f"안내: 상호대사 결과를 저장했습니다 — {args.report_path}")
    return report.exit_code


def cmd_template(args: argparse.Namespace) -> int:
    from dsd_bridge.errors import InputError
    from dsd_bridge.template import check_against, identity, scrub

    if args.scrub and args.check_path:
        raise InputError("오류: --scrub 과 --check 는 따로 실행하세요.")
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
        raise InputError("오류: --replace 와 -o 는 --scrub 과 함께 씁니다.")
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

    report = diff(args.old, args.new, rollforward=args.rollforward)
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
