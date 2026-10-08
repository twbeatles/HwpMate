"""테스트 전용 가짜 HwpMate CLI.

실제 argparse 인터페이스(--input/--format/--output/--report/--retry)를 흉내내며
ConversionSummary.to_json_dict() 형태의 보고서를 기록한다.
동작은 환경 변수로 제어한다:
  HWPMATE_FAKECLI_MODE: ok(기본) | busy | noreport | sleep
  HWPMATE_FAKECLI_SLEEP: sleep 모드 대기 초
  HWPMATE_FAKECLI_FAIL_NAMES: 실패로 처리할 입력 basename 목록 (쉼표 구분)
프로덕션 코드가 아니며 테스트에서 subprocess로만 실행한다.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="")
    parser.add_argument("--format", default="PDF")
    parser.add_argument("--output", default="")
    parser.add_argument("--report", default="")
    parser.add_argument("--retry", default="1")
    parser.add_argument("--pdf-export-mode", default="saveas_first")
    args = parser.parse_args()

    mode = os.environ.get("HWPMATE_FAKECLI_MODE", "ok")
    if mode == "busy":
        print("[오류] HwpMate(GUI 또는 다른 CLI)가 이미 실행 중입니다. 종료한 뒤 다시 실행하세요.", file=sys.stderr)
        return 1
    if mode == "nocom":
        print("[오류] 한글 COM 초기화 실패: fake", file=sys.stderr)
        return 1
    if mode == "sleep":
        time.sleep(float(os.environ.get("HWPMATE_FAKECLI_SLEEP", "30")))
    if mode == "noreport":
        return 1

    fail_names = {name.strip().lower() for name in os.environ.get("HWPMATE_FAKECLI_FAIL_NAMES", "").split(",") if name.strip()}
    input_path = Path(args.input)
    fmt = str(args.format).upper()
    ext_map = {"PDF": ".pdf", "DOCX": ".docx", "TXT": ".txt", "HWP": ".hwp", "HWPX": ".hwpx"}
    ext = ext_map.get(fmt, ".pdf")
    out_dir = Path(args.output) if args.output else input_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    output_file = out_dir / (input_path.stem + ext)

    failed = input_path.name.lower() in fail_names
    if not failed:
        output_file.write_bytes(b"%PDF-1.4 fake")
    task = {
        "input_file": str(input_path),
        "output_file": str(output_file),
        "status": "실패" if failed else "성공",
        "detail": "fake failure" if failed else "",
        "retry_count": 0,
        "backup_file": "",
        "backup_error": "",
        "created_files": [] if failed else [str(output_file)],
        "output_size": "" if failed else output_file.stat().st_size,
        "output_mtime": "",
        "save_format": fmt,
        "export_method": "saveas_2",
        "progid_used": "Fake.Hwp",
    }
    payload = {
        "summary": {
            "format_type": fmt,
            "total_requested": 1,
            "success_count": 0 if failed else 1,
            "failed_count": 1 if failed else 0,
            "skipped_count": 0,
            "canceled_count": 0,
            "elapsed_seconds": 0.01,
            "progid_used": "Fake.Hwp",
            "warnings": [],
        },
        "tasks": [task],
    }
    if args.report:
        report_path = Path(args.report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
