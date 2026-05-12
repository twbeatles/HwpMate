from __future__ import annotations

import argparse
import ctypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from hwpmate.constants import FORMAT_TYPES
from hwpmate.models import ConversionSummary, ConversionTask
from hwpmate.services.hwp_converter import HWPConverter, get_registered_hwp_progids


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run a real HWP COM smoke conversion and write a JSON result report.",
    )
    parser.add_argument("--input", required=True, help="Input .hwp or .hwpx file")
    parser.add_argument("--format", default="PDF", choices=sorted(FORMAT_TYPES), help="Output format")
    parser.add_argument("--output-dir", default="", help="Output directory. Defaults to the input folder.")
    parser.add_argument("--result-json", default="", help="Result JSON path. Defaults to the output directory.")
    parser.add_argument("--allow-non-admin", action="store_true", help="Run even when not elevated.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_file = Path(args.input).resolve()
    if not input_file.is_file():
        print(f"Input file not found: {input_file}", file=sys.stderr)
        return 2

    warnings: list[str] = []
    if not is_admin():
        message = "Administrator privileges are recommended for HWP COM smoke verification."
        if not args.allow_non_admin:
            print(message, file=sys.stderr)
            return 3
        warnings.append(message)

    registered_progids = get_registered_hwp_progids()
    if not registered_progids:
        warnings.append("No registered HWP COM ProgID was detected before conversion.")

    output_dir = Path(args.output_dir).resolve() if args.output_dir else input_file.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f"{input_file.stem}{FORMAT_TYPES[args.format].ext}"
    result_json = (
        Path(args.result_json).resolve()
        if args.result_json
        else output_dir / f"hwp_com_smoke_{int(time.time())}.json"
    )

    task = ConversionTask(input_file=input_file, output_file=output_file)
    converter = HWPConverter()
    started = time.perf_counter()
    try:
        converter.initialize()
        success, error = converter.convert_file(input_file, output_file, args.format)
        task.status = "성공" if success else "실패"
        task.error = error
        task.created_files = list(converter.last_created_files)
        task.output_size = converter.last_output_size
        task.output_mtime = converter.last_output_mtime
        task.save_format = converter.last_save_format
        task.progid_used = converter.progid_used

        if converter.security_module_registered is False:
            detail = f" 상세: {converter.security_module_error}" if converter.security_module_error else ""
            warnings.append(f"한글 보안 모듈 등록에 실패했습니다.{detail}")
        if converter.process_tracking_warning:
            warnings.append(converter.process_tracking_warning)
    except Exception as exc:
        task.status = "실패"
        task.error = str(exc)
    finally:
        converter.cleanup()

    summary = ConversionSummary(
        format_type=args.format,
        tasks=[task],
        warnings=warnings,
        elapsed_seconds=time.perf_counter() - started,
        progid_used=task.progid_used,
    )
    with result_json.open("w", encoding="utf-8") as f:
        json.dump(summary.to_json_dict(), f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"Result JSON: {result_json}")
    if task.status == "성공":
        print(f"Created files: {', '.join(str(path) for path in task.created_files) or output_file}")
        return 0

    print(f"Conversion failed: {task.detail}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
