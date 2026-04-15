#!/usr/bin/env python3

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Callable

import offline_batch_ocr as core

LAST_RUN_LOG: Path | None = None
LAST_FAILURE_REASON: str = ""
RUNTIME_ROOT: Path | None = None


def subprocess_window_kwargs() -> dict[str, object]:
    if os.name != "nt":
        return {}
    return {"creationflags": subprocess.CREATE_NO_WINDOW}


def print_line(text: str = "") -> None:
    print(text, flush=True)
    if LAST_RUN_LOG is not None:
        with LAST_RUN_LOG.open("a", encoding="utf-8") as handle:
            handle.write(text + "\n")


def set_failure_reason(reason: str) -> None:
    global LAST_FAILURE_REASON
    LAST_FAILURE_REASON = reason.strip() if reason else ""


def get_last_failure_reason() -> str:
    return LAST_FAILURE_REASON


def configure_runtime_paths() -> Path:
    local_app_data = os.environ.get("LocalAppData")
    if not local_app_data:
        return core.ROOT

    runtime_root = Path(local_app_data) / "DetailExtractOCR"
    runtime_root.mkdir(parents=True, exist_ok=True)

    core.ROOT = runtime_root
    core.INPUT_DIR = runtime_root / "input-images"
    core.OUTPUT_DIR = runtime_root / "output"
    core.RAW_COMBINED_OUTPUT = core.OUTPUT_DIR / "combined_ocr.txt"
    core.PARSED_COMBINED_OUTPUT = core.OUTPUT_DIR / "combined_customers.txt"
    global LAST_RUN_LOG, RUNTIME_ROOT
    LAST_RUN_LOG = runtime_root / "last_run.log"
    LAST_RUN_LOG.write_text("", encoding="utf-8")
    RUNTIME_ROOT = runtime_root
    return runtime_root


def configure_tesseract_environment(tesseract_path: str) -> None:
    install_dir = Path(tesseract_path).resolve().parent
    tessdata = install_dir / "tessdata"
    # Tesseract expects TESSDATA_PREFIX to be the install dir (parent of tessdata),
    # not the tessdata directory itself.
    if tessdata.exists():
        os.environ["TESSDATA_PREFIX"] = str(install_dir)


def can_execute_tesseract(exe_path: Path) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            [str(exe_path), "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
            **subprocess_window_kwargs(),
        )
    except Exception as error:
        return False, f"{error}"
    if result.returncode != 0:
        return False, (result.stderr or "").strip() or f"exit code {result.returncode}"
    return True, ((result.stdout or "").splitlines()[0].strip() if result.stdout else "ok")


def prepare_runtime_tesseract(candidate_exe: Path) -> Path | None:
    runtime = RUNTIME_ROOT or core.ROOT
    target_dir = runtime / "tesseract-runtime"
    target_exe = target_dir / "tesseract.exe"
    try:
        if target_dir.exists():
            # Keep startup fast; only refresh if exe is missing.
            if not target_exe.exists():
                shutil.rmtree(target_dir, ignore_errors=True)
                shutil.copytree(candidate_exe.parent, target_dir, dirs_exist_ok=True)
        else:
            shutil.copytree(candidate_exe.parent, target_dir, dirs_exist_ok=True)
    except Exception as error:
        print_line(f"Tesseract runtime copy failed: {error}")
        return None
    return target_exe if target_exe.exists() else None


def find_windows_tesseract() -> str | None:
    frozen_base: Path | None = None
    frozen_temp: Path | None = None
    if getattr(sys, "frozen", False):
        frozen_base = Path(sys.executable).resolve().parent
        frozen_temp_path = getattr(sys, "_MEIPASS", "")
        if frozen_temp_path:
            frozen_temp = Path(frozen_temp_path)

    bundled_candidates: list[Path] = []
    if frozen_base is not None:
        bundled_candidates.append(frozen_base / "tesseract" / "tesseract.exe")
    if frozen_temp is not None:
        bundled_candidates.append(frozen_temp / "tesseract" / "tesseract.exe")

    resolved_candidates: list[Path] = []
    for candidate in bundled_candidates:
        if candidate.exists():
            resolved_candidates.append(candidate)

    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    local_app_data = os.environ.get("LocalAppData", "")

    candidates = [
        Path(program_files) / "Tesseract-OCR" / "tesseract.exe",
        Path(program_files_x86) / "Tesseract-OCR" / "tesseract.exe",
    ]
    if local_app_data:
        candidates.append(Path(local_app_data) / "Programs" / "Tesseract-OCR" / "tesseract.exe")

    existing = core.get_tesseract_path()
    if existing:
        resolved_candidates.append(Path(existing))

    for candidate in candidates:
        if candidate.exists():
            resolved_candidates.append(candidate)

    # Deduplicate while preserving order.
    seen: set[str] = set()
    deduped: list[Path] = []
    for candidate in resolved_candidates:
        key = str(candidate.resolve())
        if key not in seen:
            seen.add(key)
            deduped.append(candidate)

    for candidate in deduped:
        ok, _ = can_execute_tesseract(candidate)
        if ok:
            os.environ["PATH"] = f"{candidate.parent}{os.pathsep}{os.environ.get('PATH', '')}"
            configure_tesseract_environment(str(candidate))
            return str(candidate)

        runtime_exe = prepare_runtime_tesseract(candidate)
        if runtime_exe is None:
            continue
        ok_runtime, _ = can_execute_tesseract(runtime_exe)
        if ok_runtime:
            os.environ["PATH"] = f"{runtime_exe.parent}{os.pathsep}{os.environ.get('PATH', '')}"
            configure_tesseract_environment(str(runtime_exe))
            return str(runtime_exe)

    return None


def verify_tesseract_ready(tesseract_path: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            [tesseract_path, "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=12,
            **subprocess_window_kwargs(),
        )
    except Exception as error:
        return False, f"unable to execute tesseract: {error}"

    if result.returncode != 0:
        stderr = (result.stderr or "").strip()
        return False, stderr or f"tesseract exited with code {result.returncode}"
    return True, (result.stdout or "").splitlines()[0].strip() if result.stdout else "tesseract ready"


def extract_code(chunk: str) -> int | None:
    match = re.match(r"^(\d{6})\b", chunk)
    if not match:
        return None
    return int(match.group(1))


def normalize_record_start_lines(text: str) -> str:
    # OCR can split "RecordNo + Title" across two lines:
    # 456845\nDr. Name...
    # Join those lines so downstream strict parser keeps 40-field output format.
    pattern = re.compile(rf"(?m)^(\d{{6}})\s*\n\s*({core.TITLE_PATTERN}(?:\s|$).*)$")
    return pattern.sub(r"\1 \2", text)


def run_ocr_with_timeout(image_path: Path, tesseract_path: str, timeout_seconds: int = 180) -> str:
    # Keep OCR accurate, but avoid a huge pass matrix that makes the UI look frozen.
    tessdata_dir = Path(tesseract_path).resolve().parent / "tessdata"
    tessdata_args = ["--tessdata-dir", str(tessdata_dir)] if tessdata_dir.exists() else []
    staging_dir = core.ROOT / "_staging"
    staging_dir.mkdir(parents=True, exist_ok=True)
    staged_path = staging_dir / image_path.name
    preferred_path = image_path
    try:
        shutil.copy2(image_path, staged_path)
        preferred_path = staged_path
    except Exception:
        pass

    candidates: list[str] = []
    errors: list[str] = []

    input_path = preferred_path
    enhanced_paths: list[Path] = []
    try:
        enhanced_paths = core.build_enhanced_images(input_path)
    except Exception:
        enhanced_paths = []

    candidate_paths = [input_path, *enhanced_paths[:1]]
    commands: list[tuple[list[str], int]] = []
    for candidate_path, psm in [
        (candidate_paths[0], "6"),
        (candidate_paths[0], "11"),
        (candidate_paths[-1], "6"),
    ]:
        commands.append(
            (
                [
                    tesseract_path,
                    str(candidate_path),
                    "stdout",
                    *tessdata_args,
                    "--oem",
                    "1",
                    "--psm",
                    psm,
                    "-c",
                    "preserve_interword_spaces=1",
                ],
                15,
            )
        )

    try:
        for command, command_timeout in commands:
            try:
                result = subprocess.run(
                    command,
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=command_timeout,
                    **subprocess_window_kwargs(),
                )
                text = result.stdout.strip()
                if text:
                    candidates.append(text)
                else:
                    errors.append(f"empty output with psm {command[command.index('--psm') + 1]}")
            except PermissionError as error:
                errors.append(f"permission error for {input_path}: {error}")
            except subprocess.TimeoutExpired:
                errors.append(f"timeout with psm {command[command.index('--psm') + 1]}")
            except subprocess.CalledProcessError as error:
                stderr = (error.stderr or "").strip()
                errors.append(stderr or f"command failed with psm {command[command.index('--psm') + 1]}")

        if not candidates and core.cv2 is not None and core.pytesseract is not None:
            for psm in ("6",):
                try:
                    cv2_text = core.run_pytesseract_cv2(input_path, tesseract_path, psm)
                    if cv2_text:
                        candidates.append(cv2_text)
                except Exception as error:
                    errors.append(f"cv2 fallback failed: {error}")
    finally:
        if enhanced_paths:
            shutil.rmtree(enhanced_paths[0].parent, ignore_errors=True)

    if candidates:
        return max(candidates, key=core.score_ocr_text)

    raise RuntimeError("; ".join(errors) if errors else "OCR failed")


def write_extraction_analytics(parsed_blocks: list[str]) -> Path:
    rows: list[dict[str, str]] = []
    for block in parsed_blocks:
        values = block.splitlines()
        if len(values) != len(core.FIELD_NAMES):
            continue
        rows.append(dict(zip(core.FIELD_NAMES, values)))

    summary_txt = core.OUTPUT_DIR / "extraction_quality_summary.txt"

    if not rows:
        summary_txt.write_text(
            "\n".join(
                [
                    "Extraction Quality Summary",
                    "--------------------------",
                    "records: 0",
                    "missing_total: 0",
                    "unclear_total: 0",
                    "confidence_percent_estimate: 0.0",
                ]
            )
            + "\n",
            encoding="utf-8",
        )
        return summary_txt

    # Optional use of pandas/numpy for richer analytics.
    try:
        import pandas as pd  # type: ignore
        import numpy as np  # type: ignore

        df = pd.DataFrame(rows)
        missing_mask = df.eq(core.MISSING_TOKEN)
        unclear_mask = df.eq(core.UNCLEAR_TOKEN)
        missing_total = int(missing_mask.sum().sum())
        unclear_total = int(unclear_mask.sum().sum())
        total_cells = int(np.prod(df.shape))
        valid_total = int(total_cells - missing_total - unclear_total)
        confidence = round((valid_total / total_cells) * 100, 2) if total_cells else 0.0

        summary_txt.write_text(
            "\n".join(
                [
                    "Extraction Quality Summary",
                    "--------------------------",
                    f"records: {int(len(df))}",
                    f"missing_total: {missing_total}",
                    f"unclear_total: {unclear_total}",
                    f"confidence_percent_estimate: {confidence}",
                    "analytics_engine: pandas+numpy",
                ]
            )
            + "\n",
            encoding="utf-8",
        )
        return summary_txt
    except Exception:
        pass

    missing_total = 0
    unclear_total = 0
    total_cells = 0
    for row in rows:
        for field in core.FIELD_NAMES:
            total_cells += 1
            if row.get(field) == core.MISSING_TOKEN:
                missing_total += 1
            elif row.get(field) == core.UNCLEAR_TOKEN:
                unclear_total += 1
    valid_total = total_cells - missing_total - unclear_total
    confidence = round((valid_total / total_cells) * 100, 2) if total_cells else 0.0

    summary_txt.write_text(
        "\n".join(
            [
                "Extraction Quality Summary",
                "--------------------------",
                f"records: {len(rows)}",
                f"missing_total: {missing_total}",
                f"unclear_total: {unclear_total}",
                f"confidence_percent_estimate: {confidence}",
                "analytics_engine: builtin",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    return summary_txt


def process_images(
    images: list[Path],
    tesseract_path: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
    message_callback: Callable[[str], None] | None = None,
) -> int:
    set_failure_reason("")

    if not images:
        print_line("No images were provided.")
        set_failure_reason("No images were provided.")
        return 1

    ready, detail = verify_tesseract_ready(tesseract_path)
    if not ready:
        print_line(f"Tesseract preflight failed: {detail}")
        set_failure_reason(f"Tesseract preflight failed: {detail}")
        if message_callback is not None:
            message_callback(f"Tesseract preflight failed: {detail}")
        return 1
    print_line(f"Tesseract preflight: {detail}")
    if message_callback is not None:
        message_callback(f"Tesseract preflight: {detail}")

    print_line(f"Found {len(images)} image(s).")
    print_line(f"Writing to:   {core.OUTPUT_DIR}")
    print_line("")
    if message_callback is not None:
        message_callback(f"Found {len(images)} image(s)")
        message_callback(f"Output folder: {core.OUTPUT_DIR}")

    max_workers = max(1, min(2, len(images), os.cpu_count() or 2))
    print_line(f"OCR workers: {max_workers}")
    if message_callback is not None:
        message_callback(f"OCR workers: {max_workers}")

    def ocr_one(idx: int, image_path: Path) -> tuple[int, Path, str]:
        print_line(f"Starting {idx}/{len(images)}: {image_path.name}")
        if message_callback is not None:
            message_callback(f"Starting {idx}/{len(images)}: {image_path.name}")
        try:
            return idx, image_path, run_ocr_with_timeout(image_path, tesseract_path)
        except subprocess.TimeoutExpired:
            print_line(f"  OCR timeout for {image_path.name} (>{180}s)")
            if message_callback is not None:
                message_callback(f"Timeout: {image_path.name} (>{180}s)")
            return idx, image_path, "[ocr timeout]"
        except RuntimeError as error:
            print_line(f"  OCR failed for {image_path.name}: {error}")
            if message_callback is not None:
                message_callback(f"Failed: {image_path.name} -> {error}")
            return idx, image_path, "[ocr failed]"
        except subprocess.CalledProcessError as error:
            if error.stderr:
                print_line(f"  OCR failed for {image_path.name}: {error.stderr.strip()}")
                if message_callback is not None:
                    message_callback(f"Failed: {image_path.name} -> {error.stderr.strip()}")
            else:
                print_line(f"  OCR failed for {image_path.name}")
                if message_callback is not None:
                    message_callback(f"Failed: {image_path.name}")
            return idx, image_path, "[ocr failed]"
        except Exception as error:
            print_line(f"  OCR failed for {image_path.name}: unexpected error: {error}")
            if message_callback is not None:
                message_callback(f"Failed: {image_path.name} -> unexpected error: {error}")
            return idx, image_path, "[ocr failed]"

    indexed: list[tuple[int, Path, str]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = [pool.submit(ocr_one, idx, image_path) for idx, image_path in enumerate(images, start=1)]
        done = 0
        for future in as_completed(futures):
            result = future.result()
            indexed.append(result)
            done += 1
            _, image_path, _ = result
            print_line(f"[{done}/{len(images)}] Finished {image_path.name}")
            if progress_callback is not None:
                progress_callback(done, len(images), image_path.name)
            if message_callback is not None:
                message_callback(f"Finished {done}/{len(images)}: {image_path.name}")

    indexed.sort(key=lambda row: row[0])
    raw_results: list[tuple[Path, str]] = [(image_path, text) for _, image_path, text in indexed]
    core.write_raw_output(raw_results)
    if message_callback is not None:
        message_callback(f"Saved raw OCR: {core.RAW_COMBINED_OUTPUT}")
        message_callback("Parsing structured customer fields...")

    parsed_blocks: list[str] = []

    for page_index, (_, page_text) in enumerate(raw_results, start=1):
        page_stream = normalize_record_start_lines(page_text)
        _, page_chunks = core.split_global_record_chunks(page_stream)
        if not page_chunks:
            continue
        for chunk in page_chunks:
            record = core.parse_record(chunk, False)
            parsed_blocks.append(core.format_record_block(0, record))
        core.write_parsed_output(parsed_blocks)
        if message_callback is not None:
            message_callback(f"Parsed image {page_index}/{len(raw_results)}")

    if not parsed_blocks:
        parsed_blocks = [core.format_record_block(0, core.default_record())]
        print_line("No structured record pattern detected. Saved raw OCR output for manual review.")
        set_failure_reason("No structured record pattern detected.")
        core.write_parsed_output(parsed_blocks)

    core.save_combined(raw_results, parsed_blocks)

    print_line("")
    print_line("Done.")
    print_line(f"Raw OCR file: {core.RAW_COMBINED_OUTPUT}")
    print_line(f"Parsed customer file: {core.PARSED_COMBINED_OUTPUT}")
    if LAST_RUN_LOG is not None:
        print_line(f"Run log: {LAST_RUN_LOG}")
    if message_callback is not None:
        message_callback("Extraction complete")
        message_callback(f"Raw OCR: {core.RAW_COMBINED_OUTPUT}")
        message_callback(f"Parsed output: {core.PARSED_COMBINED_OUTPUT}")
    return 0


def main() -> int:
    runtime_root = configure_runtime_paths()

    print_line("Offline Batch OCR (Windows)")
    print_line("---------------------------")

    tesseract_path = find_windows_tesseract()
    if not tesseract_path:
        print_line("Tesseract was not found on this PC.")
        print_line("Install from: https://github.com/UB-Mannheim/tesseract/wiki")
        print_line("After install, close this window and run the tool again.")
        return 1

    core.ensure_directories()
    images = core.list_images()
    if not images:
        print_line(f"No images found in: {core.INPUT_DIR}")
        print_line("Put your image files into the 'input-images' folder and run again.")
        return 1

    print_line(f"Found {len(images)} image(s).")
    print_line(f"Reading from: {core.INPUT_DIR}")
    print_line(f"Writing to:   {core.OUTPUT_DIR}")
    print_line(f"Runtime root: {runtime_root}")
    return process_images(images, tesseract_path)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print_line("\nStopped.")
        raise SystemExit(130)
