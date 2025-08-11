from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Dict


def write_report(results: Dict, fmt: str, out_path: str | Path) -> None:
    fmt = fmt.lower()
    if fmt == "json":
        _write_json(results, out_path)
    elif fmt == "csv":
        _write_csv(results, out_path)
    elif fmt == "md":
        _write_md(results, out_path)
    else:
        raise ValueError(f"Unsupported format: {fmt}")


def _write_json(results: Dict, out_path: str | Path) -> None:
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "gt_total": results["gt_total"],
                "eval_total": results["eval_total"],
                "correct": results["correct"],
                "misplaced": results["misplaced"],
                "missed": results["missed"],
            },
            f,
        )


def _write_csv(results: Dict, out_path: str | Path) -> None:
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["gt_total", "eval_total", "correct", "misplaced", "missed"]
        )
        writer.writeheader()
        writer.writerow(
            {
                "gt_total": results["gt_total"],
                "eval_total": results["eval_total"],
                "correct": results["correct"],
                "misplaced": results["misplaced"],
                "missed": results["missed"],
            }
        )


def _write_md(results: Dict, out_path: str | Path) -> None:
    lines = [
        "| gt_total | eval_total | correct | misplaced | missed |",
        "| --- | --- | --- | --- | --- |",
        f"| {results['gt_total']} | {results['eval_total']} | {results['correct']} | {results['misplaced']} | {results['missed']} |",
    ]
    Path(out_path).write_text("\n".join(lines), encoding="utf-8")