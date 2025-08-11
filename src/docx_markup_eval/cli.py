import argparse
import json
import os
import sys
from pathlib import Path

from .evaluator import evaluate_documents
from .report import write_report
from .annotate import generate_annotations


def validate_args(args: argparse.Namespace) -> None:
    gt = Path(args.gt)
    ev = Path(args.eval)
    if not gt.exists() or gt.suffix.lower() != ".docx":
        sys.stderr.write("Error: --gt must be an existing .docx file\n")
        sys.exit(2)
    if not ev.exists() or ev.suffix.lower() != ".docx":
        sys.stderr.write("Error: --eval must be an existing .docx file\n")
        sys.exit(2)
    fmt = args.format.lower()
    if fmt not in {"json", "csv", "md"}:
        sys.stderr.write("Error: --format must be one of json|csv|md\n")
        sys.exit(2)
    out = Path(args.out)
    if out.suffix.lower().lstrip(".") != fmt:
        sys.stderr.write("Error: --out extension must match --format\n")
        sys.exit(2)
    if args.annotate:
        if not args.ref_out:
            sys.stderr.write("Error: --ref-out is required when --annotate is set\n")
            sys.exit(2)
        ref = Path(args.ref_out)
        try:
            ref.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            sys.stderr.write(f"Error: cannot create --ref-out directory: {e}\n")
            sys.exit(2)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="docx-markup-eval",
        description="Evaluate markup placement in DOCX tables.",
    )
    parser.add_argument("--gt", required=True, help="Path to ground-truth .docx")
    parser.add_argument("--eval", required=True, help="Path to evaluated .docx")
    parser.add_argument("--format", required=True, choices=["json", "csv", "md"], help="Report format")
    parser.add_argument("--out", required=True, help="Output report file path")
    parser.add_argument("--annotate", action="store_true", help="Enable annotated reference output")
    parser.add_argument("--ref-out", help="Directory to write annotated outputs when --annotate is set")
    parser.add_argument("--debug", action="store_true", help="Print debug details")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    ns = parse_args(sys.argv[1:] if argv is None else argv)
    validate_args(ns)

    results = evaluate_documents(
        gt_path=ns.gt,
        eval_path=ns.eval,
        debug=ns.debug,
    )

    write_report(results, ns.format, ns.out)

    if ns.annotate:
        generate_annotations(
            gt_path=ns.gt,
            eval_path=ns.eval,
            evaluation=results,
            output_dir=ns.ref_out,
            debug=ns.debug,
        )

    if ns.debug:
        sys.stderr.write(json.dumps(results.get("debug", {}), indent=2) + "\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())