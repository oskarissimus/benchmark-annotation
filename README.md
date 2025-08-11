# DOCX Markup Placement Evaluator

CLI to evaluate markup placement in Word .docx tables and optionally generate annotated references.

## Install

- Requires Python 3.11+
- Runtime deps: python-docx, Pillow, pdf2image (optional)
- For PDF/PNG artifacts, either:
  - Install LibreOffice (`soffice`) and poppler (`pdf2image`), or
  - The tool falls back to a placeholder PDF/PNG using Pillow.

## Usage

```
docx-markup-eval \
  --gt path/to/gt.docx \
  --eval path/to/eval.docx \
  --format json|csv|md \
  --out path/to/report.(json|csv|md) \
  [--annotate --ref-out path/to/output_dir] \
  [--debug]
```