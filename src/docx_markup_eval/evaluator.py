from __future__ import annotations

import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

from .docx_utils import LogicalCell, LogicalCellKey, load_document, extract_logical_cells, replace_nbsp

TOKEN_RE = re.compile(r"(?i)cell_\d+")


@dataclass
class CellTokenData:
    # Original token strings with their start indices in original flattened text
    tokens: List[Tuple[int, str]]
    # Base text with tokens removed
    base_text: str
    # Base indices of token anchors (number of base chars before token starts)
    base_indices: List[int]
    # Spans of tokens in original flattened text
    spans: List[Tuple[int, int]]


def _detect_tokens_and_base(text: str) -> CellTokenData:
    tokens: List[Tuple[int, str]] = []
    spans: List[Tuple[int, int]] = []
    for m in TOKEN_RE.finditer(text):
        start, end = m.span()
        tokens.append((start, m.group(0)))
        spans.append((start, end))
    # Build base text and base index mapping
    base_parts: List[str] = []
    base_indices: List[int] = []
    cursor = 0
    for (start, end), (_, tok) in zip(spans, tokens):
        base_parts.append(text[cursor:start])
        base_indices.append(len("".join(base_parts)))
        cursor = end
    base_parts.append(text[cursor:])
    base_text = "".join(base_parts)
    return CellTokenData(tokens=tokens, base_text=base_text, base_indices=base_indices, spans=spans)


def _map_gt_base_to_eval(gt_base: str, ev_base: str, gt_indices: Sequence[int]) -> List[int]:
    sm = SequenceMatcher(a=gt_base, b=ev_base, autojunk=False)
    opcodes = sm.get_opcodes()
    mapped: List[int] = []
    for pos in gt_indices:
        mapped_pos = 0
        for tag, i1, i2, j1, j2 in opcodes:
            if i1 <= pos <= i2:
                if tag == "equal":
                    mapped_pos = j1 + (pos - i1)
                elif tag in ("replace", "delete"):
                    mapped_pos = j1
                elif tag == "insert":
                    # Insertion in eval before this gt position; keep current j1 as anchor
                    mapped_pos = j1
                break
            elif pos > i2:
                # Advance through this block
                if tag in ("equal", "replace", "delete"):
                    # consumes gt span; eval advanced by (j2-j1)
                    continue
        else:
            mapped_pos = pos
        mapped.append(mapped_pos)
    return mapped


def _map_eval_base_to_gt(gt_base: str, ev_base: str, ev_indices: Sequence[int]) -> List[int]:
    sm = SequenceMatcher(a=gt_base, b=ev_base, autojunk=False)
    opcodes = sm.get_opcodes()
    mapped: List[int] = []
    for pos in ev_indices:
        mapped_pos = 0
        for tag, i1, i2, j1, j2 in opcodes:
            if j1 <= pos <= j2:
                if tag == "equal":
                    mapped_pos = i1 + (pos - j1)
                elif tag in ("replace", "insert"):
                    mapped_pos = i1
                elif tag == "delete":
                    mapped_pos = i1
                break
        else:
            mapped_pos = pos
        mapped.append(mapped_pos)
    return mapped


def _original_index_from_base_index(spans: Sequence[Tuple[int, int]], base_index: int) -> int:
    # Given token spans removed from original, compute original index corresponding to base_index
    shift = 0
    for start, end in spans:
        if start - shift >= base_index:
            break
        shift += (end - start)
    return base_index + shift


def evaluate_documents(gt_path: str | Path, eval_path: str | Path, debug: bool = False) -> Dict:
    gt_doc = load_document(gt_path)
    ev_doc = load_document(eval_path)

    gt_cells, gt_map = extract_logical_cells(gt_doc)
    ev_cells, ev_map = extract_logical_cells(ev_doc)

    if len(gt_cells) != len(ev_cells):
        # Pairing relies on same geometry
        raise SystemExit("Error: Ground-truth and evaluated documents must have identical table shapes and orders")

    overall = {"gt_total": 0, "eval_total": 0, "correct": 0, "missed": 0, "misplaced": 0}
    debug_cells: List[Dict] = []

    cell_pairs: List[Tuple[LogicalCell, LogicalCell]] = list(zip(gt_cells, ev_cells))

    per_cell_results: List[Dict] = []

    for gt_cell, ev_cell in cell_pairs:
        gt_text = replace_nbsp(gt_cell.text)
        ev_text = replace_nbsp(ev_cell.text)
        gt_data = _detect_tokens_and_base(gt_text)
        ev_data = _detect_tokens_and_base(ev_text)

        mapped_gt_to_ev = _map_gt_base_to_eval(gt_data.base_text, ev_data.base_text, gt_data.base_indices)
        ev_base_set = set(ev_data.base_indices)
        correct = sum(1 for idx in mapped_gt_to_ev if idx in ev_base_set)
        gt_total = len(gt_data.tokens)
        ev_total = len(ev_data.tokens)
        missed = gt_total - correct
        misplaced = ev_total - correct

        overall["gt_total"] += gt_total
        overall["eval_total"] += ev_total
        overall["correct"] += correct
        overall["missed"] += missed
        overall["misplaced"] += misplaced

        cell_debug = {
            "cell": {
                "table": gt_cell.key.table_index,
                "row": gt_cell.key.row_index,
                "col": gt_cell.key.col_index,
                "merged_rect": gt_cell.merged_rect,
            },
            "gt": {
                "text": gt_text,
                "tokens": gt_data.tokens,
                "base_indices": gt_data.base_indices,
                "base_text": gt_data.base_text,
                "spans": gt_data.spans,
            },
            "eval": {
                "text": ev_text,
                "tokens": ev_data.tokens,
                "base_indices": ev_data.base_indices,
                "base_text": ev_data.base_text,
                "spans": ev_data.spans,
            },
            "mapped_gt_to_eval_base": mapped_gt_to_ev,
        }

        per_cell_results.append(
            {
                "key": {
                    "table": gt_cell.key.table_index,
                    "row": gt_cell.key.row_index,
                    "col": gt_cell.key.col_index,
                },
                "gt_total": gt_total,
                "eval_total": ev_total,
                "correct": correct,
                "missed": missed,
                "misplaced": misplaced,
                "gt_data": gt_data,
                "ev_data": ev_data,
            }
        )

        if debug:
            debug_cells.append(cell_debug)

    results = {**overall, "cells": per_cell_results}
    if debug:
        results["debug"] = {"cells": debug_cells}
    return results