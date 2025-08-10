from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from docx import Document
from docx.table import _Cell, Table

NBSP_CHARS = "\u00A0\u2007\u202F"


def replace_nbsp(text: str) -> str:
    for ch in NBSP_CHARS:
        text = text.replace(ch, " ")
    return text


@dataclass(frozen=True)
class LogicalCellKey:
    table_index: int
    row_index: int
    col_index: int


@dataclass
class LogicalCell:
    key: LogicalCellKey
    merged_rect: Tuple[int, int, int, int]  # (row_min, col_min, row_max_inclusive, col_max_inclusive)
    text: str
    # Index map from flattened text global index -> (paragraph_idx, run_idx, offset_in_run)
    index_map: List[Tuple[int, int, int]]
    # For convenience keep reference to the underlying cell object in a second mapping built during extraction


def load_document(path: str | Path) -> Document:
    return Document(str(path))


def iter_tables(doc: Document) -> Iterable[Tuple[int, Table]]:
    for t_idx, tbl in enumerate(doc.tables):
        yield t_idx, tbl


def _cell_owner_coords(table: Table) -> Dict[int, Tuple[int, int]]:
    id_to_coords: Dict[int, Tuple[int, int]] = {}
    seen_groups: Dict[int, List[Tuple[int, int]]] = {}
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            group_id = id(cell._tc)
            seen_groups.setdefault(group_id, []).append((r_idx, c_idx))
    for gid, coords in seen_groups.items():
        top_left = min(coords)
        id_to_coords[gid] = top_left
    return id_to_coords


def _merged_rect_for_group(coords: List[Tuple[int, int]]) -> Tuple[int, int, int, int]:
    rows = [r for r, _ in coords]
    cols = [c for _, c in coords]
    return (min(rows), min(cols), max(rows), max(cols))


def _flatten_cell_with_index_map(cell: _Cell) -> Tuple[str, List[Tuple[int, int, int]]]:
    parts: List[str] = []
    index_map: List[Tuple[int, int, int]] = []
    global_idx = 0
    for p_idx, para in enumerate(cell.paragraphs):
        for r_idx, run in enumerate(para.runs):
            txt = run.text or ""
            if not txt:
                continue
            norm = replace_nbsp(txt)
            parts.append(norm)
            # Map each character position of norm to (p_idx, r_idx, offset)
            for offset in range(len(norm)):
                index_map.append((p_idx, r_idx, offset))
            global_idx += len(norm)
        # Paragraph boundaries are not represented in flattened text; they are skipped
    return ("".join(parts), index_map)


def extract_logical_cells(doc: Document) -> Tuple[List[LogicalCell], Dict[LogicalCellKey, _Cell]]:
    logical_cells: List[LogicalCell] = []
    key_to_cell: Dict[LogicalCellKey, _Cell] = {}

    for t_idx, tbl in iter_tables(doc):
        id_to_coords = _cell_owner_coords(tbl)
        groups: Dict[Tuple[int, int], List[Tuple[int, int]]] = {}
        for r_idx, row in enumerate(tbl.rows):
            for c_idx, cell in enumerate(row.cells):
                gid = id(cell._tc)
                owner = id_to_coords[gid]
                groups.setdefault(owner, []).append((r_idx, c_idx))
        for (owner_r, owner_c), coords in groups.items():
            owner_cell = tbl.cell(owner_r, owner_c)
            text, index_map = _flatten_cell_with_index_map(owner_cell)
            rect = _merged_rect_for_group(coords)
            key = LogicalCellKey(t_idx, owner_r, owner_c)
            logical_cells.append(
                LogicalCell(key=key, merged_rect=rect, text=text, index_map=index_map)
            )
            key_to_cell[key] = owner_cell
    # Sort logical cells by table, row, col for stable pairing
    logical_cells.sort(key=lambda lc: (lc.key.table_index, lc.key.row_index, lc.key.col_index))
    return logical_cells, key_to_cell