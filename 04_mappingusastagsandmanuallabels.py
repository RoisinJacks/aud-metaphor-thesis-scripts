#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Union
import pandas as pd


ColSpec = Union[str, int]


def _parse_colspec(value: str) -> ColSpec:
    """
    Allow CLI args like "4" (meaning column #4, 1-based) or "Item" (meaning a column name).
    """
    value = value.strip()
    if value.isdigit():
        return int(value)
    return value


def _get_series(df: pd.DataFrame, col: Optional[ColSpec], *, fallback_none: bool = True) -> Optional[pd.Series]:
    """
    Retrieve a column from df by name (str) or 1-based index (int).
    Returns None if col is None and fallback_none is True.
    Raises a helpful error if a specified column cannot be found.
    """
    if col is None:
        return None if fallback_none else pd.Series([None] * len(df))

    if isinstance(col, int):
        idx = col - 1
        if idx < 0 or idx >= df.shape[1]:
            raise ValueError(
                f"Column index {col} is out of range for sheet with {df.shape[1]} columns."
            )
        return df.iloc[:, idx]

    # name
    if col not in df.columns:
        raise ValueError(
            f"Column '{col}' not found. Available columns: {list(df.columns)}"
        )
    return df[col]


def _build_key_columns(
    df: pd.DataFrame,
    item_col: ColSpec,
    label_col: Optional[ColSpec],
    vg_col: Optional[ColSpec],
) -> Tuple[pd.Series, Optional[pd.Series], Optional[pd.Series]]:
    item = _get_series(df, item_col, fallback_none=False).astype(str)
    label = _get_series(df, label_col)  # may be None
    vg = _get_series(df, vg_col)        # may be None

    if label is not None:
        label = label.astype(str)
    if vg is not None:
        vg = vg.astype(str)

    return item, label, vg


def build_key(item: str, label: Optional[str], vg: Optional[str]) -> Tuple[str, Optional[str], Optional[str]]:
    return (item, label, vg)


def build_mapping_unique_tags(
    sample_df: pd.DataFrame,
    item_col: ColSpec,
    tags_start_col_1_based: int,
    label_col: Optional[ColSpec] = None,
    vg_col: Optional[ColSpec] = None,
) -> Dict[Tuple[str, Optional[str], Optional[str]], List[str]]:
    """
    Build mapping: (Item, Manual_Label, VG) -> sorted unique tags.
    If label_col and/or vg_col are None, mapping collapses over those dimensions.
    """
    tags_start_idx = tags_start_col_1_based - 1
    if tags_start_idx < 0 or tags_start_idx >= sample_df.shape[1]:
        raise ValueError("tags_start_col is out of range for the sample sheet.")

    item_s, label_s, vg_s = _build_key_columns(sample_df, item_col, label_col, vg_col)

    m: Dict[Tuple[str, Optional[str], Optional[str]], set] = {}

    for i in range(len(sample_df)):
        item = item_s.iat[i]
        label = label_s.iat[i] if label_s is not None else None
        vg = vg_s.iat[i] if vg_s is not None else None

        if item == "nan" or item.strip() == "":
            continue

        key = build_key(item, label, vg)
        if key not in m:
            m[key] = set()

        row = sample_df.iloc[i, tags_start_idx:]
        for tag in row:
            if pd.notna(tag):
                m[key].add(str(tag))

    return {k: sorted(list(v)) for k, v in m.items()}


def build_mapping_tag_counts(
    sample_df: pd.DataFrame,
    item_col: ColSpec,
    tags_start_col_1_based: int,
    label_col: Optional[ColSpec] = None,
    vg_col: Optional[ColSpec] = None,
) -> Dict[Tuple[str, Optional[str], Optional[str]], Dict[str, int]]:
    """
    Build mapping: (Item, Manual_Label, VG) -> {tag: count}.
    """
    tags_start_idx = tags_start_col_1_based - 1
    if tags_start_idx < 0 or tags_start_idx >= sample_df.shape[1]:
        raise ValueError("tags_start_col is out of range for the sample sheet.")

    item_s, label_s, vg_s = _build_key_columns(sample_df, item_col, label_col, vg_col)

    counts: Dict[Tuple[str, Optional[str], Optional[str]], Dict[str, int]] = {}

    for i in range(len(sample_df)):
        item = item_s.iat[i]
        label = label_s.iat[i] if label_s is not None else None
        vg = vg_s.iat[i] if vg_s is not None else None

        if item == "nan" or item.strip() == "":
            continue

        key = build_key(item, label, vg)
        if key not in counts:
            counts[key] = {}

        row = sample_df.iloc[i, tags_start_idx:]
        for tag in row:
            if pd.notna(tag):
                tag = str(tag)
                counts[key][tag] = counts[key].get(tag, 0) + 1

    return counts


def enrich_manual_list_with_tags(
    manual_df: pd.DataFrame,
    tags_map: Dict[Tuple[str, Optional[str], Optional[str]], List[str]],
    item_col: ColSpec,
    label_col: Optional[ColSpec],
    vg_col: Optional[ColSpec],
) -> pd.DataFrame:
    """
    Append Tag1..TagN columns to the manual list.
    Looks up tags by (Item, Manual_Label, VG) when available; otherwise falls back
    to (Item, None, None).
    """
    out = manual_df.copy()
    item_s, label_s, vg_s = _build_key_columns(out, item_col, label_col, vg_col)

    # Determine max tags needed across rows
    max_tags = 0
    tags_by_row: List[List[str]] = []

    for i in range(len(out)):
        item = item_s.iat[i]
        label = label_s.iat[i] if label_s is not None else None
        vg = vg_s.iat[i] if vg_s is not None else None

        # Try exact key first
        tags = tags_map.get(build_key(item, label, vg))

        # Fall back to item-only if exact key not available
        if tags is None:
            tags = tags_map.get(build_key(item, None, None), [])

        tags_by_row.append(tags)
        max_tags = max(max_tags, len(tags))

    if max_tags == 0:
        return out

    tag_cols = [f"Tag{i+1}" for i in range(max_tags)]
    tags_df = pd.DataFrame(index=out.index, columns=tag_cols)

    for i, tags in enumerate(tags_by_row):
        for j, tag in enumerate(tags):
            tags_df.iat[i, j] = tag

    return pd.concat([out, tags_df], axis=1)


def build_counts_long_table(
    manual_df: pd.DataFrame,
    counts_map: Dict[Tuple[str, Optional[str], Optional[str]], Dict[str, int]],
    item_col: ColSpec,
    label_col: Optional[ColSpec],
    vg_col: Optional[ColSpec],
) -> pd.DataFrame:
    """
    Long table: Item | Manual_Label | VG | Tag | Count
    Uses exact (Item, Label, VG) matching with item-only fallback.
    """
    item_s, label_s, vg_s = _build_key_columns(manual_df, item_col, label_col, vg_col)

    rows: List[Tuple[str, Optional[str], Optional[str], str, int]] = []

    for i in range(len(manual_df)):
        item = item_s.iat[i]
        label = label_s.iat[i] if label_s is not None else None
        vg = vg_s.iat[i] if vg_s is not None else None

        d = counts_map.get(build_key(item, label, vg))
        if d is None:
            d = counts_map.get(build_key(item, None, None), {})

        for tag, count in d.items():
            rows.append((item, label, vg, tag, int(count)))

    return pd.DataFrame(rows, columns=["Item", "Manual_Label", "VG", "Tag", "Count"])


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Map Wmatrix USAS tags onto a manually labelled list and export Tag1..TagN columns."
    )
    parser.add_argument("--input_excel", required=True, help="Path to the input Excel workbook.")
    parser.add_argument("--manual_sheet", required=True, help="Sheet name for the manual list.")
    parser.add_argument("--sample_sheet", required=True, help="Sheet name for the USAS sample sheet.")
    parser.add_argument("--output_excel", required=True, help="Path to the output Excel file.")

    parser.add_argument("--mode", choices=["map", "count", "both"], default="map",
                        help="Outputs: map (Tag1..TagN), count (long tag counts), or both. Default: map.")

    parser.add_argument("--item_col", type=_parse_colspec, default="4",
                        help="Item/type column (name or 1-based index). Default: 4.")
    parser.add_argument("--label_col", type=_parse_colspec, default=None,
                        help="Manual label column (name or 1-based index). If omitted, mapping is item-only.")
    parser.add_argument("--vg_col", type=_parse_colspec, default=None,
                        help="Vehicle grouping (VG) column (name or 1-based index). If omitted, mapping is item-only.")

    parser.add_argument("--tags_start_col", type=int, default=5,
                        help="1-based column index where USAS tag columns begin in the sample sheet. Default: 5.")

    parser.add_argument("--out_sheet_mapped", default="manual_with_usas_tags",
                        help="Sheet name for mapped manual list output.")
    parser.add_argument("--out_sheet_counts", default="usas_tag_counts",
                        help="Sheet name for long-format counts output.")

    args = parser.parse_args()

    in_path = Path(args.input_excel)
    manual_df = pd.read_excel(in_path, sheet_name=args.manual_sheet)
    sample_df = pd.read_excel(in_path, sheet_name=args.sample_sheet)

    # Build mappings from the sample sheet
    tags_map = build_mapping_unique_tags(
        sample_df,
        item_col=args.item_col,
        tags_start_col_1_based=args.tags_start_col,
        label_col=args.label_col,
        vg_col=args.vg_col,
    )

    counts_map = None
    if args.mode in ("count", "both"):
        counts_map = build_mapping_tag_counts(
            sample_df,
            item_col=args.item_col,
            tags_start_col_1_based=args.tags_start_col,
            label_col=args.label_col,
            vg_col=args.vg_col,
        )

    # Write outputs
    out_path = Path(args.output_excel)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        if args.mode in ("map", "both"):
            mapped = enrich_manual_list_with_tags(
                manual_df, tags_map,
                item_col=args.item_col,
                label_col=args.label_col,
                vg_col=args.vg_col,
            )
            mapped.to_excel(writer, sheet_name=args.out_sheet_mapped, index=False)

        if args.mode in ("count", "both") and counts_map is not None:
            counts_df = build_counts_long_table(
                manual_df, counts_map,
                item_col=args.item_col,
                label_col=args.label_col,
                vg_col=args.vg_col,
            )
            counts_df.to_excel(writer, sheet_name=args.out_sheet_counts, index=False)

    print("Done.")
    print(f"Output written to: {out_path}")


if __name__ == "__main__":
    main()
