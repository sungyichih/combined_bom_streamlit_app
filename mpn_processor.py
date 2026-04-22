
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd


REQUIRED_COLUMNS = ["Material NO", "MFR. Name", "MFR. P/N"]


@dataclass
class ProcessResult:
    mapping_df: pd.DataFrame
    duplicate_df: pd.DataFrame
    summary: dict
    ignored_incomplete_rows: int


def _normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def detect_header_row(excel_file) -> int:
    preview = pd.read_excel(excel_file, sheet_name=0, header=None, dtype=str, nrows=25)
    for idx, row in preview.iterrows():
        labels = {_normalize_text(v).lower() for v in row.tolist()}
        if all(col.lower() in labels for col in REQUIRED_COLUMNS):
            return int(idx)
    raise ValueError(
        "找不到必要欄位：Material NO / MFR. Name / MFR. P/N。請確認來源檔格式是否與系統匯出檔一致。"
    )


def load_source_data(excel_file) -> pd.DataFrame:
    header_row = detect_header_row(excel_file)
    df = pd.read_excel(excel_file, sheet_name=0, header=header_row, dtype=str)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    data = df[REQUIRED_COLUMNS].copy()
    data.columns = ["SPN", "Manufacturer", "MPN"]

    for col in data.columns:
        data[col] = data[col].map(_normalize_text)

    ignored_incomplete_rows = int(((data["SPN"] == "") | (data["MPN"] == "")).sum())

    data = data[(data["SPN"] != "") & (data["MPN"] != "")].copy()
    data.reset_index(drop=True, inplace=True)
    return data, ignored_incomplete_rows


def build_mapping(data: pd.DataFrame) -> pd.DataFrame:
    mapping_df = data.copy()

    spn_to_mpn_count = mapping_df.groupby("SPN")["MPN"].nunique()
    mpn_to_spn_count = mapping_df.groupby("MPN")["SPN"].nunique()

    mapping_df["exact_duplicate"] = mapping_df.duplicated(["SPN", "MPN"], keep=False)
    mapping_df["duplicate_mpn"] = mapping_df["MPN"].map(mpn_to_spn_count).fillna(0).astype(int) > 1
    mapping_df["multi_mpn_spn"] = mapping_df["SPN"].map(spn_to_mpn_count).fillna(0).astype(int) > 1

    def build_remarks(row) -> str:
        flags = []
        if row["exact_duplicate"]:
            flags.append("Exact Duplicate")
        if row["duplicate_mpn"]:
            flags.append("⚠ Duplicate MPN")
        if row["multi_mpn_spn"]:
            flags.append("Multi-MPN SPN")
        return " | ".join(flags)

    mapping_df["Remarks"] = mapping_df.apply(build_remarks, axis=1)

    return mapping_df[
        ["SPN", "Manufacturer", "MPN", "Remarks", "exact_duplicate", "duplicate_mpn", "multi_mpn_spn"]
    ]


def build_duplicate_mpn_table(mapping_df: pd.DataFrame) -> pd.DataFrame:
    duplicate_mpns = (
        mapping_df.groupby("MPN")["SPN"]
        .nunique()
        .sort_values(ascending=False)
    )
    duplicate_mpns = duplicate_mpns[duplicate_mpns > 1]

    rows = []
    max_spn_count = 0

    base_cols_df = mapping_df[["MPN", "SPN", "Manufacturer"]].drop_duplicates().copy()

    for mpn in duplicate_mpns.index:
        pairs = (
            base_cols_df[base_cols_df["MPN"] == mpn]
            .sort_values(["SPN", "Manufacturer"])
            .drop_duplicates(subset=["SPN", "Manufacturer"])
        )
        record = {"MPN": mpn}
        max_spn_count = max(max_spn_count, len(pairs))
        for idx, (_, item) in enumerate(pairs.iterrows(), start=1):
            record[f"SPN {idx}"] = item["SPN"]
            record[f"Manufacturer {idx}"] = item["Manufacturer"]
        record["Note"] = f"{len(pairs)} SPNs share this MPN"
        rows.append(record)

    if not rows:
        return pd.DataFrame(columns=["MPN", "Note"])

    columns = ["MPN"]
    for idx in range(1, max_spn_count + 1):
        columns.extend([f"SPN {idx}", f"Manufacturer {idx}"])
    columns.append("Note")

    duplicate_df = pd.DataFrame(rows)
    for col in columns:
        if col not in duplicate_df.columns:
            duplicate_df[col] = ""
    duplicate_df = duplicate_df[columns].sort_values("MPN").reset_index(drop=True)
    return duplicate_df


def build_summary(mapping_df: pd.DataFrame, ignored_incomplete_rows: int) -> dict:
    duplicate_mpn_mask = mapping_df["duplicate_mpn"]
    summary = {
        "total_records": int(len(mapping_df)),
        "unique_spns": int(mapping_df["SPN"].nunique()),
        "unique_mpns": int(mapping_df["MPN"].nunique()),
        "spns_with_multiple_mpns": int((mapping_df.groupby("SPN")["MPN"].nunique() > 1).sum()),
        "duplicate_mpns": int((mapping_df.groupby("MPN")["SPN"].nunique() > 1).sum()),
        "rows_affected_by_duplicate_mpn": int(duplicate_mpn_mask.sum()),
        "exact_duplicate_rows": int(mapping_df["exact_duplicate"].sum()),
        "ignored_incomplete_rows": int(ignored_incomplete_rows),
    }
    return summary


def process_excel(excel_file) -> ProcessResult:
    data, ignored_incomplete_rows = load_source_data(excel_file)
    mapping_df = build_mapping(data)
    duplicate_df = build_duplicate_mpn_table(mapping_df)
    summary = build_summary(mapping_df, ignored_incomplete_rows)
    return ProcessResult(
        mapping_df=mapping_df,
        duplicate_df=duplicate_df,
        summary=summary,
        ignored_incomplete_rows=ignored_incomplete_rows,
    )


def create_excel_report(result: ProcessResult) -> bytes:
    output = BytesIO()

    export_mapping = result.mapping_df[["SPN", "Manufacturer", "MPN", "Remarks"]].copy().fillna("")
    export_duplicates = result.duplicate_df.copy().fillna("")

    summary_rows = [
        ("Total SPN-MPN records", result.summary["total_records"]),
        ("Unique SPNs", result.summary["unique_spns"]),
        ("Unique MPNs", result.summary["unique_mpns"]),
        ("SPNs with multiple MPNs", result.summary["spns_with_multiple_mpns"]),
        ("Duplicate MPNs (same MPN → multiple SPNs)", result.summary["duplicate_mpns"]),
        ("Rows affected by duplicate MPN", result.summary["rows_affected_by_duplicate_mpn"]),
        ("Exact duplicate rows (same SPN + MPN)", result.summary["exact_duplicate_rows"]),
        ("Ignored incomplete rows", result.summary["ignored_incomplete_rows"]),
    ]
    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        export_mapping.to_excel(writer, sheet_name="SPN-MPN Mapping", index=False)
        export_duplicates.to_excel(writer, sheet_name="Duplicate MPNs", index=False)

        workbook = writer.book
        ws_summary = writer.sheets["Summary"]
        ws_mapping = writer.sheets["SPN-MPN Mapping"]
        ws_duplicate = writer.sheets["Duplicate MPNs"]

        fmt_title = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "bg_color": "#1F4E78",
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
        })
        fmt_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1,
            "text_wrap": True,
            "valign": "vcenter",
        })
        fmt_cell = workbook.add_format({"border": 1, "valign": "vcenter"})
        fmt_multi = workbook.add_format({"bg_color": "#FFF2CC", "border": 1})
        fmt_dup = workbook.add_format({"bg_color": "#F4CCCC", "border": 1})
        fmt_exact = workbook.add_format({"bg_color": "#EA9999", "border": 1})
        fmt_legend_label = workbook.add_format({"bold": True})

        ws_summary.merge_range("A1:B1", "SPN-MPN Mapping Summary", fmt_title)
        ws_summary.set_row(0, 24)
        ws_summary.set_column("A:A", 42)
        ws_summary.set_column("B:B", 18)
        ws_summary.write("A12", "Color Legend", fmt_legend_label)
        ws_summary.write("A13", "🟡", fmt_legend_label)
        ws_summary.write("B13", "Multi-MPN SPN — one SPN maps to multiple approved MPNs")
        ws_summary.write("A14", "🔴", fmt_legend_label)
        ws_summary.write("B14", "Duplicate MPN — same MPN maps to multiple SPNs")
        ws_summary.write("A15", "🔴🔴", fmt_legend_label)
        ws_summary.write("B15", "Exact Duplicate — identical SPN + MPN row appears more than once")
        for col_num, value in enumerate(summary_df.columns.values):
            ws_summary.write(2, col_num, value, fmt_header)
        for row_num in range(3, 3 + len(summary_df)):
            ws_summary.write(row_num, 0, summary_df.iloc[row_num - 3, 0], fmt_cell)
            ws_summary.write(row_num, 1, summary_df.iloc[row_num - 3, 1], fmt_cell)

        ws_mapping.freeze_panes(1, 0)
        ws_mapping.set_column("A:A", 18)
        ws_mapping.set_column("B:B", 26)
        ws_mapping.set_column("C:C", 28)
        ws_mapping.set_column("D:D", 36)
        for idx, col in enumerate(export_mapping.columns):
            ws_mapping.write(0, idx, col, fmt_header)

        for row_idx, (_, row) in enumerate(result.mapping_df.iterrows(), start=1):
            if row["exact_duplicate"]:
                row_fmt = fmt_exact
            elif row["duplicate_mpn"]:
                row_fmt = fmt_dup
            elif row["multi_mpn_spn"]:
                row_fmt = fmt_multi
            else:
                row_fmt = fmt_cell

            ws_mapping.write(row_idx, 0, row["SPN"], row_fmt)
            ws_mapping.write(row_idx, 1, row["Manufacturer"], row_fmt)
            ws_mapping.write(row_idx, 2, row["MPN"], row_fmt)
            ws_mapping.write(row_idx, 3, row["Remarks"], row_fmt)

        ws_duplicate.freeze_panes(1, 0)
        for idx, col in enumerate(export_duplicates.columns):
            ws_duplicate.write(0, idx, col, fmt_header)
            width = 26 if "Manufacturer" in col else 20
            if col == "MPN":
                width = 28
            if col == "Note":
                width = 22
            ws_duplicate.set_column(idx, idx, width)
        for row_idx in range(1, len(export_duplicates) + 1):
            for col_idx in range(len(export_duplicates.columns)):
                ws_duplicate.write(row_idx, col_idx, export_duplicates.iloc[row_idx - 1, col_idx], fmt_cell)

    output.seek(0)
    return output.getvalue()
