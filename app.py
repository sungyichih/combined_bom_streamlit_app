
from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

import cpn_processor
import mpn_processor
import bom_mapping_core as bom


st.set_page_config(
    page_title="SAP BOM Full Mapping Tool",
    page_icon="📦",
    layout="wide",
)


def reset_file(file_obj):
    """Safely reset Streamlit uploaded file pointer before another read."""
    if file_obj is not None:
        try:
            file_obj.seek(0)
        except Exception:
            pass


def make_system_mpn_df(mpn_result: mpn_processor.ProcessResult) -> pd.DataFrame:
    """Convert MPN organizer result into the format needed by BOM mapping."""
    df = mpn_result.mapping_df[["SPN", "Manufacturer", "MPN"]].copy()
    df.columns = ["SPN", "System_MFG", "System_MPN"]
    return df.drop_duplicates().reset_index(drop=True)


def make_cpn_mapping_df(cpn_result: cpn_processor.ProcessingResult) -> pd.DataFrame:
    """Convert CPN organizer result into the format needed by BOM mapping."""
    df = cpn_result.mapping_df[["SPN", "CPN"]].copy()
    return df.drop_duplicates().reset_index(drop=True)


def render_status_legend():
    st.markdown(
        """
        **MPN Compare color priority**

        - 🔴 **Missing SPN** — highest priority
        - 🟥 **Ambiguous - same overlap**
        - 🟧 **Missing in System**
        - 🟠 **Extra in System**
        - 🟨 **Partial Match**
        - ✅ **Full Match** — no highlight
        """
    )


def compare_row_style(row: pd.Series):
    status = row.get("MPN_Compare_Status", "")
    color_map = {
        "Missing SPN": "#C00000",
        "Ambiguous - same overlap": "#E06666",
        "Missing in System": "#F4B183",
        "Extra in System": "#FCE5CD",
        "Partial Match": "#FFF2CC",
    }
    color = color_map.get(status, "")
    if not color:
        return [""] * len(row)
    text_color = "color: white;" if status == "Missing SPN" else ""
    return [f"background-color: {color}; {text_color}" for _ in row]


def cpn_row_style(row: pd.Series):
    if row.get("Is Duplicate CPN", False):
        return ["background-color: #FDE2E1"] * len(row)
    if row.get("Is Multi-CPN SPN", False):
        return ["background-color: #FFF4CC"] * len(row)
    return [""] * len(row)


def mpn_row_style(row: pd.Series):
    if row.get("exact_duplicate", False):
        return ["background-color: #EA9999"] * len(row)
    if row.get("duplicate_mpn", False):
        return ["background-color: #F4CCCC"] * len(row)
    if row.get("multi_mpn_spn", False):
        return ["background-color: #FFF2CC"] * len(row)
    return [""] * len(row)


def main():
    st.title("📦 SAP BOM Full Mapping Tool")
    st.caption(
        "Upload the system CPN/SPN data, system MPN/SPN data, and customer Original BOM. "
        "The tool will organize both system files first, then run the final BOM mapping comparison."
    )

    with st.expander("Expected input files / sheets", expanded=False):
        st.markdown(
            """
            **1. System CPN/SPN data**
            - Required source columns: `Material` and `Customer material no.`
            - Slash `/` inside CPN is automatically split into multiple CPN rows.

            **2. System MPN/SPN data**
            - Required source columns: `Material NO`, `MFR. Name`, and `MFR. P/N`

            **3. Customer Original BOM**
            - Sheet name must be `BOM`
            - Default data start row is `2`
            - Columns are expected as:
              - A = Customer CPN
              - B = Description
              - C = Qty
              - D = Location
              - E/F = Primary MFG / MPN
              - G/H, I/J... = Alternate MFG / MPN
            """
        )

    col1, col2, col3 = st.columns(3)

    with col1:
        system_cpn_file = st.file_uploader(
            "1️⃣ Upload System CPN/SPN Data",
            type=["xlsx", "xls", "xlsm"],
            key="system_cpn",
        )

    with col2:
        system_mpn_file = st.file_uploader(
            "2️⃣ Upload System MPN/SPN Data",
            type=["xlsx", "xls", "xlsm"],
            key="system_mpn",
        )

    with col3:
        original_bom_file = st.file_uploader(
            "3️⃣ Upload Customer Original BOM",
            type=["xlsx", "xls", "xlsm"],
            key="original_bom",
        )
        bom_start_row = st.number_input(
            "Original BOM data starts at row",
            min_value=1,
            value=2,
            step=1,
        )

    process = st.button("🚀 Run Full Process", type="primary", use_container_width=True)

    if not process:
        st.info("Upload all 3 files, then click **Run Full Process**.")
        return

    if system_cpn_file is None or system_mpn_file is None or original_bom_file is None:
        st.error("Please upload all 3 files first.")
        return

    try:
        progress = st.progress(0, text="Step 1/3: Organizing CPN/SPN data...")
        reset_file(system_cpn_file)
        cpn_result = cpn_processor.process_uploaded_file(system_cpn_file)
        organized_cpn_bytes = cpn_result.export_bytes
        organized_cpn_df = make_cpn_mapping_df(cpn_result)

        progress.progress(35, text="Step 2/3: Organizing MPN/SPN data...")
        reset_file(system_mpn_file)
        mpn_result = mpn_processor.process_excel(system_mpn_file)
        organized_mpn_bytes = mpn_processor.create_excel_report(mpn_result)
        organized_mpn_df = make_system_mpn_df(mpn_result)

        progress.progress(70, text="Step 3/3: Running Original BOM mapping comparison...")
        reset_file(original_bom_file)
        original_base_df, original_mpn_df = bom.read_original_bom(
            original_bom_file,
            sheet_name="BOM",
            data_start_row=int(bom_start_row),
        )

        mapped_df = bom.map_cpn_to_spn(
            original_base_df,
            original_mpn_df,
            organized_cpn_df,
            organized_mpn_df,
        )

        compare_df = bom.build_mpn_compare(
            mapped_df,
            original_mpn_df,
            organized_mpn_df,
        )

        missing_spn_df = bom.build_missing_spn_list(
            mapped_df,
            original_mpn_df,
        )

        summary_df = bom.build_summary(
            original_base_df,
            mapped_df,
            compare_df,
        )

        mapping_output = bom.make_result_excel(
            original_base_df,
            original_mpn_df,
            mapped_df,
            compare_df,
            missing_spn_df,
            summary_df,
        )

        mapping_bytes = mapping_output.getvalue()

        progress.progress(100, text="Completed.")
        st.success("All files processed successfully.")

    except Exception as exc:
        st.error(f"Processing failed: {exc}")
        return

    st.subheader("⬇️ Download Results")
    d1, d2, d3 = st.columns(3)

    with d1:
        st.download_button(
            "Download Organized CPN/SPN",
            data=organized_cpn_bytes,
            file_name="organized_cpn_spn.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d2:
        st.download_button(
            "Download Organized MPN/SPN",
            data=organized_mpn_bytes,
            file_name="organized_mpn_spn.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d3:
        st.download_button(
            "Download Final Mapping Result",
            data=mapping_bytes,
            file_name="original_bom_mapping_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.subheader("📌 Final Mapping Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Original BOM rows", len(original_base_df))
    m2.metric("Missing SPN", int((mapped_df["Selection_Status"] == "Missing SPN").sum()))
    m3.metric("Ambiguous", int((mapped_df["Selection_Status"] == "Ambiguous - same overlap").sum()))
    m4.metric("Extra in System", int((compare_df["MPN_Compare_Status"] == "Extra in System").sum()))
    m5.metric(
        "MPN differences",
        int(
            compare_df["MPN_Compare_Status"].isin(
                [
                    "Partial Match",
                    "Missing in System",
                    "Extra in System",
                    "Missing SPN",
                    "Ambiguous - same overlap",
                ]
            ).sum()
        ),
    )

    render_status_legend()

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "Final MPN Compare",
            "CPN → SPN Map",
            "Missing SPN",
            "Organized CPN/SPN",
            "Organized MPN/SPN",
            "Summary",
        ]
    )

    with tab1:
        st.subheader("Final MPN Compare")
        diff_only = st.checkbox("Show differences only", value=True)
        display_compare = compare_df.copy()
        if diff_only:
            display_compare = display_compare[display_compare["MPN_Compare_Status"] != "Full Match"]

        st.dataframe(
            display_compare.style.apply(compare_row_style, axis=1),
            use_container_width=True,
            height=520,
            hide_index=True,
        )

    with tab2:
        st.subheader("CPN → SPN Mapping")
        st.dataframe(mapped_df, use_container_width=True, height=420, hide_index=True)

    with tab3:
        st.subheader("Missing SPN / Category Suggestion")
        st.dataframe(missing_spn_df, use_container_width=True, height=420, hide_index=True)

    with tab4:
        st.subheader("Organized CPN/SPN")
        cpn_view = cpn_result.mapping_df[
            ["SPN", "CPN", "Remarks", "Is Duplicate CPN", "Is Multi-CPN SPN"]
        ].copy()
        st.dataframe(
            cpn_view.style.apply(cpn_row_style, axis=1),
            use_container_width=True,
            height=420,
            hide_index=True,
        )

    with tab5:
        st.subheader("Organized MPN/SPN")
        st.dataframe(
            mpn_result.mapping_df.style.apply(mpn_row_style, axis=1),
            use_container_width=True,
            height=420,
            hide_index=True,
        )

    with tab6:
        st.subheader("Summary")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
