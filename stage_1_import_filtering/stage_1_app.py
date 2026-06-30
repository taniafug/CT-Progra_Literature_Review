import sys
from pathlib import Path
import shutil

import streamlit as st

# ============================================================
# Project paths
# ============================================================

STAGE_1_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAGE_1_DIR.parent

INPUT_DIR_1 = STAGE_1_DIR / "input"
OUTPUT_DIR_1 = STAGE_1_DIR / "output"

INPUT_DIR_1.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_1.mkdir(parents=True, exist_ok=True)

STAGE_2_DIR = PROJECT_ROOT / "stage_2_abstract_screening"
STAGE_2_INPUT_DIR = STAGE_2_DIR / "input"
STAGE_2_INPUT_DIR.mkdir(parents=True, exist_ok=True)

if str(STAGE_1_DIR) not in sys.path:
    sys.path.insert(0, str(STAGE_1_DIR))

from stage_1_import_filtering.stage_1_logic import (
    DEFAULT_DATABASES,
    run_stage1_pipeline,
    write_stage1_output_workbooks,
)


def main():

    # ============================================================
    # Page setup
    # ============================================================

    st.set_page_config(
        page_title="Stage 1 – Import, Merge & Rule-based Filtering",
        layout="wide",
    )

    st.title("Stage 1 – Import, Merge & Rule-based Filtering")

    st.write(
        "Import database exports, standardize records, remove duplicates, "
        "apply rule-based pre-screening, and prepare the abstract screening input."
    )

    # ============================================================
    # Sidebar
    # ============================================================

    with st.sidebar:

        st.header("Stage 1 settings")

        selected_databases = st.multiselect(
            "Select databases to process",
            options=DEFAULT_DATABASES + ["Other"],
            default=DEFAULT_DATABASES,
            key="selected_databases",
        )

        custom_database = ""

        if "Other" in selected_databases:
            custom_database = st.text_input(
                "Additional database name",
                key="custom_database",
            )

        output_folder = st.text_input(
            "Output folder",
            value=str(OUTPUT_DIR_1),
            key="output_folder",
            help="Final Stage 1 workbooks will be saved here.",
        )

        st.caption(f"Default input folder: `{INPUT_DIR_1}`")


    # ============================================================
    # Main interface
    # ============================================================

    database_names = [
        db for db in selected_databases
        if db != "Other"
    ]

    if custom_database.strip():
        database_names.append(custom_database.strip())

    st.info(
        "Upload one Excel file per database. Files may contain several sheets, "
        "such as evaluation, intervention, secondary evaluation, or preprint searches."
    )

    uploaded_files = {}

    upload_columns = st.columns(2)

    for index, database in enumerate(database_names):
        with upload_columns[index % 2]:

            uploaded_files[database] = st.file_uploader(
                f"Upload Excel file for {database}",
                type=["xlsx", "xls", "ods"],
                key=f"upload_{database}",
            )

    run_button = st.button(
        "Run Stage 1",
        type="primary",
        key="run_stage1_button",
    )

    # ============================================================
    # Run Stage 1
    # ============================================================

    if run_button:

        uploaded_files = {
            database: file
            for database, file in uploaded_files.items()
            if file is not None
        }

        if not uploaded_files:
            st.error("Please upload at least one database file.")
            st.stop()

        try:
            with st.spinner("Processing Stage 1..."):

                results = run_stage1_pipeline(
                    uploaded_files_by_database=uploaded_files
                )

                output_dir = Path(output_folder)
                output_dir.mkdir(parents=True, exist_ok=True)

                generated_paths = write_stage1_output_workbooks(
                    results=results,
                    output_dir=output_dir,
                )

                stage2_input_path = (
                    STAGE_2_INPUT_DIR
                    / "stage_2_input.xlsx"
                )

                shutil.copy2(
                    generated_paths["abstract_input_path"],
                    stage2_input_path,
                )

        except Exception as error:
            st.exception(error)
            st.stop()

        st.success("Stage 1 completed successfully.")

        # ========================================================
        # Metrics
        # ========================================================

        metric_1, metric_2, metric_3, metric_4 = st.columns(4)

        metric_1.metric(
            "Imported records",
            len(results["standardized_all_records"]),
        )

        metric_2.metric(
            "Exact duplicates removed",
            int(
                (
                    results["deduplicated_marked"]["Duplicate_Status"]
                    == "Duplicate"
                ).sum()
            ),
        )

        metric_3.metric(
            "Non-screenable records",
            len(results["non_screenable_records"]),
        )

        metric_4.metric(
            "Sent to abstract screening",
            len(results["stage_2_input"]),
        )

        # ========================================================
        # Output files
        # ========================================================

        st.subheader("Generated files")

        st.success(
            f"""
Stage 1 completed successfully.

📁 Stage 1 output folder:
{output_dir}

Generated files:
• {generated_paths['processing_path'].name}

• {generated_paths['title_path'].name}

• {generated_paths['abstract_input_path'].name}

📥 Automatically copied to Stage 2 input:
{stage2_input_path}
"""
        )

        st.info(
            "The abstract screening input has been automatically copied "
            "to Stage 2 input folder."
        )

        # ========================================================
        # Preview tables
        # ========================================================

        st.subheader("PRISMA counts")

        st.dataframe(
            results["prisma_counts"].astype(str),
            use_container_width=True,
        )

        st.subheader("Database statistics")

        st.dataframe(
            results["database_statistics"].astype(str),
            use_container_width=True,
        )

        with st.expander("Exact duplicates", expanded=False):
            st.dataframe(
                results["exact_duplicates"].astype(str),
                use_container_width=True,
            )

        with st.expander("Non-screenable records", expanded=False):
            st.dataframe(
                results["non_screenable_records"].astype(str),
                use_container_width=True,
            )

        with st.expander("Manual review needed", expanded=False):
            st.dataframe(
                results["manual_review_needed"].astype(str),
                use_container_width=True,
            )

        with st.expander("Title screening results", expanded=True):
            st.dataframe(
                results["title_screening_results"].astype(str),
                use_container_width=True,
            )

        with st.expander("Abstract screening input", expanded=True):
            st.dataframe(
                results["stage_2_input"].astype(str),
                use_container_width=True,
            )


if __name__ == "__main__":
    main()