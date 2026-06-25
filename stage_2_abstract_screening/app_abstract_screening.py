import os
import sys
from pathlib import Path
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

import shutil

# ============================================================
# Project paths
# ============================================================

STAGE_2_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAGE_2_DIR.parent

INPUT_DIR_2 = STAGE_2_DIR / "input"
OUTPUT_DIR_2 = STAGE_2_DIR / "output"
TEMP_DIR_2 = STAGE_2_DIR / "temp"

INPUT_DIR_2.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_2.mkdir(parents=True, exist_ok=True)
TEMP_DIR_2.mkdir(parents=True, exist_ok=True)

STAGE_3_DIR = PROJECT_ROOT / "stage_3_fulltext_initial_pairing"
STAGE_3_INPUT_DIR = STAGE_3_DIR / "input"
STAGE_3_INPUT_DIR.mkdir(parents=True, exist_ok=True)

MANUAL_PDF_DIR = (
    STAGE_2_DIR
    / "manual_pdf_completion"
)

MANUAL_PDF_PAPERS_DIR = (
    MANUAL_PDF_DIR
    / "pdfs"
)

MANUAL_PDF_COMPLETED_DIR = (
    MANUAL_PDF_DIR
    / "completed_input"
)

MANUAL_PDF_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

MANUAL_PDF_PAPERS_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

MANUAL_PDF_COMPLETED_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

if str(STAGE_2_DIR) not in sys.path:
    sys.path.insert(0, str(STAGE_2_DIR))


from abstract_screening_app_logic import (
    DEFAULT_ABSTRACT_COL,
    DEFAULT_MODEL,
    DEFAULT_SYSTEM_ROLE,
    DEFAULT_TITLE_COL,
    build_dashboard_data,
    get_prompt_preview,
    load_input_preview,
    run_abstract_screening,
)

load_dotenv()


def configure_page() -> None:
    st.set_page_config(page_title="Abstract Screening with LLM", layout="wide")

    st.title("Abstract Screening with LLM")
    st.write("Screen abstracts for the CT–programming scoping review.")

    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.5rem;}
        div[data-testid="stMetricValue"] {font-size: 1.4rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_settings() -> Dict[str, Any]:

    with st.sidebar:

        # ==================================================
        # Pipeline navigation context
        # ==================================================
        
        st.markdown("## Scoping Review Pipeline")

        st.success(
            "Current stage:\n\n"
            "**Stage 2 – Abstract screening**"
        )

        with st.expander("Pipeline stages", expanded=False):
            st.markdown(
                """
                **1. Import, merge & rule-based filtering**

                **2. Abstract screening ← current**

                **3. Full-text screening + initial pairing**

                **4. Full pairing analysis**

                **5. Characterization**

                **6. Results analysis**
                """
            )

        st.divider()

        # ==================================================
        # LLM settings
        # ==================================================

        with st.expander("LLM settings", expanded=True):

            api_key = st.text_input(
                "OpenAI API key",
                value=os.getenv("OPENAI_API_KEY", ""),
                type="password",
            )

            model = st.text_input(
                "Model",
                value=DEFAULT_MODEL
            )

            number_of_runs = st.number_input(
                "Number of LLM runs per abstract",
                min_value=1,
                max_value=10,
                value=3,
                step=1,
                help="Use 3 for majority voting. Use 1 to reduce cost.",
            )

        # ==================================================
        # Dataset settings
        # ==================================================

        with st.expander("Dataset settings", expanded=True):

            title_col = st.text_input(
                "Paper column name",
                value=DEFAULT_TITLE_COL
            )

            abstract_col = st.text_input(
                "Abstract column name",
                value=DEFAULT_ABSTRACT_COL
            )

            min_abstract_chars = st.number_input(
                "Minimum abstract length for LLM screening",
                min_value=0,
                max_value=500,
                value=40,
                step=5,
                help=(
                    "Records with shorter or missing abstracts are counted before "
                    "abstract screening and are not sent to the LLM."
                ),
            )

            deduplicate_titles = st.checkbox(
                "Safety duplicate check before abstract screening",
                value=False,
                help=(
                    "Keep this off when the file comes from Stage 1, because Stage 1 "
                    "is the authoritative deduplication stage. Turn it on only for custom files."
                ),
            )

        # ==================================================
        # Output settings
        # ==================================================

        with st.expander("Output settings", expanded=False):

            validation_sample_percent = st.number_input(
                "Validation sample percentage",
                min_value=1,
                max_value=100,
                value=15,
                step=1,
            )

            output_name = st.text_input(
                "Base output filename",
                value="abstracts_screened_llm.xlsx",
                help=(
                    "The app will create outputs using this base name: "
                    "_results, _author_validation, _llm_costs, "
                    "_next_stage_fulltext, and _checkpoint_raw."
                ),
            )

        # ==================================================
        # Full-text preparation
        # ==================================================

        with st.expander("Full-text preparation", expanded=False):

            include_undetermined_for_fulltext = st.checkbox(
                "Send Undetermined papers to full-text screening",
                value=True,
                help=(
                    "Recommended for scoping reviews. Included papers and "
                    "manual-review papers are always sent to the next stage. "
                    "This option also keeps Undetermined papers."
                ),
            )

            auto_retrieve_pdfs = st.checkbox(
                "Try automatic PDF retrieval for next stage",
                value=False,
                help=(
                    "If enabled, the app tries to download open-access PDFs "
                    "using PDF URLs/URLs and DOI + Unpaywall. Papers not found "
                    "are marked as Manual retrieval needed."
                ),
            )

            pdf_repository_dir = st.text_input(
                "PDF repository folder",
                value=str(STAGE_3_DIR / "papers_repository" / "pdfs"),
                help=(
                    "PDFs will be saved using the paper ID, "
                    "for example P0001.pdf or R00001.pdf."
                ),
            )

            unpaywall_email = st.text_input(
                "Unpaywall email for DOI-based OA lookup",
                value=os.getenv("UNPAYWALL_EMAIL", ""),
                help=(
                    "Required only if automatic PDF retrieval should use DOI + Unpaywall."
                ),
            )

    return {
        "api_key": api_key,
        "model": model,
        "title_col": title_col,
        "abstract_col": abstract_col,
        "min_abstract_chars": min_abstract_chars,
        "number_of_runs": number_of_runs,
        "validation_sample_percent": validation_sample_percent,
        "deduplicate_titles": deduplicate_titles,
        "output_name": output_name,
        "include_undetermined_for_fulltext": include_undetermined_for_fulltext,
        "auto_retrieve_pdfs": auto_retrieve_pdfs,
        "pdf_repository_dir": pdf_repository_dir,
        "unpaywall_email": unpaywall_email,
        "output_dir": str(OUTPUT_DIR_2),
        "temp_dir": str(TEMP_DIR_2),
    }


def render_file_uploaders() -> Tuple[Any, Any]:

    AUTO_INPUT_MODE = (
        "Automatically use Abstract_Screening_Input from Stage 1"
    )

    CUSTOM_INPUT_MODE = (
        "Use custom abstracts file"
    )

    col_file_1, col_file_2 = st.columns(2)

    with col_file_1:
        rules_file = st.file_uploader(
            "Upload abstract screening rules Excel file",
            type=["xlsx"],
        )

    with col_file_2:

        input_mode = st.radio(
            "Select input source",
            [
                AUTO_INPUT_MODE,
                CUSTOM_INPUT_MODE,
            ],
        )

        default_abstract_path = (
            INPUT_DIR_2
            / "abstract_screening_input.xlsx"
        )

        abstract_file = None

        # ============================================
        # Automatic Stage 1 → Stage 2 input
        # ============================================
        if input_mode == AUTO_INPUT_MODE:

            if default_abstract_path.exists():

                st.success(
                    "Abstract input automatically loaded "
                    "from Stage 1."
                )

                abstract_file = str(
                    default_abstract_path
                )

                st.caption(
                    f"Using: {default_abstract_path}"
                )

            else:

                st.warning(
                    "No Stage 1 abstract input found."
                )

                abstract_file = st.file_uploader(
                    "Upload abstracts Excel file",
                    type=["xlsx"],
                )

        # ============================================
        # Custom file upload
        # ============================================
        else:

            abstract_file = st.file_uploader(
                "Upload custom abstracts Excel file",
                type=["xlsx"],
            )

    st.info(
        "The rules, output schema, examples, and default prompt are read from the Excel template. "
        "The abstract file should contain paper titles and abstract text."
    )

    return rules_file, abstract_file


def render_input_preview(abstract_file) -> None:
    if abstract_file is None:
        return

    try:
        preview_df = load_input_preview(abstract_file)
        st.caption("Input preview")
        st.dataframe(preview_df, use_container_width=True)
    except Exception as e:
        st.error(f"Could not read input file preview: {e}")


def render_prompt_configuration(rules_file) -> tuple[str, str]:
    st.subheader("Prompt configuration")

    if rules_file is None:
        st.info(
            "Upload the abstract screening rules Excel file to preview the generated prompt "
            "from the Prompt_Config sheet."
        )
        return "", DEFAULT_SYSTEM_ROLE

    try:
        default_prompt_instructions, system_role = get_prompt_preview(rules_file)

        prompt_instructions = st.text_area(
            "Prompt generated from Prompt_Config",
            value=default_prompt_instructions,
            height=260,
            help=(
                "Edits here affect only this run. "
                "To permanently change the default prompt, edit Prompt_Config in the Excel template."
            ),
        )

        return prompt_instructions, system_role

    except Exception as e:
        st.warning(f"Could not read Prompt_Config yet: {e}")
        return "", DEFAULT_SYSTEM_ROLE


def show_duplicate_summary(deduplicate_titles: bool, duplicate_summary: pd.DataFrame) -> None:
    if deduplicate_titles and duplicate_summary is not None and not duplicate_summary.empty:
        st.subheader("Duplicate summary")
        st.write("Duplicate summary saved in the results workbook, sheet: Duplicate_Summary")
        st.dataframe(duplicate_summary, use_container_width=True)


def render_download_button(path: Path, label: str) -> None:
    if path and Path(path).exists():
        with open(path, "rb") as f:
            st.download_button(
                label=label,
                data=f,
                file_name=Path(path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


def show_results_dashboard(workflow_output: Dict[str, Any], number_of_runs: int) -> None:
    dashboard = build_dashboard_data(
        results=workflow_output["results"],
        number_of_runs=int(number_of_runs),
    )

    included_preview = dashboard["included_preview"]
    excluded_preview = dashboard["excluded_preview"]
    undetermined_preview = dashboard["undetermined_preview"]
    manual_review_preview = dashboard["manual_review_preview"]
    usage_costs_df = dashboard["usage_costs_df"]

    st.subheader("Run summary")

    metric_1, metric_2, metric_3, metric_4, metric_5 = st.columns(5)
    metric_1.metric("Included", len(included_preview))
    metric_2.metric("Excluded", len(excluded_preview))
    metric_3.metric("Undetermined", len(undetermined_preview))
    metric_4.metric("Manual review", len(manual_review_preview))
    metric_5.metric("Cost (USD)", f"${dashboard['total_cost']:.6f}")

    st.info(
        f"Results workbook: {workflow_output['results_output_path']}  |  "
        f"Validation workbook: {workflow_output['validation_output_path']}  |  "
        f"Costs workbook: {workflow_output['costs_output_path']}  |  "
        f"Next-stage workbook: {workflow_output['next_stage_output_path']}  |  "
        f"Checkpoint: {workflow_output['checkpoint_path']}"
    )


    with st.expander("Included", expanded=True):
        st.dataframe(included_preview, use_container_width=True)

    with st.expander("Excluded", expanded=False):
        st.dataframe(excluded_preview, use_container_width=True)

    with st.expander("Undetermined", expanded=True):
        st.dataframe(undetermined_preview, use_container_width=True)

    with st.expander("Manual review needed", expanded=False):
        st.dataframe(manual_review_preview, use_container_width=True)

    if "dataframes" in workflow_output and "Non_Screenable_Records" in workflow_output["dataframes"]:
        non_screenable_df = workflow_output["dataframes"]["Non_Screenable_Records"]
        with st.expander("Excluded before abstract screening: no usable abstract", expanded=False):
            st.dataframe(non_screenable_df, use_container_width=True)

    if "dataframes" in workflow_output and "Next_Stage_Fulltext" in workflow_output["dataframes"]:
        next_stage_df = workflow_output["dataframes"]["Next_Stage_Fulltext"]
        with st.expander("Next stage: full-text preparation", expanded=True):
            if not next_stage_df.empty and "PDF_Retrieval_Status" in next_stage_df.columns:
                retrieval_summary = (
                    next_stage_df["PDF_Retrieval_Status"]
                    .value_counts(dropna=False)
                    .reset_index()
                )
                retrieval_summary.columns = ["PDF_Retrieval_Status", "Count"]
                st.dataframe(retrieval_summary, use_container_width=True)
            st.dataframe(next_stage_df, use_container_width=True)

    with st.expander("Usage and costs", expanded=False):
        st.write(f"Total API runs: {dashboard['total_api_runs']}")
        st.write(f"Total tokens used: {dashboard['total_tokens']:,}")
        st.write(f"Estimated total cost: ${dashboard['total_cost']:.6f} USD")
        st.dataframe(usage_costs_df, use_container_width=True)


def main() -> None:
    configure_page()

    settings = render_sidebar_settings()

    rules_file, abstract_file = render_file_uploaders()
    render_input_preview(abstract_file)

    prompt_instructions, system_role = render_prompt_configuration(rules_file)

    stop_file = (
        TEMP_DIR_2
        / "STOP_ABSTRACT_SCREENING.txt"
    )

    col_run, col_stop, col_clear = st.columns(3)

    with col_run:
        run_button = st.button(
            "Run Abstract Screening",
            type="primary",
            key="run_abstract_screening_button",
        )

    with col_stop:
        stop_button = st.button(
            "Stop after current paper",
            key="stop_screening_button",
        )

    with col_clear:
        clear_stop_button = st.button(
            "Clear stop request",
            key="clear_stop_button",
        )

    if stop_button:
        stop_file.write_text("stop", encoding="utf-8")
        st.warning(
            "Stop requested. Screening will stop safely after the current paper."
        )

    if clear_stop_button:
        if stop_file.exists():
            stop_file.unlink()

        st.success("Stop request cleared.")

    if run_button:
        progress = st.progress(0)
        status_box = st.empty()
        counter_box = st.empty()

        def update_progress(value):
            progress.progress(value)
            counter_box.info(f"Progress: {value * 100:.1f}%")

        def update_log(message):
            status_box.info(message)

        try:
            if stop_file.exists():
                stop_file.unlink()

            workflow_output = run_abstract_screening(
                api_key=settings["api_key"],
                rules_file=rules_file,
                abstract_file=abstract_file,
                prompt_instructions=prompt_instructions,
                initial_system_role=system_role,
                settings=settings,
                log_callback=update_log,
                progress_callback=update_progress,
            )
            # Copy Stage 2 next-stage output to Stage 3 input
            stage3_input_path = (
            STAGE_3_INPUT_DIR
            / "fulltext_screening_input.xlsx"
            )

            source_path = Path(workflow_output["next_stage_output_path"])

            if source_path.exists():
                shutil.copy2(source_path, stage3_input_path)

        except Exception as e:
            st.error(str(e))
            st.stop()
        

        # Copy checkpoint to temp folder
        checkpoint_source = Path(workflow_output["checkpoint_path"])

        if checkpoint_source.exists():
            checkpoint_target = TEMP_DIR_2 / checkpoint_source.name

            if checkpoint_source.resolve() != checkpoint_target.resolve():
                shutil.copy2(checkpoint_source, checkpoint_target)
            
            workflow_output["checkpoint_path"] = checkpoint_target

    # ============================================================
    # Copy Stage 2 output to Stage 3 input
    # ============================================================

        next_stage_source = Path(workflow_output["next_stage_output_path"])

        stage3_input_path = (
            STAGE_3_INPUT_DIR
            / "fulltext_screening_input.xlsx"
        )


        st.subheader("Stage 3 handoff check")
        st.write(f"Source file: {next_stage_source}")
        st.write(f"Source exists: "f"{next_stage_source.exists()}")
        st.write(f"Destination folder: "f"{STAGE_3_INPUT_DIR}")
        st.write(f"Destination exists: "f"{STAGE_3_INPUT_DIR.exists()}")

        try:

            if not next_stage_source.exists():

                st.error(
                    "Stage 2 next-stage file "
                    "was not created.\n\n"
                    f"Expected:\n"
                    f"{next_stage_source}"
                )

            else:

                STAGE_3_INPUT_DIR.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                shutil.copy2(
                    next_stage_source,
                    stage3_input_path,
                )

                if stage3_input_path.exists():

                    st.success(
                        "Stage 2 output copied "
                        "successfully to Stage 3."
                    )

                    st.caption(
                        f"Saved as:\n"
                        f"{stage3_input_path}"
                    )

                else:

                    st.error(
                        "Copy command finished "
                        "but the Stage 3 file "
                        "does not exist."
                    )

        except Exception as e:

            st.error(
                f"Could not copy Stage 2 "
                f"output to Stage 3:\n{e}"
            )

        workflow_output[
            "stage3_input_path"
        ] = stage3_input_path


        st.info(
            f"Original records: "
            f"{workflow_output['original_count']} | "
            f"Records after deduplication: "
            f"{workflow_output['deduplicated_count']} | "
            f"Duplicates removed: "
            f"{workflow_output['removed_duplicates_count']} | "
            f"Non-screenable before abstract screening: "
            f"{workflow_output['non_screenable_count']} | "
            f"Sent to LLM abstract screening: "
            f"{workflow_output['abstract_screening_input_count']}"
        )
        st.success(
            f"""
        Abstract screening completed successfully.

        📁 Stage 2 output folder:
        {OUTPUT_DIR_2}

        Generated files copied:
        • Results workbook
        • Validation workbook
        • Costs workbook
        • Next-stage full-text workbook

        Checkpoint saved in temp:
        {workflow_output["checkpoint_path"]}

        Stage 3 input prepared:
        {workflow_output["stage3_input_path"]}
        """
        )

        show_duplicate_summary(settings["deduplicate_titles"], workflow_output["duplicate_summary"])

        show_results_dashboard(
            workflow_output=workflow_output,
            number_of_runs=int(settings["number_of_runs"]),
        )


if __name__ == "__main__":
    main()
