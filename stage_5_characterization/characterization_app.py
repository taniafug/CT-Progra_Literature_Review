import os
from pathlib import Path
from typing import List

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
STAGE_5_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAGE_5_DIR.parent

INPUT_DIR_5 = STAGE_5_DIR / "input"
OUTPUT_DIR_5 = STAGE_5_DIR / "output"
TEMP_DIR_5 = STAGE_5_DIR / "temp"

INPUT_DIR_5.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_5.mkdir(parents=True, exist_ok=True)
TEMP_DIR_5.mkdir(parents=True, exist_ok=True)

STAGE_4_OUTPUT_PATH = (
    PROJECT_ROOT
    / "stage_4_full_pairing"
    / "output"
    / "paper_level_framework_results.xlsx"
)

STAGE_5_INPUT_PATH = (
    INPUT_DIR_5
    / "characterization_input.xlsx"
)

from characterization_logic import (
    build_dependency_rules_prompt,
    build_prompt_config_text,
    build_schema_prompt,
    get_output_fields,
    get_system_role,
    load_excel_sheet_names,
    load_input_preview,
    make_json_schema,
    prepare_paper_list,
    read_coding_schema,
    read_dependency_rules,
    read_prompt_config,
    read_section_patterns,
    run_characterization,
    write_output_workbooks,
)

load_dotenv()

DEFAULT_MODEL = "gpt-4.1-mini"
DEFAULT_INPUT_SHEET = "Sheet1"


def select_default_index(options: List[str], preferred: str) -> int:
    return options.index(preferred) if preferred in options else 0


# ==================================================
# MAIN APP
# ==================================================

def main():
    st.set_page_config(page_title="Paper Characterization LLM", layout="wide")
    st.title("Paper Characterization with LLM")
    st.write("Upload your coding schema and paper list, then export characterization files.")

    with st.sidebar:
        st.header("Configuration")
        api_key = st.text_input("OpenAI API key", value=os.getenv("OPENAI_API_KEY", ""), type="password")
        model = st.text_input("Model", value=DEFAULT_MODEL)

        max_chars = st.number_input(
            "Maximum selected text length per paper (characters)",
            min_value=20_000,
            max_value=300_000,
            value=80_000,
            step=10_000,
        )

        include_optional_sections = st.checkbox(
            "Include optional sections",
            value=True,
            help="Uses Section_Patterns rows with Use_Mode='optional'.",
        )

        allow_fallback_sections = st.checkbox(
            "Allow fallback sections",
            value=True,
            help="Uses Section_Patterns rows with Use_Mode='fallback' only if primary/optional sections are weak.",
        )

        pause_seconds = st.number_input("Pause between papers (seconds)", min_value=0.0, max_value=30.0, value=0.5, step=0.5)

        st.subheader("Validation and costs")
        sample_percent = st.number_input("Validation sample percentage", min_value=0.0, max_value=100.0, value=15.0, step=1.0)
        random_seed = st.number_input("Random seed", min_value=0, max_value=1_000_000, value=42, step=1)
        input_price_per_1m = st.number_input("Input price per 1M tokens", min_value=0.0, value=0.0, step=0.01)
        output_price_per_1m = st.number_input("Output price per 1M tokens", min_value=0.0, value=0.0, step=0.01)

        st.subheader("Output files")
        results_output_name = st.text_input("Results workbook filename", value="characterization_results.xlsx")
        validation_output_name = st.text_input("Validation workbook filename", value="characterization_author_validation.xlsx")
        costs_output_name = st.text_input("Costs workbook filename", value="characterization_llm_costs.xlsx")
        analysis_input_output_name = st.text_input("Analysis input filename", value="results_characterization.xlsx")

    schema_file = st.file_uploader("Upload corrected characterization rules Excel file",type=["xlsx"])

    AUTO_INPUT_MODE = (
        "Automatically use Next_stage_included "
        "from Stage 3"
    )

    CUSTOM_INPUT_MODE = (
        "Use custom file with Paper_ID "
        "and PDF_path"
    )

    input_mode = st.radio(
        "Select input source",
        [
            AUTO_INPUT_MODE,
            CUSTOM_INPUT_MODE,
        ],
    )

    paper_list_file = None

    if input_mode == AUTO_INPUT_MODE:

        if STAGE_5_INPUT_PATH.exists():

            st.info(
                "Characterization input "
                "automatically loaded."
            )

            st.caption(f"Using: {STAGE_5_INPUT_PATH}")

            paper_list_file = str(
                STAGE_5_INPUT_PATH
            )

        elif STAGE_4_OUTPUT_PATH.exists():

            st.info(
                "Stage 4 output detected. "
                "It will be used as "
                "characterization input."
            )

            st.caption(f"Using: {STAGE_4_OUTPUT_PATH}")

            paper_list_file = str(STAGE_4_OUTPUT_PATH)

        else:

            st.warning("No Stage 3 output found.")

            paper_list_file = st.file_uploader(
                "Upload paper list Excel file",
                type=["xlsx"],
                help=(
                    "Use Stage 3 output, "
                    "usually fulltext_results.xlsx."
                ),
            )

    else:

        paper_list_file = st.file_uploader(
            "Upload custom characterization "
            "input file",
            type=["xlsx"],
            help=(
                "Required columns: "
                "Paper_ID, PDF_Path, "
                "Paper (optional), "
                "summary (optional)."
            ),
        )

    st.info(
        "The prompt is read from Prompt_Config. The schema is read from Coding_Schema_Characterization. "
        "Dependency rules and section patterns are also read from the Excel rules file."
    )

    available_columns: List[str] = []
    selected_sheet = ""

    if paper_list_file is not None:
        try:
            sheets = load_excel_sheet_names(paper_list_file)
            selected_sheet = st.selectbox("Sheet to use", options=sheets, index=select_default_index(sheets, DEFAULT_INPUT_SHEET))
            preview_df = load_input_preview(paper_list_file, selected_sheet)
            available_columns = preview_df.columns.tolist()
            st.caption("Input preview")
            st.dataframe(preview_df, use_container_width=True)
        except Exception as e:
            st.error(f"Could not read input preview: {e}")

    if available_columns:
        id_col = st.selectbox("Paper ID column", options=available_columns, index=select_default_index(available_columns, "Paper_ID"))
        pdf_path_candidates = [
            "PDF_Path",
            "PDF_path",
            "PDF Address",
            "PDF_Address",
        ]

        default_path_col = None

        for col in pdf_path_candidates:
            if col in available_columns:
                default_path_col = col
                break

        if default_path_col is None:
            default_path_col = available_columns[0]

        path_col = st.selectbox(
            "PDF path column",
            options=available_columns,
            index=select_default_index(
                available_columns,
                default_path_col
            )
        )

        st.write("Selected PDF path column:")
        st.write(path_col)
        title_col_options = [""] + available_columns
        title_col = st.selectbox("Paper title column (optional)", options=title_col_options, index=select_default_index(title_col_options, "Paper"))
        summary_col_options = [""] + available_columns
        summary_col = st.selectbox("Previous summary column (optional)", options=summary_col_options, index=select_default_index(summary_col_options, "summary"))
    else:
        id_col = "Paper_ID"
        path_col = "PDF_Address"
        title_col = ""
        summary_col = ""

    st.subheader("Prompt configuration preview")
    if schema_file is not None:
        try:
            prompt_config = read_prompt_config(schema_file)
            prompt_config_text = build_prompt_config_text(prompt_config)
            st.text_area("Prompt instructions read from Prompt_Config", value=prompt_config_text, height=240, disabled=True)
        except Exception as e:
            st.warning(f"Could not read Prompt_Config: {e}")
    else:
        st.info("Upload the corrected rules Excel file to preview Prompt_Config.")

    run_button = st.button("Run characterization", type="primary")

    if run_button:
        if not api_key:
            st.error("Please provide an OpenAI API key.")
            st.stop()
        if schema_file is None:
            st.error("Please upload the corrected characterization rules Excel file.")
            st.stop()
        if paper_list_file is None:
            st.error("Please upload the paper list Excel file.")
            st.stop()

        try:
            prompt_config = read_prompt_config(schema_file)
            prompt_config_text = build_prompt_config_text(prompt_config)
            system_role = get_system_role(prompt_config)
            schema_df = read_coding_schema(schema_file)
            dependency_rules = read_dependency_rules(schema_file)
            section_patterns_df = read_section_patterns(schema_file)
            schema_prompt = build_schema_prompt(schema_df)
            dependency_rules_prompt = build_dependency_rules_prompt(dependency_rules)
            output_fields = get_output_fields(schema_df)
            json_schema = make_json_schema(schema_df)
        except Exception as e:
            st.error(f"Error reading rules Excel file: {e}")
            st.stop()

        try:
            papers_df = prepare_paper_list(
                paper_list_file=paper_list_file,
                sheet_name=selected_sheet,
                id_col=id_col,
                path_col=path_col,
                title_col=title_col,
                summary_col=summary_col,
            )
        except Exception as e:
            st.error(f"Error reading paper list: {e}")
            st.stop()

        if papers_df.empty:
            st.error("The paper list is empty.")
            st.stop()

        st.info(f"Papers selected for characterization: {len(papers_df)}")

        client = OpenAI(api_key=api_key)
        progress = st.progress(0)
        log_box = st.empty()
        checkpoint_path = TEMP_DIR_5 / "characterization_checkpoint.xlsx"

        results = run_characterization(
            client=client,
            model=model,
            papers_df=papers_df,
            prompt_config_text=prompt_config_text,
            system_role=system_role,
            dependency_rules_prompt=dependency_rules_prompt,
            schema_prompt=schema_prompt,
            output_fields=output_fields,
            json_schema=json_schema,
            section_patterns_df=section_patterns_df,
            max_chars=int(max_chars),
            include_optional_sections=bool(include_optional_sections),
            allow_fallback_sections=bool(allow_fallback_sections),
            pause_seconds=float(pause_seconds),
            log_callback=log_box.info,
            progress_callback=progress.progress,
            checkpoint_path=checkpoint_path,
            resume_from_checkpoint=True,
            input_price_per_1m=float(input_price_per_1m),
            output_price_per_1m=float(output_price_per_1m),
            sample_fraction=float(sample_percent) / 100.0,
            random_seed=int(random_seed),
        )

        results_output_path = OUTPUT_DIR_5 / results_output_name
        validation_output_path = OUTPUT_DIR_5 / validation_output_name
        costs_output_path = OUTPUT_DIR_5 / costs_output_name
        analysis_input_output_path = OUTPUT_DIR_5 / analysis_input_output_name

        outputs = write_output_workbooks(
            results_output_path=results_output_path,
            validation_output_path=validation_output_path,
            costs_output_path=costs_output_path,
            analysis_input_output_path=analysis_input_output_path,
            results=results,
            input_price_per_1m=float(input_price_per_1m),
            output_price_per_1m=float(output_price_per_1m),
            sample_fraction=float(sample_percent) / 100.0,
            random_seed=int(random_seed),
        )

        st.success("Done. Final output workbooks saved successfully.")
        st.subheader("Summary")
        st.dataframe(outputs["summary"], use_container_width=True)
        st.subheader("Manual review needed")
        st.dataframe(outputs["manual_review"], use_container_width=True)
        st.subheader("Cost summary")
        st.dataframe(outputs["charges_summary"], use_container_width=True)

        st.success(
            f"""
        Characterization completed successfully.

        📁 Stage 5 output folder:
        {OUTPUT_DIR_5}

        Generated files:
        • {results_output_path.name}
        • {validation_output_path.name}
        • {costs_output_path.name}
        • {analysis_input_output_path.name}

        Checkpoint saved in temp:
        {checkpoint_path}
        """
        )

if __name__ == "__main__":
    main()