import os
import time
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI


STAGE_3_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAGE_3_DIR.parent

INPUT_DIR_3 = STAGE_3_DIR / "input"
OUTPUT_DIR_3 = STAGE_3_DIR / "output"
TEMP_DIR_3 = STAGE_3_DIR / "temp"

INPUT_DIR_3.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_3.mkdir(parents=True, exist_ok=True)
TEMP_DIR_3.mkdir(parents=True, exist_ok=True)

PAPERS_REPOSITORY_DIR = (STAGE_3_DIR/ "papers_repository")

PAPERS_PDF_DIR = (PAPERS_REPOSITORY_DIR/ "pdfs")

PAPERS_PDF_DIR.mkdir(parents=True,exist_ok=True,)

NO_FULL_TEXT_PATH = (OUTPUT_DIR_3/ "no_full_text_available.xlsx")

from screening_logic import (
    append_manual_reason,
    build_prompt_instructions,
    build_relevant_text_for_screening,
    build_schema_prompt,
    classify_fulltext,
    get_output_fields,
    load_pdf_text,
    make_json_schema,
    postprocess_classification,
    prefix_input_columns,
    read_examples,
    read_output_schema,
    read_prompt_config,
    read_rules,
    read_section_patterns,
    remove_duplicates_keep_first,
    reset_excel_pointer,
    safe_str,
    write_results_excel,
    write_output_workbooks,
)

load_dotenv()


DEFAULT_MODEL = "gpt-4.1-mini"
DEFAULT_TITLE_COL = "Paper"
DEFAULT_PDF_PATH_COL = "PDF_Path"
DEFAULT_PREVIOUS_FILTER_COL = "Abstract_Decision"
DEFAULT_RESULTS_OUTPUT_NAME = "fulltext_results.xlsx"
DEFAULT_VALIDATION_OUTPUT_NAME = "fulltext_author_validation.xlsx"
DEFAULT_COSTS_OUTPUT_NAME = "fulltext_llm_costs.xlsx"


# ============================================================
# UI helpers
# ============================================================
def select_default_index(options: List[str], preferred: str) -> int:
    return options.index(preferred) if preferred in options else 0


def load_excel_sheet_names(input_file) -> List[str]:
    if isinstance(input_file, str):
        excel = pd.ExcelFile(input_file)
        return excel.sheet_names

    reset_excel_pointer(input_file)
    excel = pd.ExcelFile(input_file)
    reset_excel_pointer(input_file)
    return excel.sheet_names


def load_input_preview(input_file, sheet_name: str) -> pd.DataFrame:
    if isinstance(input_file, str):
        preview = pd.read_excel(
            input_file,
            sheet_name=sheet_name,
            nrows=5,
            dtype=str,
        )
    else:
        reset_excel_pointer(input_file)
        preview = pd.read_excel(
            input_file,
            sheet_name=sheet_name,
            nrows=5,
            dtype=str,
        )
        reset_excel_pointer(input_file)

    preview.columns = [str(c).strip() for c in preview.columns]
    return preview


def column_selectbox(
    label: str,
    columns: List[str],
    preferred: str,
    help_text: str = "",
    key: str | None = None,
) -> str:
    return st.selectbox(
        label,
        options=columns,
        index=select_default_index(columns, preferred),
        help=help_text,
        key=key,
    )


def make_checkpoint_key(paper: str, pdf_path: str) -> str:
    return f"{str(paper).strip().lower()}||{str(pdf_path).strip()}"


def validate_pdf_paths(input_file, sheet_name: str, pdf_path_col: str) -> tuple[bool, pd.DataFrame]:

    df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    if pdf_path_col not in df.columns:
        return False, pd.DataFrame(
            [{"Issue": f"PDF path column '{pdf_path_col}' not found."}]
        )

    rows = []

    for index, row in df.iterrows():
        path_text = str(row.get(pdf_path_col, "")).strip()

        if not path_text or path_text.lower() in ["nan", "none", "null", "0", "no access"]:
            rows.append(
                {
                    "Row": index + 2,
                    "Code": row.get("Full_Text_ID", ""),
                    "PDF_Filename": row.get("PDF_Filename", ""),
                    "Paper": row.get("Paper", ""),
                    "PDF_path": path_text,
                    "Issue": "Missing PDF path",
                }
            )
            continue

        pdf_path = Path(path_text).expanduser()

        if not pdf_path.exists():
            rows.append(
                {
                    "Row": index + 2,
                    "Code": row.get("Full_Text_ID", ""),
                    "PDF_Filename": row.get("PDF_Filename", ""),
                    "Paper": row.get("Paper", ""),
                    "PDF_path": path_text,
                    "Issue": "PDF file not found",
                }
            )

    issues_df = pd.DataFrame(rows)

    return issues_df.empty, issues_df

# ============================================================
# Streamlit app
# ============================================================
def main(): 
    st.set_page_config(page_title="Full-Text Screening with LLM", layout="wide")
    st.title("Full-Text Screening with LLM")
    st.write("Screen selected full-text sections for the CT–programming scoping review.")

    with st.sidebar:
        st.header("Configuration")

        api_key = st.text_input(
            "OpenAI API key",
            value=os.getenv("OPENAI_API_KEY", ""),
            type="password",
        )

        model = st.text_input("Model", value=DEFAULT_MODEL)

        max_chars = st.number_input(
            "Maximum selected-section characters sent to the model",
            min_value=20_000,
            max_value=250_000,
            value=120_000,
            step=10_000,
        )

        deduplicate_titles = st.checkbox("Remove duplicate titles before full-text screening", value=False)

        manual_sample_percent = st.number_input(
            "Random sample percentage for author/manual validation by decision category",
            min_value=0.0,
            max_value=100.0,
            value=15.0,
            step=1.0,
        )

        random_seed = st.number_input(
            "Random seed for reproducible validation samples",
            min_value=0,
            max_value=1_000_000,
            value=42,
            step=1,
        )

        st.subheader("LLM charge settings")
        input_price_per_1m = st.number_input(
            "Input price per 1M tokens",
            min_value=0.0,
            value=0.40,
            step=0.01,
        )

        output_price_per_1m = st.number_input(
            "Output price per 1M tokens",
            min_value=0.0,
            value=1.60,
            step=0.01,
        )

        results_output_name = st.text_input(
            "Results Excel filename",
            value=DEFAULT_RESULTS_OUTPUT_NAME,
            help="Main workbook: full results, decision summary, included/excluded/undetermined, and Next_stage_included.",
        )

        validation_output_name = st.text_input(
            "Author validation Excel filename",
            value=DEFAULT_VALIDATION_OUTPUT_NAME,
            help="Workbook for manual/author validation with 15% sample by decision category.",
        )

        costs_output_name = st.text_input(
            "LLM costs Excel filename",
            value=DEFAULT_COSTS_OUTPUT_NAME,
            help="Workbook with token usage and estimated LLM costs.",
        )
        resume_checkpoint = st.checkbox(
        "Resume from checkpoint",
        value=True
        )

        checkpoint_name = st.text_input(
            "Checkpoint filename",
            value="checkpoint_fulltext_screening.xlsx"
        )

        pause_seconds = st.number_input(
            "Pause between papers (seconds)",
            min_value=0.0,
            max_value=30.0,
            value=2.0,
            step=0.5,
        )


    rules_file = st.file_uploader("Upload full-text screening rules Excel file", type=["xlsx"])

    AUTO_INPUT_MODE = ("Automatically use Next_Stage_Fulltext from Stage 2")

    CUSTOM_INPUT_MODE = ("Use custom file with Paper_ID and PDF_path")

    input_mode = st.radio("Select input source",
        [
            AUTO_INPUT_MODE,
            CUSTOM_INPUT_MODE,
        ],
    )

    default_input_path = (INPUT_DIR_3/ "fulltext_screening_input.xlsx")

    if input_mode == AUTO_INPUT_MODE:

        if default_input_path.exists():

            st.success(
                "Full-text input automatically loaded "
                "from Stage 2."
            )

            st.caption(
                f"Using: {default_input_path}"
            )

            input_file = str(
                default_input_path
            )

        else:

            st.warning(
            "   No Stage 2 full-text input found."
            )

            input_file = st.file_uploader(
                "Upload Excel file with papers to full-text screen",
                type=["xlsx"],
                help=(
                    "Use the file containing previous-stage records "
                    "and PDF_Path values."
                ),
            )

    else:

        input_file = st.file_uploader(
            "Upload custom full-text screening input file",
            type=["xlsx"],
            help=(
                "Required columns: Paper_ID or Record_ID, "
                "Paper or Title, and PDF_Path."
            ),
        )

    st.info(
        "Instructions, rules, schema, examples, and section patterns are read from the rules template. "
        "The final decision is corrected deterministically after the LLM output."
    )


    available_columns: List[str] = []
    selected_input_sheet = 0

    if input_file is not None:
        try:
            input_sheets = load_excel_sheet_names(input_file)

            preferred_sheets = [
                "Next_Stage_Fulltext",
                "Next_stage_included",
                "Fulltext_ready",
                "Included",
            ]

            default_sheet_index = 0
            for preferred_sheet in preferred_sheets:
                if preferred_sheet in input_sheets:
                    default_sheet_index = input_sheets.index(preferred_sheet)
                    break

            selected_input_sheet = st.selectbox(
                "Sheet to use for full-text screening",
                options=input_sheets,
                index=default_sheet_index,
                help="Select the sheet that contains Paper and PDF_path. For abstract outputs, this is usually Next_Stage_Fulltext.",
            )

            preview_df = load_input_preview(input_file, selected_input_sheet)
            available_columns = preview_df.columns.tolist()
            st.caption("Input preview")
            st.dataframe(preview_df, use_container_width=True)
        except Exception as e:
            st.error(f"Could not read input file preview: {e}")


    if available_columns:
        title_col = column_selectbox("Title column", available_columns, DEFAULT_TITLE_COL)
        pdf_path_col = column_selectbox("PDF path column", available_columns, DEFAULT_PDF_PATH_COL)
    else:
        title_col = DEFAULT_TITLE_COL
        pdf_path_col = DEFAULT_PDF_PATH_COL

    paths_ready = False

    if input_file is not None and available_columns:

        paths_ready, pdf_issues_df = validate_pdf_paths(
            input_file=input_file,
            sheet_name=selected_input_sheet,
            pdf_path_col=pdf_path_col,
        )

        if paths_ready:
            st.success(
                "All PDF paths are complete and valid. "
                "Full-text screening can be started."
            )

        else:

            st.warning("Some PDFs are missing or unavailable.")

            st.info(
                f"""
        Missing PDFs detected.

        If you can retrieve PDFs manually,
        place them in:

        {PAPERS_PDF_DIR}

        Then rerun Stage 3.

        Papers without obtainable full text
        can be excluded below and reserved
        for reporting.
        """
            )

            st.dataframe(pdf_issues_df,use_container_width=True,)

            st.subheader(
                "Reserve papers without full text"
            )

            selected_no_full_text = st.multiselect(
                "Select papers to exclude from full-text screening",
                options=sorted(
                    pdf_issues_df["Paper"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                ),
                help=(
                    "Selected papers will be removed "
                    "from active screening and saved "
                    "for reporting."
                ),
            )

            if st.button(
                "Reserve selected papers"
            ):

                full_df = pd.read_excel(
                    input_file,
                    sheet_name=selected_input_sheet,
                    dtype=str,
                )

                full_df.columns = [
                    str(c).strip()
                    for c in full_df.columns
                ]

                reserved_df = full_df[
                    full_df["Paper"]
                    .astype(str)
                    .isin(selected_no_full_text)
                ].copy()

                remaining_df = full_df[
                    ~full_df["Paper"]
                    .astype(str)
                    .isin(selected_no_full_text)
                ].copy()

                reserved_df["Full_Text_Status"] = "No full text available"

                reserved_df["Reserved_For_Report"] = "Yes"
                reserved_df["Full_Text_Decision"] = "Not screened"
                reserved_df["Full_Text_Exclusion_Reason"] = ("Full text not available after manual retrieval attempt")

                if NO_FULL_TEXT_PATH.exists():

                    previous_reserved_df = pd.read_excel(
                        NO_FULL_TEXT_PATH,
                        dtype=str,
                    )

                    reserved_df = pd.concat(
                        [previous_reserved_df, reserved_df],
                        ignore_index=True,
                    )

                    if "Paper_ID" in reserved_df.columns:
                        reserved_df = reserved_df.drop_duplicates(
                            subset=["Paper_ID"],
                            keep="first",
                        )
                    else:
                        reserved_df = reserved_df.drop_duplicates(
                            subset=["Paper"],
                            keep="first",
                        )

                reserved_df.to_excel(
                    NO_FULL_TEXT_PATH,
                    index=False,
                )

                active_input_path = (
                    INPUT_DIR_3
                    / "fulltext_screening_input.xlsx"
                )

                remaining_df.to_excel(
                   active_input_path,
                    index=False,
                )

                st.success(
                    f"""
        Reserved papers saved:

        {NO_FULL_TEXT_PATH}

        Updated Stage 3 input:

        {active_input_path}
        """
                )

                st.rerun()

            st.stop()




    use_previous_stage_filter = st.checkbox(
        "Filter rows using a previous-stage decision column",
        value=False,
        help="Leave unchecked if the uploaded file already contains only papers to full-text screen.",
    )

    if use_previous_stage_filter and available_columns:
        previous_filter_col = column_selectbox(
            "Previous-stage decision column",
            available_columns,
            DEFAULT_PREVIOUS_FILTER_COL,
        )

        unique_values = []
        try:
            reset_excel_pointer(input_file)
            full_preview = pd.read_excel(input_file, sheet_name=selected_input_sheet, dtype=str, usecols=[previous_filter_col])
            unique_values = sorted(
                [v for v in full_preview[previous_filter_col].dropna().astype(str).str.strip().unique().tolist() if v]
            )
        except Exception:
            unique_values = []

        if unique_values:
            included_value = st.selectbox(
                "Value that means the paper passes previous-stage filtering",
                options=unique_values,
                index=select_default_index(unique_values, "Included"),
            )
        else:
            included_value = st.text_input("Value that means the paper passes previous-stage filtering", value="Included")
    else:
        previous_filter_col = ""
        included_value = ""


    st.subheader("Prompt configuration")
    if rules_file is not None:
        try:
            prompt_config_preview = read_prompt_config(rules_file)
            prompt_from_file = build_prompt_instructions(prompt_config_preview)
            st.text_area(
                "Prompt instructions read from Prompt_Config",
                value=prompt_from_file,
                height=220,
                disabled=True,
            )
        except Exception as e:
            st.error(f"Prompt_Config problem: {e}")
    else:
        st.info("Upload the rules template to preview Prompt_Config.")


    run_button = st.button("Run full-text screening", type="primary")


    # ============================================================
    # Main process
    # ============================================================
    if run_button:
        if not api_key:
            st.error("Please provide an OpenAI API key.")
            st.stop()

        if rules_file is None:
            st.error("Please upload the full-text rules Excel file.")
            st.stop()

        if input_file is None:
            st.error("Please upload the Excel file with papers to full-text screen.")
            st.stop()

        client = OpenAI(api_key=api_key)

        try:
            rules_prompt = read_rules(rules_file)
            schema_df = read_output_schema(rules_file)
            schema_prompt = build_schema_prompt(schema_df)
            examples_prompt = read_examples(rules_file)
            output_fields = get_output_fields(schema_df)
            json_schema = make_json_schema(output_fields)
            prompt_config = read_prompt_config(rules_file)
            prompt_instructions = build_prompt_instructions(prompt_config)
            section_patterns, section_pattern_source = read_section_patterns(rules_file)
        except Exception as e:
            st.error(f"Error reading rules file: {e}")
            st.stop()

        st.info(section_pattern_source)

        try:
            reset_excel_pointer(input_file)
            df = pd.read_excel(input_file, sheet_name=selected_input_sheet, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]

        except Exception as e:
            st.error(f"Error reading input file: {e}")
            st.stop()

        if title_col not in df.columns:
            st.error(f"Title column '{title_col}' not found.")
            st.stop()

        pdf_path_aliases = [
            "PDF Path",
            "PDF_Path",
            "PDF_path",
            "pdf_path",
        ]

        if pdf_path_col not in df.columns:
            found_col = None

            for col in pdf_path_aliases:
                if col in df.columns:
                    found_col = col
                    break

            if found_col:
                pdf_path_col = found_col
                st.info(f"Using detected PDF path column: {pdf_path_col}")
            else:
                st.error(
                    f"PDF path column '{pdf_path_col}' not found."
            )
                st.stop()

        original_count = len(df)

        if use_previous_stage_filter:
            if previous_filter_col not in df.columns:
                st.error(f"Previous-stage decision column '{previous_filter_col}' not found.")
                st.stop()
            filtered_df = df[df[previous_filter_col].astype(str).str.strip() == included_value].copy()
        else:
            filtered_df = df.copy()

        filtered_count = len(filtered_df)
        duplicate_summary = pd.DataFrame()

        if deduplicate_titles:
            filtered_df, duplicate_summary = remove_duplicates_keep_first(filtered_df, title_col=title_col)

        deduplicated_count = len(filtered_df)
        removed_duplicates_count = filtered_count - deduplicated_count

        st.info(
            f"Input records: {original_count} | "
            f"Records selected for full-text screening: {filtered_count} | "
            f"After deduplication: {deduplicated_count} | "
            f"Duplicates removed: {removed_duplicates_count}"
        )

        checkpoint_path = TEMP_DIR_3 / checkpoint_name

        if resume_checkpoint and checkpoint_path.exists():
            checkpoint_df = pd.read_excel(checkpoint_path, sheet_name="All_results_corrected", dtype=str)
            results = checkpoint_df.drop(columns=["Paper_ID"], errors="ignore").to_dict("records")
            processed_keys = set(
                checkpoint_df.apply(
                    lambda r: make_checkpoint_key(r.get("Paper", ""), r.get("PDF_path", "")),
                    axis=1,
                )
            )
        else:
            results = []
            processed_keys = set()

 
        results_output_path = OUTPUT_DIR_3 / results_output_name
        validation_output_path = OUTPUT_DIR_3 / validation_output_name
        costs_output_path = OUTPUT_DIR_3 / costs_output_name
    
        progress = st.progress(0)
        log_box = st.empty()
        metrics_box = st.empty()

        total_iterations = max(1, len(filtered_df))

        for i, (_, row_data) in enumerate(filtered_df.iterrows(), start=1):
            paper = safe_str(row_data.get(title_col, ""))
            raw_pdf_path = safe_str(row_data.get(pdf_path_col, ""))
            pdf_path = Path(raw_pdf_path).expanduser() if raw_pdf_path else Path("")
            log_box.info(
            f"PDF exists? {pdf_path.exists()} | path={pdf_path}"
            )

            current_key = make_checkpoint_key(paper, str(pdf_path))

            if current_key in processed_keys:
                log_box.info(f"Skipping already processed paper {i}/{len(filtered_df)}: {paper[:90]}")
                progress.progress(i / total_iterations)
                continue
            pdf_name = pdf_path.name if raw_pdf_path else ""

            result: Dict[str, Any] = {
                **prefix_input_columns(row_data),
                "Paper": paper,
                "PDF_name": pdf_name,
                "PDF_path": str(pdf_path),
                "PDF_Source": "PDF path column",
                "input_tokens": 0,
                "output_tokens": 0,
                "total_tokens": 0,
            }

            log_box.info(f"Processing paper {i}/{len(filtered_df)}: {paper[:90]}")

            skip_reason = ""
            if not raw_pdf_path or raw_pdf_path.lower() in ["nan", "none", "null", "0", "no access"]:
                skip_reason = "No valid PDF path was provided."
                result["PDF_Status"] = "Missing PDF path"
            elif not pdf_path.exists():
                skip_reason = f"PDF not found at {pdf_path}"
                result["PDF_Status"] = "PDF not found"

            if skip_reason:
                result.update(
                    {
                        "PDF_Status_Reason": skip_reason,
                        "Processing_Status": "Skipped",
                        "Error_Message": "",
                        "q1": "Undetermined",
                        "q2": "Undetermined",
                        "q3": "Undetermined",
                        "q4": "Undetermined",
                        "decision": "Undetermined",
                        "reason": skip_reason,
                        "manualCheckNeeded": "Yes for decision",
                        "manualCheckReason": skip_reason,
                    }
                )
                results.append(result)
                processed_keys.add(make_checkpoint_key(result.get("Paper", ""), result.get("PDF_path", "")))
                outputs = write_results_excel(
                    output_path=checkpoint_path,
                    results=results,
                    duplicate_summary=duplicate_summary,
                    sample_fraction=float(manual_sample_percent) / 100.0,
                    random_seed=int(random_seed),
                    input_price_per_1m=float(input_price_per_1m),
                    output_price_per_1m=float(output_price_per_1m),
                )


                progress.progress(i / total_iterations)
                continue

            full_text_raw = load_pdf_text(pdf_path)

            if not full_text_raw.strip():
                reason = "No extractable text from PDF; review manually."
                result.update(
                    {
                        "PDF_Status": "No extractable text",
                        "PDF_Status_Reason": reason,
                        "Processing_Status": "Skipped",
                        "Error_Message": "",
                        "q1": "Undetermined",
                        "q2": "Undetermined",
                        "q3": "Undetermined",
                        "q4": "Undetermined",
                        "decision": "Undetermined",
                        "reason": reason,
                        "manualCheckNeeded": "Yes for decision",
                        "manualCheckReason": reason,
                    }
                )
                results.append(result)
                processed_keys.add(make_checkpoint_key(result.get("Paper", ""), result.get("PDF_path", "")))

                outputs = write_results_excel(
                    output_path=checkpoint_path,
                    results=results,
                    duplicate_summary=duplicate_summary,
                    sample_fraction=float(manual_sample_percent) / 100.0,
                    random_seed=int(random_seed),
                    input_price_per_1m=float(input_price_per_1m),
                    output_price_per_1m=float(output_price_per_1m),
                )

                progress.progress(i / total_iterations)
                continue

            selected_text, text_was_truncated, section_detection_quality, selected_section_names = build_relevant_text_for_screening(
                full_text=full_text_raw,
                max_chars=int(max_chars),
                section_patterns=section_patterns,
            )

            result.update(
                {
                    "Section_Detection_Quality": section_detection_quality,
                    "Selected_Sections": selected_section_names,
                    "PDF_Status": "Text extracted",
                    "PDF_Status_Reason": "",
                    "Full_Text_Characters_Extracted": len(full_text_raw),
                    "Selected_Text_Characters_Sent": len(selected_text),
                    "Text_Was_Truncated": text_was_truncated,
                }
            )

            try:
                classification, usage = classify_fulltext(
                    client=client,
                    model=model,
                    paper=paper,
                    selected_text=selected_text,
                    rules_prompt=rules_prompt,
                    schema_prompt=schema_prompt,
                    examples_prompt=examples_prompt,
                    json_schema=json_schema,
                    text_was_truncated=text_was_truncated,
                    section_detection_quality=section_detection_quality,
                    prompt_instructions=prompt_instructions,
                )

                result.update(usage)

                if section_detection_quality == "Weak" and safe_str(classification.get("decision", "")) == "Undetermined":
                    current_manual = safe_str(classification.get("manualCheckNeeded", ""))
                    if current_manual in ["", "No"]:
                        classification["manualCheckNeeded"] = "Yes for decision"

                    classification["manualCheckReason"] = append_manual_reason(
                        classification.get("manualCheckReason", ""),
                        "Weak section detection; relevant sections may be missing.",
                    )

                classification = postprocess_classification(classification)

                for field in output_fields:
                    result[field] = classification.get(field, "")

                result["Processing_Status"] = "Success"
                result["Error_Message"] = ""

            except Exception as e:
                for field in output_fields:
                    result[field] = ""
                result["Processing_Status"] = "Error"
                result["Error_Message"] = str(e)

            results.append(result)
            processed_keys.add(make_checkpoint_key(result.get("Paper", ""), result.get("PDF_path", "")))

            # Save checkpoint immediately
            outputs = write_results_excel(
                output_path=checkpoint_path,
                results=results,
                duplicate_summary=duplicate_summary,
                sample_fraction=float(manual_sample_percent) / 100.0,
                random_seed=int(random_seed),
                input_price_per_1m=float(input_price_per_1m),
                output_price_per_1m=float(output_price_per_1m),
            )


            summary = outputs["summary"]
            charges_summary = outputs["charges_summary"]
            total_cost = 0.0
            if not charges_summary.empty:
                total_cost = float(charges_summary.loc[charges_summary["metric"] == "total_cost", "value"].iloc[0])

            with metrics_box.container():
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Processed", f"{i}/{len(filtered_df)}")
                c2.metric("Included", int(summary.loc[summary["decision"] == "Included", "count"].iloc[0]))
                c3.metric("Excluded", int(summary.loc[summary["decision"] == "Excluded", "count"].iloc[0]))
                c4.metric("Undetermined", int(summary.loc[summary["decision"] == "Undetermined", "count"].iloc[0]))
                c5.metric("Estimated cost", f"${total_cost:.4f}")

            progress.progress(i / total_iterations)
            time.sleep(pause_seconds)

        outputs = write_output_workbooks(
            results_output_path=results_output_path,
            validation_output_path=validation_output_path,
            costs_output_path=costs_output_path,
            results=results,
            duplicate_summary=duplicate_summary,
            sample_fraction=float(manual_sample_percent) / 100.0,
            random_seed=int(random_seed),
            input_price_per_1m=float(input_price_per_1m),
            output_price_per_1m=float(output_price_per_1m),
        )

        st.success("Done. Results saved successfully.")

        st.subheader("Decision summary")
        st.dataframe(outputs["summary"], use_container_width=True)

        if "prisma_summary" in outputs:
            st.subheader("PRISMA summary")
            st.dataframe(outputs["prisma_summary"], use_container_width=True)

        st.subheader("LLM charge summary")
        st.dataframe(outputs["charges_summary"], use_container_width=True)

        st.subheader("Author check sample")
        st.dataframe(outputs["author_sample"], use_container_width=True)

        st.subheader("Author check clean")
        st.dataframe(outputs["author_clean"], use_container_width=True)

        st.subheader("Next-stage included papers")
        st.dataframe(outputs["next_stage"], use_container_width=True)

        st.success(
            f"""
        Full-text screening completed successfully.

        📁 Stage 3 output folder:
        {OUTPUT_DIR_3}

        Generated files:
        • {results_output_path.name}
        • {validation_output_path.name}
        • {costs_output_path.name}

        Checkpoint saved in temp:
        {checkpoint_path}
        """
        )

if __name__ == "__main__":
    main()