import os
import time
from pathlib import Path
import math

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
import shutil

STAGE_4_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAGE_4_DIR.parent

INPUT_DIR_4 = STAGE_4_DIR / "input"
OUTPUT_DIR_4 = STAGE_4_DIR / "output"
TEMP_DIR_4 = STAGE_4_DIR / "temp"

INPUT_DIR_4.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_4.mkdir(parents=True, exist_ok=True)
TEMP_DIR_4.mkdir(parents=True, exist_ok=True)

STAGE_3_DIR = PROJECT_ROOT / "stage_3_fulltext_initial_pairing"
STAGE_3_INPUT_PATH = STAGE_3_DIR / "output" / "stage_3_to_stage_4.xlsx"
STAGE_4_INPUT_PATH = INPUT_DIR_4 / "stage_4_input.xlsx"

if STAGE_3_INPUT_PATH.exists() and not STAGE_4_INPUT_PATH.exists():
    shutil.copy2(
        STAGE_3_INPUT_PATH,
        STAGE_4_INPUT_PATH,
    )

from stage_4_logic import (
    AUTO_INPUT_MODE,
    CUSTOM_INPUT_MODE,
    build_dependency_rules_prompt,
    build_prompt_from_template,
    build_schema_prompt,
    build_targeted_paper_context,
    build_valid_pairs,
    call_openai_with_usage,
    explode_valid_pairs,
    extract_pdf_text,
    find_existing_next_stage_file,
    get_active_prompt_template,
    get_output_fields,
    get_processed_pair_keys,
    get_processed_paper_ids,
    load_checkpoint,
    load_paper_list,
    make_json_schema,
    make_pair_key,
    make_usage_record,
    read_coding_schema,
    read_dependency_rules,
    read_normalization_rules,
    read_sheet_preamble,
    save_usage_workbook,
    summarize_usage,
    read_broad_normalization_rules,
    build_normalized_pair_matrix,
    filter_final_stage4_outputs,
)

load_dotenv()


DEFAULT_MODEL = "gpt-4.1-mini"
DEFAULT_PAPER_OUTPUT = "stage_4_paper_analysis.xlsx"
DEFAULT_PAIR_OUTPUT = "stage_4_pairing_analysis.xlsx"
DEFAULT_COST_OUTPUT = "stage_4_llm_costs.xlsx"

def main():
    # ============================================================
    # Page setup
    # ============================================================
    st.set_page_config(page_title="CT–Programming Mapping Pipeline", layout="wide")

    st.title("CT–Programming Mapping Pipeline")
    st.write(
        "Pipeline: previous full screening input → targeted paper-level coding → Python valid pairs → pair-level mapping analysis."
    )

    st.info(
        "Prompts are read from the `Prompt_Templates` sheet in the coding schema file. "
        "The previous-stage matching location is used as a navigation hint, not as evidence."
    )


    # ============================================================
    # Sidebar configuration
    # ============================================================
    with st.sidebar:
        st.header("Configuration")

        api_key = st.text_input(
            "OpenAI API key",
            value=os.getenv("OPENAI_API_KEY", ""),
            type="password",
        )

        model = st.text_input("Model", value=DEFAULT_MODEL)

        st.subheader("Outputs")
        paper_output_name = st.text_input(
            "Paper-level output filename",
            value=DEFAULT_PAPER_OUTPUT,
        )

        pair_output_name = st.text_input(
            "Pair-level output filename",
            value=DEFAULT_PAIR_OUTPUT,
        )

        cost_output_name = st.text_input(
            "LLM costs output filename",
            value=DEFAULT_COST_OUTPUT,
        )

        st.subheader("Processing")
        max_pdf_chars = st.number_input(
            "Maximum raw PDF text before targeting",
            min_value=20_000,
            max_value=300_000,
            value=160_000,
            step=10_000,
        )

        max_context_chars = st.number_input(
            "Maximum targeted context sent to model",
            min_value=10_000,
            max_value=120_000,
            value=60_000,
            step=5_000,
        )

        additional_keywords = st.text_area(
            "Additional targeting keywords (comma-separated)",
            value="",
            height=80,
            help="Optional. Example: robotics, micro:bit, proof assistant.",
        )

        run_paper_level = st.checkbox("Run paper-level coding", value=True)
        run_pair_level = st.checkbox("Run pair-level mapping analysis", value=True)

        resume_from_checkpoints = st.checkbox(
            "Resume from checkpoints if available",
            value=True,
        )

        pause_seconds = st.number_input(
            "Pause between papers (seconds)",
            min_value=0.0,
            max_value=30.0,
            value=0.5,
            step=0.5,
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

    # ============================================================
    # Input files in main body
    # ============================================================

    st.subheader("Input files")

    schema_file = st.file_uploader(
        "Upload coding schema Excel file",
        type=["xlsx"],
    )

    input_mode = st.radio(
        "Select input source",
        [
            AUTO_INPUT_MODE,
            CUSTOM_INPUT_MODE,
        ],
    )

    existing_next_stage_path = (
        STAGE_4_INPUT_PATH
        if STAGE_4_INPUT_PATH.exists()
        else None
    )

    if input_mode == AUTO_INPUT_MODE:
        if existing_next_stage_path is not None:
            st.info("Stage 4 input automatically loaded.")
            st.caption(f"Using: {existing_next_stage_path}")
            paper_list_file = None
        else:
            st.warning("No previous screening workbook auto-detected.")
            paper_list_file = st.file_uploader(
                "Upload previous screening workbook",
                type=["xlsx"],
            )
    else:
        paper_list_file = st.file_uploader(
            "Upload custom paper list Excel file",
            type=["xlsx"],
            help="Required columns: Paper_ID and PDF_path or PDF_Address.",
        )

    run_button = st.button("Run pipeline", type="primary")


    # ============================================================
    # Main execution
    # ============================================================
    if run_button:
        if not api_key:
            st.error("Please provide an OpenAI API key.")
            st.stop()

        if schema_file is None:
            st.error("Please upload the coding schema Excel file.")
            st.stop()

        if input_mode == CUSTOM_INPUT_MODE and paper_list_file is None:
            st.error("Please upload the custom paper list Excel file.")
            st.stop()

        client = OpenAI(api_key=api_key)

        try:
            # Paper-level schema
            paper_schema_df = read_coding_schema(
                schema_file,
                sheet_name="Coding_Schema_Framework",
            )
            paper_dependency_df = read_dependency_rules(
                schema_file,
                sheet_name="Framework_Dependency_Rules",
            )
            normalization_df = read_normalization_rules(schema_file)
            broad_normalization_rules = read_broad_normalization_rules(schema_file)

            paper_schema_prompt = build_schema_prompt(paper_schema_df)
            paper_dependency_prompt = build_dependency_rules_prompt(paper_dependency_df)

            paper_output_fields = get_output_fields(
                paper_schema_df,
                exclude_fields=["Valid_CT_Programming_Pairs"],
            )
            paper_json_schema = make_json_schema(
                paper_output_fields,
                schema_name="paper_level_framework",
                schema_df=paper_schema_df,
            )

            # Pair-level schema
            pair_schema_df = read_coding_schema(
                schema_file,
                sheet_name="Mapping",
            )
            pair_dependency_df = read_dependency_rules(
                schema_file,
                sheet_name="Mapping_Dependency_Rules",
            )

            pair_schema_prompt = build_schema_prompt(pair_schema_df)
            pair_dependency_prompt = build_dependency_rules_prompt(pair_dependency_df)
            mapping_preamble = read_sheet_preamble(schema_file, "Mapping")

            pair_output_fields = get_output_fields(pair_schema_df)
            pair_json_schema = make_json_schema(
                pair_output_fields,
                schema_name="pair_level_mapping",
            )

            # Prompt templates from Excel
            paper_prompt_template = get_active_prompt_template(
                schema_file,
                stage="paper_level",
            )
            pair_prompt_template = get_active_prompt_template(
                schema_file,
                stage="pair_level",
            )

        except Exception as e:
            st.error(f"Error reading schema/rules/prompt file: {e}")
            st.stop()

        try:
            paper_list_df, excluded_before_stage4_df = load_paper_list(
                input_mode=input_mode,
                uploaded_file=paper_list_file,
                default_next_stage_path=existing_next_stage_path,
            )
        
            st.info(
                f"Papers removed before Stage 4 because candidate matching was not specified: "
                f"{len(excluded_before_stage4_df)}"
            )
        except Exception as e:
            st.error(f"Error reading input file: {e}")
            st.stop()

        if paper_list_df.empty:
            st.error("The paper list is empty.")
            st.stop()

        pdf_items = []
        for _, row in paper_list_df.iterrows():
            pdf_items.append(
                {
                    "Paper_ID": str(row["Paper_ID"]).strip(),
                    "Title": str(row.get("Title", "")).strip(),
                    "summary": str(row.get("summary", "")).strip(),
                    "q1": str(row.get("q1", "")).strip(),
                    "q2": str(row.get("q2", "")).strip(),
                    "q3": str(row.get("q3", "")).strip(),
                    "q4": str(row.get("q4", "")).strip(),
                    "decision": str(row.get("decision", "")).strip(),
                    "Matching_Evidence_Location": str(row.get("Matching_Evidence_Location", "")).strip(),
                    "Next_stage_label": str(row.get("Next_stage_label", "")).strip(),
                    "PDF_path": Path(str(row["PDF_Address"]).strip()),
                }
            )

        paper_output_path = OUTPUT_DIR_4 / paper_output_name
        pair_output_path = OUTPUT_DIR_4 / pair_output_name
        cost_output_path = OUTPUT_DIR_4 / cost_output_name

        paper_checkpoint_path = (
            TEMP_DIR_4
            / f"{paper_output_path.stem}_checkpoint.xlsx"
        )

        pair_checkpoint_path = (
            TEMP_DIR_4
            / f"{pair_output_path.stem}_checkpoint.xlsx"
        )

        if resume_from_checkpoints:
            paper_checkpoint_df = load_checkpoint(paper_checkpoint_path)
            pair_checkpoint_df = load_checkpoint(pair_checkpoint_path)
        else:
            paper_checkpoint_df = pd.DataFrame()
            pair_checkpoint_df = pd.DataFrame()

        paper_results = (
            paper_checkpoint_df.to_dict("records")
            if not paper_checkpoint_df.empty
            else []
        )
        pair_results = (
            pair_checkpoint_df.to_dict("records")
            if not pair_checkpoint_df.empty
            else []
        )

        processed_paper_ids = get_processed_paper_ids(paper_checkpoint_df)
        processed_pair_keys = get_processed_pair_keys(pair_checkpoint_df)

        usage_records = []

        progress = st.progress(0)
        log_box = st.empty()
        metrics_box = st.empty()

        total_papers = len(pdf_items)

        for i, item in enumerate(pdf_items, start=1):
            paper_id = item["Paper_ID"]
            title_from_input = item.get("Title", "")
            previous_summary = item.get("summary", "")
            matching_evidence_location = item.get("Matching_Evidence_Location", "")
            pdf_path = item["PDF_path"]

            log_box.info(f"Processing paper {i}/{total_papers}: {pdf_path.name}")

            paper_text = ""
            targeted_context = ""
            paper_result = None

            try:
                if not pdf_path.exists():
                    raise ValueError(f"PDF not found: {pdf_path}")

                paper_text = extract_pdf_text(pdf_path, max_chars=int(max_pdf_chars))

                if not paper_text:
                    raise ValueError("No extractable text found in PDF.")

                targeted_context = build_targeted_paper_context(
                    paper_text=paper_text,
                    title=title_from_input,
                    summary=previous_summary,
                    matching_evidence_location=matching_evidence_location,
                    max_chars=int(max_context_chars),
                    additional_keywords=additional_keywords,
                )

                # ----------------------------------------------------
                # Paper-level coding
                # ----------------------------------------------------
                if run_paper_level and paper_id not in processed_paper_ids:
                    paper_prompt = build_prompt_from_template(
                        paper_prompt_template,
                        dependency_rules_prompt=paper_dependency_prompt,
                        schema_prompt=paper_schema_prompt,
                        paper_filename=pdf_path.name,
                        previous_summary=previous_summary,
                        matching_evidence_location=matching_evidence_location,
                        paper_text=targeted_context,
                    )

                    paper_result, usage = call_openai_with_usage(
                        client=client,
                        model=model,
                        prompt=paper_prompt,
                        json_schema=paper_json_schema,
                        schema_name="paper_level_framework",
                    )

                    usage_records.append(
                        make_usage_record(
                            stage="paper_level",
                            paper_id=paper_id,
                            title=title_from_input or paper_result.get("Title", ""),
                            source_file=pdf_path.name,
                            model=model,
                            usage=usage,
                            input_price_per_1m=float(input_price_per_1m),
                            output_price_per_1m=float(output_price_per_1m),
                        )
                    )

                    if title_from_input and not str(paper_result.get("Title", "")).strip():
                        paper_result["Title"] = title_from_input

                    paper_result["Paper_ID"] = paper_id
                    paper_result["Source_File"] = pdf_path.name
                    paper_result["PDF_path"] = str(pdf_path)
                    paper_result["PDF_Address"] = str(pdf_path)
                    paper_result["Previous_Summary"] = previous_summary
                    paper_result["Matching_Evidence_Location"] = matching_evidence_location
                    paper_result["Targeted_Context_Chars"] = len(targeted_context)

                    matching_text = paper_result.get("Matching_CT_Programming_Elements", "")
                    valid_pairs = build_valid_pairs(matching_text, normalization_df)
                    paper_result["Valid_CT_Programming_Pairs"] = valid_pairs

                    paper_result["Processing_Status"] = "Success"
                    paper_result["Error_Message"] = ""

                    paper_results.append(paper_result)
                    processed_paper_ids.add(paper_id)

                    pd.DataFrame(paper_results).to_excel(paper_checkpoint_path, index=False)
                    pd.DataFrame(paper_results).to_excel(paper_output_path, index=False)

                elif paper_id in processed_paper_ids:
                    previous_rows = paper_checkpoint_df[
                        paper_checkpoint_df["Paper_ID"].astype(str).str.strip() == paper_id
                    ]
                    if not previous_rows.empty:
                        paper_result = previous_rows.iloc[-1].to_dict()
                    else:
                        paper_result = {field: "" for field in paper_output_fields}
                        paper_result["Paper_ID"] = paper_id
                        paper_result["Title"] = title_from_input
                        paper_result["PDF_path"] = str(pdf_path)
                        paper_result["PDF_Address"] = str(pdf_path)
                        paper_result["Valid_CT_Programming_Pairs"] = "No valid pairs"
                        paper_result["Processing_Status"] = "Skipped"
                        paper_result["Error_Message"] = "Skipped because Paper_ID was already processed."

                else:
                    paper_result = {field: "" for field in paper_output_fields}
                    paper_result["Paper_ID"] = paper_id
                    paper_result["Title"] = title_from_input
                    paper_result["Source_File"] = pdf_path.name
                    paper_result["PDF_path"] = str(pdf_path)
                    paper_result["PDF_Address"] = str(pdf_path)
                    paper_result["Previous_Summary"] = previous_summary
                    paper_result["Matching_Evidence_Location"] = matching_evidence_location
                    paper_result["Valid_CT_Programming_Pairs"] = "No valid pairs"
                    paper_result["Processing_Status"] = "Skipped"
                    paper_result["Error_Message"] = "Paper-level coding was disabled."

            except Exception as e:
                paper_result = {field: "" for field in paper_output_fields}
                paper_result["Paper_ID"] = paper_id
                paper_result["Title"] = title_from_input
                paper_result["Source_File"] = pdf_path.name
                paper_result["PDF_path"] = str(pdf_path)
                paper_result["PDF_Address"] = str(pdf_path)
                paper_result["Previous_Summary"] = previous_summary
                paper_result["Matching_Evidence_Location"] = matching_evidence_location
                paper_result["Valid_CT_Programming_Pairs"] = "No valid pairs"
                paper_result["Processing_Status"] = "Error"
                paper_result["Error_Message"] = str(e)

                paper_results.append(paper_result)
                pd.DataFrame(paper_results).to_excel(paper_checkpoint_path, index=False)
                pd.DataFrame(paper_results).to_excel(paper_output_path, index=False)

            # ----------------------------------------------------
            # Pair-level coding
            # ----------------------------------------------------
            if run_pair_level and paper_result.get("Processing_Status") in ["Success", "Skipped"]:
                valid_pairs_text = paper_result.get("Valid_CT_Programming_Pairs", "")
                exploded_pairs = explode_valid_pairs(valid_pairs_text)
                title = paper_result.get("Title", "") or title_from_input

                for ct_element, programming_element in exploded_pairs:
                    pair_key = make_pair_key(paper_id, ct_element, programming_element)

                    if resume_from_checkpoints and pair_key in processed_pair_keys:
                        continue

                    try:
                        pair_prompt = build_prompt_from_template(
                            pair_prompt_template,
                            mapping_preamble=mapping_preamble,
                            dependency_rules_prompt=pair_dependency_prompt,
                            mapping_schema_prompt=pair_schema_prompt,
                            paper_id=paper_id,
                            title=title,
                            ct_element=ct_element,
                            programming_element=programming_element,
                            previous_summary=previous_summary,
                            matching_evidence_location=matching_evidence_location,
                            paper_filename=pdf_path.name,
                            paper_text=targeted_context if targeted_context else paper_text,
                        )

                        pair_result, usage = call_openai_with_usage(
                            client=client,
                            model=model,
                            prompt=pair_prompt,
                            json_schema=pair_json_schema,
                            schema_name="pair_level_mapping",
                        )

                        usage_records.append(
                            make_usage_record(
                                stage="pair_level",
                                paper_id=paper_id,
                                title=title,
                                source_file=pdf_path.name,
                                model=model,
                                usage=usage,
                                input_price_per_1m=float(input_price_per_1m),
                                output_price_per_1m=float(output_price_per_1m),
                                ct_element=ct_element,
                                programming_element=programming_element,
                            )
                        )

                        pair_result["Paper_ID"] = paper_id
                        pair_result["Title"] = title
                        pair_result["CT_element"] = ct_element
                        pair_result["Programming_element"] = programming_element
                        pair_result["Source_File"] = pdf_path.name
                        pair_result["PDF_path"] = str(pdf_path)
                        pair_result["PDF_Address"] = str(pdf_path)
                        pair_result["Previous_Summary"] = previous_summary
                        pair_result["Matching_Evidence_Location"] = matching_evidence_location
                        pair_result["Targeted_Context_Chars"] = len(targeted_context)
                        pair_result["Processing_Status"] = "Success"
                        pair_result["Error_Message"] = ""

                    except Exception as e:
                        pair_result = {field: "" for field in pair_output_fields}
                        pair_result["Paper_ID"] = paper_id
                        pair_result["Title"] = title
                        pair_result["CT_element"] = ct_element
                        pair_result["Programming_element"] = programming_element
                        pair_result["Source_File"] = pdf_path.name
                        pair_result["PDF_path"] = str(pdf_path)
                        pair_result["PDF_Address"] = str(pdf_path)
                        pair_result["Previous_Summary"] = previous_summary
                        pair_result["Matching_Evidence_Location"] = matching_evidence_location
                        pair_result["Targeted_Context_Chars"] = len(targeted_context)
                        pair_result["Processing_Status"] = "Error"
                        pair_result["Error_Message"] = str(e)

                        usage_records.append(
                            make_usage_record(
                                stage="pair_level",
                                paper_id=paper_id,
                                title=title,
                                source_file=pdf_path.name,
                                model=model,
                                usage={
                                    "input_tokens": 0,
                                    "cached_input_tokens": 0,
                                    "output_tokens": 0,
                                    "total_tokens": 0,
                                },
                                input_price_per_1m=float(input_price_per_1m),
                                output_price_per_1m=float(output_price_per_1m),
                                ct_element=ct_element,
                                programming_element=programming_element,
                                status="Error",
                                error_message=str(e),
                            )
                        )

                    pair_results.append(pair_result)
                    processed_pair_keys.add(pair_key)

                    pd.DataFrame(pair_results).to_excel(pair_checkpoint_path, index=False)
                    pd.DataFrame(pair_results).to_excel(pair_output_path, index=False)

            # Save costs after each paper
            usage_df, usage_summary_df = save_usage_workbook(cost_output_path, usage_records)
            cost_summary = {
                row["Metric"]: row["Value"]
                for _, row in usage_summary_df.iterrows()
            }

            if cost_output_path.exists():
                cost_output_path.unlink()

            summary_for_metrics = summarize_usage(usage_records)
            total_cost_row = summary_for_metrics[
                summary_for_metrics["Metric"] == "Estimated total cost USD"
            ]
            total_cost = (
                float(total_cost_row["Value"].iloc[0])
                if not total_cost_row.empty
                else 0.0
            )

            paper_errors = sum(
                1 for r in paper_results if str(r.get("Processing_Status", "")).strip() == "Error"
            )
            pair_errors = sum(
                1 for r in pair_results if str(r.get("Processing_Status", "")).strip() == "Error"
            )

            with metrics_box.container():
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Papers processed", f"{i}/{total_papers}")
                c2.metric("Paper rows", len(paper_results))
                c3.metric("Pair rows", len(pair_results))
                c4.metric("Errors", paper_errors + pair_errors)
                c5.metric("Estimated cost", f"${total_cost:.4f}")

            progress.progress(i / total_papers)
            time.sleep(float(pause_seconds))

# ============================================================
# Final saves
# ============================================================

        paper_previous_df = pd.DataFrame(paper_results)
        pair_previous_df = pd.DataFrame(pair_results)

        paper_clean_df = build_normalized_pair_matrix(
            papers_df=paper_previous_df,
            broad_rules=broad_normalization_rules,
            mapping_col="Matching_CT_Programming_Elements",
        )

        final_paper_df, final_pair_df = filter_final_stage4_outputs(
            paper_df=paper_clean_df,
            pair_df=pair_previous_df,
        )

        paper_clean_ids = set(
            paper_clean_df["Paper_ID"].astype(str).str.strip()
        ) if "Paper_ID" in paper_clean_df.columns else set()

        pair_clean_df = pair_previous_df[
            pair_previous_df["Paper_ID"]
            .astype(str)
            .str.strip()
            .isin(paper_clean_ids)
        ].copy() if "Paper_ID" in pair_previous_df.columns else pd.DataFrame()

        # Final files for next stage
        final_paper_df.to_excel(paper_output_path, index=False)
        final_pair_df.to_excel(pair_output_path, index=False)

        # Previous and cleaning outputs in one audit file
        audit_outputs_path = OUTPUT_DIR_4 / "stage4_previous_and_cleaning_outputs.xlsx"

        with pd.ExcelWriter(audit_outputs_path, engine="openpyxl") as writer:
            paper_previous_df.to_excel(writer, sheet_name="Previous_Paper_Level", index=False)
            pair_previous_df.to_excel(writer, sheet_name="Previous_Pair_Level", index=False)
            paper_clean_df.to_excel(writer, sheet_name="Cleaned_Paper_Level", index=False)
            pair_clean_df.to_excel(writer, sheet_name="Cleaned_Pair_Level", index=False)

        usage_df, usage_summary_df = save_usage_workbook(cost_output_path, usage_records)

        # ============================================================
        # Author validation workbook
        # ============================================================

        paper_df = pd.DataFrame(paper_results)
        pair_df = pd.DataFrame(pair_results)

        paper_sample_size = max(
            1,
            math.ceil(len(paper_df) * 0.15)
        ) if not paper_df.empty else 0

        pair_sample_size = max(
            1,
            math.ceil(len(pair_df) * 0.15)
        ) if not pair_df.empty else 0

        paper_validation_df = (
            paper_df.sample(
                n=paper_sample_size,
                random_state=42,
            )
            if paper_sample_size > 0
            else pd.DataFrame()
        )

        pair_validation_df = (
            pair_df.sample(
                n=pair_sample_size,
                random_state=42,
            )
            if pair_sample_size > 0
            else pd.DataFrame()
        )

        paper_validation_df["Author_Decision"] = ""
        paper_validation_df["Author_Notes"] = ""

        pair_validation_df["Author_Decision"] = ""
        pair_validation_df["Author_Notes"] = ""

        validation_path = (
            OUTPUT_DIR_4
            / "stage4_author_validation.xlsx"
        )

        with pd.ExcelWriter(
            validation_path,
            engine="openpyxl",
        ) as writer:

            paper_validation_df.to_excel(
                writer,
                sheet_name="Paper_Level_Validation",
                index=False,
            )

            pair_validation_df.to_excel(
                writer,
                sheet_name="Pair_Level_Validation",
                index=False,
            )


       # ============================================================
        # Final display and summary counts
        # ============================================================

        broad_included_count = int(
            paper_clean_df["Pairing_decision"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("included")
            .sum()
        ) if "Pairing_decision" in paper_clean_df.columns else 0

        framework_valid_count = int(
            paper_clean_df["Valid_CT_Programming_Pairs"]
            .astype(str)
            .str.strip()
            .str.lower()
            .ne("no valid pairs")
            .sum()
        ) if "Valid_CT_Programming_Pairs" in paper_clean_df.columns else 0

        has_broad_pair = (
            paper_clean_df["Pairing_decision"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("included")
            if "Pairing_decision" in paper_clean_df.columns
            else pd.Series(False, index=paper_clean_df.index)
        )

        has_framework_pair = (
            paper_clean_df["Valid_CT_Programming_Pairs"]
            .astype(str)
            .str.strip()
            .str.lower()
            .ne("no valid pairs")
            if "Valid_CT_Programming_Pairs" in paper_clean_df.columns
            else pd.Series(False, index=paper_clean_df.index)
        )

        has_any_valid_pair = has_broad_pair | has_framework_pair

        any_valid_pair_count = int(has_any_valid_pair.sum())

        no_valid_pair_count = int(
            len(paper_clean_df) - any_valid_pair_count
        )

        br_paper_count = int(
            paper_clean_df["Use_of_Brennan_Resnick_Framework"]
            .astype(str)
            .str.strip()
            .str.lower()
            .str.startswith("yes")
            .sum()
        )

        final_paper_count = len(final_paper_df)

        removed_pairs_after_final_filter = int(
            len(pair_previous_df) - len(final_pair_df)
        )

        stage4_summary_df = pd.DataFrame([
            {
                "Metric": "Papers received from Stage 3",
                "Count": len(paper_list_df) + len(excluded_before_stage4_df),
            },
            {
                "Metric": "Removed before Stage 4 because candidate matching was not specified",
                "Count": len(excluded_before_stage4_df),
            },
            {
                "Metric": "Papers sent to Stage 4 LLM",
                "Count": len(paper_list_df),
            },
            {
                "Metric": "Paper-level results generated by LLM",
                "Count": len(paper_previous_df),
            },
            {
                "Metric": "Paper-level rows after broad normalization",
                "Count": len(paper_clean_df),
            },
            {
                "Metric": "Papers with broad normalized CT-Programming pairs",
                "Count": broad_included_count,
            },
            {
                "Metric": "Papers with framework-valid CT-Programming pairs",
                "Count": framework_valid_count,
            },
            {
                "Metric": "Papers with at least one valid pair after either filter",
                "Count": any_valid_pair_count,
            },
            {
                "Metric": "Papers excluded because no valid pair was found",
                "Count": no_valid_pair_count,
            },
            {

                "Metric": "Papers using Brennan & Resnick framework",
                "Count": br_paper_count,

            },
            {
                "Metric": "Final papers sent to characterization",
                "Count": final_paper_count,
            },
            {
                "Metric": "Pair-level mapping rows generated by LLM",
                "Count": len(pair_previous_df),
            },
            {
                "Metric": "Pair-level mapping rows after cleaning",
                "Count": len(pair_clean_df),
            },
            {
                "Metric": "Final pair-level mapping rows",
                "Count": len(final_pair_df),
            },
            {
                "Metric": "Pair-level rows removed after final paper filtering",
                "Count": removed_pairs_after_final_filter,
            },
            {
                "Metric": "Paper-level validation sample size",
                "Count": paper_sample_size,
            },
            {
                "Metric": "Pair-level validation sample size",
                "Count": pair_sample_size,
            },
            {
                "Metric": "Estimated total cost USD",
                "Count": cost_summary.get("Estimated total cost USD", 0),
            },
            {
                "Metric": "Paper-level calls",
                "Count": cost_summary.get("Paper-level calls", 0),
            },
            {
                "Metric": "Pair-level calls",
                "Count": cost_summary.get("Pair-level calls", 0),
            },
        ])

        stage4_summary_path = OUTPUT_DIR_4 / "stage4_summary_and_costs.xlsx"

        with pd.ExcelWriter(stage4_summary_path, engine="openpyxl") as writer:
            stage4_summary_df.to_excel(writer, sheet_name="Stage4_summary", index=False)
            usage_summary_df.to_excel(writer, sheet_name="Cost_Summary", index=False)
            usage_df.to_excel(writer, sheet_name="Detailed_Usage", index=False)
            excluded_before_stage4_df.to_excel(writer, sheet_name="Excluded_before_Stage4", index=False)

        st.success("Pipeline completed.")

        st.write(f"Final paper-level results saved to: {paper_output_path}")
        st.write(f"Final pair-level results saved to: {pair_output_path}")
        st.write(f"Audit outputs saved to: {audit_outputs_path}")
        st.write(f"Stage 4 summary saved to: {stage4_summary_path}")
        st.write(f"Paper-level checkpoint saved to: {paper_checkpoint_path}")
        st.write(f"Pair-level checkpoint saved to: {pair_checkpoint_path}")
        st.write(f"LLM costs saved to: {cost_output_path}")

        st.subheader("Cost summary")
        st.dataframe(usage_summary_df, use_container_width=True)

        st.subheader("Final paper-level results for characterization")
        st.dataframe(final_paper_df, use_container_width=True)

        if not final_pair_df.empty:
            st.subheader("Final pair-level results")
            st.dataframe(final_pair_df, use_container_width=True)
        else:
            st.info("No final pair-level results generated.")

if __name__ == "__main__":
    main()