from pathlib import Path
from typing import Any, Dict, List

import re
import time
import pandas as pd
from openai import OpenAI

try:
    import requests
except Exception:
    requests = None

from abstract_screening_logic import (
    DEFAULT_ABSTRACT_COL,
    DEFAULT_MODEL,
    DEFAULT_SYSTEM_ROLE,
    DEFAULT_TITLE_COL,
    build_clean_summary_df,
    build_default_prompt_instructions,
    build_usage_costs_df,
    get_decision_group,
    get_manual_review_group,
    get_script_folder,
    load_and_prepare_abstracts,
    load_screening_configuration,
    process_all_papers,
    read_prompt_config,
    build_validation_sample,
    build_decision_disagreement_summary,
    style_excel_workbook,
    find_potential_duplicates,
)


# ============================================================
# Input / configuration helpers
# ============================================================

def reset_excel_pointer(excel_file) -> None:
    try:
        excel_file.seek(0)
    except Exception:
        pass


def load_input_preview(abstract_file, nrows: int = 5) -> pd.DataFrame:
    reset_excel_pointer(abstract_file)
    preview = pd.read_excel(abstract_file, nrows=nrows, dtype=str)
    preview.columns = [str(c).strip() for c in preview.columns]
    reset_excel_pointer(abstract_file)
    return preview


def get_prompt_preview(rules_file) -> tuple[str, str]:
    if rules_file is None:
        return "", DEFAULT_SYSTEM_ROLE

    reset_excel_pointer(rules_file)
    prompt_config_preview = read_prompt_config(rules_file)
    default_prompt_instructions = build_default_prompt_instructions(prompt_config_preview)
    system_role = prompt_config_preview.get("system_role", DEFAULT_SYSTEM_ROLE)
    reset_excel_pointer(rules_file)

    return default_prompt_instructions, system_role


def validate_required_inputs(api_key, rules_file, abstract_file) -> list[str]:
    errors = []

    if not api_key:
        errors.append("Please provide an OpenAI API key.")

    if rules_file is None:
        errors.append("Please upload the abstract screening rules Excel file.")

    if abstract_file is None:
        errors.append("Please upload the abstracts Excel file.")

    return errors


# ============================================================
# Dataframe builders
# ============================================================

def add_paper_id(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.copy()

    out = df.copy()
    if "Paper_ID" not in out.columns:
        out.insert(0, "Paper_ID", range(1, len(out) + 1))
    return out


def build_decision_subsets(clean_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    if clean_df.empty or "Final_Filter_Decision" not in clean_df.columns:
        return {
            "included": pd.DataFrame(),
            "excluded": pd.DataFrame(),
            "undetermined": pd.DataFrame(),
        }

    decision = clean_df["Final_Filter_Decision"].astype(str).str.strip()

    return {
        "included": clean_df[decision == "Included"].copy(),
        "excluded": clean_df[decision == "Excluded"].copy(),
        "undetermined": clean_df[decision == "Undetermined"].copy(),
    }


def is_manual_review_value(value: Any) -> bool:
    return str(value).strip().lower() in [
        "yes",
        "yes for decision",
        "true",
        "1",
        "manual revision",
        "manual review",
    ]


def build_manual_review_df(clean_df: pd.DataFrame) -> pd.DataFrame:
    if clean_df.empty or "Final_ManualCheckNeeded" not in clean_df.columns:
        return pd.DataFrame()

    return clean_df[clean_df["Final_ManualCheckNeeded"].apply(is_manual_review_value)].copy()


def build_no_manual_review_df(clean_df: pd.DataFrame) -> pd.DataFrame:
    if clean_df.empty:
        return pd.DataFrame()

    if "Final_ManualCheckNeeded" not in clean_df.columns:
        return clean_df.copy()

    return clean_df[~clean_df["Final_ManualCheckNeeded"].apply(is_manual_review_value)].copy()


def build_publication_type_excluded_df(clean_df: pd.DataFrame) -> pd.DataFrame:
    if clean_df.empty or "Final_Publication_Type_Flag" not in clean_df.columns:
        return pd.DataFrame()

    excluded_types = ["Poster", "Conference abstract", "Proceedings-front matter"]

    return clean_df[
        clean_df["Final_Publication_Type_Flag"].astype(str).str.strip().isin(excluded_types)
    ].copy()


def split_screenable_abstracts(
    df: pd.DataFrame,
    abstract_col: str,
    min_abstract_chars: int = 40,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Separate records with usable abstracts from records that cannot be screened at abstract level."""
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame()

    if abstract_col not in df.columns:
        raise ValueError(f"Abstract column not found: {abstract_col}. Available columns: {df.columns.tolist()}")

    out = df.copy()
    abstract_text = out[abstract_col].fillna("").astype(str).str.strip()
    abstract_lower = abstract_text.str.lower()

    invalid_values = {
        "",
        "nan",
        "none",
        "null",
        "no abstract",
        "no abstract available",
        "not available",
        "n/a",
        "na",
    }

    non_screenable_mask = (
        abstract_lower.isin(invalid_values)
        | (abstract_text.str.len() < int(min_abstract_chars))
    )

    non_screenable_df = out[non_screenable_mask].copy()
    screenable_df = out[~non_screenable_mask].copy()

    if not non_screenable_df.empty:
        non_screenable_df["Pre_Abstract_Status"] = "Excluded before abstract screening"
        non_screenable_df["Pre_Abstract_Exclusion_Reason"] = "No usable abstract"

    if not screenable_df.empty:
        screenable_df["Pre_Abstract_Status"] = "Eligible for abstract screening"
        screenable_df["Pre_Abstract_Exclusion_Reason"] = ""

    return screenable_df.reset_index(drop=True), non_screenable_df.reset_index(drop=True)


def build_decision_summary_df(
    clean_df: pd.DataFrame,
    included_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
    undetermined_df: pd.DataFrame,
    manual_review_df: pd.DataFrame,
    publication_type_excluded_df: pd.DataFrame,
    duplicate_summary: pd.DataFrame,
    non_screenable_df: pd.DataFrame | None = None,
    abstract_screening_input_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    non_screenable_count = 0 if non_screenable_df is None else len(non_screenable_df)
    abstract_input_count = 0 if abstract_screening_input_df is None else len(abstract_screening_input_df)

    duplicate_removed_count = 0
    if duplicate_summary is not None and not duplicate_summary.empty:
        if "Number_of_Occurrences" in duplicate_summary.columns:
            duplicate_removed_count = int(
                pd.to_numeric(duplicate_summary["Number_of_Occurrences"], errors="coerce")
                .fillna(1)
                .sub(1)
                .clip(lower=0)
                .sum()
            )
        else:
            duplicate_removed_count = len(duplicate_summary)

    return pd.DataFrame(
        [
            {"Category": "Safety duplicates removed before abstract screening", "Count": duplicate_removed_count},
            {"Category": "Excluded before abstract screening - no usable abstract", "Count": non_screenable_count},
            {"Category": "Abstracts sent to LLM screening", "Count": abstract_input_count},
            {"Category": "Included", "Count": len(included_df)},
            {"Category": "Excluded", "Count": len(excluded_df)},
            {"Category": "Undetermined", "Count": len(undetermined_df)},
            {"Category": "Manual review needed", "Count": len(manual_review_df)},
            {"Category": "Publication-type excluded by LLM", "Count": len(publication_type_excluded_df)},
            {"Category": "Total LLM-screened records", "Count": len(clean_df)},
        ]
    )



def build_prisma_summary_df(
    clean_df: pd.DataFrame,
    included_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
    undetermined_df: pd.DataFrame,
    manual_review_df: pd.DataFrame,
    non_screenable_df: pd.DataFrame | None = None,
    abstract_screening_input_df: pd.DataFrame | None = None,
    next_stage_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Compact PRISMA-oriented counts for the abstract-screening stage."""
    received_count = 0 if abstract_screening_input_df is None else len(abstract_screening_input_df)
    no_abstract_count = 0 if non_screenable_df is None else len(non_screenable_df)
    llm_screened_count = len(clean_df)
    next_stage_count = 0 if next_stage_df is None else len(next_stage_df)

    return pd.DataFrame(
        [
            {"PRISMA_Item": "Records received from title screening", "Count": received_count + no_abstract_count},
            {"PRISMA_Item": "Records excluded before abstract screening - no usable abstract", "Count": no_abstract_count},
            {"PRISMA_Item": "Records assessed by abstract screening", "Count": llm_screened_count},
            {"PRISMA_Item": "Records excluded by abstract screening", "Count": len(excluded_df)},
            {"PRISMA_Item": "Records included after abstract screening", "Count": len(included_df)},
            {"PRISMA_Item": "Records undetermined after abstract screening", "Count": len(undetermined_df)},
            {"PRISMA_Item": "Records requiring manual review after abstract screening", "Count": len(manual_review_df)},
            {"PRISMA_Item": "Records sent to full-text stage", "Count": next_stage_count},
        ]
    )

def build_cost_summary_df(usage_costs_df: pd.DataFrame) -> pd.DataFrame:
    if usage_costs_df.empty:
        return pd.DataFrame(
            [
                {"Metric": "Total input tokens", "Value": 0},
                {"Metric": "Total output tokens", "Value": 0},
                {"Metric": "Total tokens", "Value": 0},
                {"Metric": "Total cost USD", "Value": 0.0},
            ]
        )

    def safe_sum(col: str):
        if col not in usage_costs_df.columns:
            return 0
        return pd.to_numeric(usage_costs_df[col], errors="coerce").fillna(0).sum()

    return pd.DataFrame(
        [
            {"Metric": "Total input tokens", "Value": safe_sum("Total_Input_Tokens")},
            {"Metric": "Total output tokens", "Value": safe_sum("Total_Output_Tokens")},
            {"Metric": "Total tokens", "Value": safe_sum("Total_Tokens")},
            {"Metric": "Total cost USD", "Value": safe_sum("Total_Cost_USD")},
        ]
    )


# ============================================================
# Next-stage + PDF helpers
# ============================================================

def safe_filename(value: Any) -> str:
    text = str(value).strip()
    text = re.sub(r"[^A-Za-z0-9_\-]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "paper"


def clean_doi(value: Any) -> str:
    if pd.isna(value):
        return ""

    doi = str(value).strip()
    doi = doi.replace("https://doi.org/", "")
    doi = doi.replace("http://doi.org/", "")
    doi = doi.replace("doi:", "")
    doi = doi.strip().strip(".")
    return doi


def first_existing_column(df: pd.DataFrame, candidates: List[str]) -> str | None:
    return next((col for col in candidates if col in df.columns), None)


def add_columns_from_results(
    next_stage: pd.DataFrame,
    results_df: pd.DataFrame,
    wanted_columns: List[str],
) -> pd.DataFrame:
    if next_stage.empty or results_df.empty:
        return next_stage

    merge_key = None
    for key in ["Paper", "Title", "Record_ID", "Paper_ID"]:
        if key in next_stage.columns and key in results_df.columns:
            merge_key = key
            break

    if merge_key is None:
        return next_stage

    cols_to_merge = [merge_key]
    rename_map = {}

    for wanted in wanted_columns:
        if wanted in next_stage.columns:
            continue

        possible_names = [
            wanted,
            wanted.lower(),
            wanted.upper(),
            wanted.replace("_", " "),
            wanted.replace("_", "/"),
        ]

        source_col = first_existing_column(results_df, possible_names)
        if source_col is not None and source_col not in cols_to_merge:
            cols_to_merge.append(source_col)
            rename_map[source_col] = wanted

    if len(cols_to_merge) == 1:
        return next_stage

    lookup = results_df[cols_to_merge].drop_duplicates(subset=[merge_key])
    next_stage = next_stage.merge(lookup, on=merge_key, how="left")
    next_stage = next_stage.rename(columns=rename_map)

    return next_stage


def build_next_stage_fulltext_df(
    results_df: pd.DataFrame,
    clean_df: pd.DataFrame,
    include_undetermined: bool = True,
) -> pd.DataFrame:
    if clean_df.empty or "Final_Filter_Decision" not in clean_df.columns:
        return pd.DataFrame()

    decision = clean_df["Final_Filter_Decision"].astype(str).str.strip()
    decision_mask = decision == "Included"

    if include_undetermined:
        decision_mask = decision_mask | (decision == "Undetermined")

    if "Final_ManualCheckNeeded" in clean_df.columns:
        manual_mask = clean_df["Final_ManualCheckNeeded"].apply(is_manual_review_value)
    else:
        manual_mask = pd.Series(False, index=clean_df.index)

    next_stage = clean_df[decision_mask | manual_mask].copy()

    if next_stage.empty:
        return next_stage

    next_stage = add_columns_from_results(
        next_stage=next_stage,
        results_df=results_df,
        wanted_columns=[
            "Record_ID",
            "DOI",
            "URL",
            "PDF_URL",
            "Source_Database",
            "Query_Type",
            "Sheet_Name",
            "Authors",
            "Year",
        ],
    )

    if "Record_ID" in next_stage.columns:
        id_series = next_stage["Record_ID"].fillna("").astype(str).str.strip()
    elif "Paper_ID" in next_stage.columns:
        id_series = next_stage["Paper_ID"].fillna("").astype(str).str.strip()
    else:
        id_series = pd.Series([f"P{i:05d}" for i in range(1, len(next_stage) + 1)], index=next_stage.index)
        next_stage.insert(0, "Paper_ID", id_series)

    id_series = id_series.replace("", pd.NA).fillna(
        pd.Series([f"P{i:05d}" for i in range(1, len(next_stage) + 1)], index=next_stage.index)
    )

    next_stage["Full_Text_ID"] = id_series.apply(safe_filename)
    next_stage["PDF_Filename"] = next_stage["Full_Text_ID"] + ".pdf"

    def stage3_status(row):
        if "Final_ManualCheckNeeded" in row.index and is_manual_review_value(row.get("Final_ManualCheckNeeded", "")):
            return "Manual review required after abstract screening"
        if str(row.get("Final_Filter_Decision", "")).strip() == "Undetermined":
            return "Undetermined after abstract screening"
        return "Included after abstract screening"

    next_stage["Stage3_Status"] = next_stage.apply(stage3_status, axis=1)

    if "PDF_Path" not in next_stage.columns:
        next_stage["PDF_Path"] = ""

    next_stage["PDF_Retrieval_Status"] = "Pending"
    next_stage["PDF_Retrieval_Source"] = ""
    next_stage["PDF_Retrieval_Message"] = ""
    next_stage["File_Exists"] = False

    next_stage["Full_Text_Screening_Decision"] = ""
    next_stage["Full_Text_Exclusion_Reason"] = ""
    next_stage["Full_Text_Reviewer_Notes"] = ""

    preferred_columns = [
        "Paper_ID",
        "Record_ID",
        "Full_Text_ID",
        "Paper",
        "Title",
        "Abstract",
        "Authors",
        "Year",
        "DOI",
        "URL",
        "PDF_URL",
        "Source_Database",
        "Query_Type",
        "Sheet_Name",
        "Final_Filter_Decision",
        "Final_ManualCheckNeeded",
        "Stage3_Status",
        "PDF_Filename",
        "PDF_Path",
        "PDF_Retrieval_Status",
        "PDF_Retrieval_Source",
        "PDF_Retrieval_Message",
        "File_Exists",
        "Full_Text_Screening_Decision",
        "Full_Text_Exclusion_Reason",
        "Full_Text_Reviewer_Notes",
    ]

    ordered_existing = [c for c in preferred_columns if c in next_stage.columns]
    remaining = [c for c in next_stage.columns if c not in ordered_existing]

    return next_stage[ordered_existing + remaining]


def get_unpaywall_pdf_url(doi: str, email: str, timeout: int = 20) -> tuple[str, str]:
    if requests is None:
        return "", "The requests package is not installed. Run: pip install requests"

    doi = clean_doi(doi)
    if not doi:
        return "", "No DOI available"

    if not email:
        return "", "Unpaywall email missing"

    api_url = f"https://api.unpaywall.org/v2/{doi}"

    try:
        response = requests.get(api_url, params={"email": email}, timeout=timeout)
        if response.status_code != 200:
            return "", f"Unpaywall returned status {response.status_code}"

        data = response.json()
        best = data.get("best_oa_location") or {}
        pdf_url = best.get("url_for_pdf") or ""

        if pdf_url:
            return pdf_url, "Found via Unpaywall best_oa_location"

        for location in data.get("oa_locations", []) or []:
            pdf_url = location.get("url_for_pdf") or ""
            if pdf_url:
                return pdf_url, "Found via Unpaywall oa_locations"

        return "", "No open-access PDF URL found in Unpaywall"

    except Exception as exc:
        return "", f"Unpaywall lookup failed: {exc}"


def download_pdf(url: str, output_path: Path, timeout: int = 40) -> tuple[bool, str]:
    if requests is None:
        return False, "The requests package is not installed. Run: pip install requests"

    if not url:
        return False, "No PDF URL available"

    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0 Safari/537.36"
            ),
            "Accept": (
                "application/pdf,text/html,"
                "application/xhtml+xml,application/xml;q=0.9,"
                "*/*;q=0.8"
            ),
        }

        response = requests.get(
            url,
            headers=headers,
            timeout=timeout,
            allow_redirects=True,
        )

        if response.status_code == 403:
            return (
                False,
                "Access blocked by publisher (HTTP 403). "
                "Manual download is required through browser "
                "or institutional access."
            )

        if response.status_code == 401:
            return (
                False,
                "Authentication required (HTTP 401). "
                "Manual download is required through browser "
                "or institutional access."
            )

        if response.status_code == 429:
            return (
                False,
                "Too many requests (HTTP 429). "
                "Retry later or download manually."
            )

        if response.status_code != 200:
            return False, f"Download returned status {response.status_code}"

        content_type = response.headers.get(
            "Content-Type", ""
        ).lower()

        content = response.content

        if not (
            content.startswith(b"%PDF")
            or "pdf" in content_type
        ):
            return (
                False,
                "Downloaded content is not a PDF. "
                "Manual browser download may be required."
            )

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        output_path.write_bytes(content)

        return True, "Downloaded successfully"

    except Exception as exc:
        return False, f"Download failed: {exc}"


def retrieve_pdfs_for_next_stage(
    next_stage_df: pd.DataFrame,
    pdf_repository_dir: Path,
    unpaywall_email: str = "",
    delay_seconds: float = 0.2,
) -> pd.DataFrame:
    if next_stage_df.empty:
        return next_stage_df.copy()

    out = next_stage_df.copy()
    pdf_repository_dir = Path(pdf_repository_dir)
    pdf_repository_dir.mkdir(parents=True, exist_ok=True)

    for idx, row in out.iterrows():
        filename = str(row.get("PDF_Filename", "")).strip()
        if not filename:
            filename = f"{safe_filename(row.get('Full_Text_ID', idx + 1))}.pdf"
            out.at[idx, "PDF_Filename"] = filename

        target_path = pdf_repository_dir / filename
        out.at[idx, "PDF_Path"] = str(target_path)

        if target_path.exists() and target_path.stat().st_size > 0:
            out.at[idx, "PDF_Retrieval_Status"] = "Already exists"
            out.at[idx, "PDF_Retrieval_Source"] = "Local repository"
            out.at[idx, "PDF_Retrieval_Message"] = "File already exists"
            out.at[idx, "File_Exists"] = True
            continue

        candidate_urls = []
        for col in ["PDF_URL", "PDF URL", "pdf_url", "URL", "Url", "url", "Link", "link"]:
            if col in out.columns:
                value = str(row.get(col, "")).strip()
                if value and value.lower() not in ["nan", "none"]:
                    candidate_urls.append((col, value))

        downloaded = False
        messages = []

        for source_col, url in candidate_urls:
            ok, message = download_pdf(url, target_path)
            messages.append(f"{source_col}: {message}")
            if ok:
                out.at[idx, "PDF_Retrieval_Status"] = "Downloaded"
                out.at[idx, "PDF_Retrieval_Source"] = source_col
                out.at[idx, "PDF_Retrieval_Message"] = message
                out.at[idx, "File_Exists"] = True
                downloaded = True
                break

        if downloaded:
            time.sleep(delay_seconds)
            continue

        doi_col = first_existing_column(out, ["DOI", "doi", "DOI_Clean", "Doi"])
        doi = clean_doi(row.get(doi_col, "")) if doi_col else ""

        if doi and unpaywall_email:
            pdf_url, lookup_message = get_unpaywall_pdf_url(doi, unpaywall_email)
            if pdf_url:
                ok, message = download_pdf(pdf_url, target_path)
                if ok:
                    out.at[idx, "PDF_Retrieval_Status"] = "Downloaded"
                    out.at[idx, "PDF_Retrieval_Source"] = "Unpaywall"
                    out.at[idx, "PDF_Retrieval_Message"] = f"{lookup_message}; {message}"
                    out.at[idx, "File_Exists"] = True
                    time.sleep(delay_seconds)
                    continue
                messages.append(f"Unpaywall download: {message}")
            else:
                messages.append(lookup_message)

        out.at[idx, "PDF_Retrieval_Status"] = "Manual retrieval needed"
        out.at[idx, "PDF_Retrieval_Source"] = ""
        out.at[idx, "PDF_Retrieval_Message"] = "; ".join(messages) if messages else "No usable URL or OA PDF found"
        out.at[idx, "File_Exists"] = False
        time.sleep(delay_seconds)

    return out


# ============================================================
# Dataframe output organization
# ============================================================

def build_all_output_dataframes(
    results: List[Dict[str, Any]],
    duplicate_summary: pd.DataFrame | None,
    number_of_runs: int,
    validation_sample_percent: int,
    include_undetermined_for_fulltext: bool = True,
    non_screenable_df: pd.DataFrame | None = None,
    abstract_screening_input_df: pd.DataFrame | None = None,
) -> Dict[str, pd.DataFrame]:
    results_df = pd.DataFrame(results)
    results_df = add_paper_id(results_df)

    if duplicate_summary is None:
        duplicate_summary = pd.DataFrame()
    if non_screenable_df is None:
        non_screenable_df = pd.DataFrame()
    if abstract_screening_input_df is None:
        abstract_screening_input_df = pd.DataFrame()

    clean_df = build_clean_summary_df(results_df)
    clean_df = add_paper_id(clean_df)

    subsets = build_decision_subsets(clean_df)
    included_df = subsets["included"]
    excluded_df = subsets["excluded"]
    undetermined_df = subsets["undetermined"]

    manual_review_df = build_manual_review_df(clean_df)
    no_manual_review_df = build_no_manual_review_df(clean_df)
    publication_type_excluded_df = build_publication_type_excluded_df(clean_df)

    summary_df = build_decision_summary_df(
        clean_df=clean_df,
        included_df=included_df,
        excluded_df=excluded_df,
        undetermined_df=undetermined_df,
        manual_review_df=manual_review_df,
        publication_type_excluded_df=publication_type_excluded_df,
        duplicate_summary=duplicate_summary,
        non_screenable_df=non_screenable_df,
        abstract_screening_input_df=abstract_screening_input_df,
    )

    usage_costs_df = build_usage_costs_df(results_df, int(number_of_runs))
    cost_summary_df = build_cost_summary_df(usage_costs_df)

    decision_disagreements_df = build_decision_disagreement_summary(results_df, int(number_of_runs))

    validation_with_llm_df, manual_template_df, sampling_summary_df = build_validation_sample(
        results_df=results_df,
        number_of_runs=int(number_of_runs),
        sample_percent=int(validation_sample_percent),
    )

    next_stage_df = build_next_stage_fulltext_df(
        results_df,
        clean_df,
        include_undetermined=include_undetermined_for_fulltext,
    )

    prisma_summary_df = build_prisma_summary_df(
        clean_df=clean_df,
        included_df=included_df,
        excluded_df=excluded_df,
        undetermined_df=undetermined_df,
        manual_review_df=manual_review_df,
        non_screenable_df=non_screenable_df,
        abstract_screening_input_df=abstract_screening_input_df,
        next_stage_df=next_stage_df,
    )

    return {
        "All_Clean_Summary": clean_df,
        "Summary": summary_df,
        "PRISMA_summary": prisma_summary_df,
        "Included": included_df,
        "Excluded": excluded_df,
        "Undetermined": undetermined_df,
        "Manual_Review": manual_review_df,
        "No_Manual_Review": no_manual_review_df,
        "Publication_Type_Excluded": publication_type_excluded_df,
        "Duplicate_Summary": duplicate_summary,
        "Non_Screenable_Records": non_screenable_df,
        "Abstract_Screening_Input": abstract_screening_input_df,

        "Validation_Sample_With_LLM": validation_with_llm_df,
        "Manual_Coding_Template": manual_template_df,
        "Sampling_Summary": sampling_summary_df,
        "Manual_Review_Needed": manual_review_df,
        "Decision_Disagreements": decision_disagreements_df,

        "Usage_Costs": usage_costs_df,
        "Cost_Summary": cost_summary_df,

        "Next_Stage_Fulltext": next_stage_df,

        "Complete_Audit_All_Columns": results_df,
    }


# ============================================================
# Workbook writers
# ============================================================

def write_workbook(path: Path, sheets: Dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_sheet_name = sheet_name[:31]
            if df is None:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

    style_excel_workbook(path)


def write_reorganized_output_workbooks(
    base_output_path: Path,
    results: List[Dict[str, Any]],
    duplicate_summary: pd.DataFrame | None,
    number_of_runs: int,
    validation_sample_percent: int,
    include_undetermined_for_fulltext: bool = True,
    non_screenable_df: pd.DataFrame | None = None,
    abstract_screening_input_df: pd.DataFrame | None = None,
    auto_retrieve_pdfs: bool = False,
    pdf_repository_dir: str | Path | None = None,
    unpaywall_email: str = "",
    potential_duplicates_df: pd.DataFrame | None = None,
) -> Dict[str, Any]:
    stem = base_output_path.stem

    results_output_path = base_output_path.with_name(f"{stem}_results.xlsx")
    validation_output_path = base_output_path.with_name(f"{stem}_author_validation.xlsx")
    costs_output_path = base_output_path.with_name(f"{stem}_llm_costs.xlsx")
    next_stage_output_path = base_output_path.with_name(f"{stem}_next_stage_fulltext.xlsx")

    dfs = build_all_output_dataframes(
        results=results,
        duplicate_summary=duplicate_summary,
        number_of_runs=int(number_of_runs),
        validation_sample_percent=int(validation_sample_percent),
        include_undetermined_for_fulltext=include_undetermined_for_fulltext,
        non_screenable_df=non_screenable_df,
        abstract_screening_input_df=abstract_screening_input_df,
    )

    if potential_duplicates_df is None:
        potential_duplicates_df = pd.DataFrame()

    if "Summary" in dfs:
        duplicate_count_row = pd.DataFrame([
            {
                "Category": "Potential duplicates detected during abstract screening",
                "Count": len(potential_duplicates_df),
            }
        ])
        dfs["Summary"] = pd.concat([dfs["Summary"], duplicate_count_row], ignore_index=True)

    if auto_retrieve_pdfs:
        if pdf_repository_dir is None or str(pdf_repository_dir).strip() == "":
            pdf_repository_dir = base_output_path.parent / "papers_repository" / "pdfs"
        dfs["Next_Stage_Fulltext"] = retrieve_pdfs_for_next_stage(
            next_stage_df=dfs["Next_Stage_Fulltext"],
            pdf_repository_dir=Path(pdf_repository_dir),
            unpaywall_email=unpaywall_email,
        )

    write_workbook(
        results_output_path,
        {
            "Summary": dfs["Summary"],
            "PRISMA_summary": dfs["PRISMA_summary"],
            "Included": dfs["Included"],
            "Excluded": dfs["Excluded"],
            "Undetermined": dfs["Undetermined"],
            "Manual_Review": dfs["Manual_Review"],
            "No_Manual_Review": dfs["No_Manual_Review"],
            "Publication_Type_Excluded": dfs["Publication_Type_Excluded"],
            "Non_Screenable_Records": dfs["Non_Screenable_Records"],
            "Abstract_Screening_Input": dfs["Abstract_Screening_Input"],
            "All_Clean_Summary": dfs["All_Clean_Summary"],
            "Duplicate_Summary": dfs["Duplicate_Summary"],
        },
    )

    write_workbook(
        validation_output_path,
        {
            "Validation_Sample_With_LLM": dfs["Validation_Sample_With_LLM"],
            "Manual_Coding_Template": dfs["Manual_Coding_Template"],
            "Sampling_Summary": dfs["Sampling_Summary"],
            "Manual_Review_Needed": dfs["Manual_Review_Needed"],
            "Decision_Disagreements": dfs["Decision_Disagreements"],
            "Potential_Duplicates_During_Abstract": potential_duplicates_df,
        },
    )

    write_workbook(
        costs_output_path,
        {
            "Usage_Costs": dfs["Usage_Costs"],
            "Cost_Summary": dfs["Cost_Summary"],
        },
    )

    write_workbook(
        next_stage_output_path,
        {
            "Next_Stage_Fulltext": dfs["Next_Stage_Fulltext"],
        },
    )

    return {
        "results_output_path": results_output_path,
        "validation_output_path": validation_output_path,
        "costs_output_path": costs_output_path,
        "next_stage_output_path": next_stage_output_path,
        "dataframes": dfs,
    }


# ============================================================
# Processing workflow
# ============================================================

def run_abstract_screening(
    api_key: str,
    rules_file,
    abstract_file,
    prompt_instructions: str,
    initial_system_role: str,
    settings: Dict[str, Any],
    log_callback=None,
    progress_callback=None,
) -> Dict[str, Any]:
    errors = validate_required_inputs(api_key, rules_file, abstract_file)
    if errors:
        raise ValueError(" ".join(errors))

    client = OpenAI(api_key=api_key)

    reset_excel_pointer(rules_file)
    config = load_screening_configuration(rules_file)

    system_role = config["prompt_config"].get(
        "system_role",
        initial_system_role or DEFAULT_SYSTEM_ROLE,
    )

    reset_excel_pointer(abstract_file)
    df, duplicate_summary, original_count, deduplicated_count, removed_duplicates_count = load_and_prepare_abstracts(
        abstract_file=abstract_file,
        title_col=settings["title_col"],
        abstract_col=settings["abstract_col"],
        deduplicate_titles=settings["deduplicate_titles"],
    )

    abstract_screening_input_df, non_screenable_df = split_screenable_abstracts(
        df=df,
        abstract_col=settings["abstract_col"],
        min_abstract_chars=int(settings.get("min_abstract_chars", 40)),
    )

    potential_duplicates_df = find_potential_duplicates(df)

    output_dir = Path(settings.get("output_dir", get_script_folder() / "output"))
    temp_dir = Path(settings.get("temp_dir", get_script_folder() / "temp"))

    output_dir.mkdir(parents=True,exist_ok=True,)

    temp_dir.mkdir(parents=True,exist_ok=True,)

    base_output_path = (output_dir/ settings["output_name"])

    checkpoint_path = (temp_dir/ f"{base_output_path.stem}_checkpoint_raw.xlsx")

    results = process_all_papers(
        df=abstract_screening_input_df,
        client=client,
        model=settings["model"],
        title_col=settings["title_col"],
        abstract_col=settings["abstract_col"],
        prompt_instructions=prompt_instructions,
        system_role=system_role,
        rules_prompt=config["rules_prompt"],
        schema_prompt=config["schema_prompt"],
        examples_prompt=config["examples_prompt"],
        json_schema=config["json_schema"],
        output_fields=config["output_fields"],
        number_of_runs=int(settings["number_of_runs"]),
        checkpoint_path=checkpoint_path,
        potential_duplicates_df=potential_duplicates_df,
        log_callback=log_callback,
        progress_callback=progress_callback,
    )

    output_paths = write_reorganized_output_workbooks(
        base_output_path=base_output_path,
        results=results,
        duplicate_summary=duplicate_summary,
        number_of_runs=int(settings["number_of_runs"]),
        validation_sample_percent=int(settings["validation_sample_percent"]),
        include_undetermined_for_fulltext=bool(settings.get("include_undetermined_for_fulltext", True)),
        non_screenable_df=non_screenable_df,
        abstract_screening_input_df=abstract_screening_input_df,
        auto_retrieve_pdfs=bool(settings.get("auto_retrieve_pdfs", False)),
        pdf_repository_dir=settings.get("pdf_repository_dir", ""),
        unpaywall_email=settings.get("unpaywall_email", ""),
        potential_duplicates_df=potential_duplicates_df,
    )

    return {
        "results": results,
        "duplicate_summary": duplicate_summary,
        "original_count": original_count,
        "deduplicated_count": deduplicated_count,
        "removed_duplicates_count": removed_duplicates_count,
        "non_screenable_count": len(non_screenable_df),
        "abstract_screening_input_count": len(abstract_screening_input_df),
        "non_screenable_df": non_screenable_df,
        "abstract_screening_input_df": abstract_screening_input_df,
        "base_output_path": base_output_path,
        "checkpoint_path": checkpoint_path,
        **output_paths,
    }


# ============================================================
# Dashboard preparation
# ============================================================

def build_dashboard_data(results: List[Dict[str, Any]], number_of_runs: int) -> Dict[str, Any]:
    results_df = pd.DataFrame(results)

    clean_results_df = build_clean_summary_df(results_df)
    usage_costs_df = build_usage_costs_df(results_df, int(number_of_runs))

    included_preview = get_decision_group(clean_results_df, "Included")
    excluded_preview = get_decision_group(clean_results_df, "Excluded")
    undetermined_preview = get_decision_group(clean_results_df, "Undetermined")
    manual_review_preview = get_manual_review_group(clean_results_df)

    total_api_runs = len(results_df) * int(number_of_runs)
    if "Total_Tokens" in results_df.columns:
        total_tokens = (
            pd.to_numeric(
                results_df["Total_Tokens"],
                errors="coerce"
            )
            .fillna(0)
            .sum()
        )
    else:
        total_tokens = 0

    if "Total_Cost_USD" in results_df.columns:
        total_cost = (
            pd.to_numeric(
                results_df["Total_Cost_USD"],
                errors="coerce"
            )
            .fillna(0)
            .sum()
        )
    else:
        total_cost = 0.0

    return {
        "results_df": results_df,
        "clean_results_df": clean_results_df,
        "usage_costs_df": usage_costs_df,
        "included_preview": included_preview,
        "excluded_preview": excluded_preview,
        "undetermined_preview": undetermined_preview,
        "manual_review_preview": manual_review_preview,
        "total_api_runs": total_api_runs,
        "total_tokens": total_tokens,
        "total_cost": total_cost,
    }
