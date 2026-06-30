
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

import fitz  # PyMuPDF
import pandas as pd
from openai import OpenAI


# ============================================================
# Constants
# ============================================================
AUTO_INPUT_MODE = "Automatically use Next_stage_included from previous screening"
CUSTOM_INPUT_MODE = "Use custom file with Paper_ID and PDF_path"


# ============================================================
# Small utilities
# ============================================================
def reset_excel_pointer(excel_file) -> None:
    try:
        excel_file.seek(0)
    except Exception:
        pass


def safe_str(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_active(value: Any) -> bool:
    return safe_str(value).lower() in ["yes", "true", "1", "active"]


# ============================================================
# PDF reading and targeted context
# ============================================================
def extract_pdf_text(pdf_path: Path, max_chars: int = 120_000) -> str:
    text_parts = []

    with fitz.open(pdf_path) as doc:
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text("text")
            if text.strip():
                text_parts.append(f"\n--- Page {page_num} ---\n{text}")

    full_text = "\n".join(text_parts).strip()

    if len(full_text) > max_chars:
        full_text = full_text[:max_chars] + "\n\n[TEXT TRUNCATED DUE TO LENGTH]"

    return full_text


def split_text_by_pages(paper_text: str) -> List[Tuple[int | None, str]]:
    pattern = r"\n?--- Page (\d+) ---\n"
    parts = re.split(pattern, paper_text)
    pages: List[Tuple[int | None, str]] = []

    if parts and parts[0].strip():
        pages.append((None, parts[0].strip()))

    for i in range(1, len(parts), 2):
        page_number = int(parts[i])
        page_text = parts[i + 1].strip() if i + 1 < len(parts) else ""
        if page_text:
            pages.append((page_number, page_text))

    return pages


def extract_location_terms(location_hint: str) -> List[str]:
    if not isinstance(location_hint, str) or not location_hint.strip():
        return []

    text = location_hint.strip()

    patterns = [
        r"Section\s+\d+(?:\.\d+)*\s*[^;\],)]*",
        r"Table\s+\d+[A-Za-z]?",
        r"Figure\s+\d+[A-Za-z]?",
        r"Appendix\s+[A-Za-z0-9.]+",
    ]

    terms: List[str] = []

    for pattern in patterns:
        for match in re.findall(pattern, text, flags=re.IGNORECASE):
            cleaned = match.strip(" ;,])")
            if cleaned and cleaned.lower() not in [t.lower() for t in terms]:
                terms.append(cleaned)

    soft_parts = re.split(r"[;\[\]\(\)]", text)
    for part in soft_parts:
        cleaned = part.strip(" ;,.-")
        if 4 <= len(cleaned) <= 90 and cleaned.lower() not in [t.lower() for t in terms]:
            terms.append(cleaned)

    return terms


def extract_relevant_sections(
    paper_text: str,
    location_hint: str = "",
    max_chars: int = 60_000,
    additional_keywords: str = "",
) -> str:
    pages = split_text_by_pages(paper_text)

    if not pages:
        return paper_text[:max_chars]

    priority_keywords = [
        term.lower() for term in extract_location_terms(location_hint) if term.strip()
    ]

    broader_keywords = [
        "computational thinking",
        "programming",
        "coding",
        "scratch",
        "block-based",
        "block based",
        "algorithm",
        "abstraction",
        "decomposition",
        "pattern recognition",
        "debugging",
        "iteration",
        "loops",
        "conditionals",
        "variables",
        "framework",
        "mapping",
        "curriculum",
        "intervention",
        "method",
        "methodology",
        "materials and methods",
        "procedure",
        "activities",
        "learning activities",
        "assessment",
        "results",
        "findings",
        "discussion",
        "table",
        "figure",
    ]

    if additional_keywords:
        broader_keywords.extend(
            [k.strip().lower() for k in additional_keywords.split(",") if k.strip()]
        )

    selected_pages: List[Tuple[int | None, str, str]] = []
    selected_page_numbers = set()

    def add_page(page_number: int | None, page_text: str, reason: str) -> None:
        key = page_number if page_number is not None else f"no_page_{len(selected_pages)}"
        if key in selected_page_numbers:
            return
        selected_page_numbers.add(key)
        selected_pages.append((page_number, page_text, reason))

    # First: previous-stage location hint
    if priority_keywords:
        for page_number, page_text in pages:
            lower = page_text.lower()
            if any(keyword in lower for keyword in priority_keywords):
                add_page(page_number, page_text, "matched previous-stage location hint")

    # Second: broader CT/programming/methods/results/discussion context
    if len("\n".join(page for _, page, _ in selected_pages)) < 8_000:
        for page_number, page_text in pages:
            lower = page_text.lower()
            if any(keyword in lower for keyword in broader_keywords):
                add_page(
                    page_number,
                    page_text,
                    "matched broader CT/programming or paper-section keyword",
                )

    # Fallback
    if not selected_pages:
        for page_number, page_text in pages[:8]:
            add_page(page_number, page_text, "fallback first pages")

    context_parts = []
    for page_number, page_text, reason in selected_pages:
        page_label = f"Page {page_number}" if page_number is not None else "Unnumbered text"
        context_parts.append(f"--- {page_label} | Selection reason: {reason} ---\n{page_text}")

    targeted_text = "\n\n".join(context_parts).strip()

    if len(targeted_text) > max_chars:
        targeted_text = targeted_text[:max_chars] + "\n\n[TARGETED TEXT TRUNCATED DUE TO LENGTH]"

    return targeted_text


def build_targeted_paper_context(
    paper_text: str,
    title: str = "",
    summary: str = "",
    matching_evidence_location: str = "",
    max_chars: int = 60_000,
    additional_keywords: str = "",
) -> str:
    sections = []

    if title:
        sections.append(f"TITLE FROM PREVIOUS SCREENING:\n{title}")

    if summary:
        sections.append(f"PREVIOUS SCREENING SUMMARY:\n{summary}")

    if matching_evidence_location:
        sections.append(
            "MATCHING EVIDENCE LOCATION FROM PREVIOUS SCREENING "
            "(navigation hint only; verify independently in the paper):\n"
            f"{matching_evidence_location}"
        )

    targeted_text = extract_relevant_sections(
        paper_text=paper_text,
        location_hint=matching_evidence_location,
        max_chars=max_chars,
        additional_keywords=additional_keywords,
    )

    sections.append(
        "TARGETED PAPER TEXT SELECTED FOR THIS PAIRING STAGE:\n"
        f"{targeted_text}"
    )

    return "\n\n".join(sections).strip()


# ============================================================
# Schema, dependency rules and prompt templates
# ============================================================
def read_sheet_preamble(excel_file, sheet_name: str) -> str:
    reset_excel_pointer(excel_file)
    raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    header_row = None
    for i, row in raw_df.iterrows():
        values = [str(v).strip() for v in row.dropna().tolist()]
        if "Field" in values and "Description" in values and "Allowed values" in values:
            header_row = i
            break

    if header_row is None:
        return ""

    lines = []
    for i in range(header_row):
        row_text = " ".join(
            str(v).strip()
            for v in raw_df.iloc[i].dropna().tolist()
            if str(v).strip()
        )
        if row_text:
            lines.append(row_text)

    return "\n".join(lines)


def read_coding_schema(excel_file, sheet_name: str) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    required_headers = {
        "Field",
        "Description",
        "Allowed values",
        "Definition",
        "Coding note / Decision rules",
    }

    header_row = None
    for i, row in raw_df.iterrows():
        row_values = set(str(v).strip() for v in row.dropna().tolist())
        if required_headers.issubset(row_values):
            header_row = i
            break

    if header_row is None:
        raise ValueError(f"Could not find header row in {sheet_name}. Expected columns: {required_headers}")

    reset_excel_pointer(excel_file)
    schema_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row)
    schema_df.columns = [str(c).strip() for c in schema_df.columns]

    missing = required_headers - set(schema_df.columns)
    if missing:
        raise ValueError(f"Missing required columns in {sheet_name}: {missing}")

    schema_df["Field"] = schema_df["Field"].ffill()
    schema_df["Description"] = schema_df["Description"].ffill()
    return schema_df.dropna(subset=["Field"])


def build_schema_prompt(schema_df: pd.DataFrame) -> str:
    sections = []

    for field, group in schema_df.groupby("Field", sort=False):
        description = str(group["Description"].iloc[0]).strip()
        allowed_notes = []

        for _, row in group.iterrows():
            allowed = str(row["Allowed values"]).strip()
            definition = str(row["Definition"]).strip()
            coding_note = str(row["Coding note / Decision rules"]).strip()

            if allowed.lower() in ["nan", ""]:
                continue

            text = f"- {allowed}"
            if definition.lower() not in ["nan", ""]:
                text += f"\n  Definition: {definition}"
            if coding_note.lower() not in ["nan", ""]:
                text += f"\n  Decision rules: {coding_note}"

            allowed_notes.append(text)

        sections.append(
            f"FIELD: {field}\n"
            f"DESCRIPTION: {description}\n"
            f"ALLOWED VALUES / NOTES:\n"
            + "\n".join(allowed_notes)
        )

    return "\n\n".join(sections)


def get_output_fields(schema_df: pd.DataFrame, exclude_fields: List[str] | None = None) -> List[str]:
    exclude_fields = exclude_fields or []
    fields = list(dict.fromkeys(schema_df["Field"].dropna().astype(str).str.strip().tolist()))
    return [field for field in fields if field not in exclude_fields]


def make_json_schema(
    fields: List[str],
    schema_name: str,
    schema_df: pd.DataFrame | None = None,
) -> Dict[str, Any]:

    properties = {}

    for field in fields:
        field_schema = {"type": "string"}

        if schema_df is not None:
            group = schema_df[
                schema_df["Field"].astype(str).str.strip() == field
            ]

            allowed_values = (
                group["Allowed values"]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
            )

            allowed_values = [
                value for value in allowed_values
                if value and value.lower() != "nan"
            ]

            is_open_text = any(
                "open text" in value.lower()
                or value.lower() in ["open", "text", "free text"]
                for value in allowed_values
            )

            if allowed_values and not is_open_text:
                field_schema["enum"] = list(dict.fromkeys(allowed_values))

        properties[field] = field_schema

    return {
        "type": "object",
        "properties": properties,
        "required": fields,
        "additionalProperties": False,
    }



def read_dependency_rules(excel_file, sheet_name: str) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    rules_df = pd.read_excel(excel_file, sheet_name=sheet_name)
    rules_df.columns = [str(c).strip() for c in rules_df.columns]

    required = {
        "Rule_ID",
        "If_Field",
        "If_Value",
        "Then_Field",
        "Expected_Value",
        "Rule_Type",
        "Message",
    }

    missing = required - set(rules_df.columns)
    if missing:
        raise ValueError(f"Missing required columns in {sheet_name}: {missing}")

    return rules_df


def build_dependency_rules_prompt(rules_df: pd.DataFrame) -> str:
    rules = []

    for _, row in rules_df.iterrows():
        rule_id = str(row["Rule_ID"]).strip()
        if rule_id.lower() == "nan":
            continue

        rules.append(
            f"- {rule_id}: IF {str(row['If_Field']).strip()} = {str(row['If_Value']).strip()}, "
            f"THEN {str(row['Then_Field']).strip()} must be {str(row['Expected_Value']).strip()}. "
            f"Rule type: {str(row['Rule_Type']).strip()}. Note: {str(row['Message']).strip()}"
        )

    return "\n".join(rules)


def read_prompt_templates(excel_file) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    df = pd.read_excel(excel_file, sheet_name="Prompt_Templates")
    df.columns = [str(c).strip() for c in df.columns]

    required = {"Prompt_ID", "Stage", "Prompt_Text", "Active"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Prompt_Templates: {missing}")

    df = df.dropna(subset=["Stage", "Prompt_Text"]).copy()
    return df


def get_active_prompt_template(excel_file, stage: str) -> str:
    df = read_prompt_templates(excel_file)

    active = df[
        (df["Stage"].astype(str).str.strip() == stage)
        & (df["Active"].apply(normalize_active))
    ]

    if active.empty:
        raise ValueError(f"No active prompt found for stage '{stage}' in Prompt_Templates.")

    if len(active) > 1:
        ids = ", ".join(active["Prompt_ID"].astype(str).tolist())
        raise ValueError(
            f"More than one active prompt found for stage '{stage}'. Active prompts: {ids}"
        )

    return str(active.iloc[0]["Prompt_Text"]).strip()


def build_prompt_from_template(template: str, **kwargs: Any) -> str:
    safe_kwargs = {
        key: ("Not available" if value is None or str(value).strip() == "" else value)
        for key, value in kwargs.items()
    }

    try:
        return template.format(**safe_kwargs).strip()
    except KeyError as e:
        missing = str(e).strip("'")
        raise ValueError(
            f"Prompt template contains an unknown placeholder: {{{missing}}}. "
            "Check Prompt_Templates in the rules Excel file."
        )


# ============================================================
# Normalization and pair parsing
# ============================================================
def read_normalization_rules(excel_file) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    df = pd.read_excel(excel_file, sheet_name="Normalization_Rules")
    df.columns = [str(c).strip() for c in df.columns]

    required = {"Type", "Raw_Term", "Normalized_Value"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Normalization_Rules: {missing}")

    df = df.dropna(subset=["Type", "Raw_Term", "Normalized_Value"]).copy()
    df["Type"] = df["Type"].astype(str).str.strip()
    df["Raw_Term"] = df["Raw_Term"].astype(str).str.strip().str.lower()
    df["Normalized_Value"] = df["Normalized_Value"].astype(str).str.strip()

    return df


def clean_term(term: str) -> str:
    term = str(term).strip().lower().replace("_", " ").replace("-", " ")
    term = re.sub(r"[^\w\s/]", " ", term)
    term = re.sub(r"\s+", " ", term)
    return term.strip()


def normalize_terms_from_phrase(phrase: str, term_type: str, rules_df: pd.DataFrame) -> List[str]:
    clean_phrase = clean_term(phrase)

    subset = rules_df[
        rules_df["Type"].str.lower() == term_type.lower()
    ].copy()

    subset["Clean_Raw_Term"] = subset["Raw_Term"].apply(clean_term)

    matches: List[str] = []

    for _, row in subset.iterrows():
        raw = str(row["Clean_Raw_Term"]).strip()
        normalized = str(row["Normalized_Value"]).strip()

        if not raw or not normalized:
            continue

        # Conservative rule: only exact dictionary match
        if raw == clean_phrase:
            if normalized not in matches:
                matches.append(normalized)

    return matches


def parse_matching_pairs(mapping_text: str) -> List[Tuple[str, str]]:
    if not isinstance(mapping_text, str):
        return []

    text = mapping_text.strip()

    invalid_values = {
        "",
        "nan",
        "no existing mapping",
        "no valid pairs",
        "undetermined",
        "manual revision",
        "not specified",
    }

    if text.lower() in invalid_values:
        return []

    if text.startswith("[") and text.endswith("]"):
        text = text[1:-1]

    pairs: List[Tuple[str, str]] = []
    chunks = [chunk.strip() for chunk in text.split(";") if chunk.strip()]

    for chunk in chunks:
        match = re.match(r"(.+?)\s*:\s*\((.*?)\)", chunk)
        if not match:
            continue

        ct_raw = match.group(1).strip()
        prog_items = [p.strip() for p in match.group(2).strip().split(",") if p.strip()]
        pairs.extend((ct_raw, prog_raw) for prog_raw in prog_items)

    return pairs


def build_valid_pairs(mapping_text: str, rules_df: pd.DataFrame) -> str:
    raw_pairs = parse_matching_pairs(mapping_text)
    grouped: Dict[str, List[str]] = {}

    for ct_raw, prog_raw in raw_pairs:
        ct_norms = normalize_terms_from_phrase(ct_raw, "CT", rules_df)
        prog_norms = normalize_terms_from_phrase(prog_raw, "Programming", rules_df)

        for ct_norm in ct_norms:
            for prog_norm in prog_norms:
                grouped.setdefault(ct_norm, [])
                if prog_norm not in grouped[ct_norm]:
                    grouped[ct_norm].append(prog_norm)

    if not grouped:
        return "No valid pairs"

    return "[" + "; ".join(f"{ct}:({', '.join(progs)})" for ct, progs in grouped.items()) + "]"


def explode_valid_pairs(valid_pairs_text: str) -> List[Tuple[str, str]]:
    return parse_matching_pairs(valid_pairs_text)


# ============================================================
# OpenAI call and costs
# ============================================================
def extract_usage(response: Any) -> Dict[str, int]:
    usage = getattr(response, "usage", None)

    if usage is None:
        return {
            "input_tokens": 0,
            "cached_input_tokens": 0,
            "output_tokens": 0,
            "total_tokens": 0,
        }

    input_tokens = int(getattr(usage, "input_tokens", 0) or 0)
    output_tokens = int(getattr(usage, "output_tokens", 0) or 0)
    total_tokens = int(getattr(usage, "total_tokens", input_tokens + output_tokens) or (input_tokens + output_tokens))

    cached_input_tokens = 0
    details = getattr(usage, "input_tokens_details", None)
    if details is not None:
        cached_input_tokens = int(getattr(details, "cached_tokens", 0) or 0)

    return {
        "input_tokens": input_tokens,
        "cached_input_tokens": cached_input_tokens,
        "output_tokens": output_tokens,
        "total_tokens": total_tokens,
    }


def estimate_cost_usd(
    input_tokens: int,
    output_tokens: int,
    input_price_per_1m: float,
    output_price_per_1m: float,
) -> float:
    return (input_tokens / 1_000_000 * input_price_per_1m) + (
        output_tokens / 1_000_000 * output_price_per_1m
    )


def call_openai_with_usage(
    client: OpenAI,
    model: str,
    prompt: str,
    json_schema: Dict[str, Any],
    schema_name: str,
    temperature: float = 0.0,
) -> Tuple[Dict[str, Any], Dict[str, int]]:
    response = client.responses.create(
        model=model,
        input=[
            {
                "role": "system",
                "content": "You are a careful research assistant coding academic papers into a fixed schema.",
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        temperature=temperature,
        text={
            "format": {
                "type": "json_schema",
                "name": schema_name,
                "strict": True,
                "schema": json_schema,
            }
        },
    )

    return json.loads(response.output_text), extract_usage(response)


# Backward-compatible wrapper
def call_openai(
    client: OpenAI,
    model: str,
    prompt: str,
    json_schema: Dict[str, Any],
    schema_name: str,
    temperature: float = 0.0,
) -> Dict[str, Any]:
    result, _ = call_openai_with_usage(
        client=client,
        model=model,
        prompt=prompt,
        json_schema=json_schema,
        schema_name=schema_name,
        temperature=temperature,
    )
    return result


def make_usage_record(
    stage: str,
    paper_id: str,
    title: str,
    source_file: str,
    model: str,
    usage: Dict[str, int],
    input_price_per_1m: float,
    output_price_per_1m: float,
    status: str = "Success",
    ct_element: str = "",
    programming_element: str = "",
    error_message: str = "",
) -> Dict[str, Any]:
    cost = estimate_cost_usd(
        input_tokens=usage.get("input_tokens", 0),
        output_tokens=usage.get("output_tokens", 0),
        input_price_per_1m=input_price_per_1m,
        output_price_per_1m=output_price_per_1m,
    )

    return {
        "Stage": stage,
        "Paper_ID": paper_id,
        "Title": title,
        "Source_File": source_file,
        "CT_element": ct_element,
        "Programming_element": programming_element,
        "Model": model,
        "Input_Tokens": usage.get("input_tokens", 0),
        "Cached_Input_Tokens": usage.get("cached_input_tokens", 0),
        "Output_Tokens": usage.get("output_tokens", 0),
        "Total_Tokens": usage.get("total_tokens", 0),
        "Estimated_Cost_USD": cost,
        "Status": status,
        "Error_Message": error_message,
    }


def summarize_usage(usage_records: List[Dict[str, Any]]) -> pd.DataFrame:
    if not usage_records:
        return pd.DataFrame(
            [
                {"Metric": "Total input tokens", "Value": 0},
                {"Metric": "Total output tokens", "Value": 0},
                {"Metric": "Total tokens", "Value": 0},
                {"Metric": "Estimated total cost USD", "Value": 0.0},
                {"Metric": "Paper-level calls", "Value": 0},
                {"Metric": "Pair-level calls", "Value": 0},
            ]
        )

    df = pd.DataFrame(usage_records)

    return pd.DataFrame(
        [
            {"Metric": "Total input tokens", "Value": float(df["Input_Tokens"].sum())},
            {"Metric": "Total cached input tokens", "Value": float(df["Cached_Input_Tokens"].sum())},
            {"Metric": "Total output tokens", "Value": float(df["Output_Tokens"].sum())},
            {"Metric": "Total tokens", "Value": float(df["Total_Tokens"].sum())},
            {"Metric": "Estimated total cost USD", "Value": float(df["Estimated_Cost_USD"].sum())},
            {"Metric": "Paper-level calls", "Value": int((df["Stage"] == "paper_level").sum())},
            {"Metric": "Pair-level calls", "Value": int((df["Stage"] == "pair_level").sum())},
            {"Metric": "Error calls", "Value": int((df["Status"] == "Error").sum())},
        ]
    )


def save_usage_workbook(path: Path, usage_records: List[Dict[str, Any]]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    usage_df = pd.DataFrame(usage_records)
    summary_df = summarize_usage(usage_records)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        usage_df.to_excel(writer, sheet_name="Detailed_Usage", index=False)
        summary_df.to_excel(writer, sheet_name="Cost_Summary", index=False)

    return usage_df, summary_df


# ============================================================
# Input loading and checkpoints
# ============================================================
def find_existing_next_stage_file(default_filename: str = "full_screening_llm_corrected.xlsx") -> Path | None:
    candidates = [
        Path.cwd() / default_filename,
        Path.cwd() / "full_screening_llm_correctd.xlsx",
    ]

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return None


def standardize_paper_list_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {}

    for pdf_col in [
        "PDF_path",
        "PDF_Path",
        "PDF Path",
        "pdf_path",
        "pdf path",
        "Path",
    ]:
        if pdf_col in df.columns and "PDF_Address" not in df.columns:
            rename_map[pdf_col] = "PDF_Address"
            break
    if "Paper" in df.columns and "Title" not in df.columns:
        rename_map["Paper"] = "Title"

    df = df.rename(columns=rename_map)

    required_cols = {"Paper_ID", "PDF_Address"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(
            "Input file must contain Paper_ID and either PDF_path or PDF_Address. "
            f"Missing: {missing}"
        )

    for col, default in {
        "Title": "",
        "summary": "",
        "Next_stage_label": "",
        "Matching_Evidence_Location": "",
        "q1": "",
        "q2": "",
        "q3": "",
        "q4": "",
        "decision": "",
    }.items():
        if col not in df.columns:
            df[col] = default

    df = df.dropna(subset=["Paper_ID", "PDF_Address"]).copy()

    for col in df.columns:
        df[col] = df[col].astype(str).replace("nan", "").str.strip()

    return df


def load_paper_list(
    input_mode: str,
    uploaded_file,
    default_next_stage_path: Path | None
) -> tuple[pd.DataFrame, pd.DataFrame]:

    def normalize_paper_id(value) -> str:
        text = str(value).strip()
        if text.endswith(".0"):
            text = text[:-2]
        return text

    if input_mode == AUTO_INPUT_MODE:
        if uploaded_file is not None:
            source = uploaded_file
        elif default_next_stage_path is not None and default_next_stage_path.exists():
            source = default_next_stage_path
        else:
            raise ValueError(
                "Could not find the Stage 4 input workbook automatically. "
                "Please upload the previous screening workbook."
            )

        xls = pd.ExcelFile(source)

        preferred_sheets = [
            "Next_stage_included",
            "Next_Stage_Fulltext",
            "Included",
            "fulltext_screening_input",
        ]

        selected_sheet = next(
            (sheet for sheet in preferred_sheets if sheet in xls.sheet_names),
            xls.sheet_names[0],
        )

        df = pd.read_excel(source, sheet_name=selected_sheet)
        df = standardize_paper_list_columns(df)

        excluded_before_stage4 = pd.DataFrame()

        if "Included" in xls.sheet_names:
            included_df_raw = pd.read_excel(source, sheet_name="Included")
            included_df_raw.columns = [str(c).strip() for c in included_df_raw.columns]

            if "Candidate_Matching_CT_Programming_Elements" in included_df_raw.columns:
                included_df = standardize_paper_list_columns(included_df_raw)

                candidate_col = (
                    included_df["Candidate_Matching_CT_Programming_Elements"]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )

                excluded_ids = set(
                    included_df.loc[
                        candidate_col.eq("not specified"),
                        "Paper_ID"
                    ]
                    .apply(normalize_paper_id)
                )

                df["_Paper_ID_norm"] = df["Paper_ID"].apply(normalize_paper_id)

                excluded_before_stage4 = df[
                    df["_Paper_ID_norm"].isin(excluded_ids)
                ].copy()

                df = df[
                    ~df["_Paper_ID_norm"].isin(excluded_ids)
                ].copy()

                df = df.drop(columns=["_Paper_ID_norm"], errors="ignore")
                excluded_before_stage4 = excluded_before_stage4.drop(
                    columns=["_Paper_ID_norm"],
                    errors="ignore",
                )

    else:
        df = pd.read_excel(uploaded_file)
        df = standardize_paper_list_columns(df)
        excluded_before_stage4 = pd.DataFrame()

    return df, excluded_before_stage4

def load_checkpoint(path: Path) -> pd.DataFrame:
    return pd.read_excel(path) if path.exists() else pd.DataFrame()


def make_pair_key(paper_id: str, ct_element: str, programming_element: str) -> str:
    return f"{str(paper_id).strip()}||{str(ct_element).strip()}||{str(programming_element).strip()}"


def get_processed_paper_ids(checkpoint_df: pd.DataFrame) -> set:
    if checkpoint_df.empty or "Paper_ID" not in checkpoint_df.columns:
        return set()

    successful = checkpoint_df
    if "Processing_Status" in checkpoint_df.columns:
        successful = checkpoint_df[
            checkpoint_df["Processing_Status"].astype(str).str.strip() == "Success"
        ]

    return set(successful["Paper_ID"].astype(str).str.strip())


def get_processed_pair_keys(checkpoint_df: pd.DataFrame) -> set:
    required = {"Paper_ID", "CT_element", "Programming_element"}
    if checkpoint_df.empty or not required.issubset(checkpoint_df.columns):
        return set()

    successful = checkpoint_df
    if "Processing_Status" in checkpoint_df.columns:
        successful = checkpoint_df[
            checkpoint_df["Processing_Status"].astype(str).str.strip() == "Success"
        ]

    return set(
        make_pair_key(row["Paper_ID"], row["CT_element"], row["Programming_element"])
        for _, row in successful.iterrows()
    )

# ----------------------------------------------------
# Bulding a new Matrix to define the Included and Excluded Operationalization papers
# -----------------------------------------------------
def read_broad_normalization_rules(excel_file) -> dict:
    reset_excel_pointer(excel_file)

    ct_df = pd.read_excel(excel_file, sheet_name="CT_Normalization")
    reset_excel_pointer(excel_file)
    prog_df = pd.read_excel(excel_file, sheet_name="Programming_Normalization")

    for sheet_name, df in {
        "CT_Normalization": ct_df,
        "Programming_Normalization": prog_df,
    }.items():
        df.columns = [str(c).strip() for c in df.columns]
        required = {"Raw_Term", "Header_Name"}
        missing = required - set(df.columns)
        if missing:
            raise ValueError(f"Missing columns in {sheet_name}: {missing}")

    ct_df = ct_df.dropna(subset=["Raw_Term", "Header_Name"]).copy()
    prog_df = prog_df.dropna(subset=["Raw_Term", "Header_Name"]).copy()

    ct_df["Clean_Raw_Term"] = ct_df["Raw_Term"].apply(clean_term)
    prog_df["Clean_Raw_Term"] = prog_df["Raw_Term"].apply(clean_term)

    return {
        "CT": dict(zip(ct_df["Clean_Raw_Term"], ct_df["Header_Name"].astype(str).str.strip())),
        "Programming": dict(zip(prog_df["Clean_Raw_Term"], prog_df["Header_Name"].astype(str).str.strip())),
    }

def build_normalized_pair_matrix(
    papers_df: pd.DataFrame,
    broad_rules: dict,
    mapping_col: str = "Matching_CT_Programming_Elements",
) -> pd.DataFrame:
    """
    Builds a clean broad CT-Programming pairing matrix from raw LLM mappings.

    Rules:
    - The left side of each mapping must match CT_Normalization.
    - The right side of each mapping must match Programming_Normalization.
    - A pair is valid only if BOTH sides are matched.
    - If only one side is matched, nothing is marked.
    - Comma-separated programming elements generate multiple candidate pairs.
    - Matches_Normalized stores only valid normalized pairs.
    - Pairing_decision = Included if at least one valid pair exists; otherwise Excluded.
    """

    result_df = papers_df.copy()

    ct_headers = sorted(set(broad_rules["CT"].values()))
    prog_headers = sorted(set(broad_rules["Programming"].values()))

    for col in ct_headers + prog_headers:
        result_df[col] = "No"

    result_df["Matches_Normalized"] = "No valid matches"
    result_df["Pairing_decision"] = "Excluded"

    for idx, row in result_df.iterrows():
        mapping_text = row.get(mapping_col, "")
        raw_pairs = parse_matching_pairs(mapping_text)

        grouped: Dict[str, List[str]] = {}

        for ct_raw, prog_raw in raw_pairs:
            ct_clean = clean_term(ct_raw)
            prog_clean = clean_term(prog_raw)

            ct_header = broad_rules["CT"].get(ct_clean)
            prog_header = broad_rules["Programming"].get(prog_clean)

            # Valid only as a complete pair
            if not ct_header or not prog_header:
                continue

            result_df.at[idx, ct_header] = "Yes"
            result_df.at[idx, prog_header] = "Yes"

            grouped.setdefault(ct_header, [])
            if prog_header not in grouped[ct_header]:
                grouped[ct_header].append(prog_header)

        if grouped:
            matches_normalized = "[" + "; ".join(
                f"{ct}:({', '.join(progs)})"
                for ct, progs in grouped.items()
            ) + "]"

            result_df.at[idx, "Matches_Normalized"] = matches_normalized
            result_df.at[idx, "Pairing_decision"] = "Included"

    return result_df
# ------------------------------------------------
# Final files
# ------------------------------------------------
def filter_final_stage4_outputs(
    paper_df: pd.DataFrame,
    pair_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:

    paper_df = paper_df.copy()
    pair_df = pair_df.copy()

    has_broad_pair = (
        paper_df["Pairing_decision"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("included")
        if "Pairing_decision" in paper_df.columns
        else pd.Series(False, index=paper_df.index)
    )

    has_framework_pair = (
        paper_df["Valid_CT_Programming_Pairs"]
        .astype(str)
        .str.strip()
        .str.lower()
        .ne("no valid pairs")
        if "Valid_CT_Programming_Pairs" in paper_df.columns
        else pd.Series(False, index=paper_df.index)
    )

    not_brennan_resnick = (
        paper_df["Use_of_Brennan_Resnick_Framework"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("no")
        if "Use_of_Brennan_Resnick_Framework" in paper_df.columns
        else pd.Series(False, index=paper_df.index)
    )

    final_paper_df = paper_df[
        has_broad_pair | has_framework_pair
    ].copy()

    if "Paper_ID" not in final_paper_df.columns or "Paper_ID" not in pair_df.columns:
        return final_paper_df, pd.DataFrame()

    final_ids = set(
        final_paper_df["Paper_ID"]
        .astype(str)
        .str.strip()
    )

    final_pair_df = pair_df[
        pair_df["Paper_ID"]
        .astype(str)
        .str.strip()
        .isin(final_ids)
    ].copy()

    return final_paper_df, final_pair_df