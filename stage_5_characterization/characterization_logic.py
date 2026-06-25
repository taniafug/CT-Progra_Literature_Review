import json
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional

import fitz
import pandas as pd
from openai import OpenAI
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


# -----------------------------
# Utility funcitons
# -----------------------------
def reset_excel_pointer(excel_file) -> None:
    try:
        excel_file.seek(0)
    except Exception:
        pass


def safe_str(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def clean_sheet_name(name: str) -> str:
    invalid_chars = ["\\", "/", "*", "?", ":", "[", "]"]
    name = str(name)
    for ch in invalid_chars:
        name = name.replace(ch, "_")
    return name[:31]


def make_checkpoint_key(paper_id: str, pdf_path: str) -> str:
    return f"{safe_str(paper_id)}||{safe_str(pdf_path)}"


def load_excel_sheet_names(excel_file) -> List[str]:
    if isinstance(excel_file, str):
        excel = pd.ExcelFile(excel_file)
        return excel.sheet_names

    reset_excel_pointer(excel_file)
    excel = pd.ExcelFile(excel_file)
    reset_excel_pointer(excel_file)
    return excel.sheet_names


def load_input_preview(excel_file, sheet_name: str, nrows: int = 5) -> pd.DataFrame:
    if isinstance(excel_file, str):
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            nrows=nrows,
            dtype=str,
        )
    else:
        reset_excel_pointer(excel_file)
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            nrows=nrows,
            dtype=str,
        )
        reset_excel_pointer(excel_file)

    df = normalize_colnames(df)
    return df

def clean_excel_value(value):
    if value is None:
        return value
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
        value = re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", "", value)
        return value
    return value


def clean_results_for_excel(results):
    cleaned = []
    for row in results:
        cleaned.append({
            key: clean_excel_value(value)
            for key, value in row.items()
        })
    return cleaned
# -----------------------------
# Excel rules readers
# -----------------------------
def read_prompt_config(excel_file) -> Dict[str, str]:
    reset_excel_pointer(excel_file)
    df = pd.read_excel(excel_file, sheet_name="Prompt_Config")
    df = normalize_colnames(df)

    key_col = "key" if "key" in df.columns else "Field" if "Field" in df.columns else None
    value_col = "value" if "value" in df.columns else "Value" if "Value" in df.columns else None

    if key_col is None or value_col is None:
        raise ValueError("Prompt_Config must contain columns 'key' and 'value' or 'Field' and 'Value'.")

    config: Dict[str, str] = {}
    for _, row in df.iterrows():
        key = safe_str(row.get(key_col, ""))
        value = safe_str(row.get(value_col, ""))
        if key and value:
            config[key] = value

    if not config:
        raise ValueError("Prompt_Config is empty.")

    return config


def build_prompt_config_text(prompt_config: Dict[str, str]) -> str:
    preferred_order = [
        "system_role",
        "task_description",
        "schema_intro",
        "general_behavior",
        "general_behavior_1",
        "general_behavior_2",
        "not_applicable_rule",
        "uncertainty_rules",
        "uncertainty_rule_undetermined",
        "uncertainty_rule_manual_revision",
        "other_format_rule",
        "open_text_rule",
        "anti_invention_rule",
        "missing_value_rule",
        "metadata_rule",
        "paper_metadata_rule",
        "evidence_scope",
        "section_use_rules",
        "dependency_rules_intro",
    ]

    parts: List[str] = []
    used = set()
    for key in preferred_order:
        value = prompt_config.get(key, "")
        if safe_str(value):
            parts.append(safe_str(value))
            used.add(key)

    for key, value in prompt_config.items():
        if key not in used and safe_str(value):
            parts.append(safe_str(value))

    return "\n\n".join(parts)


def get_system_role(prompt_config: Dict[str, str]) -> str:
    return prompt_config.get(
        "system_role",
        "You are a careful research assistant coding academic papers into a fixed schema.",
    )


def read_coding_schema(excel_file) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    schema_df = pd.read_excel(excel_file, sheet_name="Coding_Schema_Characterization")
    schema_df = normalize_colnames(schema_df)

    required = {"Field", "Description", "Allowed values", "Definition / coding note"}
    missing = required - set(schema_df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Coding_Schema_Characterization: {missing}")

    schema_df["Field"] = schema_df["Field"].ffill()
    schema_df["Description"] = schema_df["Description"].ffill()
    schema_df = schema_df.dropna(subset=["Field"])
    return schema_df


def read_dependency_rules(excel_file) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    rules_df = pd.read_excel(excel_file, sheet_name="Dependency_Rules")
    rules_df = normalize_colnames(rules_df)

    required = {"Rule_ID", "Rule_Text"}
    missing = required - set(rules_df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Dependency_Rules: {missing}")

    return rules_df.dropna(subset=["Rule_ID", "Rule_Text"])


def build_dependency_rules_prompt(rules_df: pd.DataFrame) -> str:
    rules = []
    for _, row in rules_df.iterrows():
        rule_id = safe_str(row.get("Rule_ID", ""))
        rule_text = safe_str(row.get("Rule_Text", ""))
        if rule_id and rule_text:
            rules.append(f"- {rule_id}: {rule_text}")
    return "\n".join(rules)


def build_schema_prompt(schema_df: pd.DataFrame) -> str:
    sections = []
    for field, group in schema_df.groupby("Field", sort=False):
        description = safe_str(group["Description"].iloc[0])
        allowed_notes = []
        for _, row in group.iterrows():
            allowed = safe_str(row.get("Allowed values", ""))
            note = safe_str(row.get("Definition / coding note", ""))
            if not allowed:
                continue
            normalized_allowed = allowed.lower().strip()
            is_open_text = ("open text" in normalized_allowed or "free text" in normalized_allowed or normalized_allowed == "text" or normalized_allowed == "open")
            if is_open_text:
                if note:
                    allowed_notes.append(f"- Open text field. {note}")
                else:
                    allowed_notes.append("- Open text field. Extract directly from the paper.")
            else:
                allowed_notes.append(f"- {allowed}: {note}" if note else f"- {allowed}")   

        sections.append(
            f"FIELD: {field}\nDESCRIPTION: {description}\nALLOWED VALUES / NOTES:\n" + "\n".join(allowed_notes)
        )
    return "\n\n".join(sections)


def get_output_fields(schema_df: pd.DataFrame) -> List[str]:
    return list(dict.fromkeys(schema_df["Field"].dropna().astype(str).str.strip().tolist()))


def make_json_schema(
    schema_df: pd.DataFrame
) -> Dict[str, Any]:

    properties = {}

    for field, group in schema_df.groupby(
        "Field",
        sort=False
    ):

        allowed_values = (
            group["Allowed values"]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )

        # Remove empty values
        allowed_values = [
            value
            for value in allowed_values
            if value
        ]

        # Detect open text fields
        normalized_allowed = [
            value.lower().strip()
            for value in allowed_values
        ]

        is_open_text = any(
            (
                "open text" in value
                or "free text" in value
                or value == "text"
                or value == "open"
            )
            for value in normalized_allowed
        )

        # Open text fields should NOT become enums
        if allowed_values and not is_open_text:

            properties[field] = {
                "type": "string",
                "enum": allowed_values,
            }

        else:

            properties[field] = {
                "type": "string"
            }

    return {
        "type": "object",
        "properties": properties,
        "required": list(properties.keys()),
        "additionalProperties": False,
    }


# -----------------------------
# Section patterns
# -----------------------------
def default_section_patterns() -> pd.DataFrame:
    rows = [
        ("methods", "method", 1, "primary"),
        ("methods", "methods", 1, "primary"),
        ("methods", "methodology", 1, "primary"),
        ("methods", "research design", 1, "primary"),
        ("participants_context", "participants", 2, "primary"),
        ("participants_context", "sample", 2, "primary"),
        ("participants_context", "context", 2, "primary"),
        ("intervention_procedure", "intervention", 3, "primary"),
        ("intervention_procedure", "procedure", 3, "primary"),
        ("intervention_procedure", "learning activities", 3, "primary"),
        ("instruments_assessment", "assessment", 4, "primary"),
        ("instruments_assessment", "evaluation", 4, "primary"),
        ("instruments_assessment", "data collection", 4, "primary"),
        ("results_findings", "results", 5, "primary"),
        ("results_findings", "findings", 5, "primary"),
        ("discussion_conclusion", "discussion", 6, "optional"),
        ("discussion_conclusion", "conclusion", 6, "optional"),
        ("fallback_intro", "introduction", 99, "fallback"),
        ("fallback_intro", "background", 99, "fallback"),
        ("fallback_abstract", "abstract", 100, "fallback"),
    ]
    return pd.DataFrame(rows, columns=["Section_Group", "Heading", "Priority", "Use_Mode"])


def read_section_patterns(excel_file) -> pd.DataFrame:
    reset_excel_pointer(excel_file)
    try:
        df = pd.read_excel(excel_file, sheet_name="Section_Patterns")
    except Exception:
        return default_section_patterns()

    df = normalize_colnames(df)
    required = {"Section_Group", "Heading", "Priority", "Use_Mode"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Section_Patterns: {missing}")

    df = df.dropna(subset=["Section_Group", "Heading"])
    df["Priority"] = pd.to_numeric(df["Priority"], errors="coerce").fillna(999).astype(int)
    df["Use_Mode"] = df["Use_Mode"].astype(str).str.strip().str.lower()
    return df


def section_patterns_to_dict(patterns_df: pd.DataFrame) -> Dict[str, List[str]]:
    patterns: Dict[str, List[str]] = {}
    for _, row in patterns_df.iterrows():
        group = safe_str(row.get("Section_Group", ""))
        heading = safe_str(row.get("Heading", ""))
        if group and heading:
            patterns.setdefault(group, []).append(heading)
    return patterns


def find_section_positions(text: str, patterns: Dict[str, List[str]]) -> List[Tuple[str, int]]:
    positions = []
    for section_name, headings in patterns.items():
        for heading in headings:
            pattern = rf"(^|\n)\s*((\d+(\.\d+)*)|[IVX]+)?\.?\s*{re.escape(heading)}[:\s]*($|\n)"
            match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
            if match:
                positions.append((section_name, match.start()))
                break
    return sorted(positions, key=lambda x: x[1])


def extract_sections(text: str, patterns_df: pd.DataFrame) -> Dict[str, str]:
    patterns = section_patterns_to_dict(patterns_df)
    positions = find_section_positions(text, patterns)
    if not positions:
        return {"full_text_excerpt": text}

    sections: Dict[str, str] = {}
    for idx, (section_name, start) in enumerate(positions):
        end = positions[idx + 1][1] if idx + 1 < len(positions) else len(text)
        sections[section_name] = text[start:end].strip()
    return sections


def remove_bibliography(text: str) -> str:
    """
    Removes reference/bibliography sections from a PDF text.

    This is a technical safeguard for fallback extraction. It does not decide
    what evidence is valid; evidence priorities remain controlled by the Excel
    file through Prompt_Config and Section_Patterns.
    """
    if not text:
        return ""

    bibliography_headings = [
        r"references",
        r"reference list",
        r"bibliography",
        r"works cited",
        r"literature cited",
        r"cited references",
    ]

    cut_positions: List[int] = []

    for heading in bibliography_headings:
        pattern = rf"(^|\n)\s*((\d+(\.\d+)*)|[IVX]+)?\.?\s*{heading}\s*[:\n]"
        match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        if match:
            cut_positions.append(match.start())

    if cut_positions:
        return text[: min(cut_positions)].strip()

    return text.strip()


def build_rest_of_document_without_bibliography(
    full_text: str,
    already_selected_text: str = "",
) -> str:
    """
    Returns a bibliography-free fallback text.

    If selected sections were already extracted but too short, this returns the
    rest of the document without bibliography. It keeps the implementation
    simple and conservative: the model receives non-reference paper content,
    not the bibliography.
    """
    text_no_bibliography = remove_bibliography(full_text)

    if not text_no_bibliography:
        return ""

    # Avoid duplicating the exact selected block when possible.
    selected = safe_str(already_selected_text)
    if selected and selected in text_no_bibliography:
        text_no_bibliography = text_no_bibliography.replace(selected, " ").strip()

    return text_no_bibliography.strip()


def build_selected_text(
    full_text: str,
    patterns_df: pd.DataFrame,
    max_chars: int,
    include_optional: bool = True,
    allow_fallback: bool = True,
    weak_threshold_chars: int = 5000,
) -> Tuple[str, str, str, str]:
    """
    Builds the text sent to the LLM.

    Flow:
    1. Use Section_Patterns primary sections.
    2. Add optional sections if enabled.
    3. If selected text is weak, use Section_Patterns fallback sections.
    4. If it is still weak, use the rest of the document without bibliography.
    """
    full_text_no_bibliography = remove_bibliography(full_text)
    sections = extract_sections(full_text_no_bibliography, patterns_df)

    meta = (
        patterns_df[["Section_Group", "Priority", "Use_Mode"]]
        .drop_duplicates(subset=["Section_Group"])
        .sort_values(["Priority", "Section_Group"])
    )
    # Include beginning of paper for metadata
    metadata_excerpt = full_text_no_bibliography[:5000]

    # Include beginning of paper for metadata
    metadata_excerpt = full_text_no_bibliography[:5000]

    selected_parts: List[str] = ["\n\n===== PAPER METADATA / FIRST PAGE =====\n" + metadata_excerpt]

    selected_groups: List[str] = ["paper_metadata_first_page"]

    for _, row in meta.iterrows():
        group = safe_str(row["Section_Group"])
        use_mode = safe_str(row["Use_Mode"]).lower()

        if group not in sections or not sections[group].strip():
            continue

        if use_mode == "primary" or (use_mode == "optional" and include_optional):
            selected_parts.append(f"\n\n===== {group.upper()} =====\n{sections[group]}")
            selected_groups.append(group)

    used_fallback = "No"

    # Fallback 1: use configured fallback sections from Excel.
    if (not selected_parts or len("\n".join(selected_parts)) < weak_threshold_chars) and allow_fallback:
        fallback_meta = meta[meta["Use_Mode"].astype(str).str.lower() == "fallback"]

        for _, row in fallback_meta.iterrows():
            group = safe_str(row["Section_Group"])

            if group in sections and sections[group].strip():
                selected_parts.append(f"\n\n===== FALLBACK {group.upper()} =====\n{sections[group]}")
                selected_groups.append(group)
                used_fallback = "Yes"
                break

    selected_text = "\n".join(selected_parts).strip()

    # Fallback 2: if still weak, use rest of document without bibliography.
    if (not selected_text or len(selected_text) < weak_threshold_chars) and allow_fallback:
        rest_text = build_rest_of_document_without_bibliography(
            full_text=full_text,
            already_selected_text=selected_text,
        )

        if rest_text:
            selected_parts.append(
                "\n\n===== FALLBACK REST OF DOCUMENT WITHOUT BIBLIOGRAPHY =====\n"
                + rest_text
            )
            selected_groups.append("rest_of_document_without_bibliography")
            used_fallback = "Yes"
            selected_text = "\n".join(selected_parts).strip()

    # Last defensive fallback: only if nothing else was selected.
    if not selected_text and "full_text_excerpt" in sections:
        selected_text = sections["full_text_excerpt"].strip()
        selected_groups.append("full_text_excerpt")
        used_fallback = "Yes"

    was_truncated = "No"
    if len(selected_text) > max_chars:
        selected_text = selected_text[:max_chars] + "\n\n[TEXT TRUNCATED DUE TO LENGTH]"
        was_truncated = "Yes"

    return selected_text, was_truncated, ", ".join(selected_groups), used_fallback


# -----------------------------
# PDF reading
# -----------------------------
def extract_pdf_full_text(pdf_path: Path) -> str:
    text_parts = []
    with fitz.open(pdf_path) as doc:
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text("text")
            if text.strip():
                text_parts.append(f"\n--- Page {page_num} ---\n{text}")
    return "\n".join(text_parts).strip()


def extract_pdf_text(pdf_path: Path, max_chars: int = 120_000) -> str:
    full_text = extract_pdf_full_text(pdf_path)
    if len(full_text) > max_chars:
        full_text = full_text[:max_chars] + "\n\n[TEXT TRUNCATED DUE TO LENGTH]"
    return full_text


# -----------------------------
# Prompt construction + LLM
# -----------------------------
def build_prompt(
    prompt_config_text: str,
    dependency_rules_prompt: str,
    schema_prompt: str,
    paper_text: str,
    paper_filename: str,
    paper_id: str = "",
    paper_title: str = "",
    selected_sections: str = "",
    text_was_truncated: str = "",
    previous_summary: str = "",
) -> str:
    metadata_lines = [f"PAPER FILE NAME: {paper_filename}"]
    if paper_id:
        metadata_lines.append(f"PAPER ID: {paper_id}")
    if paper_title:
        metadata_lines.append(f"PAPER TITLE: {paper_title}")
    if selected_sections:
        metadata_lines.append(f"SELECTED SECTIONS: {selected_sections}")
    if text_was_truncated:
        metadata_lines.append(f"TEXT TRUNCATED: {text_was_truncated}")

    summary_block = f"\n\nPREVIOUS SCREENING SUMMARY:\n{previous_summary}" if previous_summary else ""

    return f"""
{prompt_config_text}

DEPENDENCY RULES:
{dependency_rules_prompt}

CODING SCHEMA:
{schema_prompt}

PAPER METADATA:
{chr(10).join(metadata_lines)}
{summary_block}

PAPER TEXT:
{paper_text}
""".strip()


def call_openai(
    client: OpenAI,
    model: str,
    prompt: str,
    json_schema: Dict[str, Any],
    system_role: str,
    temperature: float = 0.0,
) -> Tuple[Dict[str, Any], Dict[str, int]]:
    response = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": system_role},
            {"role": "user", "content": prompt},
        ],
        temperature=temperature,
        text={
            "format": {
                "type": "json_schema",
                "name": "paper_characterization",
                "strict": True,
                "schema": json_schema,
            }
        },
    )
    usage = getattr(response, "usage", None)
    input_tokens = getattr(usage, "input_tokens", 0) if usage else 0
    output_tokens = getattr(usage, "output_tokens", 0) if usage else 0
    total_tokens = getattr(usage, "total_tokens", input_tokens + output_tokens) if usage else 0
    return json.loads(response.output_text), {
        "input_tokens": int(input_tokens or 0),
        "output_tokens": int(output_tokens or 0),
        "total_tokens": int(total_tokens or 0),
    }


# -----------------------------
# Post-processing
# -----------------------------
def flag_empty_values(result: Dict[str, Any]) -> Dict[str, Any]:
    empty_fields = []
    for field, value in result.items():
        if field == "Error_Message":
            continue
        if value is None or str(value).strip() == "":
            empty_fields.append(field)
    if empty_fields:
        current_note = result.get("Note", "")
        result["Note"] = (safe_str(current_note) + f" | Manual check needed: empty fields found: {', '.join(empty_fields)}").strip(" |")
    return result


def apply_dependency_rules(result: Dict[str, Any]) -> Dict[str, Any]:
    # Preserved from your working script as deterministic safeguards.
    if result.get("Programming_Type (paradigm)") == "Unplugged (no programming)":
        result["Programming_Tool/Language"] = "Not applicable"
        result["Programming_Kind"] = "Not applicable"
    if result.get("Type_of_Contribution") not in ["Learning model application", "Mixed"]:
        result["Activities_Conducted_for_Learning"] = "Not applicable"
    if result.get("Type_of_Contribution") not in ["Evaluation", "Mixed"]:
        result["Evaluation_Process_Conducted"] = "Not applicable"
    if result.get("Disciplinary_Focus") != "Interdisciplinary":
        result["It_is_interdisciplinary"] = "Not applicable"
    return result


# -----------------------------
# Input preparation
# -----------------------------
def prepare_paper_list(
    paper_list_file,
    sheet_name: str,
    id_col: str,
    path_col: str,
    title_col: str = "",
    summary_col: str = "",
) -> pd.DataFrame:

    # ---------------------------------------
    # Support uploaded file OR string path
    # ---------------------------------------
    if not isinstance(paper_list_file, str):
        reset_excel_pointer(paper_list_file)

    df = pd.read_excel(
        paper_list_file,
        sheet_name=sheet_name,
        dtype=str,
    )

    df = normalize_colnames(df)

    # ---------------------------------------
    # Validate required columns
    # ---------------------------------------
    if id_col not in df.columns:
        raise ValueError(
            f"Paper ID column '{id_col}' not found."
        )

    if path_col not in df.columns:
        raise ValueError(
            f"PDF path column '{path_col}' not found."
        )

    # ---------------------------------------
    # Keep selected columns
    # ---------------------------------------
    keep_cols = [id_col, path_col]

    if title_col and title_col in df.columns:
        keep_cols.append(title_col)

    if summary_col and summary_col in df.columns:
        keep_cols.append(summary_col)

    out = (
        df[keep_cols]
        .copy()
        .dropna(subset=[id_col, path_col])
    )

    # ---------------------------------------
    # Standardize column names
    # ---------------------------------------
    rename_map = {
        id_col: "Paper_ID",
        path_col: "PDF_Path",
    }

    if title_col and title_col in out.columns:
        rename_map[title_col] = "Paper"

    if summary_col and summary_col in out.columns:
        rename_map[summary_col] = (
            "Previous_Summary"
        )

    out = out.rename(columns=rename_map)

    # ---------------------------------------
    # Ensure optional columns exist
    # ---------------------------------------
    if "Paper" not in out.columns:
        out["Paper"] = ""

    if "Previous_Summary" not in out.columns:
        out["Previous_Summary"] = ""

    # ---------------------------------------
    # Clean values
    # ---------------------------------------
    out["Paper_ID"] = (
        out["Paper_ID"]
        .astype(str)
        .str.strip()
    )

    out["PDF_Path"] = (
        out["PDF_Path"]
        .astype(str)
        .str.strip()
    )

    out["Paper"] = (
        out["Paper"]
        .astype(str)
        .str.strip()
    )

    out["Previous_Summary"] = (
        out["Previous_Summary"]
        .astype(str)
        .str.strip()
    )

    return out

# -----------------------------
# Outputs
# -----------------------------
def calculate_costs(results_df: pd.DataFrame, input_price_per_1m: float, output_price_per_1m: float) -> Tuple[pd.DataFrame, pd.DataFrame]:
    charges = results_df.copy()
    for col in ["input_tokens", "output_tokens", "total_tokens"]:
        if col not in charges.columns:
            charges[col] = 0
        charges[col] = pd.to_numeric(charges[col], errors="coerce").fillna(0)
    charges["input_cost"] = charges["input_tokens"] / 1_000_000 * input_price_per_1m
    charges["output_cost"] = charges["output_tokens"] / 1_000_000 * output_price_per_1m
    charges["total_cost"] = charges["input_cost"] + charges["output_cost"]
    keep_cols = ["Paper_ID", "Paper", "Source_File", "PDF_Address", "PDF_Path", "Processing_Status", "input_tokens", "output_tokens", "total_tokens", "input_cost", "output_cost", "total_cost"]
    charges = charges[[c for c in keep_cols if c in charges.columns]].copy()
    summary = pd.DataFrame([
        {"metric": "input_tokens", "value": charges["input_tokens"].sum() if "input_tokens" in charges else 0},
        {"metric": "output_tokens", "value": charges["output_tokens"].sum() if "output_tokens" in charges else 0},
        {"metric": "total_tokens", "value": charges["total_tokens"].sum() if "total_tokens" in charges else 0},
        {"metric": "input_cost", "value": charges["input_cost"].sum() if "input_cost" in charges else 0},
        {"metric": "output_cost", "value": charges["output_cost"].sum() if "output_cost" in charges else 0},
        {"metric": "total_cost", "value": charges["total_cost"].sum() if "total_cost" in charges else 0},
    ])
    return charges, summary


def manual_review_mask(results_df: pd.DataFrame) -> pd.Series:
    if results_df.empty:
        return pd.Series(dtype=bool)

    text = results_df.astype(str).agg(" ".join, axis=1).str.lower()
    return (
        text.str.contains("manual revision", na=False)
        | text.str.contains("undetermined", na=False)
        | text.str.contains("error", na=False)
        | text.str.contains("manual check needed", na=False)
    )


def make_analysis_input(success_df: pd.DataFrame) -> pd.DataFrame:
    """
    Creates the clean characterization dataset for analysis.

    This sheet is intentionally different from Success/All_results:
    it removes technical LLM/execution columns and keeps a characterization-level
    review flag derived from the characterization output itself.
    """
    if success_df.empty:
        return pd.DataFrame()

    out = success_df.copy()

    review_mask = manual_review_mask(out)
    out["Characterization_Manual_Review"] = review_mask.map({True: "Yes", False: "No"})

    technical_cols = [
        "Source_File",
        "PDF_Address",
        "PDF_Path",
        "Selected_Sections",
        "Used_Fallback_Sections",
        "Text_Was_Truncated",
        "Processing_Status",
        "Error_Message",
        "input_tokens",
        "output_tokens",
        "total_tokens",
        "input_cost",
        "output_cost",
        "total_cost",
    ]

    out = out.drop(
        columns=[c for c in technical_cols if c in out.columns],
        errors="ignore",
    )

    return out


def sample_author_validation(results_df: pd.DataFrame, sample_fraction: float, random_seed: int) -> pd.DataFrame:
    if results_df.empty or "Processing_Status" not in results_df.columns or sample_fraction <= 0:
        return pd.DataFrame()
    samples = []
    for idx, status in enumerate(["Success", "Error"]):
        subset = results_df[results_df["Processing_Status"] == status].copy()
        if subset.empty:
            continue
        n = max(1, round(len(subset) * sample_fraction))
        n = min(n, len(subset))
        sample = subset.sample(n=n, random_state=random_seed + idx).copy()
        sample.insert(0, "Sample_Group", f"{status} random sample")
        sample.insert(1, "Sample_Size_Source_N", len(subset))
        sample.insert(2, "Sample_Percentage", sample_fraction)
        samples.append(sample)
    if not samples:
        return pd.DataFrame()
    return pd.concat(samples, ignore_index=True)


def make_author_check_clean(author_sample: pd.DataFrame) -> pd.DataFrame:
    if author_sample.empty:
        return pd.DataFrame()
    base_cols = ["Paper_ID", "Paper", "Source_File", "PDF_Address", "PDF_Path", "Processing_Status", "Error_Message"]
    clean = author_sample[[c for c in base_cols if c in author_sample.columns]].copy()
    clean["Author_revision_needed"] = ""
    clean["Author_notes"] = ""
    return clean


def build_output_dataframes(results: List[Dict[str, Any]], input_price_per_1m: float = 0.0, output_price_per_1m: float = 0.0, sample_fraction: float = 0.15, random_seed: int = 42) -> Dict[str, pd.DataFrame]:
    results_df = pd.DataFrame(results)
    if results_df.empty:
        empty = pd.DataFrame()
        return {"results": empty, "summary": empty, "success": empty, "errors": empty, "manual_review": empty, "charges": empty, "charges_summary": empty, "author_sample": empty, "author_clean": empty, "analysis_input": empty}
    success_df = results_df[results_df["Processing_Status"] == "Success"].copy() if "Processing_Status" in results_df.columns else pd.DataFrame()
    errors_df = results_df[results_df["Processing_Status"] == "Error"].copy() if "Processing_Status" in results_df.columns else pd.DataFrame()
    manual_review_df = results_df[manual_review_mask(results_df)].copy()
    charges_df, charges_summary = calculate_costs(results_df, input_price_per_1m, output_price_per_1m)
    author_sample = sample_author_validation(results_df, sample_fraction, random_seed)
    author_clean = make_author_check_clean(author_sample)
    analysis_input = make_analysis_input(success_df)

    # Paper-level summary statistics

    unique_papers_processed = (
        results_df["Paper_ID"]
        .astype(str)
        .str.strip()
        .nunique()
        if "Paper_ID" in results_df.columns
        else 0
    )

    unique_papers_exported = (
        analysis_input["Paper_ID"]
        .astype(str)
        .str.strip()
        .nunique()
        if (
            not analysis_input.empty
            and "Paper_ID" in analysis_input.columns
        )
        else 0
    )

    summary_df = pd.DataFrame([
        {"category": "Successfully characterized", "count": len(success_df)},
        {"category": "Error during characterization", "count": len(errors_df)},
        {"category": "Manual review needed", "count": len(manual_review_df)},
        {"category": "Validation sample selected", "count": len(author_sample)},
        {"category": "Total included studies processed", "count": len(results_df)},
        {"category": "Unique papers characterized", "count": unique_papers_processed},
        {"category": "Papers exported to Stage 6", "count": unique_papers_exported},
    ])

    return {
        "results": results_df,
        "summary": summary_df,
        "success": success_df,
        "errors": errors_df,
        "manual_review": manual_review_df,
        "charges": charges_df,
        "charges_summary": charges_summary,
        "author_sample": author_sample,
        "author_clean": author_clean,
        "analysis_input": analysis_input,
    }

def clean_excel_value(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def clean_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.map(clean_excel_value)

def write_workbook(output_path: Path, sheets: Dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if df is None:
                df = pd.DataFrame()

            # remove illegal Excel characters
            df = clean_dataframe_for_excel(df)

            df.to_excel(
                writer,
                sheet_name=clean_sheet_name(sheet_name),
                index=False
            )


def write_single_results_workbook(output_path, results):
    results = clean_results_for_excel(results)
    # Preserves original behavior: one Excel with one row per paper, saved after every paper.
    pd.DataFrame(results).to_excel(output_path, index=False)


def write_checkpoint_workbook(
    checkpoint_path: Path,
    results: List[Dict[str, Any]],
    input_price_per_1m: float = 0.0,
    output_price_per_1m: float = 0.0,
    sample_fraction: float = 0.15,
    random_seed: int = 42,
) -> None:
    """
    Technical checkpoint for restart/recovery.
    """
    outputs = build_output_dataframes(
        results,
        input_price_per_1m,
        output_price_per_1m,
        sample_fraction,
        random_seed,
    )

    write_workbook(
        checkpoint_path,
        {
            "All_results": outputs["results"],
            "Summary": outputs["summary"],
            "LLM_charges": outputs["charges"],
            "LLM_charge_summary": outputs["charges_summary"],
        },
    )


def load_checkpoint_results(checkpoint_path: Path) -> List[Dict[str, Any]]:
    if not checkpoint_path.exists():
        return []

    checkpoint_df = pd.read_excel(checkpoint_path, sheet_name="All_results", dtype=str)
    checkpoint_df = normalize_colnames(checkpoint_df)
    return checkpoint_df.to_dict("records")


def write_output_workbooks(
    results_output_path: Path,
    validation_output_path: Path,
    costs_output_path: Path,
    results: List[Dict[str, Any]],
    input_price_per_1m: float,
    output_price_per_1m: float,
    sample_fraction: float,
    random_seed: int,
    analysis_input_output_path: Optional[Path] = None,
) -> Dict[str, pd.DataFrame]:
    """
    Writes the final characterization outputs.

    Final files:
    1. characterization_results.xlsx
       - All_results
       - Summary
       - Success
       - Errors
       - Manual_review_needed
       - Analysis_input

    2. characterization_author_validation.xlsx
       - Author_check_sample
       - Author_check_clean
       - Manual_review_needed

    3. characterization_llm_costs.xlsx
       - LLM_charges
       - LLM_charge_summary

    Backward compatibility:
    If analysis_input_output_path is provided by an older app, the clean analysis
    input is also written there. This preserves previous functionality.
    """
    outputs = build_output_dataframes(
        results,
        input_price_per_1m,
        output_price_per_1m,
        sample_fraction,
        random_seed,
    )

    write_workbook(
        results_output_path,
        {
            "All_results": outputs["results"],
            "Summary": outputs["summary"],
            "Success": outputs["success"],
            "Errors": outputs["errors"],
            "Manual_review_needed": outputs["manual_review"],
            "Analysis_input": outputs["analysis_input"],
        },
    )

    write_workbook(
        validation_output_path,
        {
            "Author_check_sample": outputs["author_sample"],
            "Author_check_clean": outputs["author_clean"],
            "Manual_review_needed": outputs["manual_review"],
        },
    )

    write_workbook(
        costs_output_path,
        {
            "LLM_charges": outputs["charges"],
            "LLM_charge_summary": outputs["charges_summary"],
        },
    )

    if analysis_input_output_path is not None:
        write_workbook(
            analysis_input_output_path,
            {"data_cleaned": outputs["analysis_input"]},
        )

    return outputs


# -----------------------------
# Main processing
# -----------------------------
def run_characterization(
    client: OpenAI,
    model: str,
    papers_df: pd.DataFrame,
    prompt_config_text: str,
    system_role: str,
    dependency_rules_prompt: str,
    schema_prompt: str,
    output_fields: List[str],
    json_schema: Dict[str, Any],
    section_patterns_df: pd.DataFrame,
    max_chars: int,
    include_optional_sections: bool,
    allow_fallback_sections: bool,
    pause_seconds: float = 0.5,
    log_callback=None,
    progress_callback=None,
    checkpoint_path: Optional[Path] = None,
    resume_from_checkpoint: bool = False,
    input_price_per_1m: float = 0.0,
    output_price_per_1m: float = 0.0,
    sample_fraction: float = 0.15,
    random_seed: int = 42,
) -> List[Dict[str, Any]]:
    results: List[Dict[str, Any]] = []
    processed_keys = set()

    if (
        resume_from_checkpoint
        and checkpoint_path is not None
        and checkpoint_path.exists()
    ):

        checkpoint_results = load_checkpoint_results(
            checkpoint_path
        )

        results = checkpoint_results

        processed_keys = {
            make_checkpoint_key(
                safe_str(row.get("Paper_ID", "")),
                safe_str(
                    row.get(
                        "PDF_Path",
                        row.get("PDF_Address", "")
                    )
                ),
            )
            for row in checkpoint_results
            if safe_str(
                row.get(
                    "Processing_Status",
                    ""
                )
            ) == "Success"
        }

        if log_callback:
            log_callback(
                f"Recovered {len(processed_keys)} successful papers from checkpoint."
            )

    total = len(papers_df)

    if log_callback:
        log_callback(
        f"Recovered {len(processed_keys)} successful papers from checkpoint."
        )

    for i, (_, row) in enumerate(papers_df.iterrows(), start=1):
        paper_id = safe_str(row.get("Paper_ID", ""))
        paper_title = safe_str(row.get("Paper", ""))
        previous_summary = safe_str(row.get("Previous_Summary", ""))
        raw_pdf_path = safe_str(row.get("PDF_Path", ""))
        pdf_path = Path(raw_pdf_path).expanduser()
        current_key = make_checkpoint_key(paper_id, str(pdf_path))

        if current_key in processed_keys:
            if log_callback:
                log_callback(f"Skipping already processed {i}/{total}: {paper_id} - {pdf_path.name}")
            if progress_callback:
                progress_callback(i / max(1, total))
            continue

        if log_callback:
            log_callback(f"Processing {i}/{total}: {pdf_path.name}")

        try:
            if not pdf_path.exists():
                raise ValueError(f"PDF not found: {pdf_path}")

            full_text = extract_pdf_full_text(pdf_path)

            if not full_text:
                raise ValueError("No extractable text found in PDF.")

            paper_text, text_was_truncated, selected_sections, used_fallback = build_selected_text(
                full_text,
                section_patterns_df,
                max_chars,
                include_optional_sections,
                allow_fallback_sections,
            )

            if not paper_text:
                raise ValueError("No selected text found in PDF.")

            prompt = build_prompt(
                prompt_config_text,
                dependency_rules_prompt,
                schema_prompt,
                paper_text,
                pdf_path.name,
                paper_id,
                paper_title,
                selected_sections,
                text_was_truncated,
                previous_summary,
            )

            result, usage = call_openai(client, model, prompt, json_schema, system_role)
            result = apply_dependency_rules(result)
            result["Paper_ID"] = paper_id
            result["Paper"] = paper_title
            result["Source_File"] = pdf_path.name
            result["PDF_Address"] = str(pdf_path)
            result["PDF_Path"] = str(pdf_path)
            result["Selected_Sections"] = selected_sections
            result["Used_Fallback_Sections"] = used_fallback
            result["Text_Was_Truncated"] = text_was_truncated
            result["Processing_Status"] = "Success"
            result["Error_Message"] = ""
            result.update(usage)
            result = flag_empty_values(result)

        except Exception as e:
            result = {field: "" for field in output_fields}
            result["Paper_ID"] = paper_id
            result["Paper"] = paper_title
            result["Source_File"] = pdf_path.name
            result["PDF_Address"] = str(pdf_path)
            result["PDF_Path"] = str(pdf_path)
            result["Selected_Sections"] = ""
            result["Used_Fallback_Sections"] = ""
            result["Text_Was_Truncated"] = ""
            result["Processing_Status"] = "Error"
            result["Error_Message"] = str(e)
            result["input_tokens"] = 0
            result["output_tokens"] = 0
            result["total_tokens"] = 0

        results.append(result)
        processed_keys.add(current_key)

        # Save only the technical checkpoint during the run.
        # Final result files are written by write_output_workbooks() after the run finishes.
        if checkpoint_path is not None:
            write_checkpoint_workbook(
                checkpoint_path=checkpoint_path,
                results=results,
                input_price_per_1m=input_price_per_1m,
                output_price_per_1m=output_price_per_1m,
                sample_fraction=sample_fraction,
                random_seed=random_seed,
            )

        if progress_callback:
            progress_callback(i / max(1, total))

        time.sleep(pause_seconds)

    return results
