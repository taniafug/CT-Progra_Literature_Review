from pathlib import Path
import pandas as pd


def safe_read_excel(path, sheet_name=0):
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()

    try:
        data = pd.read_excel(path, sheet_name=sheet_name)

        if isinstance(data, dict):
            if not data:
                return pd.DataFrame()
            return next(iter(data.values()))

        return data

    except Exception:
        return pd.DataFrame()


def count_rows(path, sheet_name=0):
    df = safe_read_excel(path, sheet_name=sheet_name)
    return len(df) if isinstance(df, pd.DataFrame) and not df.empty else 0


def find_file(folder, patterns):
    folder = Path(folder)
    for pattern in patterns:
        matches = list(folder.glob(pattern))
        if matches:
            return matches[0]
    return None


def read_prisma_summary(path, stage_name):
    path = Path(path) if path is not None else None
    if path is None or not path.exists():
        return []

    df = safe_read_excel(path, sheet_name="PRISMA_summary")
    if df.empty:
        df = safe_read_excel(path, sheet_name="PRISMA_Summary")

    if df.empty:
        return []

    df.columns = [str(c).strip() for c in df.columns]

    metric_col = None
    for col in ["PRISMA Metric", "Metric", "Category", "Stage"]:
        if col in df.columns:
            metric_col = col
            break

    count_col = None
    for col in ["Count", "count", "Value", "value"]:
        if col in df.columns:
            count_col = col
            break

    if metric_col is None or count_col is None:
        return []

    rows = []
    for _, row in df.iterrows():
        rows.append({
            "Stage": stage_name,
            "Metric": row.get(metric_col, ""),
            "Count": row.get(count_col, 0),
            "Source_File": str(path),
        })

    return rows


def count_rows_with_fallback(path, preferred_sheet=None):
    if preferred_sheet:
        count = count_rows(path, preferred_sheet)
        if count > 0:
            return count
    return count_rows(path)




def get_framework_pair_counts(path):
    df = safe_read_excel(path)
    if df.empty:
        return {
            "framework_papers": 0,
            "framework_pairs": 0,
            "unique_pair_types": 0,
        }

    df.columns = [str(c).strip() for c in df.columns]

    if "Processing_Status" in df.columns:
        df = df[
            df["Processing_Status"]
            .fillna("")
            .astype(str)
            .str.strip()
            == "Success"
        ].copy()

    if df.empty:
        return {
            "framework_papers": 0,
            "framework_pairs": 0,
            "unique_pair_types": 0,
        }

    framework_papers = (
        df["Paper_ID"].astype(str).str.strip().nunique()
        if "Paper_ID" in df.columns
        else 0
    )

    pair_cols = [
        col for col in ["Paper_ID", "CT_element", "Programming_element"]
        if col in df.columns
    ]

    if pair_cols:
        framework_pairs = len(df.drop_duplicates(subset=pair_cols))
    else:
        framework_pairs = len(df)

    if {"CT_element", "Programming_element"}.issubset(df.columns):
        unique_pair_types = len(
            df.drop_duplicates(subset=["CT_element", "Programming_element"])
        )
    else:
        unique_pair_types = 0

    return {
        "framework_papers": int(framework_papers),
        "framework_pairs": int(framework_pairs),
        "unique_pair_types": int(unique_pair_types),
    }


def summarize_cost_file(path, stage_name):
    if path is None or not Path(path).exists():
        return {
            "Stage": stage_name,
            "Input_Tokens": 0,
            "Output_Tokens": 0,
            "Total_Tokens": 0,
            "Estimated_Cost_USD": 0.0,
            "Source_File": "",
        }

    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception:
        return {
            "Stage": stage_name,
            "Input_Tokens": 0,
            "Output_Tokens": 0,
            "Total_Tokens": 0,
            "Estimated_Cost_USD": 0.0,
            "Source_File": f"Could not read: {path}",
        }

    frames = []

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
            frames.append(df)
        except Exception:
            continue

    if not frames:
        return {
            "Stage": stage_name,
            "Input_Tokens": 0,
            "Output_Tokens": 0,
            "Total_Tokens": 0,
            "Estimated_Cost_USD": 0.0,
            "Source_File": f"No readable sheets: {path}",
        }

    df = pd.concat(frames, ignore_index=True)

    def sum_possible(cols):
        for col in cols:
            if col in df.columns:
                return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
        return 0

    return {
        "Stage": stage_name,
        "Input_Tokens": sum_possible(["Input_Tokens", "input_tokens"]),
        "Output_Tokens": sum_possible(["Output_Tokens", "output_tokens"]),
        "Total_Tokens": sum_possible(["Total_Tokens", "total_tokens"]),
        "Estimated_Cost_USD": sum_possible(
            ["Estimated_Cost_USD", "Total_Cost_USD", "Cost_USD", "estimated_cost_usd", "total_cost"]
        ),
        "Source_File": str(path),
    }


def build_final_review_summary(
    project_root,
    output_dir,
    total_characterization_papers=0,
    total_mapped_papers=0,
    final_analysis_papers=0,
    missing_ids=None,
):
    if missing_ids is None:
        missing_ids = set()
        
    project_root = Path(project_root)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # ==================================================
    # Stage folders
    # ==================================================

    stage1 = project_root / "stage_1_import_filtering"
    stage2 = project_root / "stage_2_abstract_screening"
    stage3 = project_root / "stage_3_fulltext_initial_pairing"
    stage4 = project_root / "stage_4_full_pairing"
    stage5 = project_root / "stage_5_characterization"

    # ==================================================
    # Key files
    # ==================================================

    stage1_processing = stage1 / "output" / "review_pipeline_processing.xlsx"
    stage2_input = stage2 / "input" / "abstract_screening_input.xlsx"
    stage3_input = stage3 / "input" / "fulltext_screening_input.xlsx"
    stage2_results = find_file(stage2 / "output", ["*results*.xlsx", "*screened*.xlsx"])
    stage3_results = stage3 / "output" / "fulltext_results.xlsx"
    stage4_paper_results = stage4 / "output" / "paper_level_framework_results.xlsx"
    stage4_pair_results = stage4 / "output" / "pair_level_mapping_results.xlsx"
    stage5_results = stage5 / "output" / "results_characterization.xlsx"

    # costs
    stage2_costs = find_file(stage2 / "output", ["*cost*.xlsx", "*costs*.xlsx"])
    stage3_costs = find_file(stage3 / "output", ["*cost*.xlsx", "*costs*.xlsx"])
    stage4_costs = find_file(stage4 / "output", ["*cost*.xlsx", "*costs*.xlsx"])
    stage5_costs = find_file(stage5 / "output", ["*cost*.xlsx", "*costs*.xlsx"])

    # ==================================================
    # PRISMA / pipeline counts
    # ==================================================

    prisma_rows = []

    stage1_prisma = safe_read_excel(
        stage1_processing,
        sheet_name="09_prisma_counts",
    )

    if not stage1_prisma.empty:
        for _, row in stage1_prisma.iterrows():
            prisma_rows.append({
                "Stage": "Stage 1",
                "Metric": row.get("Stage", ""),
                "Count": row.get("Count", 0),
                "Source_File": str(stage1_processing),
            })

    stage2_prisma_rows = read_prisma_summary(stage2_results, "Stage 2")
    stage3_prisma_rows = read_prisma_summary(stage3_results, "Stage 3")

    if stage2_prisma_rows:
        prisma_rows.extend(stage2_prisma_rows)
    else:
        prisma_rows.append({
            "Stage": "Stage 2",
            "Metric": "Records received for abstract screening",
            "Count": count_rows_with_fallback(stage2_input, "Abstract_Screening_Input"),
            "Source_File": str(stage2_input),
        })

    if stage3_prisma_rows:
        prisma_rows.extend(stage3_prisma_rows)
    else:
        prisma_rows.append({
            "Stage": "Stage 3",
            "Metric": "Records received for full-text screening",
            "Count": count_rows(stage3_input),
            "Source_File": str(stage3_input),
        })

    prisma_rows.extend([
        {
            "Stage": "Stage 4",
            "Metric": "Paper-level pairing records",
            "Count": count_rows(stage4_paper_results),
            "Source_File": str(stage4_paper_results),
        },
        {
            "Stage": "Stage 4",
            "Metric": "Pair-level CT-programming records",
            "Count": count_rows(stage4_pair_results),
            "Source_File": str(stage4_pair_results),
        },
        {
            "Stage": "Stage 5",
            "Metric": "Characterized papers",
            "Count": count_rows(stage5_results),
            "Source_File": str(stage5_results),
        },
    ])

    framework_counts = get_framework_pair_counts(stage4_pair_results)

    prisma_rows.extend([
        {
            "Stage": "Stage 4",
            "Metric": "Framework-following papers with at least one CT-programming pair",
            "Count": framework_counts["framework_papers"],
            "Source_File": str(stage4_pair_results),
        },
        {
            "Stage": "Stage 4",
            "Metric": "Total CT-programming pair records following the framework",
            "Count": framework_counts["framework_pairs"],
            "Source_File": str(stage4_pair_results),
        },
        {
            "Stage": "Stage 4",
            "Metric": "Unique CT-programming pair types following the framework",
            "Count": framework_counts["unique_pair_types"],
            "Source_File": str(stage4_pair_results),
        },
    ])

    prisma_rows.extend([
        {
            "Stage": "Stage 6",
            "Metric": "Mapped papers in paper-level dataset",
            "Count": total_mapped_papers,
            "Source_File": "Stage 6 runtime",
        },
        {
            "Stage": "Stage 6",
            "Metric": "Characterization papers loaded",
            "Count": total_characterization_papers,
            "Source_File": "Stage 6 runtime",
        },
        {
            "Stage": "Stage 6",
            "Metric": "Papers kept for final analysis",
            "Count": final_analysis_papers,
            "Source_File": "Stage 6 runtime",
        },
        {
            "Stage": "Stage 6",
            "Metric": "Mapped Paper_IDs missing from characterization dataset",
            "Count": len(missing_ids),
            "Source_File": "Stage 6 runtime",
        },
    ])

    prisma_summary = pd.DataFrame(prisma_rows)

    # ==================================================
    # Cost summary
    # ==================================================

    cost_rows = [
        summarize_cost_file(stage2_costs, "Stage 2 – Abstract screening"),
        summarize_cost_file(stage3_costs, "Stage 3 – Full-text screening"),
        summarize_cost_file(stage4_costs, "Stage 4 – Pairing analysis"),
        summarize_cost_file(stage5_costs, "Stage 5 – Characterization"),
    ]

    cost_summary = pd.DataFrame(cost_rows)

    total_row = {
        "Stage": "TOTAL",
        "Input_Tokens": cost_summary["Input_Tokens"].sum(),
        "Output_Tokens": cost_summary["Output_Tokens"].sum(),
        "Total_Tokens": cost_summary["Total_Tokens"].sum(),
        "Estimated_Cost_USD": cost_summary["Estimated_Cost_USD"].sum(),
        "Source_File": "",
    }

    cost_summary = pd.concat(
        [cost_summary, pd.DataFrame([total_row])],
        ignore_index=True,
    )

    # ==================================================
    # File registry
    # ==================================================

    stage_files = pd.DataFrame([
        {"Stage": "Stage 1", "File": str(stage1_processing), "Exists": stage1_processing.exists()},
        {"Stage": "Stage 2", "File": str(stage2_input), "Exists": stage2_input.exists()},
        {"Stage": "Stage 2", "File": str(stage2_results), "Exists": stage2_results is not None and Path(stage2_results).exists()},
        {"Stage": "Stage 3", "File": str(stage3_input), "Exists": stage3_input.exists()},
        {"Stage": "Stage 3", "File": str(stage3_results), "Exists": stage3_results.exists()},
        {"Stage": "Stage 4", "File": str(stage4_paper_results), "Exists": stage4_paper_results.exists()},
        {"Stage": "Stage 4", "File": str(stage4_pair_results), "Exists": stage4_pair_results.exists()},
        {"Stage": "Stage 5", "File": str(stage5_results), "Exists": stage5_results.exists()},
    ])

    warnings = stage_files[stage_files["Exists"] == False].copy()

    output_path = output_dir / "final_review_summary.xlsx"

    missing_ids_df = pd.DataFrame(
        {"Missing_Paper_ID": sorted(list(missing_ids))}
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        prisma_summary.to_excel(writer, sheet_name="PRISMA_Summary", index=False)
        cost_summary.to_excel(writer, sheet_name="Cost_Summary", index=False)
        stage_files.to_excel(writer, sheet_name="Stage_Files", index=False)
        warnings.to_excel(writer, sheet_name="Warnings", index=False)
        missing_ids_df.to_excel(writer, sheet_name="Missing_IDs", index=False)

    return {
        "output_path": output_path,
        "prisma_summary": prisma_summary,
        "cost_summary": cost_summary,
        "warnings": warnings,
    }