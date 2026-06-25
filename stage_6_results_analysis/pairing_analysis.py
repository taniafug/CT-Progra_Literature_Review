from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# ==================================================
# Helpers
# ==================================================

def clean_sheet_name(name):
    invalid_chars = ["\\", "/", "*", "?", ":", "[", "]"]
    name = str(name)
    for ch in invalid_chars:
        name = name.replace(ch, "_")
    return name[:31]


def safe_filename(text):
    return (
        str(text)
        .lower()
        .replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace("(", "")
        .replace(")", "")
        .replace(":", "")
        .replace("–", "-")
        .replace("—", "-")
    )


def save_figure(fig, figures_dir, filename):
    figures_dir = Path(figures_dir)
    figures_dir.mkdir(parents=True, exist_ok=True)

    png_path = figures_dir / f"{filename}.png"
    pdf_path = figures_dir / f"{filename}.pdf"

    fig.savefig(png_path, dpi=300, bbox_inches="tight")
    fig.savefig(pdf_path, bbox_inches="tight")

    plt.close(fig)

    return {
        "png": png_path,
        "pdf": pdf_path,
    }

def clean_pair_data(pair_data):
    excluded_values = {
        "",
        "nan",
        "None",
        "Undetermined",
        "Not applicable",
        "Manual revision",
    }

    data = pair_data.copy()

    data["CT_element"] = data["CT_element"].astype(str).str.strip()
    data["Programming_element"] = data["Programming_element"].astype(str).str.strip()

    data = data[
        ~data["CT_element"].isin(excluded_values)
    ].copy()

    data = data[
        ~data["Programming_element"].isin(excluded_values)
    ].copy()

    return data

# ==================================================
# Summary builders
# ==================================================

def build_mapped_ct_summary(paper_data, total_papers):
    ct_name_map = {
        "Abstraction": "Abstraction",
        "AlgorithmicThinking": "Algorithmic Thinking",
        "PatternRecognition": "Pattern Recognition",
        "Decomposition": "Decomposition",
        "Generalization": "Generalization",
        "Evaluation": "Evaluation",
        "Debugging": "Debugging",
        "Critical Thinking": "Critical Thinking",
        "Cooperativity": "Cooperativity",
        "Creativity": "Creativity",
    }

    ct_columns = [
        col for col in paper_data.columns
        if str(col).startswith("CT_element_")
    ]

    rows = []

    for col in ct_columns:
        raw_name = col.replace("CT_element_", "").replace("_", " ")
        ct_name = ct_name_map.get(raw_name, raw_name)

        frequency = int((paper_data[col].astype(str).str.strip() == "Yes").sum())

        rows.append({
            "CT element": ct_name,
            f"Frequency in the {total_papers} analysed papers": frequency,
        })

    summary = pd.DataFrame(rows)

    if not summary.empty:
        summary = summary.sort_values(
            by=f"Frequency in the {total_papers} analysed papers",
            ascending=False,
        )

    return summary


def build_mapped_programming_summary(paper_data, grouping, total_papers):
    rows = []

    for element, group in grouping.groupby("Programming_Element", sort=False):

        columns = (
            group["Column_Name"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        existing_columns = [
            col for col in columns
            if col in paper_data.columns
        ]

        if existing_columns:
            used_in_paper = paper_data[existing_columns].astype(str).eq("Yes").any(axis=1)
            frequency = int(used_in_paper.sum())

            total_instances = 0
            instance_counts = []

            for _, item in group.iterrows():
                column_name = str(item["Column_Name"]).strip()
                instance_name = str(item["Possible_Instance"]).strip()

                if column_name in paper_data.columns:
                    count = int((paper_data[column_name].astype(str).str.strip() == "Yes").sum())
                else:
                    count = 0

                total_instances += count

                if count > 0:
                    instance_counts.append(f"{instance_name}: {count}")

            notes = "; ".join(instance_counts)

        else:
            frequency = 0
            total_instances = 0
            notes = "No matching columns"

        category = group["Category"].iloc[0] if "Category" in grouping.columns else ""

        rows.append({
            "Category": category,
            "Programming element": element,
            "Instances reported in studies": total_instances,
            f"Frequency in the {total_papers} analysed papers": frequency,
            "Notes": notes,
        })

    return pd.DataFrame(rows)


def build_framework_ct_summary(framework_pair_data, total_papers):
    if framework_pair_data.empty:
        return pd.DataFrame(columns=[
            "CT element",
            f"Frequency in the {total_papers} analysed papers",
        ])

    return (
        framework_pair_data
        .drop_duplicates(subset=["Paper_ID", "CT_element"])
        .groupby("CT_element")["Paper_ID"]
        .nunique()
        .reset_index(name=f"Frequency in the {total_papers} analysed papers")
        .sort_values(
            f"Frequency in the {total_papers} analysed papers",
            ascending=False,
        )
        .rename(columns={"CT_element": "CT element"})
    )


def build_framework_programming_summary_from_pairs(framework_pair_data, grouping, total_papers):
    freq_col = f"Frequency in the {total_papers} analysed papers"

    grouping_order = (
        grouping[["Programming_Element", "Category"]]
        .dropna(subset=["Programming_Element"])
        .drop_duplicates(subset=["Programming_Element"])
        .reset_index(drop=True)
    )

    if framework_pair_data.empty:
        summary = grouping_order.copy()
        summary[freq_col] = 0
    else:
        pair_summary = (
            framework_pair_data
            .drop_duplicates(subset=["Paper_ID", "Programming_element"])
            .groupby("Programming_element")["Paper_ID"]
            .nunique()
            .reset_index(name=freq_col)
        )

        summary = grouping_order.merge(
            pair_summary,
            left_on="Programming_Element",
            right_on="Programming_element",
            how="left",
        )

        summary[freq_col] = summary[freq_col].fillna(0).astype(int)

    summary["Instances reported in studies"] = summary[freq_col]
    summary["Notes"] = ""

    return summary[
        [
            "Category",
            "Programming_Element",
            "Instances reported in studies",
            freq_col,
            "Notes",
        ]
    ].rename(columns={
        "Programming_Element": "Programming element",
    })


def build_framework_extra_summaries(framework_pair_data, grouping, total_papers):
    freq_col = f"Frequency in the {total_papers} analysed papers"

    if "Category" in grouping.columns:
        category_lookup = (
            grouping[["Programming_Element", "Category"]]
            .dropna(subset=["Programming_Element"])
            .drop_duplicates(subset=["Programming_Element"])
            .set_index("Programming_Element")["Category"]
            .to_dict()
        )
    else:
        category_lookup = {}

    if framework_pair_data.empty:
        pair_summary = pd.DataFrame(columns=[
            "CT element",
            "Category",
            "Programming element",
            freq_col,
        ])
        ct_per_paper = pd.DataFrame(columns=["Paper_ID", "Number_of_CT_elements"])
        prog_per_paper = pd.DataFrame(columns=["Paper_ID", "Number_of_programming_elements"])
        return pair_summary, ct_per_paper, prog_per_paper

    pair_summary = (
        framework_pair_data
        .groupby(["CT_element", "Programming_element"])["Paper_ID"]
        .nunique()
        .reset_index(name=freq_col)
        .sort_values(["CT_element", freq_col], ascending=[True, False])
    )

    pair_summary["Category"] = (
        pair_summary["Programming_element"]
        .map(category_lookup)
        .fillna("Not classified")
    )

    pair_summary = pair_summary[
        [
            "CT_element",
            "Category",
            "Programming_element",
            freq_col,
        ]
    ].rename(columns={
        "CT_element": "CT element",
        "Programming_element": "Programming element",
    })

    ct_per_paper = (
        framework_pair_data
        .groupby("Paper_ID")["CT_element"]
        .nunique()
        .reset_index(name="Number_of_CT_elements")
    )

    prog_per_paper = (
        framework_pair_data
        .groupby("Paper_ID")["Programming_element"]
        .nunique()
        .reset_index(name="Number_of_programming_elements")
    )

    return pair_summary, ct_per_paper, prog_per_paper


# ==================================================
# Matrix and heatmap data
# ==================================================

def build_ct_programming_matrix(
    framework_pair_data,
    grouping,
    exclude_algorithmic=False,
    exclude_generalization=False,
):
    ct_order = [
        "Abstraction",
        "Algorithmic Thinking",
        "Decomposition",
        "Pattern Recognition",
        "Evaluation",
        "Generalization",
    ]

    if exclude_algorithmic:
        ct_order = [ct for ct in ct_order if ct != "Algorithmic Thinking"]

    if exclude_generalization:
        ct_order = [ct for ct in ct_order if ct != "Generalization"]

    pair_data = framework_pair_data.copy()

    if exclude_algorithmic:
        pair_data = pair_data[pair_data["CT_element"] != "Algorithmic Thinking"].copy()

    if exclude_generalization:
        pair_data = pair_data[pair_data["CT_element"] != "Generalization"].copy()

    matrix_base = (
        grouping[["Category", "Programming_Element"]]
        .drop_duplicates()
        .copy()
        .rename(columns={"Programming_Element": "Programming element"})
    )

    for ct in ct_order:
        values = []

        for _, row in matrix_base.iterrows():
            programming_element = row["Programming element"]

            cell_data = pair_data[
                (pair_data["CT_element"] == ct) &
                (pair_data["Programming_element"] == programming_element)
            ]

            paper_ids = sorted(cell_data["Paper_ID"].astype(str).unique())
            count = len(paper_ids)

            values.append("" if count == 0 else f"{count} papers: {', '.join(paper_ids)}")

        matrix_base[ct] = values

    return matrix_base


def build_ct_programming_heatmap_data(
    framework_pair_data,
    grouping,
    exclude_algorithmic=False,
    exclude_generalization=False,
):
    ct_order = [
        "Abstraction",
        "Algorithmic Thinking",
        "Decomposition",
        "Pattern Recognition",
        "Evaluation",
        "Generalization",
    ]

    if exclude_algorithmic:
        ct_order = [
            ct for ct in ct_order
            if ct != "Algorithmic Thinking"
        ]

    if exclude_generalization:
        ct_order = [
            ct for ct in ct_order
            if ct != "Generalization"
        ]

    programming_order = (
        grouping["Programming_Element"]
        .dropna()
        .astype(str)
        .drop_duplicates()
        .tolist()
    )

    # DEFINE FIRST
    pair_data = framework_pair_data.copy()

    if exclude_algorithmic:
        pair_data = pair_data[
            pair_data["CT_element"] != "Algorithmic Thinking"
        ].copy()

    if exclude_generalization:
        pair_data = pair_data[
            pair_data["CT_element"] != "Generalization"
        ].copy()

    if pair_data.empty:
        return (
            pd.DataFrame(
                index=programming_order,
                columns=ct_order
            )
            .fillna(0)
            .astype(int)
        )

    heatmap_table = (
        pair_data
        .drop_duplicates(
            subset=[
                "Paper_ID",
                "CT_element",
                "Programming_element",
            ]
        )
        .groupby(
            ["Programming_element", "CT_element"]
        )["Paper_ID"]
        .nunique()
        .reset_index(name="Count")
        .pivot(
            index="Programming_element",
            columns="CT_element",
            values="Count"
        )
        .fillna(0)
        .astype(int)
    )

    heatmap_table = heatmap_table.reindex(
        index=[
            p for p in programming_order
            if p in heatmap_table.index
        ],
        columns=[
            ct for ct in ct_order
            if ct in heatmap_table.columns
        ],
        fill_value=0,
    )

    return heatmap_table


# ==================================================
# Figures
# ==================================================

def frequency_bar(df, label_col, value_col, title, xlabel):
    if df.empty or label_col not in df.columns or value_col not in df.columns:
        return None

    plot_df = df.sort_values(value_col, ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(plot_df[label_col].astype(str), plot_df[value_col], color="#1192b5")

    for container in ax.containers:
        ax.bar_label(container, label_type="edge", padding=4)

    max_value = plot_df[value_col].max()
    ax.set_ylim(0, max_value + max(1, max_value * 0.18))

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel("Number of papers")

    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()

    return fig


def heatmap_figure(heatmap_data, title):
    if heatmap_data.empty:
        return None

    fig_height = max(6, len(heatmap_data.index) * 0.45)
    fig_width = max(9, len(heatmap_data.columns) * 1.6)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))

    sns.heatmap(
        heatmap_data,
        annot=True,
        fmt="d",
        cmap="YlGnBu",
        linewidths=0.5,
        linecolor="white",
        cbar_kws={"label": "Number of papers"},
        ax=ax,
    )

    ax.set_title(title)
    ax.set_xlabel("CT element")
    ax.set_ylabel("Programming element")

    plt.xticks(rotation=45, ha="right")
    plt.yticks(rotation=0)
    plt.tight_layout()

    return fig


def network_plot_for_data(pair_data, title):
    if pair_data.empty:
        return None

    network_data = (
        pair_data
        .groupby(["CT_element", "Programming_element"])["Paper_ID"]
        .nunique()
        .reset_index(name="Weight")
    )

    if network_data.empty:
        return None

    ct_nodes = sorted(network_data["CT_element"].unique())
    prog_nodes = sorted(network_data["Programming_element"].unique())

    max_len = max(len(ct_nodes), len(prog_nodes))
    fig, ax = plt.subplots(figsize=(13, max(6, max_len * 0.45)))

    ct_positions = {
        node: (0, i / max(1, len(ct_nodes) - 1))
        for i, node in enumerate(ct_nodes)
    }

    prog_positions = {
        node: (1, i / max(1, len(prog_nodes) - 1))
        for i, node in enumerate(prog_nodes)
    }

    max_weight = network_data["Weight"].max()

    for _, row in network_data.iterrows():
        ct = row["CT_element"]
        prog = row["Programming_element"]
        weight = row["Weight"]

        x1, y1 = ct_positions[ct]
        x2, y2 = prog_positions[prog]

        ax.plot(
            [x1, x2],
            [y1, y2],
            linewidth=1 + (weight / max_weight) * 5,
            alpha=0.45,
            color="#1192b5",
        )

    for node, (x, y) in ct_positions.items():
        ax.scatter(x, y, s=120, color="#1192b5")
        ax.text(x - 0.03, y, node, ha="right", va="center", fontsize=10)

    for node, (x, y) in prog_positions.items():
        ax.scatter(x, y, s=120, color="#f28e2b")
        ax.text(x + 0.03, y, node, ha="left", va="center", fontsize=10)

    ax.set_xlim(-0.45, 1.45)
    ax.set_ylim(-0.08, 1.08)
    ax.axis("off")
    ax.set_title(title)

    plt.tight_layout()

    return fig


def network_plot_filtered(pair_data, title, exclude_algorithmic=False, ct_filter=None):
    plot_data = pair_data.copy()

    if exclude_algorithmic:
        plot_data = plot_data[
            plot_data["CT_element"] != "Algorithmic Thinking"
        ].copy()

    if ct_filter is not None:
        plot_data = plot_data[
            plot_data["CT_element"] == ct_filter
        ].copy()

    return network_plot_for_data(plot_data, title)


def grouped_frequency_bar(mapped_df, framework_df, label_col, mapped_value_col, framework_value_col, title):
    if mapped_df.empty or framework_df.empty:
        return None

    comparison = mapped_df[[label_col, mapped_value_col]].merge(
        framework_df[[label_col, framework_value_col]],
        on=label_col,
        how="outer",
    ).fillna(0)

    comparison[mapped_value_col] = comparison[mapped_value_col].astype(int)
    comparison[framework_value_col] = comparison[framework_value_col].astype(int)

    comparison = comparison.sort_values(mapped_value_col, ascending=False)

    comparison = comparison[
    ~comparison[label_col].astype(str).str.strip().isin([
        "",
        "nan",
        "None",
        "Undetermined",
        "Not applicable",
        "Manual revision",
    ])
    ].copy()

    x = range(len(comparison))
    width = 0.38

    fig, ax = plt.subplots(figsize=(12, 6))

    bars1 = ax.bar(
        [i - width / 2 for i in x],
        comparison[mapped_value_col],
        width,
        label="Mapped papers",
        color="#1192b5",
    )

    bars2 = ax.bar(
        [i + width / 2 for i in x],
        comparison[framework_value_col],
        width,
        label="Framework-following papers",
        color="#f28e2b",
    )

    ax.bar_label(bars1, padding=3)
    ax.bar_label(bars2, padding=3)

    max_value = max(
        comparison[mapped_value_col].max(),
        comparison[framework_value_col].max(),
    )

    ax.set_ylim(0, max_value + max(1, max_value * 0.18))
    ax.set_title(title)
    ax.set_xlabel(label_col)
    ax.set_ylabel("Number of papers")
    ax.set_xticks(list(x))
    ax.set_xticklabels(comparison[label_col].astype(str), rotation=45, ha="right")
    ax.legend()

    plt.tight_layout()

    return fig

def build_pair_activity_heatmap_data(framework_pair_data, top_n=20):
    required_cols = {"CT_element", "Programming_element", "Activity_Category", "Paper_ID"}

    if framework_pair_data.empty or not required_cols.issubset(framework_pair_data.columns):
        return pd.DataFrame()

    excluded_values = {
        "",
        "nan",
        "None",
        "Undetermined",
        "Not applicable",
        "Manual revision",
    }

    activity_order = [
        "Data representation",
        "Control-flow design",
        "Modular problem solving",
        "Pattern recognition / repetition",
        "Debugging and testing",
        "Algorithm construction",
        "Simulation / modeling",
        "Game or interactive logic",
        "Robotics / physical interaction",
        "Data manipulation",
        "Collaborative programming",
        "Assessment task",
        "Reflection / metacognitive activity",
        "Other",
    ]

    data = framework_pair_data.copy()

    data["Activity_Category"] = (
        data["Activity_Category"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    data = data[
        ~data["Activity_Category"].isin(excluded_values)
    ].copy()

    if data.empty:
        return pd.DataFrame()

    data["Pair"] = data["CT_element"] + " × " + data["Programming_element"]

    top_pairs = (
        data.drop_duplicates(subset=["Paper_ID", "CT_element", "Programming_element"])
        .groupby("Pair")["Paper_ID"]
        .nunique()
        .sort_values(ascending=False)
        .head(top_n)
        .index
    )

    data = data[data["Pair"].isin(top_pairs)].copy()

    heatmap_data = (
        data.drop_duplicates(subset=["Paper_ID", "Pair", "Activity_Category"])
        .groupby(["Pair", "Activity_Category"])["Paper_ID"]
        .nunique()
        .reset_index(name="Count")
        .pivot(index="Pair", columns="Activity_Category", values="Count")
        .fillna(0)
        .astype(int)
    )

    heatmap_data = heatmap_data.reindex(
        index=top_pairs,
        columns=activity_order,
        fill_value=0,
    )

    heatmap_data = heatmap_data.loc[
        :,
        (heatmap_data.sum(axis=0) > 0)
    ]

    return heatmap_data

# ==================================================
# Main reusable function
# ==================================================

def run_pairing_analysis(
    paper_level_file,
    pairing_file,
    grouping_file,
    output_dir,
    analysis_mode="Both: mapped papers and framework-following papers",
    output_name="programming_elements_summary.xlsx",
):
    """
    Run CT–Programming pairing analysis without a separate Streamlit button.

    Parameters
    ----------
    paper_level_file : str or Path
        Stage 4 paper-level framework results file.
    pairing_file : str or Path
        Stage 4 pair-level mapping results file.
    grouping_file : str or Path
        Analysis_details.xlsx file containing the sheet 'Programming_Grouping'.
    output_dir : str or Path
        Stage 6 output directory.
    analysis_mode : str
        One of:
        - 'Mapped papers only'
        - 'Framework-following papers only'
        - 'Both: mapped papers and framework-following papers'
    output_name : str
        Name of the Excel output file.

    Returns
    -------
    dict
        Summary paths and counts.
    """

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    figures_dir = output_dir / "pairing_figures"
    figures_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / output_name

    # ---------- Load data ----------

    paper_data = pd.read_excel(paper_level_file)
    grouping = pd.read_excel(grouping_file, sheet_name="Programming_Grouping")

    paper_data.columns = paper_data.columns.str.strip()
    grouping.columns = grouping.columns.str.strip()

    required_paper_cols = ["Paper_ID", "Type_of_Mapping"]
    required_grouping_cols = ["Programming_Element", "Column_Name", "Possible_Instance"]

    missing_paper = [
        col for col in required_paper_cols
        if col not in paper_data.columns
    ]

    missing_grouping = [
        col for col in required_grouping_cols
        if col not in grouping.columns
    ]

    if missing_paper:
        raise ValueError(f"Missing columns in paper-level dataset: {missing_paper}")

    if missing_grouping:
        raise ValueError(f"Missing columns in grouping file: {missing_grouping}")

    paper_data["Paper_ID"] = paper_data["Paper_ID"].astype(str).str.strip()

    # ---------- Mapped papers ----------

    total_loaded = paper_data["Paper_ID"].nunique()

    mapped_paper_data = paper_data[
        paper_data["Type_of_Mapping"]
        .fillna("")
        .astype(str)
        .str.strip() != "No Mapping"
    ].copy()

    total_mapped_papers = mapped_paper_data["Paper_ID"].nunique()

    if total_mapped_papers == 0:
        raise ValueError("No mapped papers left after excluding 'No Mapping'.")

    outputs = {}
    generated_figures = []

    # ---------- Framework pair data ----------

    needs_framework = analysis_mode in [
        "Framework-following papers only",
        "Both: mapped papers and framework-following papers",
    ]

    framework_pair_data = pd.DataFrame()
    framework_paper_data = pd.DataFrame()
    prada_like_count = 0
    prada_like_papers = pd.DataFrame()
    prada_like_details = pd.DataFrame()

    if needs_framework:
        if pairing_file is None:
            raise ValueError("pairing_file is required for framework-following analysis.")

        pairing_data = pd.read_excel(pairing_file)
        pairing_data.columns = pairing_data.columns.str.strip()

        required_pair_cols = [
            "Paper_ID",
            "CT_element",
            "Programming_element",
        ]

        missing_pair_cols = [
            col for col in required_pair_cols
            if col not in pairing_data.columns
        ]

        if missing_pair_cols:
            raise ValueError(f"Missing columns in pairing dataset: {missing_pair_cols}")

        pairing_data["Paper_ID"] = pairing_data["Paper_ID"].astype(str).str.strip()
        pairing_data["CT_element"] = pairing_data["CT_element"].astype(str).str.strip()
        pairing_data["Programming_element"] = pairing_data["Programming_element"].astype(str).str.strip()

        if "Processing_Status" in pairing_data.columns:
            pairing_data = pairing_data[
                pairing_data["Processing_Status"]
                .fillna("")
                .astype(str)
                .str.strip() == "Success"
            ].copy()
        
        pairing_data = clean_pair_data(pairing_data)

        framework_pair_data = pairing_data.drop_duplicates(
            subset=["Paper_ID", "CT_element", "Programming_element"]
        )

        framework_paper_ids = framework_pair_data["Paper_ID"].unique()

        framework_paper_data = mapped_paper_data[
            mapped_paper_data["Paper_ID"].isin(framework_paper_ids)
        ].copy()

        framework_pair_data = framework_pair_data[
            framework_pair_data["Paper_ID"].isin(framework_paper_data["Paper_ID"])
        ].copy()

        total_framework_papers = framework_paper_data["Paper_ID"].nunique()

        # ---------- PRADA-like papers ----------

        prada_elements = {
            "Abstraction",
            "Algorithmic Thinking",
            "Pattern Recognition",
            "Decomposition",
        }

        ct_sets_per_paper = (
            framework_pair_data
            .groupby("Paper_ID")["CT_element"]
            .apply(lambda x: set(x.dropna().astype(str).str.strip()))
            .reset_index(name="CT_elements_present")
        )

        if not ct_sets_per_paper.empty:
            ct_sets_per_paper["Is_PRADA_like"] = ct_sets_per_paper["CT_elements_present"].apply(
                lambda ct_set: prada_elements.issubset(ct_set)
            )

            prada_like_papers = ct_sets_per_paper[
                ct_sets_per_paper["Is_PRADA_like"]
            ].copy()

            prada_like_papers["CT_elements_present"] = (
                prada_like_papers["CT_elements_present"]
                .apply(lambda x: ", ".join(sorted(x)))
            )

            prada_like_count = prada_like_papers["Paper_ID"].nunique()

            prada_like_details = framework_pair_data[
                framework_pair_data["Paper_ID"].isin(prada_like_papers["Paper_ID"])
            ].copy()

    # ---------- Mapped outputs ----------

    if analysis_mode in [
        "Mapped papers only",
        "Both: mapped papers and framework-following papers",
    ]:
        mapped_total = mapped_paper_data["Paper_ID"].nunique()
        mapped_freq_col = f"Frequency in the {mapped_total} analysed papers"

        mapped_ct_summary = build_mapped_ct_summary(
            mapped_paper_data,
            mapped_total,
        )

        mapped_programming_summary = build_mapped_programming_summary(
            mapped_paper_data,
            grouping,
            mapped_total,
        )

        outputs["mapped"] = {
            "total": mapped_total,
            "freq_col": mapped_freq_col,
            "paper_data": mapped_paper_data,
            "ct_summary": mapped_ct_summary,
            "programming_summary": mapped_programming_summary,
        }

    # ---------- Framework outputs ----------

    if needs_framework:
        framework_total = framework_paper_data["Paper_ID"].nunique()
        framework_freq_col = f"Frequency in the {framework_total} analysed papers"

        framework_ct_summary = build_framework_ct_summary(
            framework_pair_data,
            framework_total,
        )

        framework_programming_summary = build_mapped_programming_summary(
            framework_paper_data,
            grouping,
            framework_total,
        )

        framework_programming_summary_from_pairs = build_framework_programming_summary_from_pairs(
            framework_pair_data,
            grouping,
            framework_total,
        )

        framework_pair_summary, ct_per_paper_summary, prog_per_paper_summary = (
            build_framework_extra_summaries(
                framework_pair_data,
                grouping,
                framework_total,
            )
        )

        framework_matrix = build_ct_programming_matrix(
            framework_pair_data,
            grouping,
            exclude_algorithmic=False,
        )

        framework_matrix_no_algorithmic = build_ct_programming_matrix(
            framework_pair_data,
            grouping,
            exclude_algorithmic=True,
        )

        framework_heatmap_data = build_ct_programming_heatmap_data(
            framework_pair_data,
            grouping,
            exclude_algorithmic=False,
        )

        framework_heatmap_no_algorithmic_data = build_ct_programming_heatmap_data(
            framework_pair_data,
            grouping,
            exclude_algorithmic=True,
        )

        framework_matrix_no_alg_no_gen = build_ct_programming_matrix(
            framework_pair_data,
            grouping,
            exclude_algorithmic=True,
            exclude_generalization=True,
        )

        framework_heatmap_no_alg_no_gen_data = build_ct_programming_heatmap_data(
            framework_pair_data,
            grouping,
            exclude_algorithmic=True,
            exclude_generalization=True,
        )

        framework_pair_data_no_alg_no_gen = framework_pair_data[
            ~framework_pair_data["CT_element"].isin([
                "Algorithmic Thinking",
                "Generalization",
            ])
        ].copy()

        pair_activity_heatmap_data = build_pair_activity_heatmap_data(
            framework_pair_data,
            top_n=10,
        )


        framework_pair_data_no_algorithmic = framework_pair_data[
            framework_pair_data["CT_element"] != "Algorithmic Thinking"
        ].copy()

        framework_pair_data_no_generalization = framework_pair_data[
            framework_pair_data["CT_element"] != "Generalization"
        ].copy()

        framework_matrix_no_generalization = build_ct_programming_matrix(
            framework_pair_data_no_generalization,
            grouping,
            exclude_algorithmic=False,
        )

        framework_heatmap_no_generalization_data = build_ct_programming_heatmap_data(
            framework_pair_data_no_generalization,
            grouping,
            exclude_algorithmic=False,
        )

        ct_per_paper_no_generalization = (
            framework_pair_data_no_generalization
            .groupby("Paper_ID")["CT_element"]
            .nunique()
            .reset_index(name="Number_of_CT_elements_without_Generalization")
        )

        prog_per_paper_no_generalization = (
            framework_pair_data_no_generalization
            .groupby("Paper_ID")["Programming_element"]
            .nunique()
            .reset_index(name="Number_of_programming_elements_without_Generalization")
        )

        framework_counts_summary = pd.DataFrame([
            {
                "Analysis": "Framework including Algorithmic Thinking",
                "Number_of_papers": framework_pair_data["Paper_ID"].nunique(),
            },
            {
                "Analysis": "Framework excluding Algorithmic Thinking",
                "Number_of_papers": framework_pair_data_no_algorithmic["Paper_ID"].nunique(),
            },
            {
                "Analysis": "Framework excluding Generalization",
                "Number_of_papers": framework_pair_data_no_generalization["Paper_ID"].nunique(),
            },
        ])

        outputs["framework"] = {
            "total": framework_total,
            "freq_col": framework_freq_col,
            "paper_data": framework_paper_data,
            "ct_summary": framework_ct_summary,
            "programming_summary": framework_programming_summary,
            "programming_summary_from_pairs": framework_programming_summary_from_pairs,
            "pair_data": framework_pair_data,
            "pair_summary": framework_pair_summary,
            "ct_per_paper": ct_per_paper_summary,
            "prog_per_paper": prog_per_paper_summary,
            "matrix": framework_matrix,
            "matrix_no_algorithmic": framework_matrix_no_algorithmic,
            "heatmap_data": framework_heatmap_data,
            "heatmap_no_algorithmic_data": framework_heatmap_no_algorithmic_data,
            "pair_data_no_algorithmic": framework_pair_data_no_algorithmic,
            "paper_counts_summary": framework_counts_summary,
            "pair_data_no_generalization": framework_pair_data_no_generalization,
            "matrix_no_generalization": framework_matrix_no_generalization,
            "heatmap_no_generalization_data": framework_heatmap_no_generalization_data,
            "pair_activity_heatmap_data": pair_activity_heatmap_data,
            "ct_per_paper_no_generalization": ct_per_paper_no_generalization,
            "prog_per_paper_no_generalization": prog_per_paper_no_generalization,
            "matrix_no_alg_no_gen": framework_matrix_no_alg_no_gen,
            "heatmap_no_alg_no_gen_data": framework_heatmap_no_alg_no_gen_data,
            "pair_data_no_alg_no_gen": framework_pair_data_no_alg_no_gen,
            "prada_like_summary": pd.DataFrame({
                "Criterion": [
                    "Contains Abstraction + Algorithmic Thinking + Pattern Recognition + Decomposition"
                ],
                "Number_of_papers": [prada_like_count],
            }),
            "prada_like_papers": prada_like_papers,
            "prada_like_details": prada_like_details,
        }

    # ==================================================
    # Export Excel
    # ==================================================

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        if "mapped" in outputs:
            outputs["mapped"]["paper_data"].to_excel(
                writer,
                sheet_name="Mapped_papers",
                index=False,
            )

            outputs["mapped"]["programming_summary"].to_excel(
                writer,
                sheet_name="Mapped_Programming",
                index=False,
            )

            outputs["mapped"]["ct_summary"].to_excel(
                writer,
                sheet_name="Mapped_CT",
                index=False,
            )

        if "framework" in outputs:
            fw = outputs["framework"]

            fw["paper_data"].to_excel(
                writer,
                sheet_name="Framework_papers",
                index=False,
            )

            fw["programming_summary"].to_excel(
                writer,
                sheet_name="Framework_Programming",
                index=False,
            )

            fw["programming_summary_from_pairs"].to_excel(
                writer,
                sheet_name="Framework_Prog_pairs",
                index=False,
            )

            fw["ct_summary"].to_excel(
                writer,
                sheet_name="Framework_CT",
                index=False,
            )

            fw["pair_data"].to_excel(
                writer,
                sheet_name="Framework_pairs",
                index=False,
            )

            fw["pair_summary"].to_excel(
                writer,
                sheet_name="Framework_pair_summary",
                index=False,
            )

            fw["ct_per_paper"].to_excel(
                writer,
                sheet_name="CT_per_paper",
                index=False,
            )

            fw["prog_per_paper"].to_excel(
                writer,
                sheet_name="Prog_per_paper",
                index=False,
            )

            fw["matrix"].to_excel(
                writer,
                sheet_name="Framework_matrix",
                index=False,
            )

            fw["matrix_no_algorithmic"].to_excel(
                writer,
                sheet_name="Matrix_no_algorithmic",
                index=False,
            )

            fw["heatmap_data"].to_excel(
                writer,
                sheet_name="Heatmap_matrix",
            )

            fw["heatmap_no_algorithmic_data"].to_excel(
                writer,
                sheet_name="Heatmap_no_algorithmic",
            )

            fw["paper_counts_summary"].to_excel(
                writer,
                sheet_name="Framework_counts",
                index=False,
            )

            fw["pair_data_no_generalization"].to_excel(
                writer,
                sheet_name="Pairs_no_generalization",
                index=False,
            )

            fw["matrix_no_generalization"].to_excel(
                writer,
                sheet_name="Matrix_no_generalization",
                index=False,
            )

            fw["heatmap_no_generalization_data"].to_excel(
                writer,
                sheet_name="Heatmap_no_generalization",
            )

            fw["heatmap_no_generalization_data"].to_excel(
            writer,
            sheet_name="Heatmap_no_generalization",
            )

            fw["pair_data_no_alg_no_gen"].to_excel(
                writer,
                sheet_name="Pairs_no_alg_no_gen",
                index=False,
            )

            fw["matrix_no_alg_no_gen"].to_excel(
                writer,
                sheet_name="Matrix_no_alg_no_gen",
                index=False,
            )

            fw["heatmap_no_alg_no_gen_data"].to_excel(
                writer,
                sheet_name="Heatmap_no_alg_no_gen",
            )

            fw["ct_per_paper_no_generalization"].to_excel(
                writer,
                sheet_name="CT_no_generalization",
                index=False,
            )

            fw["prog_per_paper_no_generalization"].to_excel(
                writer,
                sheet_name="Prog_no_generalization",
                index=False,
            )

            fw["prada_like_summary"].to_excel(
                writer,
                sheet_name="PRADA_like_summary",
                index=False,
            )

            fw["prada_like_papers"].to_excel(
                writer,
                sheet_name="PRADA_like_papers",
                index=False,
            )

            fw["prada_like_details"].to_excel(
                writer,
                sheet_name="PRADA_like_details",
                index=False,
            )

            

            # One sheet per CT element
            for ct in sorted(fw["pair_data"]["CT_element"].dropna().unique()):
                ct_df = fw["pair_data"][
                    fw["pair_data"]["CT_element"] == ct
                ].copy()

                ct_df.to_excel(
                    writer,
                    sheet_name=clean_sheet_name(f"Pairs_{ct}"),
                    index=False,
                )

    # ==================================================
    # Generate figures
    # ==================================================

    if "mapped" in outputs:
        mapped = outputs["mapped"]

        fig = frequency_bar(
            mapped["ct_summary"],
            "CT element",
            mapped["freq_col"],
            "CT Element Distribution in Mapped Papers",
            "CT element",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "mapped_ct_distribution")
            )

        fig = frequency_bar(
            mapped["programming_summary"],
            "Programming element",
            mapped["freq_col"],
            "Programming Element Distribution in Mapped Papers",
            "Programming element",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "mapped_programming_distribution")
            )

    if "framework" in outputs:
        fw = outputs["framework"]

        fig = frequency_bar(
            fw["ct_summary"],
            "CT element",
            fw["freq_col"],
            "CT Element Distribution in Framework-Following Papers",
            "CT element",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "framework_ct_distribution")
            )

        fig = frequency_bar(
            fw["programming_summary_from_pairs"],
            "Programming element",
            fw["freq_col"],
            "Programming Element Distribution from CT–Programming Pairs",
            "Programming element",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "framework_programming_distribution_from_pairs")
            )

        fig = heatmap_figure(
            fw["heatmap_data"],
            "Frequency of Co-occurrence Between CT Elements and Programming Elements",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "ct_programming_heatmap")
            )

        fig = heatmap_figure(
            fw["heatmap_no_algorithmic_data"],
            "Frequency of Co-occurrence Excluding Algorithmic Thinking",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "ct_programming_heatmap_no_algorithmic")
            )

        fig = heatmap_figure(
            fw["heatmap_no_generalization_data"],
            "Frequency of Co-occurrence Excluding Generalization",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "ct_programming_heatmap_no_generalization")
            )

        fig = network_plot_for_data(
            fw["pair_data"],
            "CT–Programming Relationship Network",
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "ct_programming_network")
            )

        fig = network_plot_filtered(
            fw["pair_data"],
            "CT–Programming Relationship Network Excluding Algorithmic Thinking",
            exclude_algorithmic=True,
        )
        if fig is not None:
            generated_figures.append(
                save_figure(fig, figures_dir, "ct_programming_network_no_algorithmic")
            )

        # One network figure per CT element
        for ct in sorted(fw["pair_data"]["CT_element"].dropna().unique()):
            fig = network_plot_filtered(
                fw["pair_data"],
                f"CT–Programming Relationship Network: {ct}",
                ct_filter=ct,
            )

            if fig is not None:
                generated_figures.append(
                    save_figure(
                        fig,
                        figures_dir,
                        f"ct_programming_network_{safe_filename(ct)}",
                    )
                )

        # Optional comparison charts if both outputs exist
        if "mapped" in outputs:
            mapped = outputs["mapped"]

            fig = grouped_frequency_bar(
                mapped["ct_summary"],
                fw["ct_summary"],
                "CT element",
                mapped["freq_col"],
                fw["freq_col"],
                "CT Element Distribution: Mapped vs Framework-Following Papers",
            )
            if fig is not None:
                generated_figures.append(
                    save_figure(fig, figures_dir, "ct_distribution_comparison")
                )

            fig = grouped_frequency_bar(
                mapped["programming_summary"],
                fw["programming_summary_from_pairs"],
                "Programming element",
                mapped["freq_col"],
                fw["freq_col"],
                "Programming Element Distribution: Mapped vs Framework-Following Papers",
            )
            if fig is not None:
                generated_figures.append(
                    save_figure(fig, figures_dir, "programming_distribution_comparison")
                )

    return {
        "output_path": output_path,
        "figures_dir": figures_dir,
        "generated_figures": generated_figures,
        "total_loaded_papers": total_loaded,
        "total_mapped_papers": total_mapped_papers,
        "total_framework_papers": (
            outputs.get("framework", {}).get("total", 0)
            if "framework" in outputs
            else 0
        ),
        "prada_like_count": prada_like_count,
        "outputs": outputs,
    }