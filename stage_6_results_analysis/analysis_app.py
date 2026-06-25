import os
from pathlib import Path

import pandas as pd
import streamlit as st
import seaborn as sns
import matplotlib.pyplot as plt
import plotly.express as px
import shutil
from pipeline_summary import build_final_review_summary
from pairing_analysis import (
    run_pairing_analysis,
    grouped_frequency_bar,
    heatmap_figure,
)


# ==================================================
# Paths
# ==================================================

STAGE_6_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAGE_6_DIR.parent

INPUT_DIR_6 = STAGE_6_DIR / "input"
OUTPUT_DIR_6 = STAGE_6_DIR / "output"
TEMP_DIR_6 = STAGE_6_DIR / "temp"

INPUT_DIR_6.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_6.mkdir(parents=True, exist_ok=True)
TEMP_DIR_6.mkdir(parents=True, exist_ok=True)

STAGE_6_ANALYSIS_DETAILS_PATH = INPUT_DIR_6 / "Analysis_details.xlsx"


# ==================================================
# Previous stage outputs
# ==================================================

STAGE_4_PAIR_MAPPING_PATH = (
    PROJECT_ROOT
    / "stage_4_full_pairing"
    / "output"
    / "pair_level_mapping_results.xlsx"
)

STAGE_5_CHARACTERIZATION_PATH = (
    PROJECT_ROOT
    / "stage_5_characterization"
    / "output"
    / "results_characterization.xlsx"
)

STAGE_4_PAPER_MAPPING_PATH = (
    PROJECT_ROOT
    / "stage_4_full_pairing"
    / "output"
    / "paper_level_framework_results.xlsx"
)
# --------------------------------------------------
# Copying outputs as inputs to stage 6
# --------------------------------------------------

STAGE_6_CHARACTERIZATION_INPUT_PATH = INPUT_DIR_6 / "results_characterization.xlsx"
STAGE_6_PAPER_MAPPING_INPUT_PATH = INPUT_DIR_6 / "paper_level_framework_results.xlsx"
STAGE_6_PAIR_MAPPING_INPUT_PATH = INPUT_DIR_6 / "pair_level_mapping_results.xlsx"



def copy_previous_stage_file(source_path, target_path):
    source_path = Path(source_path)
    target_path = Path(target_path)

    if source_path.exists():
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, target_path)
        return str(target_path)

    return None

def main():

# --------------------------------------------------
# Streamlit setup
# --------------------------------------------------


    st.set_page_config(
        page_title="Stage 6 – Results Analysis",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊 Stage 6 – Results Analysis")

    st.write(
        """
    Final analysis and visualization of characterization
    and CT–programming relationship results.
    """
    )

    # --------------------------------------------------
    # Helper functions
    # --------------------------------------------------

    def clean_sheet_name(name):
        invalid_chars = ["\\", "/", "*", "?", ":", "[", "]"]
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
        )


    def frequency_table(df, column_name):
        counts = df[column_name].value_counts(dropna=False)
        result = counts.reset_index()
        result.columns = [column_name, "Count"]
        result["Percentage"] = (result["Count"] / len(df) * 100).round(2)

        total_row = pd.DataFrame({
            column_name: ["TOTAL"],
            "Count": [result["Count"].sum()],
            "Percentage": [100.0]
        })

        return pd.concat([result, total_row], ignore_index=True)


    def collect_rule_violations(data, rule_id, description, mask):
        issues = data[mask].copy()
        if not issues.empty:
            issues["Rule_ID"] = rule_id
            issues["Rule_Description"] = description
        return issues


    def load_excel_or_tsv(input_file):
        try:
            df = pd.read_excel(input_file)

            if len(df.columns) == 1:
                raise ValueError("File looks like text, not real Excel")

            return df

        except Exception:
            if not isinstance(input_file, str):
                input_file.seek(0)

            return pd.read_csv(
                input_file,
                sep="\t",
                encoding="latin1",
            )

    def clean_other_text(value):
        value = str(value).strip()
        if value.startswith("Other:"):
            return "Other"
        return value


    def group_not_allowed(value, allowed_values):
        value = clean_other_text(value)
        if value not in allowed_values:
            return "Other"
        return value


    def order_counts(series, preferred_order=None):
        counts = series.value_counts()

        special_categories = [
            "Other",
            "Undetermined",
            "Manual revision",
            "Not applicable"
        ]

        if preferred_order is None:
            normal_labels = [
                x for x in counts.index
                if x not in special_categories
            ]
            ordered_labels = normal_labels
        else:
            ordered_labels = [
                x for x in preferred_order
                if x in counts.index and x not in special_categories
            ]

        remaining = [
            x for x in counts.index
            if x not in ordered_labels and x not in special_categories
        ]

        ordered_labels.extend(sorted(remaining))

        for special in special_categories:
            if special in counts.index:
                ordered_labels.append(special)

        return counts.reindex(ordered_labels)



    def make_bar_chart(series, title, xlabel, filename, figures_dir, preferred_order=None):
        counts = order_counts(series, preferred_order)

        if counts.empty:
            return

        fig, ax = plt.subplots(figsize=(10, 5))
        counts.plot(kind="bar", ax=ax, color="#1192b5")

        for container in ax.containers:
            ax.bar_label(container, label_type="edge", padding=4)

        max_count = counts.max()
        ax.set_ylim(0, max_count + max(1, max_count * 0.18))

        ax.set_title(title)
        ax.set_xlabel(xlabel)
        ax.set_ylabel("Number of studies")

        for label in ax.get_xticklabels():
            label.set_rotation(45)
            label.set_horizontalalignment("right")

        fig.tight_layout()

        st.pyplot(fig)

        fig.savefig(figures_dir / f"{filename}.png", dpi=300, bbox_inches="tight")
        fig.savefig(figures_dir / f"{filename}.pdf", bbox_inches="tight")

        plt.close(fig)

    def classify_tool_language(value, tools_grouping):

        value = str(value).strip()

        # Preserve existing categories
        if value == "Other":
            return "Other"

        if value == "Not applicable":
            return "Not applicable"

        if value in ["", "nan", "None"]:
            return "Undetermined"

        for _, row in tools_grouping.iterrows():

            category = str(row["Category"]).strip()
            examples = str(row["Examples"]).strip()

            example_list = [
                item.strip()
                for item in examples.split(",")
                if item.strip()
            ]

            for example in example_list:
                if example.lower() in value.lower():
                    return category

        # Any unmatched value goes to existing "Other"
        return "Other"
    
    def grouped_characterization_bar(
            full_data,
            framework_data,
            column,
            title,
            xlabel,
            figures_dir,
            filename,
            preferred_order=None,
        ):
            full_counts = full_data[column].value_counts()
            framework_counts = framework_data[column].value_counts()

            labels = list(full_counts.index)

            if preferred_order is not None:
                labels = [x for x in preferred_order if x in full_counts.index]

            for label in framework_counts.index:
                if label not in labels:
                    labels.append(label)

            comparison = pd.DataFrame({
                column: labels,
                "All characterized papers": [
                    int(full_counts.get(label, 0)) for label in labels
                ],
                "Framework papers": [
                    int(framework_counts.get(label, 0)) for label in labels
                ],
            })

            comparison = comparison[
                ~comparison[column].astype(str).str.strip().isin([
                    "Undetermined",
                    "Not applicable",
                    "Manual revision",
                    "",
                ])
            ].copy()

            if comparison.empty:
                return

            x = range(len(comparison))
            width = 0.38

            fig, ax = plt.subplots(figsize=(12, 6))

            bars1 = ax.bar(
                [i - width / 2 for i in x],
                comparison["All characterized papers"],
                width,
                label="All characterized papers",
                color="#1192b5",
            )

            bars2 = ax.bar(
                [i + width / 2 for i in x],
                comparison["Framework papers"],
                width,
                label="Framework papers",
                color="#f28e2b",
            )

            ax.bar_label(bars1, padding=3)
            ax.bar_label(bars2, padding=3)

            max_value = max(
                comparison["All characterized papers"].max(),
                comparison["Framework papers"].max(),
            )

            ax.set_ylim(0, max_value + max(1, max_value * 0.18))
            ax.set_title(title)
            ax.set_xlabel(xlabel)
            ax.set_ylabel("Number of papers")
            ax.set_xticks(list(x))
            ax.set_xticklabels(comparison[column].astype(str), rotation=45, ha="right")
            ax.legend()

            fig.tight_layout()

            st.pyplot(fig)

            fig.savefig(figures_dir / f"{filename}.png", dpi=300, bbox_inches="tight")
            fig.savefig(figures_dir / f"{filename}.pdf", bbox_inches="tight")

            plt.close(fig)


    # --------------------------------------------------
    # Interface
    # --------------------------------------------------

    AUTO_INPUT_MODE = ("Automatically use outputs from previous stages")

    CUSTOM_INPUT_MODE = ("Use custom uploaded files")

    input_mode = st.radio(
        "Select input source",
        [
            AUTO_INPUT_MODE,
            CUSTOM_INPUT_MODE,
        ],
    )

    characterization_file = None
    paper_level_file = None
    pair_level_file = None
    tools_grouping_file = None

    # ==========================================
    # Automatic loading
    # ==========================================
    if input_mode == AUTO_INPUT_MODE:

        characterization_file = copy_previous_stage_file(
            STAGE_5_CHARACTERIZATION_PATH,
            STAGE_6_CHARACTERIZATION_INPUT_PATH,
        )

        paper_level_file = copy_previous_stage_file(
            STAGE_4_PAPER_MAPPING_PATH,
            STAGE_6_PAPER_MAPPING_INPUT_PATH,
        )

        pair_level_file = copy_previous_stage_file(
            STAGE_4_PAIR_MAPPING_PATH,
            STAGE_6_PAIR_MAPPING_INPUT_PATH,
        )

        if characterization_file:
            st.info("Characterization dataset copied to Stage 6 input.")
            st.caption(f"Using: {characterization_file}")
        else:
            st.warning("Stage 5 characterization output not found.")

        if paper_level_file:
            st.info("Paper-level mapping dataset copied to Stage 6 input.")
            st.caption(f"Using: {paper_level_file}")
        else:
            st.warning("Stage 4 paper-level mapping output not found.")

        if pair_level_file:
            st.info("Pair-level mapping dataset copied to Stage 6 input.")
            st.caption(f"Using: {pair_level_file}")
        else:
            st.warning("Stage 4 pair-level mapping output not found.")

        st.subheader("CT–Programming Pair Analysis")

    # ==========================================
    # Manual upload
    # ==========================================
    else:

        characterization_file = st.file_uploader("Upload characterization dataset", type=["xlsx", "csv", "tsv"],)

        paper_level_file = st.file_uploader("Upload paper-level mapping dataset", type=["xlsx"],)

        pair_level_file = st.file_uploader("Upload pair-level CT–Programming mapping dataset",type=["xlsx"],)

    if input_mode == AUTO_INPUT_MODE:

        if STAGE_6_ANALYSIS_DETAILS_PATH.exists():

            st.info(
                "Analysis details file "
                "automatically loaded."
            )

            st.caption(
                f"Using: "
                f"{STAGE_6_ANALYSIS_DETAILS_PATH}"
            )

            tools_grouping_file = str(
                STAGE_6_ANALYSIS_DETAILS_PATH
            )

        else:

            st.warning(
                "Analysis_details.xlsx "
                "not found in Stage 6 input."
            )

            tools_grouping_file = st.file_uploader(
                "Upload tools grouping file",
                type=["xlsx"]
            )

    else:

        tools_grouping_file = st.file_uploader(
            "Upload tools grouping file",
            type=["xlsx"]
        )

    output_name = st.text_input("Output Excel filename", value="characterization_analysis.xlsx")

    run_button = st.button("Run analysis")

    def display_pairing_analysis_ordered(pairing_summary):
        st.header("RQ1–RQ2: CT–Programming elements and framework papers")

        outputs = pairing_summary.get("outputs", {})

        if "mapped" not in outputs or "framework" not in outputs:
            st.warning("Mapped/framework comparison is not available.")
            return

        mapped = outputs["mapped"]
        fw = outputs["framework"]

        mapped_freq_col = mapped["freq_col"]
        framework_freq_col = fw["freq_col"]

        st.info(
            f"Mapped papers analysed: {mapped['total']} | "
            f"Framework-following papers analysed: {fw['total']}"
        )

        # -------------------------------
        # RQ1: CT elements
        # -------------------------------
        st.subheader("RQ1. CT elements: total mapped papers vs framework papers")

        fig = grouped_frequency_bar(
            mapped_df=mapped["ct_summary"],
            framework_df=fw["ct_summary"],
            label_col="CT element",
            mapped_value_col=mapped_freq_col,
            framework_value_col=framework_freq_col,
            title="CT elements in mapped papers and framework-following papers",
        )

        if fig is not None:
            st.pyplot(fig)
            plt.close(fig)

        with st.expander("See CT element tables"):
            st.markdown("Mapped papers")
            st.dataframe(mapped["ct_summary"], use_container_width=True)

            st.markdown("Framework-following papers")
            st.dataframe(fw["ct_summary"], use_container_width=True)

        # -------------------------------
        # RQ2: Programming elements
        # -------------------------------
        st.subheader("RQ2. Programming elements: total mapped papers vs framework papers")

        fig = grouped_frequency_bar(
            mapped_df=mapped["programming_summary"],
            framework_df=fw["programming_summary_from_pairs"],
            label_col="Programming element",
            mapped_value_col=mapped_freq_col,
            framework_value_col=framework_freq_col,
            title="Programming elements in mapped papers and framework-following papers",
        )

        if fig is not None:
            st.pyplot(fig)
            plt.close(fig)

        with st.expander("See programming element tables"):
            st.markdown("Mapped papers")
            st.dataframe(mapped["programming_summary"], use_container_width=True)

            st.markdown("Framework-following papers")
            st.dataframe(fw["programming_summary_from_pairs"], use_container_width=True)

        # -------------------------------
        # Heatmaps
        # -------------------------------
        st.header("CT–Programming heatmaps")

        st.subheader("CT × Programming element heatmap")

        fig = heatmap_figure(
            fw["heatmap_data"],
            "CT–Programming pair frequencies",
        )

        if fig is not None:
            st.pyplot(fig)
            plt.close(fig)

        with st.expander("Alternative heatmaps"):
            st.markdown("Excluding Algorithmic Thinking")
            fig = heatmap_figure(
                fw["heatmap_no_algorithmic_data"],
                "CT–Programming pair frequencies excluding Algorithmic Thinking",
            )
            if fig is not None:
                st.pyplot(fig)
                plt.close(fig)

            if "heatmap_no_generalization_data" in fw:
                st.markdown("Excluding Generalization")
                fig = heatmap_figure(
                    fw["heatmap_no_generalization_data"],
                    "CT–Programming pair frequencies excluding Generalization",
                )
                if fig is not None:
                    st.pyplot(fig)
                    plt.close(fig)

            if "heatmap_no_alg_no_gen_data" in fw:
                st.markdown("Excluding Algorithmic Thinking and Generalization")
                fig = heatmap_figure(
                    fw["heatmap_no_alg_no_gen_data"],
                    "CT–Programming pair frequencies excluding Algorithmic Thinking and Generalization",
                )
                if fig is not None:
                    st.pyplot(fig)
                    plt.close(fig)

        # -------------------------------
        # Operationalizations
        # -------------------------------
        st.header("Operationalizations")

        if (
            "pair_activity_heatmap_data" in fw
            and not fw["pair_activity_heatmap_data"].empty
        ):
            st.subheader("Operationalization activities for recurrent CT–Programming pairs")

            fig = heatmap_figure(
                fw["pair_activity_heatmap_data"],
                "Operationalization activities for recurrent CT–Programming pairs",
            )

            if fig is not None:
                st.pyplot(fig)
                plt.close(fig)

        st.subheader("CT–Programming pair summary")
        st.dataframe(fw["pair_summary"], use_container_width=True)

        st.subheader("Papers using all 5 CT elements")

        five_ct_papers = fw["ct_per_paper"][
            fw["ct_per_paper"]["Number_of_CT_elements"] >= 5
        ].copy()

        st.write(f"Papers with at least 5 CT elements: {len(five_ct_papers)}")
        st.dataframe(five_ct_papers, use_container_width=True)

        if not five_ct_papers.empty:
            five_ct_ids = set(
                five_ct_papers["Paper_ID"]
                .astype(str)
                .str.strip()
            )

            five_ct_pair_details = fw["pair_data"][
                fw["pair_data"]["Paper_ID"]
                .astype(str)
                .str.strip()
                .isin(five_ct_ids)
            ].copy()

            st.subheader("Pair details for papers using all 5 CT elements")
            st.dataframe(five_ct_pair_details, use_container_width=True)


        with st.expander("Matrices and paper-level operationalization details"):
            st.markdown("Full CT–Programming matrix")
            st.dataframe(fw["matrix"], use_container_width=True)

            st.markdown("Matrix excluding Algorithmic Thinking")
            st.dataframe(fw["matrix_no_algorithmic"], use_container_width=True)

            if "matrix_no_generalization" in fw:
                st.markdown("Matrix excluding Generalization")
                st.dataframe(fw["matrix_no_generalization"], use_container_width=True)

            if "matrix_no_alg_no_gen" in fw:
                st.markdown("Matrix excluding Algorithmic Thinking and Generalization")
                st.dataframe(fw["matrix_no_alg_no_gen"], use_container_width=True)

            st.markdown("CT elements per paper")
            st.dataframe(fw["ct_per_paper"], use_container_width=True)

            st.markdown("Programming elements per paper")
            st.dataframe(fw["prog_per_paper"], use_container_width=True)

        if "prada_like_summary" in fw:
            with st.expander("PRADA-like papers"):
                st.dataframe(fw["prada_like_summary"], use_container_width=True)
                st.dataframe(fw["prada_like_papers"], use_container_width=True)
                st.dataframe(fw["prada_like_details"], use_container_width=True)
        # --------------------------------------------------
    # Main analysis
    # --------------------------------------------------

    if run_button:

        if characterization_file is None:
            st.error("Please upload the characterization dataset.")
            st.stop()

        if paper_level_file is None:
            st.error("Please upload the paper-level mapping dataset.")
            st.stop()
        
        if pair_level_file is None:
            st.error("Please provide the pair-level CT–Programming mapping dataset.")
            st.stop()

        if tools_grouping_file is None:
            st.error("Please upload the tools grouping file.")
            st.stop()
        

        # ---------- Load data ----------

        data = load_excel_or_tsv(characterization_file)
        paper_data = pd.read_excel(paper_level_file)
        pair_data = pd.read_excel(pair_level_file)

        data.columns = data.columns.str.strip()
        paper_data.columns = paper_data.columns.str.strip()
        pair_data.columns = pair_data.columns.str.strip()

        data = data.loc[:, ~data.columns.str.contains("^Unnamed")]
        paper_data = paper_data.loc[:, ~paper_data.columns.str.contains("^Unnamed")]
        pair_data = pair_data.loc[:, ~pair_data.columns.str.contains("^Unnamed")]

        data["Paper_ID"] = data["Paper_ID"].astype(str).str.strip()
        paper_data["Paper_ID"] = paper_data["Paper_ID"].astype(str).str.strip()
        pair_data["Paper_ID"] = pair_data["Paper_ID"].astype(str).str.strip()

        if "Processing_Status" in pair_data.columns:
            pair_data = pair_data[
                pair_data["Processing_Status"]
                .fillna("")
                .astype(str)
                .str.strip()
                == "Success"
            ].copy()

        framework_ids = set(pair_data["Paper_ID"].dropna().astype(str).str.strip())

        data["Is_Framework_Paper"] = data["Paper_ID"].isin(framework_ids)

        framework_characterization_df = data[
            data["Is_Framework_Paper"] == True
        ].copy()

        output_path = OUTPUT_DIR_6 / output_name

        figures_dir = OUTPUT_DIR_6 / "figures"
        figures_dir.mkdir(parents=True, exist_ok=True)

        st.info(
            f"Characterization papers that are part of our framework: "
            f"{framework_characterization_df['Paper_ID'].nunique()}"
        )

        st.dataframe(
            framework_characterization_df[["Paper_ID", "Paper"]]
            if "Paper" in framework_characterization_df.columns
            else framework_characterization_df,
            use_container_width=True
        )
        tools_grouping = pd.read_excel(tools_grouping_file, sheet_name="Tools")
        tools_order = tools_grouping["Category"].dropna().astype(str).str.strip().tolist()    

        data.columns = data.columns.str.strip()
        paper_data.columns = paper_data.columns.str.strip()
        tools_grouping.columns = tools_grouping.columns.str.strip()

        data = data.loc[:, ~data.columns.str.contains("^Unnamed")]
        paper_data = paper_data.loc[:, ~paper_data.columns.str.contains("^Unnamed")]

        required_data_cols = ["Paper_ID"]
        required_paper_cols = ["Paper_ID", "Type_of_Mapping"]

        missing_data_cols = [col for col in required_data_cols if col not in data.columns]
        missing_paper_cols = [col for col in required_paper_cols if col not in paper_data.columns]

        if missing_data_cols:
            st.error(f"Missing columns in characterization dataset: {missing_data_cols}")
            st.stop()

        if missing_paper_cols:
            st.error(f"Missing columns in paper-level mapping dataset: {missing_paper_cols}")
            st.write("Available columns:")
            st.write(paper_data.columns.tolist())
            st.stop()

        data["Paper_ID"] = data["Paper_ID"].astype(str).str.strip()
        paper_data["Paper_ID"] = paper_data["Paper_ID"].astype(str).str.strip()

        # ---------- Filter only mapped papers ----------

        mapped_papers = paper_data[
            paper_data["Type_of_Mapping"].fillna("").astype(str).str.strip() != "No Mapping"
        ].copy()

        mapped_ids = mapped_papers["Paper_ID"].unique()

        total_characterization_papers = len(data)
        total_mapped_papers = len(mapped_ids)

        data_all = data.copy()
        data_mapped = data[data["Paper_ID"].isin(mapped_ids)].copy()

        # Use all characterized papers for characterization analysis
        data = data_all

        missing_ids = set(mapped_ids) - set(data["Paper_ID"])

        st.info(f"Total characterization papers loaded: {total_characterization_papers}")
        st.info(f"Mapped papers in paper-level dataset: {total_mapped_papers}")
        st.success(f"Characterization papers kept for characterization analysis: {len(data)}")
        st.info(f"Mapped characterization papers available for mapping-specific analysis: {len(data_mapped)}")


        st.header("Disciplinary focus")

        grouped_characterization_bar(
            full_data=data,
            framework_data=framework_characterization_df,
            column="Disciplinary_Focus",
            title="Disciplinary focus: all papers vs framework papers",
            xlabel="Disciplinary focus",
            figures_dir=OUTPUT_DIR_6,
            filename="Disciplinary_focus_framework_comparison",
        )
        if len(data) == 0:
            st.error("No papers left after filtering. Check Paper_ID values or mapping file.")
            st.stop()

        if missing_ids:
            st.warning(f"Mapped Paper_IDs not found in characterization dataset: {sorted(missing_ids)}")

        # ---------- Clean category variations ----------

        data = data.replace({
            "Case study/descriptive": "Case study / descriptive",
            "Pre-post": "Pre–post",
            "pre-post": "Pre–post",
        })

        # ---------- Rule checking ----------

        rule_issues = []

        mask_r1 = (
            (data["Programming_Type (paradigm)"] == "Unplugged (no programming)") &
            (
                (data["Programming_Tool/Language"] != "Not applicable") |
                (data["Programming_Kind"] != "Not applicable")
            )
        )
        issue = collect_rule_violations(
            data, "R1", "Unplugged requires no tool or programming kind", mask_r1
        )
        if not issue.empty:
            rule_issues.append(issue)

        mask_r2 = (
            ~data["Type_of_Contribution"].isin(["Learning model application", "Mixed"]) &
            (data["Activities_Conducted_for_Learning"] != "Not applicable")
        )
        issue = collect_rule_violations(
            data, "R2", "Activities should be NA if no learning contribution", mask_r2
        )
        if not issue.empty:
            rule_issues.append(issue)

        mask_r3 = (
            ~data["Type_of_Contribution"].isin(["Evaluation", "Mixed"]) &
            (data["Evaluation_Process_Conducted"] != "Not applicable")
        )
        issue = collect_rule_violations(
            data, "R3", "Evaluation process should be NA if no evaluation", mask_r3
        )
        if not issue.empty:
            rule_issues.append(issue)

        mask_r4 = (
            (data["Disciplinary_Focus"] != "Interdisciplinary") &
            (data["It_is_interdisciplinary"] != "Not applicable")
        )
        issue = collect_rule_violations(
            data, "R4", "Discipline detail only applies to interdisciplinary studies", mask_r4
        )
        if not issue.empty:
            rule_issues.append(issue)

        if rule_issues:
            rule_issues_df = pd.concat(rule_issues, ignore_index=True)
        else:
            rule_issues_df = pd.DataFrame()

        # ---------- Derived analytical variables ----------

        data["Has_Rule_Violation"] = False

        for mask in [mask_r1, mask_r2, mask_r3, mask_r4]:
            data.loc[mask, "Has_Rule_Violation"] = True

        data["Code_Based_Programming_Present"] = ~data["Programming_Type (paradigm)"].isin([
            "Not applicable",
            "Unplugged (no programming)",
            "Undetermined",
            "Manual revision"
        ])

        data["CT_with_Programming"] = (
            (data["Learning_Focus"] == "Computational Thinking") &
            (data["Code_Based_Programming_Present"] == True)
        )

        data["Evaluation_Present"] = data["Type_of_Contribution"].isin([
            "Evaluation",
            "Mixed"
        ])

        data["Learning_Application_Present"] = data["Type_of_Contribution"].isin([
            "Learning model application",
            "Mixed"
        ])

        data["Programming_Tool_Grouped"] = data["Programming_Tool/Language"].apply(
        lambda x: classify_tool_language(x, tools_grouping)
        )

        data["Programming_Environment_Grouped"] = data["Programming_Environment"].apply(
            lambda x: "Other" if str(x).strip().startswith("Other:") else x
        )
        # ------------------------------------------
        st.header("Disciplinary focus vs educational context")

        discipline_education_table = pd.crosstab(
            data["Disciplinary_Focus"],
            data["Educational_Context"]
        )

        excluded_values = [
            "",
            "Undetermined",
            "Not applicable",
            "Manual revision",
        ]

        discipline_education_table = discipline_education_table.drop(
            index=[
                x for x in excluded_values
                if x in discipline_education_table.index
            ],
            errors="ignore",
        )

        discipline_education_table = discipline_education_table.drop(
            columns=[
                x for x in excluded_values
                if x in discipline_education_table.columns
            ],
            errors="ignore",
        )

        fig, ax = plt.subplots(figsize=(10, 6))

        sns.heatmap(
            discipline_education_table,
            annot=True,
            fmt="d",
            cmap="YlGnBu",
            linewidths=0.5,
            linecolor="white",
            cbar_kws={"label": "Number of papers"},
            ax=ax,
        )

        ax.set_title("Disciplinary Focus vs Educational Context")
        ax.set_xlabel("Educational Context")
        ax.set_ylabel("Disciplinary Focus", labelpad=20)

        plt.xticks(rotation=45, ha="right")
        plt.yticks(rotation=0)
        plt.tight_layout()
        plt.subplots_adjust(left=0.25)

        st.pyplot(fig)

        fig.savefig(
            figures_dir / "disciplinary_focus_vs_educational_context_heatmap.png",
            dpi=300,
            bbox_inches="tight",
        )

        fig.savefig(
            figures_dir / "disciplinary_focus_vs_educational_context_heatmap.pdf",
            bbox_inches="tight",
        )

        st.header("Disciplinary focus vs educational context – framework papers")

        framework_discipline_education_table = pd.crosstab(
            framework_characterization_df["Disciplinary_Focus"],
            framework_characterization_df["Educational_Context"]
        )

        excluded_values = [
            "",
            "Undetermined",
            "Not applicable",
            "Manual revision",
        ]

        framework_discipline_education_table = framework_discipline_education_table.drop(
            index=[x for x in excluded_values if x in framework_discipline_education_table.index],
            errors="ignore",
        )

        framework_discipline_education_table = framework_discipline_education_table.drop(
            columns=[x for x in excluded_values if x in framework_discipline_education_table.columns],
            errors="ignore",
        )

        fig, ax = plt.subplots(figsize=(10, 6))

        sns.heatmap(
            framework_discipline_education_table,
            annot=True,
            fmt="d",
            cmap="YlGnBu",
            linewidths=0.5,
            linecolor="white",
            cbar_kws={"label": "Number of framework papers"},
            ax=ax,
        )

        ax.set_title("Disciplinary Focus vs Educational Context – Framework Papers")
        ax.set_xlabel("Educational Context")
        ax.set_ylabel("Disciplinary Focus", labelpad=20)

        plt.xticks(rotation=45, ha="right")
        plt.yticks(rotation=0)
        plt.tight_layout()
        plt.subplots_adjust(left=0.25)

        st.pyplot(fig)

        fig.savefig(
            figures_dir / "disciplinary_focus_vs_educational_context_framework_heatmap.png",
            dpi=300,
            bbox_inches="tight",
        )

        fig.savefig(
            figures_dir / "disciplinary_focus_vs_educational_context_framework_heatmap.pdf",
            bbox_inches="tight",
        )

        plt.close(fig)

        # ---------- Frequency summaries ----------

        duration_order = [
            "Short-term",
            "Medium-term",
            "Long-term",
            "Other",
            "Undetermined",
            "Manual revision",
            "Not applicable"
        ]

        scale_order = [
            "Pilot (initial)",
            "Pilot (iterative)",
            "Classroom-level",
            "Program-level",
            "Large-scale",
            "Other",
            "Undetermined",
            "Manual revision",
            "Not applicable"
        ]

        programming_order = [
            "Block-based",
            "Text-based",
            "Hybrid (block + text)",
            "Unplugged (no programming)",
            "Unplugged + Block-based",
            "Unplugged + Text-based",
            "Unplugged + Hybrid (block + text)",
            "Other",
            "Undetermined",
            "Manual revision",
            "Not applicable"
        ]

        columns_to_describe = [
            "Study_Type",
            "Type_of_Contribution",
            "Research_Methodology",
            "Disciplinary_Focus",
            "Learning_Focus",
            "Educational_Context",
            "Target_Population",
            "Intervention_Duration",
            "Scale_of_Intervention",
            "Study_Design",
            "Study_Design_2",
            "Geographical_Region",
            "Programming_Type (paradigm)",
            "Programming_Environment",
            "Programming_Tool/Language",
            "Programming_Kind",
            "Interaction_Modality",
            "Evidence_Strength",
            "Evaluation_Rigor",
            "Sample_Size_Reported",
            "Limitations_Reported",
            "Has_Rule_Violation",
            "Code_Based_Programming_Present",
            "CT_with_Programming",
            "Evaluation_Present",
            "Learning_Application_Present",
            "Programming_Tool_Grouped",
            "Programming_Environment_Grouped"
        ]

        frequency_tables = {}

        for col in columns_to_describe:
            if col in data.columns:
                frequency_tables[col] = frequency_table(data, col)

        # ---------- Geographical / Educational / Disciplinary table ----------

        region_order = [
            "Africa",
            "Asia",
            "Europe",
            "Latin America",
            "North America",
            "Oceania",
            "Multiple countries"
        ]

        education_order = [
            "Primary education",
            "Secondary education",
            "Higher education",
            "Multiple",
            "Not applicable"
        ]

        geo_context_discipline_table = (
            data
            .groupby([
                "Geographical_Region",
                "Educational_Context",
                "Disciplinary_Focus"
            ])
            .size()
            .reset_index(name="Number_of_papers")
        )

        geo_context_discipline_table["Region_order"] = (
            geo_context_discipline_table["Geographical_Region"]
            .apply(lambda x: region_order.index(x) if x in region_order else len(region_order) - 1)
        )

        geo_context_discipline_table["Education_order"] = (
            geo_context_discipline_table["Educational_Context"]
            .apply(lambda x: education_order.index(x) if x in education_order else len(education_order))
        )

        geo_context_discipline_table = (
            geo_context_discipline_table
            .sort_values([
                "Region_order",
                "Education_order",
                "Disciplinary_Focus"
            ])
            .drop(columns=["Region_order", "Education_order"])
        )

        total_row = pd.DataFrame({
            "Geographical_Region": ["TOTAL"],
            "Educational_Context": [""],
            "Disciplinary_Focus": [""],
            "Number_of_papers": [geo_context_discipline_table["Number_of_papers"].sum()]
        })

        geo_context_discipline_table = pd.concat(
            [geo_context_discipline_table, total_row],
            ignore_index=True
        )

        # ---------- Scale / Duration / Disciplinary Focus table ----------

        scale_duration_discipline_table = (
            data
            .groupby([
                "Scale_of_Intervention",
                "Intervention_Duration",
                "Disciplinary_Focus"
            ])
            .size()
            .reset_index(name="Number_of_papers")
        )

        scale_duration_discipline_table["Scale_order"] = (
            scale_duration_discipline_table["Scale_of_Intervention"]
            .apply(
                lambda x: scale_order.index(x)
                if x in scale_order
                else len(scale_order)
            )
        )

        scale_duration_discipline_table["Duration_order"] = (
            scale_duration_discipline_table["Intervention_Duration"]
            .apply(
                lambda x: duration_order.index(x)
                if x in duration_order
                else len(duration_order)
            )
        )

        scale_duration_discipline_table = (
            scale_duration_discipline_table
            .sort_values([
                "Scale_order",
                "Duration_order",
                "Disciplinary_Focus"
            ])
            .drop(columns=["Scale_order", "Duration_order"])
        )

        total_row = pd.DataFrame({
            "Scale_of_Intervention": ["TOTAL"],
            "Intervention_Duration": [""],
            "Disciplinary_Focus": [""],
            "Number_of_papers": [
                scale_duration_discipline_table["Number_of_papers"].sum()
            ]
        })

        scale_duration_discipline_table = pd.concat(
            [scale_duration_discipline_table, total_row],
            ignore_index=True
        )

        # ---------- Educational Context / Disciplinary Focus / Programming Paradigm ----------

        education_programming_discipline_table = (
            data
            .groupby([
                "Educational_Context",
                "Disciplinary_Focus",
                "Programming_Type (paradigm)"
            ])
            .size()
            .reset_index(name="Number_of_papers")
        )

        education_programming_discipline_table["Education_order"] = (
            education_programming_discipline_table["Educational_Context"]
            .apply(
                lambda x: education_order.index(x)
                if x in education_order
                else len(education_order)
            )
        )

        education_programming_discipline_table["Programming_order"] = (
            education_programming_discipline_table["Programming_Type (paradigm)"]
            .apply(
                lambda x: programming_order.index(x)
                if x in programming_order
                else len(programming_order)
            )
        )

        education_programming_discipline_table = (
            education_programming_discipline_table
            .sort_values([
                "Education_order",
                "Disciplinary_Focus",
                "Programming_order"
            ])
            .drop(columns=[
                "Education_order",
                "Programming_order"
            ])
        )

        total_row = pd.DataFrame({
            "Educational_Context": ["TOTAL"],
            "Disciplinary_Focus": [""],
            "Programming_Type (paradigm)": [""],
            "Number_of_papers": [
                education_programming_discipline_table["Number_of_papers"].sum()
            ]
        })

        education_programming_discipline_table = pd.concat(
            [education_programming_discipline_table, total_row],
            ignore_index=True
        )
        # ---------- Cross-tabs ----------

        cross_tables = {
            "study_type_vs_evidence": pd.crosstab(
                data["Study_Type"], data["Evidence_Strength"], margins=True, margins_name="TOTAL"
            ),
            "study_type_vs_methodology": pd.crosstab(
                data["Study_Type"], data["Research_Methodology"], margins=True, margins_name="TOTAL"
            ),
            "learning_focus_vs_programming_type": pd.crosstab(
                data["Learning_Focus"], data["Programming_Type (paradigm)"], margins=True, margins_name="TOTAL"
            ),
            "education_vs_programming": pd.crosstab(
                data["Educational_Context"], data["Programming_Type (paradigm)"], margins=True, margins_name="TOTAL"
            ),
            "contribution_vs_rigor": pd.crosstab(
                data["Type_of_Contribution"], data["Evaluation_Rigor"], margins=True, margins_name="TOTAL"
            ),
            "region_vs_context": pd.crosstab(
                data["Geographical_Region"], data["Educational_Context"], margins=True, margins_name="TOTAL"
            ),
            "programming_present_vs_learning_focus": pd.crosstab(
                data["Learning_Focus"], data["Code_Based_Programming_Present"], margins=True, margins_name="TOTAL"
            ),
            "ct_with_programming_vs_evidence": pd.crosstab(
                data["CT_with_Programming"], data["Evidence_Strength"], margins=True, margins_name="TOTAL"
            ),
            "programming_type_vs_rigor": pd.crosstab(
                data["Programming_Type (paradigm)"], data["Evaluation_Rigor"], margins=True, margins_name="TOTAL"
            ),
            "rule_violation_vs_study_type": pd.crosstab(
                data["Has_Rule_Violation"], data["Study_Type"], margins=True, margins_name="TOTAL"
            ),
            "rule_violation_vs_rigor": pd.crosstab(
                data["Has_Rule_Violation"], data["Evaluation_Rigor"], margins=True, margins_name="TOTAL"
            ),
            "contribution_vs_evidence": pd.crosstab(
                data["Type_of_Contribution"], data["Evidence_Strength"], margins=True, margins_name="TOTAL"
            )
        }

        if not rule_issues_df.empty:
            rule_summary = rule_issues_df["Rule_ID"].value_counts().reset_index()
            rule_summary.columns = ["Rule_ID", "Count"]
            rule_summary["Percentage_of_papers"] = (
                rule_summary["Count"] / len(data) * 100
            ).round(2)
        else:
            rule_summary = pd.DataFrame(columns=["Rule_ID", "Count", "Percentage_of_papers"])


        # ---------- Heatmap: Educational Context vs Programming Paradigm ----------

        heatmap_excluded_values = [
            "Undetermined",
            "Not applicable",
            "Manual revision",
            "",
        ]

        heatmap_data_clean = data[
            ~data["Educational_Context"].fillna("").astype(str).str.strip().isin(heatmap_excluded_values)
            &
            ~data["Programming_Type (paradigm)"].fillna("").astype(str).str.strip().isin(heatmap_excluded_values)
        ].copy()

        education_vs_programming_heatmap = pd.crosstab(
            heatmap_data_clean["Educational_Context"],
            heatmap_data_clean["Programming_Type (paradigm)"]
        )

 
        education_order_heatmap = [
            "Primary education",
            "Secondary education",
            "Higher education",
            "Teacher education",
            "Informal learning",
            "Mixed / multiple levels",
        ]

        programming_order_heatmap = [
            "Block-based",
            "Text-based",
            "Hybrid (block + text)",
            "Unplugged (no programming)",
            "Unplugged + Block-based",
            "Unplugged + Text-based",
            "Unplugged + Hybrid (block + text)",
            "Other",

        ]

        education_vs_programming_heatmap = (
            education_vs_programming_heatmap
            .reindex(
                index=education_order_heatmap,
                columns=programming_order_heatmap,
                fill_value=0,
            )
        )

        fig, ax = plt.subplots(figsize=(12, 6))

        sns.heatmap(
            education_vs_programming_heatmap,
            annot=True,
            fmt="d",
            cmap="YlGnBu",
            linewidths=0.5,
            linecolor="white",
            cbar_kws={"label": "Number of studies"},
            ax=ax
        )

        ax.set_title(
            "Educational Context Distribution vs Programming Paradigm"
        )
        ax.set_xlabel("Programming Paradigm")
        ax.set_ylabel("Educational Context")

        ax.set_ylabel("Educational Context", labelpad=20)

        ax.tick_params(axis="y", pad=10)

        for label in ax.get_yticklabels():
            label.set_rotation(0)
            label.set_horizontalalignment("right")

        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()

        st.pyplot(fig)

        fig.savefig(
            figures_dir / "education_vs_programming_heatmap.png",
            dpi=300,
            bbox_inches="tight"
        )

        fig.savefig(
            figures_dir / "education_vs_programming_heatmap.pdf",
            bbox_inches="tight"
        )

        plt.close(fig)
    
        st.header("Educational context vs Programming type – framework papers")

        framework_table = pd.crosstab(
            framework_characterization_df["Educational_Context"],
            framework_characterization_df["Programming_Type (paradigm)"]
        )

        excluded_values = [
            "",
            "Undetermined",
            "Not applicable",
            "Manual revision",
        ]

        framework_table = framework_table.drop(
            index=[
                x for x in excluded_values
                if x in framework_table.index
            ],
            errors="ignore"
        )

        framework_table = framework_table.drop(
            columns=[
                x for x in excluded_values
                if x in framework_table.columns
            ],
            errors="ignore"
        )

        framework_table = (
            framework_table
            .reindex(
                index=education_order_heatmap,
                columns=programming_order_heatmap,
                fill_value=0,
            )
        )

        fig, ax = plt.subplots(figsize=(12, 6))

        sns.heatmap(
            framework_table,
            annot=True,
            fmt="d",
            cmap="YlGnBu",
            linewidths=0.5,
            linecolor="white",
            cbar_kws={"label": "Number of framework papers"},
            ax=ax,
        )

        ax.set_title(
            "Educational Context Distribution vs Programming Paradigm – Framework Papers"
        )

        ax.set_xlabel("Programming Paradigm")
        ax.set_ylabel("Educational Context", labelpad=20)

        plt.xticks(rotation=45, ha="right")
        plt.yticks(rotation=0)

        plt.tight_layout()
        plt.subplots_adjust(left=0.25)

        st.pyplot(fig)

        fig.savefig(
            figures_dir / "educational_context_vs_programming_framework_heatmap.png",
            dpi=300,
            bbox_inches="tight",
        )

        fig.savefig(
            figures_dir / "educational_context_vs_programming_framework_heatmap.pdf",
            bbox_inches="tight",
        )

        plt.close(fig)
        
        # ---------- Export Excel ----------



        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

            rule_summary.to_excel(writer, sheet_name="rule_summary", index=False)
            data.to_excel(writer, sheet_name="data_cleaned", index=False)

            if not rule_issues_df.empty:
                rule_issues_df.to_excel(writer, sheet_name="rule_issues", index=False)

            for name, table in frequency_tables.items():
                table.to_excel(writer, sheet_name=clean_sheet_name(name), index=False)

            for name, table in cross_tables.items():
                table.to_excel(writer, sheet_name=clean_sheet_name(name))
            
            geo_context_discipline_table.to_excel(
                writer,
                sheet_name="Geo_Context_Discipline",
                index=False
            )
            scale_duration_discipline_table.to_excel(
                writer,
                sheet_name="Scale_Duration_Discipline",
                index=False
            )
            education_programming_discipline_table.to_excel(
                writer,
                sheet_name="Edu_Discipline_Programming",
                index=False
            )

        summary = build_final_review_summary(
            project_root=PROJECT_ROOT,
            output_dir=OUTPUT_DIR_6,
            total_characterization_papers=total_characterization_papers,
            total_mapped_papers=total_mapped_papers,
            final_analysis_papers=len(data),
            missing_ids=missing_ids,
        )



        st.subheader("Final PRISMA summary")
        st.dataframe(summary["prisma_summary"], use_container_width=True)

        st.subheader("Pipeline cost summary")
        st.dataframe(summary["cost_summary"], use_container_width=True)
        if not summary["warnings"].empty:

            st.warning(
                "Some expected files were not found."
            )

            st.dataframe(
                summary["warnings"],
                use_container_width=True,
            )




        st.subheader("Rule summary")
        st.dataframe(rule_summary)

        st.subheader("Filtered dataset preview")
        st.dataframe(data.head())

        st.subheader("Table: Geographical distribution, educational context, and disciplinary focus")
        st.dataframe(geo_context_discipline_table, use_container_width=True)

        st.subheader("Table: Scale of intervention, intervention duration, and disciplinary focus")
        st.dataframe(scale_duration_discipline_table, use_container_width=True)

        st.subheader(
            "Table: Educational context, disciplinary focus, and programming paradigm"
        )



        st.dataframe(
            education_programming_discipline_table,
            use_container_width=True
        )

        st.success(
            f"""
        Analysis completed successfully.

        📁 Stage 6 output folder:
        {OUTPUT_DIR_6}

        Generated file:
        • {output_path.name}

        Figures saved in:
        {figures_dir}
        """
        )


        st.success(f"Final review summary saved to: {summary['output_path']}")

        st.header("CT–Programming relationship analysis")

        pairing_summary = run_pairing_analysis(
            paper_level_file=paper_level_file,
            pairing_file=pair_level_file,
            grouping_file=tools_grouping_file,
            output_dir=OUTPUT_DIR_6,
            output_name="programming_elements_summary.xlsx",
        )

        display_pairing_analysis_ordered(pairing_summary)

        # --------------------------------------------------
        # Visualization data preparation
        # --------------------------------------------------

        plot_data = data.copy()

        framework_plot_data = plot_data[
            plot_data["Is_Framework_Paper"] == True
        ].copy()

        plot_data = plot_data.replace(
            to_replace=r"^Other:.*",
            value="Other",
            regex=True
        )


        study_design_order = [
            "Cross-sectional",
            "Pre–post",
            "Longitudinal",
            "Other",
            "Undetermined",
            "Manual revision",
            "Not applicable"
        ]


        plot_data["Intervention_Duration"] = plot_data["Intervention_Duration"].apply(
            lambda x: group_not_allowed(x, duration_order)
        )

        plot_data["Scale_of_Intervention"] = plot_data["Scale_of_Intervention"].apply(
            lambda x: group_not_allowed(x, scale_order)
        )

        plot_data["Study_Design"] = plot_data["Study_Design"].apply(
            lambda x: group_not_allowed(x, study_design_order)
        )

        plot_data["Programming_Tool_Grouped"] = plot_data["Programming_Tool/Language"].apply(
            lambda x: classify_tool_language(x, tools_grouping)
        )



        # --------------------------------------------------
        # Bar charts
        # --------------------------------------------------

        st.subheader("Visualizations")
        st.markdown("## Frequency bar charts")

        st.header("General characterization of included studies")

        st.markdown("## Characterization comparison: all papers vs framework papers")

        comparison_bar_specs = [
            (
                "Educational_Context",
                "Educational Context: all papers vs framework papers",
                "Educational context",
                None,
            ),
            (
                "Study_Type",
                "Study Type: all papers vs framework papers",
                "Study type",
                None,
            ),
            (
                "Research_Methodology",
                "Research Methodology: all papers vs framework papers",
                "Research methodology",
                None,
            ),
            (
                "Programming_Type (paradigm)",
                "Programming Paradigm: all papers vs framework papers",
                "Programming paradigm",
                programming_order,
            ),
                (
                "Programming_Tool_Grouped",
                "Programming Tool/Language: all papers vs framework papers",
                "Programming tool/language",
                tools_order,
            ),
        ]

        for column, title, xlabel, preferred_order in comparison_bar_specs:
            if column in plot_data.columns and column in framework_plot_data.columns:
                st.markdown(f"### {title}")
                grouped_characterization_bar(
                    full_data=plot_data,
                    framework_data=framework_plot_data,
                    column=column,
                    title=title,
                    xlabel=xlabel,
                    figures_dir=figures_dir,
                    filename=f"grouped_{safe_filename(column)}",
                    preferred_order=preferred_order,
                )

        st.markdown("## Main descriptive bar charts")

        main_bar_specs = [
            ("Geographical_Region", "Geographical Region Distribution", "Region", None),
            ("Educational_Context", "Educational Context Distribution", "Educational context", None),
            ("Target_Population", "Target Population Distribution", "Target population", None),
            ("Study_Type", "Study Type Distribution", "Study type", None),
            ("Type_of_Contribution", "Type of Contribution Distribution", "Contribution type", None),
            ("Research_Methodology", "Research Methodology Distribution", "Research methodology", None),
            ("Study_Design", "Study Design Distribution", "Study design", study_design_order),
            ("Programming_Type (paradigm)", "Programming Type Paradigm Distribution", "Programming type", programming_order),
            ("Programming_Tool_Grouped", "Programming Tool/Language Distribution – Grouped Values", "Programming tool/language", tools_order),
            ("Programming_Environment_Grouped", "Programming Environment Distribution", "Programming environment", None),
            ("Programming_Kind", "Programming Kind Distribution", "Programming kind", None),
            ("Intervention_Duration", "Intervention Duration Distribution", "Intervention duration", duration_order),
            ("Scale_of_Intervention", "Scale of Intervention Distribution", "Scale of intervention", scale_order),
        ]

        for column, title, xlabel, preferred_order in main_bar_specs:
            if column in plot_data.columns:
                st.markdown(f"### {title}")
                make_bar_chart(
                    plot_data[column],
                    title,
                    xlabel,
                    f"bar_{safe_filename(column)}",
                    figures_dir,
                    preferred_order=preferred_order,
                )
                if column == "Programming_Tool_Grouped":
                    st.markdown("### Programming Tool/Language: all papers vs framework papers")

                    grouped_characterization_bar(
                        full_data=plot_data,
                        framework_data=framework_plot_data,
                        column="Programming_Tool_Grouped",
                        title="Programming Tool/Language: all papers vs framework papers",
                        xlabel="Programming tool/language",
                        figures_dir=figures_dir,
                        filename="grouped_programming_tool_language",
                        preferred_order=tools_order,
                    )
                if column == "Programming_Type (paradigm)":
                    st.markdown("### Programming Paradigm: all papers vs framework papers")

                    grouped_characterization_bar(
                        full_data=plot_data,
                        framework_data=framework_plot_data,
                        column="Programming_Type (paradigm)",
                        title="Programming Paradigm: all papers vs framework papers",
                        xlabel="Programming paradigm",
                        figures_dir=figures_dir,
                        filename="grouped_programming_paradigm",
                        preferred_order=programming_order,
                    )

                if column == "Educational_Context":
                    st.markdown("### Educational_Context: all papers vs framework papers")

                    grouped_characterization_bar(
                        full_data=plot_data,
                        framework_data=framework_plot_data,
                        column="Educational_Context",
                        title="Educational_Context: all papers vs framework papers",
                        xlabel="Educational_Context",
                        figures_dir=figures_dir,
                        filename="Educational_Context",
                        preferred_order=programming_order,
                    )

                if column == "Programming_Tool_Grouped":
                    st.markdown("### Programming_Tool_Grouped: all papers vs framework papers")

                    grouped_characterization_bar(
                        full_data=plot_data,
                        framework_data=framework_plot_data,
                        column="Programming_Tool_Grouped",
                        title="Programming_Tool_Groupedt: all papers vs framework papers",
                        xlabel="Programming_Tool_Grouped",
                        figures_dir=figures_dir,
                        filename="Programming_Tool_Grouped",
                        preferred_order=programming_order,
                    )


        with st.expander("Technical / quality-control bar charts"):
            technical_bar_specs = [
                ("Evidence_Strength", "Evidence Strength Distribution", "Evidence strength", None),
                ("Evaluation_Rigor", "Evaluation Rigor Distribution", "Evaluation rigor", None),
                ("Sample_Size_Reported", "Sample Size Reported Distribution", "Sample size reported", None),
                ("Limitations_Reported", "Limitations Reported Distribution", "Limitations reported", None),
                ("Has_Rule_Violation", "Rule Violation Distribution", "Rule violation", None),
                ("Code_Based_Programming_Present", "Code-Based Programming Presence Distribution", "Code-based programming present", None),
                ("CT_with_Programming", "CT with Programming Distribution", "CT with programming", None),
                ("Evaluation_Present", "Evaluation Presence Distribution", "Evaluation present", None),
                ("Learning_Application_Present", "Learning Application Presence Distribution", "Learning application present", None),
            ]

            for column, title, xlabel, preferred_order in technical_bar_specs:
                if column in plot_data.columns:
                    st.markdown(f"### {title}")
                    make_bar_chart(
                        plot_data[column],
                        title,
                        xlabel,
                        f"bar_{safe_filename(column)}",
                        figures_dir,
                        preferred_order=preferred_order,
                    )


        # --------------------------------------------------
        # Geographical map
        # --------------------------------------------------

        st.markdown("## Geographical distribution map")

        region_counts = plot_data["Geographical_Region"].value_counts().to_dict()

        world_regions = pd.DataFrame({
            "Region": ["North America", "Latin America", "Europe", "Africa", "Asia", "Oceania"],
            "Count": [
                region_counts.get("North America", 0),
                region_counts.get("Latin America", 0),
                region_counts.get("Europe", 0),
                region_counts.get("Africa", 0),
                region_counts.get("Asia", 0),
                region_counts.get("Oceania", 0),
            ],
            "lat": [50, -15, 54, 0, 34, -25],
            "lon": [-100, -60, 15, 20, 90, 135],
        })

        fig_map = px.scatter_geo(
            world_regions,
            lat="lat",
            lon="lon",
            size="Count",
            size_max=70,
            text="Count",
            hover_name="Region",
            hover_data={"Count": True, "lat": False, "lon": False},
            projection="natural earth",
            title="Geographical Distribution of Studies by Region"
        )

        fig_map.update_traces(

            text=world_regions["Count"],
            textposition="middle center",

            mode="markers+text",

            marker=dict(
                color="#1192b5",
                opacity=0.8,
                line=dict(width=1, color="white")
            ),

            textfont=dict(
                color="white",
                size=13
            )
        )
        label_offsets = {
            "North America": (0, 15),
            "Latin America": (0, -10),
            "Europe": (12, 12),
            "Asia": (14, 0),
            "Africa": (-10, -2),
            "Oceania": (12, -6)
        }
        for _, row in world_regions.iterrows():

            lon_offset, lat_offset = label_offsets.get(
                row["Region"],
                (0, 8)
            )

            fig_map.add_scattergeo(

                lon=[row["lon"] + lon_offset],
                lat=[row["lat"] + lat_offset],

                text=[row["Region"]],

                mode="text",

                showlegend=False,

                textfont=dict(
                    size=12,
                    color="black"
                )
            )

        fig_map.update_geos(
            showland=True,
            landcolor="rgb(235, 235, 235)",
            showcountries=False,
            showcoastlines=False,
            showocean=True,
            oceancolor="rgb(200, 230, 255)",
            showlakes=False,
            lataxis_showgrid=False,
            lonaxis_showgrid=False,
            projection_scale=1.15,

            center=dict(
            lat=20,
            lon=10
            ),
            lataxis_range=[-45, 75],
            lonaxis_range=[-170, 170]
        )


        fig_map.update_layout(
            height=500,
            margin=dict(l=0, r=0, t=50, b=0)
        )

        st.plotly_chart(fig_map, width="stretch")
        fig_map.write_html(str(figures_dir / "geo_map.html"))

        # --------------------------------------------------
        # Heatmaps
        # --------------------------------------------------

        st.markdown("## Heatmaps")

        heatmap_vars = {
            "Educational context vs Research methodology": (
                "Educational_Context",
                "Research_Methodology",
            ),
            "Educational context vs Study type": (
                "Educational_Context",
                "Study_Type",
            ),
            "Educational context vs Study design": (
                "Educational_Context",
                "Study_Design",
            ),
            "Educational context vs Programming type": (
                "Educational_Context",
                "Programming_Type (paradigm)",
            ),
            "Geographical region vs Educational context": (
                "Geographical_Region",
                "Educational_Context",
            ),
            "Geographical region vs Programming type": (
                "Geographical_Region",
                "Programming_Type (paradigm)",
            ),
            "Geographical region vs Study type": (
                "Geographical_Region",
                "Study_Type",
            ),
            "Type of contribution vs Research methodology": (
                "Type_of_Contribution",
                "Research_Methodology",
            ),
        }

        for title, (row_col, col_col) in heatmap_vars.items():

            if row_col not in plot_data.columns or col_col not in plot_data.columns:
                st.warning(f"Skipping {title}: missing columns.")
                continue

            invalid_values = {
                "Undetermined",
                "Not applicable",
                "Not Applicable",
                "Manual revision",
                "Not reported",
                "Unknown",
                "N/A",
                "",
            }

            heatmap_data = plot_data[
                ~plot_data[row_col].astype(str).str.strip().isin(invalid_values)
                &
                ~plot_data[col_col].astype(str).str.strip().isin(invalid_values)
            ].copy()

            table = pd.crosstab(
                heatmap_data[row_col],
                heatmap_data[col_col]
            )

            if table.empty:
                st.warning(f"Skipping {title}: empty table after removing undetermined/not applicable values.")
                continue

            st.markdown(f"### {title}")

            fig, ax = plt.subplots(figsize=(14, 7))
            sns.heatmap(
                table,
                annot=True,
                fmt="d",
                cmap="YlGnBu",
                linewidths=0.5,
                linecolor="white",
                cbar_kws={"label": "Number of studies"},
                ax=ax
            )

            ax.set_title(title)
            ax.set_xlabel(col_col.replace("_", " "))
            ax.set_ylabel(row_col.replace("_", " "), labelpad=20)

            ax.tick_params(axis="y", pad=10)

            for label in ax.get_yticklabels():
                label.set_rotation(0)
                label.set_horizontalalignment("right")

            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            plt.subplots_adjust(left=0.22)

            for label in ax.get_yticklabels():
                label.set_rotation(0)

            fig.tight_layout()

            st.pyplot(fig)

            fig_name = f"heatmap_{safe_filename(title)}"
            fig.savefig(figures_dir / f"{fig_name}.png", dpi=300, bbox_inches="tight")
            fig.savefig(figures_dir / f"{fig_name}.pdf", bbox_inches="tight")

            plt.close(fig)

        # --------------------------------------------------
        # Sunburst: Selected regions only
        # --------------------------------------------------

        st.markdown("## Sunburst: Region, educational context, and programming paradigm")

        selected_regions = [
            "Asia",
            "North America",
            "Latin America",
            "Europe"
        ]

        sunburst_data = plot_data[
            plot_data["Geographical_Region"].isin(selected_regions)
        ][
            [
                "Geographical_Region",
                "Educational_Context",
                "Programming_Type (paradigm)"
            ]
        ].copy()

        sunburst_data = sunburst_data.fillna("Not reported")

        sunburst_counts = (
            sunburst_data
            .groupby([
                "Geographical_Region",
                "Educational_Context",
                "Programming_Type (paradigm)"
            ])
            .size()
            .reset_index(name="Count")
        )

        fig_sunburst = px.sunburst(
            sunburst_counts,
            path=[
                "Geographical_Region",
                "Educational_Context",
                "Programming_Type (paradigm)"
            ],
            values="Count",
            title="Regional and Educational Distribution of Programming Paradigms"
        )

        fig_sunburst.update_layout(
            height=750,
            margin=dict(t=60, l=10, r=10, b=10)
        )

        st.plotly_chart(fig_sunburst, width="stretch")

        fig_sunburst.write_html(
            str(figures_dir / "geo_context_programming_sunburst.html")
        )

        st.markdown("## Sunburst: Region, educational context, and disciplinary focus")

        sunburst_context_discipline = plot_data[
            [
                "Geographical_Region",
                "Educational_Context",
                "Disciplinary_Focus"
            ]
        ].copy()

        sunburst_context_discipline = sunburst_context_discipline.fillna("Not reported")

        sunburst_context_discipline_counts = (
            sunburst_context_discipline
            .groupby([
                "Geographical_Region",
                "Educational_Context",
                "Disciplinary_Focus"
            ])
            .size()
            .reset_index(name="Count")
        )

        fig_sunburst_context_discipline = px.sunburst(
            sunburst_context_discipline_counts,
            path=[
                "Geographical_Region",
                "Educational_Context",
                "Disciplinary_Focus"
            ],
            values="Count",
            title="Regional and Educational Distribution by Disciplinary Focus"
        )

        st.plotly_chart(fig_sunburst_context_discipline, width="stretch")

        fig_sunburst_context_discipline.write_html(
        str(figures_dir / "geo_context_disciplinary_focus_sunburst.html")
        )

        # --------------------------------------------------
        # Radar chart
        # --------------------------------------------------

        st.markdown("## Radar charts by variable")

        radar_variables = {
            "Type_of_Contribution": "Contribution Type",
            "Educational_Context": "Educational Context",
            "Disciplinary_Focus": "Disciplinary Focus",
        }

        for var, label in radar_variables.items():

            if var not in plot_data.columns:
                continue

            st.markdown(f"### {label} by region")

            radar_rows = []

            for region, region_df in plot_data.groupby("Geographical_Region"):

                total = len(region_df)
                if total == 0:
                    continue

                categories = region_df[var].dropna().unique()

                row = {"Geographical_Region": region}

                for cat in categories:
                    row[cat] = round((region_df[var] == cat).sum() / total * 100, 1)

                radar_rows.append(row)

            radar_df = pd.DataFrame(radar_rows)

            if radar_df.empty:
                continue

            radar_long = radar_df.melt(
                id_vars="Geographical_Region",
                var_name="Category",
                value_name="Percentage"
            )

            fig = px.line_polar(
                radar_long,
                r="Percentage",
                theta="Category",
                color="Geographical_Region",
                line_close=True,
                range_r=[0, 100],
                title=f"{label} distribution by region"
            )

            fig.update_traces(fill="toself")

            st.plotly_chart(fig, width="stretch")

            fig.write_html(str(figures_dir / f"radar_{safe_filename(var)}.html"))

        st.success(f"Figures saved in: {figures_dir}")


if __name__ == "__main__":
    main()
