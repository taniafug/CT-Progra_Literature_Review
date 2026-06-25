
import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="User Guide",
    page_icon="📖",
    layout="wide",
)

st.title("📖 User Guide")

st.markdown("""
This guide explains how to use the Scoping Review CT Programming Tool,
including the purpose of each stage, required inputs, generated outputs,
and common troubleshooting steps.
""")

# ============================================================
# Welcome
# ============================================================

st.header("Welcome")

st.markdown("""
### What the tool does

The Scoping Review CT Programming Tool supports the execution of a
semi-automated scoping review workflow using Large Language Models (LLMs).

Researchers remain responsible for all final decisions.

### Overall workflow

The workflow is organized into six stages:

1. Title Screening
2. Abstract Screening
3. Full-Text Screening
4. Data Extraction
5. Characterization
6. Analysis and Visualization

### Requirements

- Python environment configured
- OpenAI API key
- PDF repository available
- Input Excel files for each stage
""")

# ============================================================
# Download Templates
# ============================================================

st.header("📥 Download Templates and Configuration Files")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"

if TEMPLATES_DIR.exists():

    template_files = sorted(TEMPLATES_DIR.glob("*.xlsx"))

    if template_files:

        for template in template_files:

            with open(template, "rb") as file:

                st.download_button(
                    label=f"📥 {template.stem}",
                    data=file,
                    file_name=template.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    else:
        st.info("No template files were found.")

else:
    st.warning("Templates folder not found.")

# ============================================================
# Workflow Overview
# ============================================================

st.header("Workflow Overview")

st.markdown("""
Stage 1 – Title Screening
        ↓
Stage 2 – Abstract Screening
        ↓
Stage 3 – Full-Text Screening
        ↓
Stage 4 – Data Extraction
        ↓
Stage 5 – Characterization
        ↓
Stage 6 – Analysis and Visualization
""")

# ============================================================
# Stage 1
# ============================================================

with st.expander("Stage 1 – Title Screening"):

    st.subheader("Purpose")
    st.write("Screen studies using title information only.")

    st.subheader("Required Inputs")
    st.markdown("""
    - Database export file
    - Title screening input file
    """)

    st.subheader("Generated Outputs")
    st.markdown("""
    - Screening results
    - Deduplicated records
    - Records passed to Stage 2
    - PRISMA summary
    """)

    st.subheader("Notes")
    st.write("Only selected records proceed to abstract screening.")

# ============================================================
# Stage 2
# ============================================================

with st.expander("Stage 2 – Abstract Screening"):

    st.subheader("Purpose")
    st.write("Screen studies using title and abstract.")

    st.subheader("Required Inputs")
    st.markdown("""
    - Abstract screening input workbook
    - stage2_abstract_screening_rules.xlsx
    """)

    st.subheader("Generated Outputs")
    st.markdown("""
    - Included
    - Excluded
    - Undetermined
    - Next_stage_included
    - PRISMA summary
    """)

    st.subheader("Notes")
    st.write(
        "Final decisions are controlled by deterministic screening rules."
    )

# ============================================================
# Stage 3
# ============================================================

with st.expander("Stage 3 – Full-Text Screening"):

    st.subheader("Purpose")
    st.write("Assess study eligibility using PDF full texts.")

    st.subheader("Required Inputs")
    st.markdown("""
    - Full-text screening input workbook
    - stage3_fulltext_screening_rules.xlsx
    - PDF repository
    """)

    st.subheader("PDF Repository")
    st.write(
        "PDF files must be available and linked correctly in the input file."
    )

    st.subheader("No Full Text Available")
    st.write(
        "Studies without accessible full texts can be reserved and stored in "
        "no_full_text_available.xlsx."
    )

    st.subheader("Meaning of Decisions")

    st.markdown("""
    **Included**
    - Meets all inclusion criteria.

    **Excluded**
    - Fails one or more inclusion criteria.

    **Undetermined**
    - Insufficient information for a reliable decision.
    """)

    st.subheader("PRISMA Counts")
    st.write(
        "The PRISMA summary includes screened papers and papers reserved "
        "because full text was unavailable."
    )

# ============================================================
# Stage 4
# ============================================================

with st.expander("Stage 4 – Data Extraction"):

    st.subheader("Purpose")
    st.write(
        "Extract study characteristics and review variables."
    )

    st.subheader("Required Inputs")
    st.markdown("""
    - Included papers from Stage 3
    - stage4_data_extraction_rules.xlsx
    """)

    st.subheader("Generated Outputs")
    st.markdown("""
    - Extraction results
    - Validation workbook
    - Cost report
    """)

    st.subheader("Validation Workbook")
    st.write(
        "A random sample is generated for manual validation."
    )

# ============================================================
# Stage 5
# ============================================================

with st.expander("Stage 5 – Characterization"):

    st.subheader("Purpose")
    st.write(
        "Characterize studies according to the review framework."
    )

    st.subheader("Required Inputs")
    st.markdown("""
    - Included studies
    - stage5_characterization_rules.xlsx
    - PDF repository
    """)

    st.subheader("Generated Outputs")
    st.markdown("""
    - characterization_results.xlsx
    - characterization_author_validation.xlsx
    - characterization_llm_costs.xlsx
    """)

    st.subheader("Checkpoint Recovery")
    st.write(
        "The stage creates checkpoints that allow interrupted runs to resume."
    )

# ============================================================
# Stage 6
# ============================================================

with st.expander("Stage 6 – Analysis and Visualization"):

    st.subheader("Purpose")
    st.write(
        "Generate final analyses, figures, and summary tables."
    )

    st.subheader("Required Inputs")
    st.markdown("""
    - Characterization results
    - stage6_framework_normalization.xlsx
    """)

    st.subheader("Generated Outputs")
    st.markdown("""
    - Summary tables
    - Figures
    - Pairwise framework analysis
    """)

# ============================================================
# Output Files Reference
# ============================================================

st.header("Output Files Reference")

st.table({
    "File": [
        "results.xlsx",
        "author_validation.xlsx",
        "llm_costs.xlsx",
        "no_full_text_available.xlsx",
        "characterization_results.xlsx",
    ],
    "Description": [
        "Main screening results",
        "Manual validation sample",
        "Token and cost report",
        "Papers unavailable for screening",
        "Characterization output",
    ],
})

# ============================================================
# Recovery and Troubleshooting
# ============================================================

st.header("Recovery and Troubleshooting")

with st.expander("Interrupted Execution"):
    st.write(
        "Restart the stage and resume from the latest checkpoint."
    )

with st.expander("Missing PDFs"):
    st.write(
        "Verify PDF paths and repository locations."
    )

with st.expander("API Errors"):
    st.write(
        "Verify the API key and internet connection."
    )

with st.expander("Empty Outputs"):
    st.write(
        "Check that the previous stage generated the expected input files."
    )

with st.expander("Resuming from Checkpoints"):
    st.write(
        "Do not delete files stored in the temp folder until the stage finishes."
    )
