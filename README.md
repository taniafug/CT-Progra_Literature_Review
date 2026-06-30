# CT–Programming Scoping Review Pipeline

A Streamlit-based pipeline supporting a scoping review on how Computational Thinking (CT) is operationalized through programming in educational research.

The pipeline combines configurable Excel templates, Large Language Model (LLM)-assisted analysis, manual validation, and automated reporting to support a transparent and reproducible review process.

> **Important**
>
> This tool supports the review process but **does not replace the researcher**. All inclusion, exclusion, coding, characterization, and interpretation decisions remain the responsibility of the researcher.

---

## Features

- Six-stage Streamlit workflow
- Configurable Excel-based screening and extraction rules
- LLM-assisted title, abstract, and full-text analysis
- Manual validation support
- CT–programming pairing extraction
- Study characterization
- Automated figures and summary tables
- PRISMA-style summaries
- Checkpoint recovery for long-running analyses
- Interactive User Guide integrated into the application

---

## Workflow

The pipeline supports a six-stage review workflow.

| Stage | Description |
|------|-------------|
| **Stage 1** | Import database exports, standardize records, remove duplicates, and prepare the abstract screening dataset. |
| **Stage 2** | Screen titles and abstracts using configurable rules and LLM-assisted classification. |
| **Stage 3** | Read full-text PDFs, identify relevant sections, perform full-text screening, and prepare papers for pairing. |
| **Stage 4** | Extract CT and programming elements, normalize terminology, and generate paper-level and pair-level mappings. |
| **Stage 5** | Characterize included studies, including educational context, methodology, programming tools, interventions, and assessment details. |
| **Stage 6** | Generate figures, summary tables, heatmaps, PRISMA-style summaries, and final review statistics. |

---

## Repository Structure

```text
.
├── Home.py
├── pages/
│   ├── 0_User_Guide.py
│   ├── 1_Stage_1_Import_Filtering.py
│   ├── 2_Stage_2_Abstract_Screening.py
│   ├── 3_Stage_3_Fulltext_Initial_Pairing.py
│   ├── 4_Stage_4_Full_Pairing.py
│   ├── 5_Stage_5_Characterization.py
│   └── 6_Stage_6_Analysis.py
├── stage_1_import_filtering/
├── stage_2_abstract_screening/
├── stage_3_fulltext_initial_pairing/
├── stage_4_full_pairing/
├── stage_5_characterization/
├── stage_6_results_analysis/
├── templates/
├── documentation/
│   ├── User_Guide.pdf
│   └── Search_Queries.pdf
├── requirements.txt
├── DEVELOPER_GUIDE.md
├── CITATION.cff
├── LICENSE
└── README.md
```

Each stage contains its own `input`, `output`, and, when required, `temp` folders. Long-running LLM stages automatically generate checkpoints that allow interrupted executions to be resumed.

---

## Installation

Create and activate a Python virtual environment.

```bash
python -m venv llm_project_env
source llm_project_env/bin/activate
```

Install the required packages.

```bash
pip install -r requirements.txt
```

Main dependencies include:

- Streamlit
- pandas
- openpyxl
- OpenAI Python SDK
- PyMuPDF
- pypdf
- matplotlib
- plotly

---

## OpenAI API Key

Stages using LLMs require an OpenAI API key.

The key can either be entered directly in the Streamlit sidebar or stored locally inside a `.env` file.

```text
OPENAI_API_KEY=your_api_key_here
```

The `.env` file should **never** be committed to the repository.

---

## Running the Application

From the project root, execute:

```bash
streamlit run Home.py
```

The application provides an integrated navigation menu containing the User Guide and all six pipeline stages.

Each stage can be executed sequentially or independently using manually supplied input files.

---

## Configuration Files

The pipeline is designed to be configurable without modifying the source code.

Most screening rules, prompts, normalization tables, extraction schemas, and characterization templates are stored as Excel files.

Examples include:

```text
templates/stage_1_rules.xlsx
templates/stage_2_rules.xlsx
templates/stage_3_rules.xlsx
templates/stage4_pairing_rules.xlsx
templates/stage5_characterization_rules.xlsx
templates/Analysis_details.xlsx
```

The integrated User Guide provides download buttons for the available templates.

---

## Documentation

The repository includes additional documentation.

| File | Description |
|------|-------------|
| **User_Guide.pdf** | Complete user manual describing every pipeline stage. |
| **Search_Queries.pdf** | Complete search strategy, including the search query used for each academic database. |
| **DEVELOPER_GUIDE.md** | Technical documentation for developers and contributors. |
| **README.md** | General overview and installation instructions. |

---

## Full-Text PDF Repository

Stages **3** and **4** require access to the full-text PDF files of the selected studies.

**The PDF files are intentionally not included in this repository.**

Most scientific publications are protected by copyright and cannot legally be redistributed through GitHub, Zenodo, or this repository unless their licenses explicitly permit redistribution.

Users should obtain the full-text articles through:

- institutional subscriptions,
- publisher websites,
- or legally available Open Access repositories.

Once obtained, the PDFs should be placed inside the local repository, for example:

```text
papers_repository/
└── pdfs/
```

The pipeline expects this local PDF repository during full-text screening and CT–programming pairing.

---

## Stage Outputs

Each stage generates its own outputs and prepares the input required for the following stage whenever applicable.

Typical Stage 3 outputs include:

```text
stage_3_results.xlsx
stage_3_author_validation.xlsx
stage_3_costs.xlsx
stage_3_no_full_text_available.xlsx
stage_3_to_stage_4_input.xlsx
```

Stage 4 generates:

- paper-level mappings
- pair-level mappings
- normalization summaries
- author validation files
- cost reports

Stage 5 generates characterization outputs.

Stage 6 generates:

- publication-ready figures
- heatmaps
- summary tables
- PRISMA-style summaries
- final review statistics

---

## Testing

The complete workflow does **not** need to be executed during testing.

Recommended strategy:

- Stage 1 can be tested using the provided example database exports.
- Stage 6 can be tested independently because it only requires outputs generated by previous stages.
- For complete end-to-end testing, use a small subset of studies instead of the full corpus.

---

## Checkpoints and Recovery

Long-running LLM stages automatically generate checkpoints.

If execution is interrupted, restart the same stage and enable the resume option whenever available.

Do not delete temporary checkpoint files until the stage has successfully completed.

---

## Reproducibility

This repository contains everything required to reproduce the review workflow, except the copyrighted full-text publications.

Included:

- Source code
- Configuration templates
- Prompt templates
- Analysis scripts
- Example input files
- Example output files
- Documentation

Excluded:

- Full-text scientific publications
- Local API credentials
- Temporary files
- Checkpoints

The original full-text scientific publications are **not included** because they are protected by copyright and cannot generally be redistributed. Users should obtain these articles through legitimate sources before executing the full-text stages of the pipeline.

---

## Files Not Intended for Version Control

The following files should normally remain outside version control.

```text
.env
**/temp/
**/__pycache__/
*.pyc
*.Zone.Identifier
**/papers_repository/
**/pdfs/
```

---

## Citation

If you use this pipeline in academic work, please cite the repository using the information provided in **CITATION.cff**.

---

## License

This project is distributed under the license provided in the `LICENSE` file.

---

## Acknowledgements

This tool was developed to support a reproducible scoping review on Computational Thinking and Programming in Computing Education.

It combines configurable rule-based processing, Large Language Models, and researcher-driven validation to improve transparency, reproducibility, and efficiency while preserving human oversight throughout the review process.
