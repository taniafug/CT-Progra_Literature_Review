# CT–Programming Scoping Review Pipeline

This repository contains a Streamlit pipeline used to support a scoping review on how computational thinking (CT) is connected to programming in educational research. The tool combines rule-based processing, Excel configuration files, LLM-assisted screening and extraction, manual validation outputs, and final analysis scripts.

The pipeline is meant to support the review process, not replace the researcher. All final screening, coding, and interpretation decisions remain the responsibility of the researcher.

## What the pipeline does

The tool supports a six-stage workflow:

1. **Import and filtering**: imports database exports, standardizes records, removes duplicates, and prepares the abstract screening input.
2. **Abstract screening**: screens titles and abstracts using configurable rules and LLM-assisted classification.
3. **Full-text screening**: reads PDFs, extracts relevant sections, applies full-text screening rules, and prepares the input for CT–programming pairing.
4. **Full pairing**: extracts paper-level CT and programming information, normalizes CT–programming terminology, and produces paper-level and pair-level mapping outputs.
5. **Characterization**: extracts study characteristics such as educational context, methodology, programming tools, and intervention details.
6. **Results analysis**: generates summary tables, PRISMA-style summaries, cost summaries, figures, heatmaps, and CT–programming analysis outputs.

## Repository structure

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
├── requirements.txt
├── DEVELOPER_GUIDE.md
├── CITATION.cff
└── LICENSE
```

Each stage folder contains its own `input`, `output`, and, when needed, `temp` folder. Long-running LLM stages use checkpoint files in the corresponding `temp` folder.

## Installation

Create and activate a Python environment. For example:

```bash
python -m venv llm_project_env
source llm_project_env/bin/activate
```

Install the required packages:

```bash
pip install -r requirements.txt
```

The main dependencies include Streamlit, pandas, openpyxl, the OpenAI Python client, PyMuPDF, pypdf, matplotlib, seaborn, and plotly.

## API key

Stages that use an LLM require an OpenAI API key. You can provide it in the Streamlit sidebar, or create a local `.env` file in the project root:

```text
OPENAI_API_KEY=your_api_key_here
```

Do not commit `.env` files to the repository.

## Running the tool

From the project root, run:

```bash
streamlit run Home.py
```

The navigation menu contains the user guide and the six pipeline stages. The stages can be run sequentially, but most stages also allow a custom input file to be uploaded manually.

## Configuration files

Most rules and prompts are stored in Excel files rather than hard-coded in Python. The templates folder contains the configuration files used by the stages, including screening rules, coding schemas, normalization files, and analysis grouping files.

Examples include:

```text
templates/stage2_abstract_screening_rules.xlsx
templates/stage_3_fulltex_screening_rules.xlsx
templates/stage4_pairing_rules.xlsx
templates/stage5_characterization_rules.xlsx
templates/Analysis_details.xlsx
```

The user guide page provides download buttons for the available templates in the `templates` folder.

## Stage outputs and handoff files

Each stage writes its own results and prepares a clean input file for the next stage when applicable.

Typical outputs include:

```text
stage_3_fulltext_initial_pairing/output/stage_3_results.xlsx
stage_3_fulltext_initial_pairing/output/stage_3_author_validation.xlsx
stage_3_fulltext_initial_pairing/output/stage_3_costs.xlsx
stage_3_fulltext_initial_pairing/output/stage_3_no_full_text_available.xlsx
stage_3_fulltext_initial_pairing/output/stage_3_to_stage_4_input.xlsx
```

The Stage 3 handoff file is also copied to:

```text
stage_4_full_pairing/input/stage_4_input.xlsx
```

Stage 4 produces the main CT–programming mapping outputs:

```text
stage_4_full_pairing/output/stage_4# CT–Programming Scoping Review Pipeline

A Streamlit-based pipeline supporting a scoping review on how Computational Thinking (CT) is operationalized through programming in educational research.

The pipeline combines configurable Excel templates, Large Language Model (LLM)-assisted analysis, manual validation, and automated reporting to support a transparent and reproducible review process.

> **Important**
>
> This tool supports the review process but **does not replace the researcher**. All inclusion, exclusion, coding, characterization, and interpretation decisions remain the responsibility of the researcher.

---

# Features

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

# Workflow

The pipeline supports a six-stage review workflow.

| Stage | Description |
|--------|-------------|
| **Stage 1** | Import database exports, standardize records, remove duplicates, and prepare the abstract screening dataset. |
| **Stage 2** | Screen titles and abstracts using configurable rules and LLM-assisted classification. |
| **Stage 3** | Read full-text PDFs, identify relevant sections, perform full-text screening, and prepare papers for pairing. |
| **Stage 4** | Extract CT and programming elements, normalize terminology, and generate paper-level and pair-level mappings. |
| **Stage 5** | Characterize included studies (educational context, methodology, programming tools, interventions, etc.). |
| **Stage 6** | Generate figures, summary tables, heatmaps, PRISMA summaries, and final review statistics. |

---

# Repository Structure

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

Each stage contains its own `input`, `output`, and (when required) `temp` folders.

Long-running LLM stages automatically generate checkpoints that allow interrupted executions to be resumed.

---

# Installation

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

# OpenAI API Key

Stages using LLMs require an OpenAI API key.

The key can either be entered directly in the Streamlit sidebar or stored locally inside a `.env` file.

```text
OPENAI_API_KEY=your_api_key_here
```

The `.env` file should **never** be committed to the repository.

---

# Running the Application

From the project root execute

```bash
streamlit run Home.py
```

The application provides an integrated navigation menu containing the User Guide together with all six pipeline stages.

Each stage can be executed sequentially or independently using manually supplied input files.

---

# Configuration Files

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

The integrated User Guide provides download buttons for all available templates.

---

# Documentation

The repository includes additional documentation.

| File | Description |
|------|-------------|
| **User_Guide.pdf** | Complete user manual describing every pipeline stage. |
| **Search_Queries.pdf** | Complete search strategy, including the search query used for each academic database. |
| **DEVELOPER_GUIDE.md** | Technical documentation for developers and contributors. |
| **README.md** | General overview and installation instructions. |

---

# Full-Text PDF Repository

Stages **3** and **4** require access to the full-text PDF files of the selected studies.

The PDF files are **not distributed** with this repository.

Most scientific publications are protected by copyright and therefore cannot legally be redistributed through GitHub, Zenodo, or this repository.

Users should obtain the full-text articles through:

- institutional subscriptions,
- publisher websites,
- or legally available Open Access repositories.

Once obtained, the PDFs should be placed inside the local repository, for example:

```text
papers_repository/
└── pdfs/
```

The pipeline expects this local repository during full-text screening and CT–programming pairing.

---

# Stage Outputs

Each stage generates its own outputs and prepares the input required for the following stage whenever applicable.

Typical outputs include

```text
stage_3_results.xlsx
stage_3_author_validation.xlsx
stage_3_costs.xlsx
stage_3_no_full_text_available.xlsx
stage_3_to_stage_4_input.xlsx
```

Stage 4 generates

- paper-level mappings
- pair-level mappings
- normalization summaries
- author validation files
- cost reports

Stage 5 generates

- characterization outputs

Stage 6 generates

- publication-ready figures
- heatmaps
- summary tables
- PRISMA summaries
- statistical summaries

---

# Testing

The complete workflow does **not** need to be executed during testing.

Recommended strategy:

- Stage 1 can be tested using the provided example database exports.
- Stage 6 can be tested independently because it only requires the outputs generated by previous stages.
- For complete end-to-end testing, use a small subset of studies instead of the full corpus.

---

# Checkpoints and Recovery

Long-running LLM stages automatically generate checkpoints.

If execution is interrupted, restart the same stage and enable the resume option whenever available.

Do not delete temporary checkpoint files until the stage has successfully completed.

---

# Reproducibility

This repository contains everything required to reproduce the review workflow except the copyrighted publications.

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

---


This repository contains everything required to reproduce the review workflow, including:

- Source code
- Configuration templates
- Prompt templates
- Analysis scripts
- Example input files
- Example output files
- Documentation

The original full-text scientific publications are **not included** because they are protected by copyright and cannot generally be redistributed. Users should obtain these articles through legitimate sources before executing the full-text stages of the pipeline.

# Files Not Intended for Version Control

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

# Citation

If you use this pipeline in academic work, please cite the repository using the information provided in **CITATION.cff**.

---

# License

This project is distributed under the license provided in the `LICENSE` file.

---

# Acknowledgements

This tool was developed to support a reproducible scoping review on Computational Thinking and Programming in Computing Education.

It combines configurable rule-based processing, Large Language Models, and researcher-driven validation to improve transparency, reproducibility, and efficiency while preserving human oversight throughout the review process.