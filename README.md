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
stage_4_full_pairing/output/stage_4_paper_analysis.xlsx
stage_4_full_pairing/output/stage_4_pairing_analysis.xlsx
stage_4_full_pairing/output/stage_4_previous_and_cleaning_outputs.xlsx
stage_4_full_pairing/output/stage_4_author_validation.xlsx
stage_4_full_pairing/output/stage_4_summary_and_costs.xlsx
```

Stage 5 produces the characterization output and the file used by Stage 6:

```text
stage_5_characterization/output/stage_5_to_stage_6_input.xlsx
```

Stage 6 produces the final analysis outputs and figures:

```text
stage_6_results_analysis/output/characterization_analysis.xlsx
stage_6_results_analysis/output/programming_elements_summary.xlsx
stage_6_results_analysis/output/final_review_summary.xlsx
stage_6_results_analysis/output/figures/
stage_6_results_analysis/output/pairing_figures/
```

## Testing without running the full LLM pipeline

Some parts of the tool can be tested without calling the LLM:

- Stage 1 can be tested using the example input files provided in the corresponding input folder.
- Stage 6 does not require an LLM. It analyzes the outputs generated by the previous stages.

For testing the complete LLM pipeline, use a small sample dataset rather than the full corpus.

## Checkpoints and recovery

Long LLM stages save checkpoints during execution. If the process is interrupted, restart the same stage and enable the resume option when available. Do not delete files in the `temp` folders until the stage has finished successfully.

## Notes on generated files

Generated outputs, checkpoints, PDFs, and local API files should not normally be committed to Git. Keep only small example files or templates needed to reproduce the workflow.

Recommended exclusions include:

```text
.env
**/temp/
**/papers_repository/
**/__pycache__/
```

## Documentation

- `README.md`: general user-facing overview and installation instructions.
- `DEVELOPER_GUIDE.md`: internal design notes for maintainers and collaborators.
- Streamlit User Guide page: interactive guide inside the application.

## Citation

Please cite this repository using the information provided in `CITATION.cff`.

## License

See `LICENSE` for license information.