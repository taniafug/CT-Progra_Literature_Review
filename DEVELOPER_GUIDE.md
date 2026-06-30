# Developer Guide

This document explains how the CT–Programming Scoping Review Pipeline is organized internally.

It is intended for future maintainers or collaborators who want to adapt the pipeline, add new stages, or reuse the architecture for another review.

---

## Design idea

The pipeline is organized as six independent but connected stages. Each stage has its own Streamlit app, logic file, input folder, output folder, and temporary folder.

The output of one stage is usually copied or used as the input for the next stage.

```text
Stage 1 → Stage 2 → Stage 3 → Stage 4 → Stage 5 → Stage 6
```

The pipeline uses a mix of:

- deterministic Python processing
- Excel-based configuration
- LLM-assisted classification or extraction
- post-processing rules
- manual validation files
- checkpoint recovery

---

## Stage overview

### Stage 1 — Import and rule-based preparation

No LLM is used here.

Responsibilities:

- read database exports
- standardize metadata
- merge records
- detect duplicates
- flag non-screenable records
- run rule-based title screening
- prepare abstract screening input

Main output for next stage:

```text
stage_2_abstract_screening/input/abstract_screening_input.xlsx
```

---

### Stage 2 — Abstract screening

Responsibilities:

- read abstracts
- read screening rules from Excel
- run one or more LLM classifications per abstract
- combine runs into a final decision
- create manual validation files
- create cost files
- prepare full-text input

Main output for next stage:

```text
stage_3_fulltext_initial_pairing/input/fulltext_screening_input.xlsx
```

---

### Stage 3 — Full-text screening

Responsibilities:

- read PDFs
- extract relevant full-text sections
- screen the full text
- apply deterministic decision correction
- reserve papers with no full text
- prepare the next included set

Main output for next stage:

```text
stage_4_full_pairing/input/full_pairing_input.xlsx
```

---

### Stage 4 — Framework-based CT–Programming pairing

Responsibilities:

- extract paper-level CT and programming information
- extract candidate CT–programming mappings
- normalize CT and programming terminology
- create valid framework pairs
- run pair-level analysis
- produce paper-level and pair-level outputs
- filter final papers with at least one valid relationship

Main outputs:

```text
paper_level_framework_results.xlsx
pair_level_mapping_results.xlsx
```

---

### Stage 5 — Characterization

Responsibilities:

- characterize included studies
- extract study design, educational context, tools, environment, methods and other descriptors
- apply dependency rules
- generate clean analysis input

Main output for next stage:

```text
stage_6_results_analysis/input/results_characterization.xlsx
```

---

### Stage 6 — Analysis and visualization

Responsibilities:

- combine characterization and pairing outputs
- generate frequency tables
- generate CT–programming summaries
- generate heatmaps, bar charts and network figures
- create final summary workbooks

Important output folders:

```text
stage_6_results_analysis/output/figures/
stage_6_results_analysis/output/pairing_figures/
```

---

## Data flow

```text
Stage 1
review_pipeline_processing.xlsx
title_screening_results.xlsx
abstract_screening_input.xlsx
        ↓
Stage 2
abstracts_screened_llm_results.xlsx
abstracts_screened_llm_next_stage_fulltext.xlsx
        ↓
Stage 3
fulltext_results.xlsx
no_full_text_available.xlsx
        ↓
Stage 4
paper_level_framework_results.xlsx
pair_level_mapping_results.xlsx
        ↓
Stage 5
results_characterization.xlsx
        ↓
Stage 6
characterization_analysis.xlsx
final_review_summary.xlsx
figures/
pairing_figures/
```

---

## Configuration system

Most stage behavior is controlled through Excel files rather than hard-coded values.

This makes it easier to reuse the pipeline for another review.

Common configuration sheets:

### Prompt_Config

Contains the main instructions sent to the LLM.

Typical fields:

- system role
- task description
- general behavior
- uncertainty rules
- output restrictions

### Screening_Rules

Contains inclusion and exclusion rules.

Used mainly in Stages 2 and 3.

### Output_Schema

Defines the fields that must be returned by the LLM.

Usually converted into a strict JSON schema.

### Coding_Schema

Defines allowed values and coding notes.

Used mainly in Stages 4 and 5.

### Dependency_Rules

Defines deterministic corrections or consistency rules.

Example:

```text
If programming type is unplugged, programming tool should be Not applicable.
```

### Section_Patterns

Defines which sections of a paper should be prioritized.

Typical groups:

- methods
- results
- framework
- discussion
- intervention
- assessment
- appendix

### Normalization sheets

Used in Stage 4.

Typical sheets:

- Normalization_Rules
- CT_Normalization
- Programming_Normalization

These map raw terms to standardized framework labels.

### Analysis_details

Used in Stage 6.

It defines grouping rules for final analysis, such as programming tool groupings and programming element groupings.

---

## Checkpoint logic

Long LLM stages write checkpoint files during execution.

The purpose is to avoid losing work if the process is interrupted.

Typical checkpoint files:

```text
checkpoint_fulltext_screening.xlsx
pair_level_mapping_results_checkpoint.xlsx
paper_level_framework_results_checkpoint.xlsx
characterization_checkpoint.xlsx
```

When resuming, the app reads the checkpoint and skips already processed papers.

---

## LLM use

The pipeline currently uses the OpenAI API.

Most calls use:

```text
gpt-4.1-mini
```

The model can be changed in the Streamlit sidebar.

The pipeline uses structured outputs where possible, usually through JSON schema constraints.

---

## Cost tracking

LLM stages save token and cost information.

The cost calculation usually uses:

```text
input_tokens
output_tokens
total_tokens
estimated_cost
```

The token price can be adjusted in the app interface.

---

## Manual validation

Each LLM stage can generate a validation workbook.

These files are useful for checking:

- random samples
- manual review cases
- undetermined records
- inconsistent outputs
- errors
- missing information

The default validation sample is usually 15%.

---

## Adding a new review topic

To adapt the pipeline to another review, it is usually enough to modify:

1. Screening rules
2. Prompt configuration
3. Output schemas
4. Coding schemas
5. Dependency rules
6. Normalization dictionaries
7. Analysis grouping file

The Python code should only need changes if the structure of the review is very different.

---

## Adding a new stage

A new stage should follow the same structure:

```text
stage_x_name/
├── input/
├── output/
├── temp/
├── app_stage_x.py
└── stage_x_logic.py
```

Recommended design:

- keep UI code in the app file
- keep processing functions in the logic file
- save checkpoints during long runs
- create clear Excel outputs
- prepare the next stage input automatically
- avoid hard-coding review-specific rules in Python when possible

---

## Notes for maintainers

- Keep generated outputs out of Git unless they are small example files.
- Do not commit `.env` files.
- Do not commit downloaded PDFs.
- Keep templates small and clean.
- Prefer configuration changes over code changes.
- Keep deterministic post-processing rules visible and documented.
