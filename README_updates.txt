Characterization logic updates:

- Keeps Excel-driven prompt: Prompt_Config, Coding_Schema_Characterization, Dependency_Rules, Section_Patterns.
- Keeps original partial one-sheet output after every paper.
- Adds bibliography removal.
- Adds fallback to rest_of_document_without_bibliography if section extraction is weak.
- Adds clean Analysis_input sheet inside characterization_results.xlsx.
- Keeps validation and cost workbooks separate.
- Adds technical checkpoint workbook and optional resume_from_checkpoint in run_characterization().
- write_output_workbooks now writes 3 final files. If an older app still passes analysis_input_output_path, it remains supported as optional backward compatibility.
