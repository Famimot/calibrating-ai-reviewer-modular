# Calibrating AI Reviewer: A Modular Architecture
DOI

## About the Project

This repository contains prompts, software, and validation results for the modular architecture of a calibrating AI reviewer. The architecture consists of six independent modules that sequentially implement the determination of the review mode (STEM/Humanities), assessment of scientific novelty, methodological rigor, practical value, visualization quality, and reproducibility.

A detailed description of the method and validation results are presented in the article:

> Kravtsov G.G. Calibrating AI Reviewer: A Modular Architecture for Evaluation and Ranking of Scientific Papers // Scientific and Technical Information. 2026. (under review)

## Repository Structure

    calibrating-ai-reviewer-modular/
    │
    ├── README.md
    ├── LICENSE
    │
    ├── prompts/
    │   ├── v.7.3_monolithic/
    │   │   ├── prompt_v73_rus.txt
    │   │   └── prompt_v73_eng.txt
    │   │
    │   └── v.8.8_modular/
    │       ├── module_1_mode_detector.txt
    │       ├── module_2_novelty_analyzer.txt
    │       ├── module_3_methodology_analyzer.txt
    │       ├── module_4_value_analyzer.txt
    │       ├── module_5_visualization_reproducibility.txt
    │       └── module_6_aggregator_categorizer.txt
    │
    ├── software/
    │   ├── llm_review_calibration_eng.exe
    │   └── llm_review_calibration_eng.py
    │
    ├── results/
    │   ├── Appendix_B_Categorization_Module_6.xlsx
    │   └── Appendix_C_Modular_10_runs_raw_data.xlsx
    │
    └── archive/
        └── ...

## Usage

### Using Prompts

1. Open a new dialog with an LLM (recommended: DeepSeek-V3)
2. Load Module 1 with the selected mode (USER_MODE = 0 or 1) and the full texts of the papers (recommended: no more than 7). Execute the first module request.
3. Then sequentially, after receiving the results of the previous module, load Modules 2–6

### Using the Software

1. Download `llm_review_calibration_eng.exe` or `llm_review_calibration_eng.py`
2. Run the program and load the file `results/Appendix_C_Modular_10_runs_raw_data.xlsx`
3. The program will automatically analyze 10 runs, generate tables and charts, and save the results to Excel

**Requirements for the Python script:**
- Python 3.8+
- Installed dependencies: `pandas`, `numpy`, `matplotlib`, `openpyxl`

## Validation Results

The `results/` folder contains:

| File | Description |
|------|-------------|
| `Appendix_B_Categorization_Module_6.xlsx` | Categorization of 7 papers (A/B/C/D) with justification |
| `Appendix_C_Modular_10_runs_raw_data.xlsx` | Raw data from 10 runs for all papers |

## Reproducibility

Any user can:
1. Take arbitrary manuscripts (at least 2)
2. Use the prompts from this repository
3. Repeat the algorithm for comparing and ranking papers

## License

This project is distributed under the MIT License.

Copyright (c) 2026 Gennady Kravtsov

This license permits any use, copying, modification, merger, publication, distribution, sublicensing, and/or sale of copies of the software without restriction, provided that the copyright notice and the text of the license are retained.

## Citation

If you use this repository in your research, please cite the article:

```bibtex
@article{Kravtsov2026_AI_Reviewer,
  author  = {Kravtsov, G.G.},
  title   = {Calibrating AI Reviewer: A Modular Architecture for Evaluation and Ranking of Scientific Papers},
  journal = {Scientific and Technical Information},
  note    = {under review},
  year    = {2026}
}