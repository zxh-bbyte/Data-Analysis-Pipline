# sp-ICP-TOF-MS Data Analysis Pipeline

A small pipeline for processing single-particle ICP-TOF-MS (sp-ICP-TOF-MS) exports.
It extracts per-measurement particle tables, classifies particles into fingerprint
(FP) fraction classes, computes particle number concentrations (PNC), and produces
several summary workbooks for downstream interpretation.

> A Chinese manual is available in [`README_使用手册.md`](README_使用手册.md).

---

## Overview

Each raw measurement is a `particle.csv` table where every column is an isotope
(plus an `embedding` column that is ignored) and every cell is the fraction of that
isotope in a particle, in `[0, 1]` (blank = 0). The pipeline turns these tables into:

- Counts of each isotope per fraction class (`smFPs`, `mmFPs`, `0.1FPs`, …)
- Dominant-element statistics per particle
- Particle number concentration (PNC) tables with live Excel formulas
- Adsorbed-state host matrix per sample (as PNC)
- A parallel-replicate consistency screen
- A toxic-element-only summary

### Pipeline at a glance

```
raw TOF folders
      │  extract_particles.py            (step 0: gather + rename)
      ▼
Particle/particle_<sample>X<dilution>-<meas>.csv
      │
      ├─► particle_fp_pipeline.py  ─► smFPs / mmFPs / 0.1FPs / 0.1-0.4FPs /
      │        (1, core entry)         0.4-1FPs / main / FPs_summary.xlsx + cleaned CSV
      │             │
      │             ├─► fp_pnc_summary.py ─► FPs_PNC_summary.xlsx        (2)
      │             │           │
      │             │           └─► adsorbed_state_host_summary.py ─►
      │             │                   0.1FPs_adsorbed_host_summary.xlsx (3)
      │             │
      │             └─► parallel_outlier_screen.py ─►
      │                     parallel_outlier_screen.xlsx                  (4)
      │
      └─► toxic_fps_summary.py ─► Toxic/ToxicFPs_summary.xlsx +
               (5, standalone)      Toxic/ToxicFPs_PNC_summary.xlsx
```

---

## Requirements

```bash
pip install -r requirements.txt      # pandas, numpy, openpyxl
```

Python 3.10+ is recommended (the scripts use `X | Y` type hints).

---

## Quick start (bundled example)

A small `example_data/` folder is included so you can see the first step run
immediately. It holds three raw measurement folders (`CFPP2X1000_iteration`,
`COP11X1000_iteration`, `EAF10X5000_iteration`). `extract_particles.py` is
preconfigured to read them, so just run:

```bash
python extract_particles.py
```

This writes the renamed tables to `Particle/`:

```
example_data/CFPP2X1000_iteration/particle.csv  ->  Particle/particle_CFPP2X1000-1.csv
example_data/COP11X1000_iteration/particle.csv  ->  Particle/particle_COP11X1000-1.csv
example_data/EAF10X5000_iteration/particle.csv  ->  Particle/particle_EAF10X5000-1.csv
```

Each output is named `particle_<sample>X<dilution>-<measurement>.csv`, parsed from
the folder name. A missing `X<dilution>` tag or `-<measurement>` number is each
defaulted to `1` (e.g. a folder `CFPP2` -> `particle_CFPP2X1-1.csv`). To run the rest
of the pipeline on this output, point `particle_fp_pipeline.py`'s `INPUT_DIR` at the
`Particle/` folder.

---

## Input data format

- **Location:** `C:\path\to\data\Particle\` (configurable).
- **Filename:** `particle_<sample>X<dilution>-<measurement>.csv` (produced by
  `extract_particles.py`), e.g. `particle_CFPP2X1000-1.csv`. The reported `sampleID`
  strips the dilution tag (`CFPP2X1000` → `CFPP2`), and the dilution factor is taken
  from that `X<dilution>` tag (`X1000` → 1000, `X5000` → 5000). A file that does not
  match `particle_<...>-<measurement>.csv` will stop the pipeline with an error.
- **Columns:** one column per isotope, plus an optional `embedding` column (dropped).
- **Cell values:** fraction of the isotope in the particle, `0 ≤ v ≤ 1` (blank = 0).

### FP fraction classes

| Class          | Fraction range | Meaning                              |
| -------------- | -------------- | ------------------------------------ |
| `smFPs`        | exactly `= 1`  | single-metal particle                |
| `mmFPs`        | `(0, 1)`       | multi-metal particle (sum of below)  |
| `0.1FPs`       | `(0, 0.1]`     | adsorbed-state fraction              |
| `0.1-0.4FPs`   | `(0.1, 0.4)`   | intermediate fraction                |
| `0.4-1FPs`     | `[0.4, 1)`     | major fraction                       |
| `main`         | —              | dominant element per particle, after dropping any row that contains a `= 1` value |

---

## Run order

Scripts have dependencies and must be run in this order (5 is standalone):

| # | Script                            | Depends on                    | Main output |
| - | --------------------------------- | ----------------------------- | ----------- |
| 0 | `extract_particles.py`            | raw TOF folders               | `Particle/particle_*.csv` |
| 1 | `particle_fp_pipeline.py`         | step 0                        | FP tables + `FPs_summary.xlsx` + cleaned CSV |
| 2 | `fp_pnc_summary.py`               | `FPs_summary.xlsx`            | `FPs_PNC_summary.xlsx` |
| 3 | `adsorbed_state_host_summary.py`  | `FPs_PNC_summary.xlsx` + raw CSV | `0.1FPs_adsorbed_host_summary.xlsx` |
| 4 | `parallel_outlier_screen.py`      | step 1 FP tables              | `parallel_outlier_screen.xlsx` |
| 5 | `toxic_fps_summary.py`            | raw CSV only (standalone)     | `Toxic/ToxicFPs_summary.xlsx` + `Toxic/ToxicFPs_PNC_summary.xlsx` |

Run each with, e.g.:

```bash
python extract_particles.py
python particle_fp_pipeline.py
python fp_pnc_summary.py
python adsorbed_state_host_summary.py
python parallel_outlier_screen.py
python toxic_fps_summary.py
```

---

## Configuration

Every script starts with a clearly delimited block of user-adjustable parameters:

```python
# =============================================================================
# CONFIGURATION -- user-adjustable parameters
# =============================================================================
...
# =============================================================================
# End of configuration
# =============================================================================
```

You normally only edit that block. The parameters per script:

### 0 · `extract_particles.py`
- `BATCH_LABEL` — subfolder under the raw-data root to process (defaults to `example_data`).
- `SOURCE_ROOT` — root folder of raw TOF measurement subfolders (defaults to the bundled example next to the script).
- `OUTPUT_DIR` — where renamed `particle_*.csv` files are written (defaults to `Particle/` next to the script).
- `SOURCE_CSV_NAME` — raw file name to look for in each subfolder.
- `DEFAULT_MEASUREMENT` — measurement number used when a folder name has none (default `1`).
- `DEFAULT_DILUTION` — dilution label used when a folder name has no `X<n>` tag (default `1`).

### 1 · `particle_fp_pipeline.py`
- `INPUT_DIR` / `ANALYSIS_DIR` / `CLEANED_DIR` — I/O directories.
- `TASKS_TO_RUN` — which tasks to run (`smFPs`, `mmFPs`, `0.1FPs`, `0.1-0.4FPs`, `0.4-1FPs`, `clean`, `main`, `summary`).
- `RANGE_RULES` — the fraction-interval bounds for each class.
- Output workbook / sheet names, and `FILENAME_PATTERN` (input naming).

### 2 · `fp_pnc_summary.py`
- `INPUT_PATH` / `OUTPUT_PATH`.
- `TE` — transport efficiency of the batch being processed (process one TE batch per run).
- `SECONDS_PER_MINUTE` / `SAMPLE_FLOW_RATE_ML_MIN` / `ACQUISITION_TIME_SECONDS` / `CONSTANT_VOLUME_ML` / `SAMPLE_MASS_MG` — PNC formula constants (see below).
- The dilution factor is read per sample from the input workbook's `Dilutionfactor` column (parsed upstream from each sample's `X<n>` tag), not configured here.

### 3 · `adsorbed_state_host_summary.py`
- `PNS_WORKBOOK_PATH` / `PARTICLE_DIR` / `OUTPUT_PATH`.
- `ADSORBED_SHEET_NAME` — worksheet holding the `≤0.1` counts.
- PNC formula constants (same set as `fp_pnc_summary.py`).

### 4 · `parallel_outlier_screen.py`
- `ANALYSIS_DIR` / `OUTPUT_PATH`.
- `MIN_VALUE_THRESHOLD` (default `100`) and `RATIO_THRESHOLD` (default `2`).
- `TABLES_TO_CHECK` / `TABLE_FILE_MAP` — which result tables to screen.

### 5 · `toxic_fps_summary.py`
- `INPUT_DIR` / `OUTPUT_DIR` and the two output paths.
- `TOXIC_ELEMENTS` — the toxic-element list.
- `RANGE_RULES`, `TE`, and PNC formula constants. (Dilution is parsed per sample from the `X<n>` tag, not configured here.)

---

## PNC formula

Particle number concentration is computed (as a live Excel formula) as:

```
PNC (particles/mg) = counts × DilutionFactor × SECONDS_PER_MINUTE × CONSTANT_VOLUME_ML
                     / (TE × SAMPLE_FLOW_RATE_ML_MIN × ACQUISITION_TIME_SECONDS × SAMPLE_MASS_MG)
```

Defaults: `SECONDS_PER_MINUTE = 60`, `SAMPLE_FLOW_RATE_ML_MIN = 0.02`,
`ACQUISITION_TIME_SECONDS = 150`, `CONSTANT_VOLUME_ML = 50`, `SAMPLE_MASS_MG = 20`.
`TE` is a single value you set per run. `DilutionFactor` is parsed per sample from the
`X<n>` tag in the sample name, so samples measured at different dilutions (e.g. 1000
and 5000) are each handled correctly. Adjust all constants to match your acquisition setup.

---

## Outputs

All analysis outputs go to `C:\path\to\data\analysis\` (configurable):

| File | Produced by | Sheets / contents |
| ---- | ----------- | ----------------- |
| `smFPs.xlsx`, `mmFPs.xlsx`, `0.1FPs.xlsx`, `0.1-0.4FPs.xlsx`, `0.4-1FPs.xlsx`, `main.xlsx` | 1 | per-class isotope counts (per measurement + sample mean) |
| `FPs_summary.xlsx` | 1 | `smFPs / mmFPs / Total / 0.1FPs / 0.1-0.4FPs / 0.4-1FPs / mainFPs` |
| `FPs_PNC_summary.xlsx` | 2 | same sheets + `TE` column and PNC formula block |
| `0.1FPs_adsorbed_host_summary.xlsx` | 3 | `host_counts` + `host_pnc` — per-sample adsorbed-element × host-element (counts and PNC) |
| `parallel_outlier_screen.xlsx` | 4 | `sample_summary / evidence_detail` |
| `Toxic/ToxicFPs_summary.xlsx`, `Toxic/ToxicFPs_PNC_summary.xlsx` | 5 | toxic-element-only summaries and PNC |

---

## Notes

- Filenames must match `particle_<sample>-<measurement>.csv`, otherwise
  `particle_fp_pipeline.py` raises an error and stops.
- PNC workbooks store Excel **formulas** — open them in Excel and it recalculates PNC.
- A pandas `PerformanceWarning` (DataFrame fragmentation) may appear; it is a
  performance hint, not an error, and can be ignored.
