"""Step 0 of the pipeline: extract per-measurement particle tables.

For every subfolder under SOURCE_ROOT that contains a raw `particle.csv`, this
script copies that file into OUTPUT_DIR, renamed following the convention

    particle_<sample>X<dilution>-<measurement>.csv

parsed from the folder name (after dropping the trailing `_iteration` tag), e.g.
`CFPP2X1000_iteration` -> `particle_CFPP2X1000-1.csv`. A folder name missing the
`X<dilution>` tag or the `-<measurement>` number has each defaulted to 1
(e.g. `CFPP2` -> `particle_CFPP2X1-1.csv`). Run this before any analysis script.

By default the paths below point at the bundled `example_data/` folder, so you
can run the script immediately and compare the input folders with the generated
`Particle/` output. Change SOURCE_ROOT / OUTPUT_DIR to your own locations for
real processing.
"""

import os
import re

import pandas as pd

# =============================================================================
# CONFIGURATION -- user-adjustable parameters
# =============================================================================
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Batch label = name of the subfolder under the raw-data root to process.
BATCH_LABEL = "example_data"

# Root folder holding the raw TOF measurement subfolders (one folder per run).
# Defaults to the bundled example data located next to this script.
SOURCE_ROOT = os.path.join(_SCRIPT_DIR, BATCH_LABEL)

# Destination folder for the renamed particle CSV files.
OUTPUT_DIR = os.path.join(_SCRIPT_DIR, "Particle")

# Name of the raw per-measurement particle file to look for in each subfolder.
SOURCE_CSV_NAME = "particle.csv"

# Measurement number to use when a folder name has no "-<n>" measurement token.
DEFAULT_MEASUREMENT = 1

# Dilution factor to use when a folder name has no "X<n>" dilution tag.
DEFAULT_DILUTION = 1
# =============================================================================
# End of configuration
# =============================================================================

# Folder naming convention (after dropping "_iteration"):
#   <sample>X<dilution>[-<measurement>]   e.g. "CFPP2X1000" or "CFPP2X1000-2"
# A missing dilution tag or measurement number is replaced with 1.
_MEASUREMENT_RE = re.compile(r"-(\d+)$")
_DILUTION_RE = re.compile(r"[Xx](\d+)$")


def build_output_name(dir_name: str) -> str:
    """Return the particle_<sample>X<dilution>-<measurement>.csv output name.

    Parses <sample>X<dilution>[-<measurement>] from the folder name (after
    stripping "_iteration"); a missing dilution or measurement defaults to 1.
    """
    core = dir_name.replace("_iteration", "")

    measurement_match = _MEASUREMENT_RE.search(core)
    if measurement_match:
        measurement = measurement_match.group(1)
        core = core[: measurement_match.start()]
    else:
        measurement = str(DEFAULT_MEASUREMENT)

    dilution_match = _DILUTION_RE.search(core)
    if dilution_match:
        sample = core[: dilution_match.start()]
        dilution = dilution_match.group(1)
    else:
        sample = core
        dilution = str(DEFAULT_DILUTION)

    return f"particle_{sample}X{dilution}-{measurement}.csv"


def main() -> None:
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    extracted = 0
    # Walk every subfolder under the source root and copy its particle.csv.
    for root, dirs, _files in os.walk(SOURCE_ROOT):
        for dir_name in dirs:
            dir_path = os.path.join(root, dir_name)
            if SOURCE_CSV_NAME not in os.listdir(dir_path):
                continue

            source_csv_path = os.path.join(dir_path, SOURCE_CSV_NAME)
            dataframe = pd.read_csv(source_csv_path)

            new_csv_path = os.path.join(OUTPUT_DIR, build_output_name(dir_name))
            dataframe.to_csv(new_csv_path, index=False)
            extracted += 1

    print(f"Extraction complete: {extracted} particle file(s) written to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
