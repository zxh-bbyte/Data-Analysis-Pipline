from __future__ import annotations

import math
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


# =============================================================================
# CONFIGURATION -- user-adjustable parameters
# =============================================================================
# Input/output directories. They default to folders next to this script so the
# bundled example runs out-of-the-box; set them to your own absolute paths
# (e.g. Path(r"C:\path\to\data\Particle")) for real processing.
_SCRIPT_DIR = Path(__file__).resolve().parent
INPUT_DIR = _SCRIPT_DIR / "Particle"
ANALYSIS_DIR = _SCRIPT_DIR / "analysis"
CLEANED_DIR = _SCRIPT_DIR / "particle_cleaned_no1"

# Tasks to run (comment out any you do not need). Valid task names:
# "smFPs", "mmFPs", "0.1FPs", "0.1-0.4FPs", "0.4-1FPs", "clean", "main", "summary"
TASKS_TO_RUN = [
    "smFPs",
    "mmFPs",
    "0.1FPs",
    "0.1-0.4FPs",
    "0.4-1FPs",
    "clean",
    "main",
    "summary",
]

# Fraction-interval definitions for each FPs class. Edit the bounds freely; use
# either "exact" (a single value) OR min/max with include_min / include_max.
RANGE_RULES: dict[str, dict[str, Any]] = {
    "smFPs": {
        "output_name": "smFPs.xlsx",
        "sheet_name": "smFPs",
        "total_column": "TotalsmFPs",
        "exact": 1.0,
    },
    "mmFPs": {
        "output_name": "mmFPs.xlsx",
        "sheet_name": "mmFPs",
        "total_column": "TotalmmFPs",
        "min": 0.0,
        "max": 1.0,
        "include_min": False,
        "include_max": False,
    },
    "0.1FPs": {
        "output_name": "0.1FPs.xlsx",
        "sheet_name": "0.1FPs",
        "total_column": "Total0.1FPs",
        "min": 0.0,
        "max": 0.1,
        "include_min": False,
        "include_max": True,
    },
    "0.1-0.4FPs": {
        "output_name": "0.1-0.4FPs.xlsx",
        "sheet_name": "0.1-0.4FPs",
        "total_column": "Total0.1-0.4FPs",
        "min": 0.1,
        "max": 0.4,
        "include_min": False,
        "include_max": False,
    },
    "0.4-1FPs": {
        "output_name": "0.4-1FPs.xlsx",
        "sheet_name": "0.4-1FPs",
        "total_column": "Total0.4-1FPs",
        "min": 0.4,
        "max": 1.0,
        "include_min": True,
        "include_max": False,
    },
}

# Output workbook / sheet names.
MAIN_OUTPUT_NAME = "main.xlsx"
MAIN_SHEET_NAME = "mainFPs"
TOTAL_SHEET_NAME = "Total"
TOTAL_TOTAL_COLUMN = "TotalFPs"
SUMMARY_OUTPUT_NAME = "FPs_summary.xlsx"

# Expected input filename pattern: particle_<sampleID>-<measurement>.csv
# (extract_particles.py encodes dilution into the sampleID as "<sample>X<dilution>";
# a legacy trailing "-<n>x" dilution tag is also tolerated and ignored). The dilution
# factor is parsed per sample from the "X<n>" tag in the sample name.
FILENAME_PATTERN = re.compile(
    r"^particle_(?P<sample_id>.+?)-(?P<measurement>\d+)(?:-\d+(?:\.\d+)?x)?\.csv$",
    re.IGNORECASE,
)
# =============================================================================
# End of configuration
# =============================================================================


@dataclass(frozen=True)
class ParticleFile:
    path: Path
    file_name: str
    sample_id: str
    measurement_index: int
    dilution_factor: int | float
    dilution_label: str
    sample_prefix: str
    sample_number: int

    @property
    def measurement_sample_id(self) -> str:
        return f"{self.sample_id}-{self.measurement_index}"

    @property
    def main_measurement_column(self) -> str:
        return f"{self.sample_id}-{self.measurement_index}-{self.dilution_label}"

    @property
    def group_key(self) -> tuple[str, int | float]:
        return self.sample_id, self.dilution_factor

    @property
    def sort_key(self) -> tuple[str, int, str, float, int]:
        return (
            self.sample_prefix.lower(),
            self.sample_number,
            self.sample_id.lower(),
            float(self.dilution_factor),
            self.measurement_index,
        )


def split_sample_for_sort(sample_id: str) -> tuple[str, int]:
    matched = re.match(r"^(?P<prefix>[A-Za-z]+)(?P<number>\d+)$", sample_id)
    if matched:
        return matched.group("prefix"), int(matched.group("number"))
    return sample_id, -1


def parse_dilution_from_sample(sample_id: str) -> int:
    """Dilution factor from the sample name's trailing "X<n>" tag (default 1)."""
    matched = re.search(r"[Xx](\d+)$", sample_id)
    return int(matched.group(1)) if matched else 1


def strip_dilution_tag(sample_id: str) -> str:
    """Drop the trailing "X<n>" dilution tag from a sample name (CFPP2X1000 -> CFPP2)."""
    return re.sub(r"[Xx]\d+$", "", sample_id)


def round_half_up(value: float) -> int:
    return int(math.floor(value + 0.5))


def discover_particle_files(input_dir: Path) -> list[ParticleFile]:
    files: list[ParticleFile] = []
    for file_path in input_dir.glob("particle_*.csv"):
        matched = FILENAME_PATTERN.match(file_path.name)
        if not matched:
            raise ValueError(f"Cannot parse file name: {file_path.name}")

        raw_sample_id = matched.group("sample_id")
        dilution_factor = parse_dilution_from_sample(raw_sample_id)
        sample_id = strip_dilution_tag(raw_sample_id)
        prefix, number = split_sample_for_sort(sample_id)

        files.append(
            ParticleFile(
                path=file_path,
                file_name=file_path.name,
                sample_id=sample_id,
                measurement_index=int(matched.group("measurement")),
                dilution_factor=dilution_factor,
                dilution_label=f"{dilution_factor}x",
                sample_prefix=prefix,
                sample_number=number,
            )
        )

    files.sort(key=lambda item: item.sort_key)
    if not files:
        raise FileNotFoundError(f"No particle CSV files found in input directory: {input_dir}")
    return files


def read_particle_dataframe(file_path: Path, expected_columns: list[str] | None = None) -> pd.DataFrame:
    dataframe = pd.read_csv(file_path)
    if "embedding" in dataframe.columns:
        dataframe = dataframe.drop(columns=["embedding"])

    dataframe = dataframe.apply(pd.to_numeric, errors="coerce")

    if expected_columns is None:
        return dataframe

    missing_columns = [column for column in expected_columns if column not in dataframe.columns]
    extra_columns = [column for column in dataframe.columns if column not in expected_columns]
    if missing_columns or extra_columns:
        raise ValueError(
            f"Columns do not match the expected set: {file_path.name}; "
            f"missing={missing_columns}, extra={extra_columns}"
        )
    return dataframe[expected_columns]


def ensure_output_dir(directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)


def build_range_mask(dataframe: pd.DataFrame, rule: dict[str, Any]) -> pd.DataFrame:
    if "exact" in rule:
        return dataframe.eq(rule["exact"])

    min_value = rule.get("min")
    max_value = rule.get("max")
    include_min = rule.get("include_min", False)
    include_max = rule.get("include_max", False)

    if min_value is None and max_value is None:
        raise ValueError("Range rule has no valid bounds.")

    mask = pd.DataFrame(True, index=dataframe.index, columns=dataframe.columns)
    if min_value is not None:
        if include_min:
            mask &= dataframe.ge(min_value)
        else:
            mask &= dataframe.gt(min_value)
    if max_value is not None:
        if include_max:
            mask &= dataframe.le(max_value)
        else:
            mask &= dataframe.lt(max_value)
    return mask


def create_measurement_record(
    meta: ParticleFile,
    counts: pd.Series,
    total_column: str,
) -> dict[str, Any]:
    record: dict[str, Any] = {
        "_sample_id": meta.sample_id,
        "_sample_prefix": meta.sample_prefix.lower(),
        "_sample_number": meta.sample_number,
        "_dilution_sort": float(meta.dilution_factor),
        "_measurement_index": meta.measurement_index,
        "sampleID": meta.measurement_sample_id,
        "Dilutionfactor": meta.dilution_factor,
        total_column: int(counts.sum()),
    }
    record.update(counts.astype(int).to_dict())
    return record


def create_average_record(
    measurement_group: pd.DataFrame,
    sample_id: str,
    dilution_factor: int | float,
    total_column: str,
    element_columns: list[str],
) -> dict[str, Any]:
    average_counts = measurement_group[element_columns].mean(axis=0)
    rounded_counts = average_counts.apply(round_half_up).astype(int)

    record: dict[str, Any] = {
        "_sample_id": sample_id,
        "_sample_prefix": measurement_group.iloc[0]["_sample_prefix"],
        "_sample_number": int(measurement_group.iloc[0]["_sample_number"]),
        "_dilution_sort": float(dilution_factor),
        "_measurement_index": math.inf,
        "sampleID": sample_id,
        "Dilutionfactor": dilution_factor,
        total_column: int(rounded_counts.sum()),
    }
    record.update(rounded_counts.to_dict())
    return record


def finalize_stat_table(
    records: list[dict[str, Any]],
    total_column: str,
    element_columns: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not records:
        raise ValueError("No statistics records; cannot build the result table.")

    measurement_df = pd.DataFrame(records)
    measurement_df = measurement_df.sort_values(
        by=[
            "_sample_prefix",
            "_sample_number",
            "_sample_id",
            "_dilution_sort",
            "_measurement_index",
        ],
        kind="stable",
    ).reset_index(drop=True)

    average_records: list[dict[str, Any]] = []
    full_output_rows: list[dict[str, Any]] = []

    group_columns = [
        "_sample_prefix",
        "_sample_number",
        "_sample_id",
        "_dilution_sort",
        "Dilutionfactor",
    ]

    for _, measurement_group in measurement_df.groupby(group_columns, sort=False):
        sample_id = str(measurement_group.iloc[0]["_sample_id"])
        dilution_factor = measurement_group.iloc[0]["Dilutionfactor"]
        average_record = create_average_record(
            measurement_group=measurement_group,
            sample_id=sample_id,
            dilution_factor=dilution_factor,
            total_column=total_column,
            element_columns=element_columns,
        )
        average_records.append(average_record)

        full_output_rows.extend(
            measurement_group[["sampleID", "Dilutionfactor", total_column] + element_columns]
            .to_dict(orient="records")
        )
        full_output_rows.append(
            {column: average_record[column] for column in ["sampleID", "Dilutionfactor", total_column] + element_columns}
        )

    output_columns = ["sampleID", "Dilutionfactor", total_column] + element_columns
    full_output_df = pd.DataFrame(full_output_rows, columns=output_columns)

    average_df = pd.DataFrame(average_records)
    average_df = average_df.sort_values(
        by=["_sample_prefix", "_sample_number", "_sample_id", "_dilution_sort"],
        kind="stable",
    )
    average_output_df = average_df[output_columns].reset_index(drop=True)
    return full_output_df, average_output_df


def compute_range_statistics(
    particle_files: list[ParticleFile],
) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame], list[str]]:
    records_by_rule: dict[str, list[dict[str, Any]]] = {name: [] for name in RANGE_RULES}
    element_columns: list[str] | None = None

    for particle_file in particle_files:
        dataframe = read_particle_dataframe(particle_file.path, expected_columns=element_columns)
        if element_columns is None:
            element_columns = dataframe.columns.tolist()

        for rule_name, rule in RANGE_RULES.items():
            mask = build_range_mask(dataframe, rule)
            counts = mask.sum(axis=0).reindex(element_columns, fill_value=0).astype(int)
            records_by_rule[rule_name].append(
                create_measurement_record(
                    meta=particle_file,
                    counts=counts,
                    total_column=rule["total_column"],
                )
            )

    if element_columns is None:
        raise ValueError("No element columns were detected.")

    full_tables: dict[str, pd.DataFrame] = {}
    average_tables: dict[str, pd.DataFrame] = {}
    for rule_name, records in records_by_rule.items():
        full_table, average_table = finalize_stat_table(
            records=records,
            total_column=RANGE_RULES[rule_name]["total_column"],
            element_columns=element_columns,
        )
        full_tables[rule_name] = full_table
        average_tables[rule_name] = average_table

    return full_tables, average_tables, element_columns


def clean_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    return dataframe.loc[~dataframe.eq(1).any(axis=1)].copy()


def save_cleaned_dataframe(dataframe: pd.DataFrame, target_path: Path) -> None:
    ensure_output_dir(target_path.parent)
    dataframe.to_csv(target_path, index=False)


def compute_dominant_element_counts(
    cleaned_dataframe: pd.DataFrame,
    element_columns: list[str],
) -> pd.Series:
    filled_dataframe = cleaned_dataframe[element_columns].fillna(0)
    dominant_elements = filled_dataframe.idxmax(axis=1)
    return dominant_elements.value_counts().reindex(element_columns, fill_value=0).astype(int)


def process_clean_and_main(
    particle_files: list[ParticleFile],
    element_columns: list[str],
    write_cleaned_csv: bool,
    generate_main_tables: bool,
) -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    dominant_records: list[tuple[ParticleFile, pd.Series]] = []
    grouped_counts: dict[tuple[str, int | float], list[pd.Series]] = defaultdict(list)

    for particle_file in particle_files:
        dataframe = read_particle_dataframe(particle_file.path, expected_columns=element_columns)
        cleaned_dataframe = clean_dataframe(dataframe)

        if write_cleaned_csv:
            save_cleaned_dataframe(cleaned_dataframe, CLEANED_DIR / particle_file.file_name)

        if generate_main_tables:
            counts = compute_dominant_element_counts(cleaned_dataframe, element_columns)
            dominant_records.append((particle_file, counts))
            grouped_counts[particle_file.group_key].append(counts)

    if not generate_main_tables:
        return None, None

    full_main = pd.DataFrame(index=element_columns)
    full_main.index.name = "Element"

    average_only = pd.DataFrame(index=element_columns)
    average_only.index.name = "Element"

    grouped_metadata: dict[tuple[str, int | float], list[ParticleFile]] = defaultdict(list)
    for particle_file, counts in dominant_records:
        full_main[particle_file.main_measurement_column] = counts.reindex(element_columns, fill_value=0).astype(int)
        grouped_metadata[particle_file.group_key].append(particle_file)

    ordered_group_keys = sorted(grouped_metadata, key=lambda key: grouped_metadata[key][0].sort_key)
    interleaved_main = pd.DataFrame(index=element_columns)
    interleaved_main.index.name = "Element"

    for group_key in ordered_group_keys:
        particle_group = sorted(grouped_metadata[group_key], key=lambda item: item.sort_key)
        for particle_file in particle_group:
            interleaved_main[particle_file.main_measurement_column] = full_main[particle_file.main_measurement_column].astype(int)

        average_series = pd.concat(grouped_counts[group_key], axis=1).mean(axis=1)
        average_series = average_series.apply(round_half_up).astype(int)
        average_column_name = group_key[0]

        interleaved_main[average_column_name] = average_series.reindex(element_columns, fill_value=0).astype(int)
        average_only[average_column_name] = average_series.reindex(element_columns, fill_value=0).astype(int)

    return interleaved_main.reset_index(), average_only.reset_index()


def write_excel_table(dataframe: pd.DataFrame, output_path: Path, sheet_name: str) -> None:
    ensure_output_dir(output_path.parent)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)


def build_total_summary_table(average_stat_tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    required_tables = {"smFPs", "mmFPs"}
    missing_tables = required_tables - set(average_stat_tables)
    if missing_tables:
        raise ValueError(f"Missing tables required to build Total: {sorted(missing_tables)}")

    sm_table = average_stat_tables["smFPs"].copy()
    mm_table = average_stat_tables["mmFPs"].copy()

    sm_total_column = RANGE_RULES["smFPs"]["total_column"]
    mm_total_column = RANGE_RULES["mmFPs"]["total_column"]
    key_columns = ["sampleID", "Dilutionfactor"]

    sm_element_columns = [column for column in sm_table.columns if column not in key_columns + [sm_total_column]]
    mm_element_columns = [column for column in mm_table.columns if column not in key_columns + [mm_total_column]]
    if sm_element_columns != mm_element_columns:
        raise ValueError("smFPs and mmFPs element columns differ; cannot build Total.")

    merged = sm_table.merge(
        mm_table,
        on=key_columns,
        how="inner",
        suffixes=("_sm", "_mm"),
        validate="one_to_one",
    )

    if len(merged) != len(sm_table) or len(merged) != len(mm_table):
        raise ValueError("smFPs and mmFPs sampleID / Dilutionfactor do not map one-to-one.")

    total_table = merged[key_columns].copy()
    total_table[TOTAL_TOTAL_COLUMN] = merged[sm_total_column] + merged[mm_total_column]
    for element in sm_element_columns:
        total_table[element] = merged[f"{element}_sm"] + merged[f"{element}_mm"]

    return total_table


def write_summary_workbook(
    average_stat_tables: dict[str, pd.DataFrame],
    average_main_table: pd.DataFrame,
) -> None:
    output_path = ANALYSIS_DIR / SUMMARY_OUTPUT_NAME
    ensure_output_dir(output_path.parent)
    total_summary_table = build_total_summary_table(average_stat_tables)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        ordered_row_sheets = [
            ("smFPs", average_stat_tables["smFPs"]),
            ("mmFPs", average_stat_tables["mmFPs"]),
            (TOTAL_SHEET_NAME, total_summary_table),
            ("0.1FPs", average_stat_tables["0.1FPs"]),
            ("0.1-0.4FPs", average_stat_tables["0.1-0.4FPs"]),
            ("0.4-1FPs", average_stat_tables["0.4-1FPs"]),
        ]
        for sheet_name, dataframe in ordered_row_sheets:
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
        average_main_table.to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)


def normalize_tasks(tasks: list[str]) -> list[str]:
    valid_tasks = set(RANGE_RULES) | {"clean", "main", "summary"}
    invalid_tasks = [task for task in tasks if task not in valid_tasks]
    if invalid_tasks:
        raise ValueError(f"Invalid task name(s): {invalid_tasks}")
    return tasks


def run_pipeline() -> None:
    selected_tasks = normalize_tasks(TASKS_TO_RUN)
    particle_files = discover_particle_files(INPUT_DIR)

    need_stat_tables = any(task in RANGE_RULES for task in selected_tasks) or ("summary" in selected_tasks)
    full_stat_tables: dict[str, pd.DataFrame] = {}
    average_stat_tables: dict[str, pd.DataFrame] = {}
    element_columns: list[str] | None = None

    if need_stat_tables:
        full_stat_tables, average_stat_tables, element_columns = compute_range_statistics(particle_files)
        for rule_name, dataframe in full_stat_tables.items():
            if rule_name in selected_tasks:
                write_excel_table(
                    dataframe=dataframe,
                    output_path=ANALYSIS_DIR / RANGE_RULES[rule_name]["output_name"],
                    sheet_name=RANGE_RULES[rule_name]["sheet_name"],
                )
                print(f"Wrote {rule_name}: {ANALYSIS_DIR / RANGE_RULES[rule_name]['output_name']}")

    need_main = ("main" in selected_tasks) or ("summary" in selected_tasks)
    need_clean = "clean" in selected_tasks

    full_main_table: pd.DataFrame | None = None
    average_main_table: pd.DataFrame | None = None

    if need_clean or need_main:
        if element_columns is None:
            first_dataframe = read_particle_dataframe(particle_files[0].path)
            element_columns = first_dataframe.columns.tolist()

        full_main_table, average_main_table = process_clean_and_main(
            particle_files=particle_files,
            element_columns=element_columns,
            write_cleaned_csv=need_clean,
            generate_main_tables=need_main,
        )

        if need_clean:
            print(f"Wrote cleaned CSV folder: {CLEANED_DIR}")

        if ("main" in selected_tasks) and (full_main_table is not None):
            write_excel_table(
                dataframe=full_main_table,
                output_path=ANALYSIS_DIR / MAIN_OUTPUT_NAME,
                sheet_name=MAIN_SHEET_NAME,
            )
            print(f"Wrote dominant-element statistics: {ANALYSIS_DIR / MAIN_OUTPUT_NAME}")

    if "summary" in selected_tasks:
        if average_main_table is None:
            raise ValueError("The summary task requires main statistics, which were not generated.")
        if not average_stat_tables:
            raise ValueError("The summary task requires FPs average tables, which were not generated.")

        write_summary_workbook(
            average_stat_tables=average_stat_tables,
            average_main_table=average_main_table,
        )
        print(f"Wrote summary workbook: {ANALYSIS_DIR / SUMMARY_OUTPUT_NAME}")


if __name__ == "__main__":
    run_pipeline()
