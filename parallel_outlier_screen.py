from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd


# =============================================================================
# CONFIGURATION -- user-adjustable parameters
# =============================================================================
# Defaults to the example analysis folder next to this script; change for real use.
ANALYSIS_DIR = Path(__file__).resolve().parent / "analysis"
OUTPUT_PATH = ANALYSIS_DIR / "parallel_outlier_screen.xlsx"

MIN_VALUE_THRESHOLD = 100
RATIO_THRESHOLD = 2
TABLES_TO_CHECK = ["smFPs", "mmFPs", "0.1FPs", "0.1-0.4FPs", "0.4-1FPs", "mainFPs"]

TABLE_FILE_MAP = {
    "smFPs": ANALYSIS_DIR / "smFPs.xlsx",
    "mmFPs": ANALYSIS_DIR / "mmFPs.xlsx",
    "0.1FPs": ANALYSIS_DIR / "0.1FPs.xlsx",
    "0.1-0.4FPs": ANALYSIS_DIR / "0.1-0.4FPs.xlsx",
    "0.4-1FPs": ANALYSIS_DIR / "0.4-1FPs.xlsx",
    "mainFPs": ANALYSIS_DIR / "main.xlsx",
}

# =============================================================================
# End of configuration
# =============================================================================

MAIN_MEASUREMENT_PATTERN = re.compile(
    r"^(?P<sample_id>.+)-(?P<measurement>\d+)-(?P<dilution>\d+(?:\.\d+)?)x$",
    re.IGNORECASE,
)


def split_sample_for_sort(sample_id: str) -> tuple[str, int, str]:
    matched = re.match(r"^(?P<prefix>[A-Za-z]+)(?P<number>\d+)$", sample_id)
    if matched:
        return matched.group("prefix").lower(), int(matched.group("number")), sample_id.lower()
    return sample_id.lower(), -1, sample_id.lower()


def sort_dataframe_by_sample(dataframe: pd.DataFrame, sample_column: str = "sampleID") -> pd.DataFrame:
    if dataframe.empty:
        return dataframe

    sort_keys = dataframe[sample_column].astype(str).map(split_sample_for_sort)
    sortable = dataframe.copy()
    sortable["_prefix"] = sort_keys.map(lambda item: item[0])
    sortable["_number"] = sort_keys.map(lambda item: item[1])
    sortable["_sample_lower"] = sort_keys.map(lambda item: item[2])
    sortable = sortable.sort_values(
        by=["_prefix", "_number", "_sample_lower"],
        kind="stable",
    ).drop(columns=["_prefix", "_number", "_sample_lower"])
    return sortable.reset_index(drop=True)


def prepare_positive_values(series: pd.Series) -> pd.Series:
    numeric_series = pd.to_numeric(series, errors="coerce").dropna()
    return numeric_series[numeric_series > 0]


def evaluate_positive_group(
    values: pd.Series,
    measurement_names: pd.Series,
    sample_id: str,
    table_name: str,
    element: str,
) -> dict[str, object] | None:
    positive_values = prepare_positive_values(values)
    if len(positive_values) != len(values.dropna()):
        return None
    if len(positive_values) < 2:
        return None

    min_value = float(positive_values.min())
    max_value = float(positive_values.max())
    if min_value < MIN_VALUE_THRESHOLD:
        return None

    ratio = max_value / min_value
    if ratio < RATIO_THRESHOLD:
        return None

    min_index = positive_values.idxmin()
    max_index = positive_values.idxmax()

    return {
        "sampleID": sample_id,
        "tableName": table_name,
        "element": element,
        "measurementCount": int(len(positive_values)),
        "minMeasurement": str(measurement_names.loc[min_index]),
        "maxMeasurement": str(measurement_names.loc[max_index]),
        "minValue": int(min_value) if min_value.is_integer() else round(min_value, 4),
        "maxValue": int(max_value) if max_value.is_integer() else round(max_value, 4),
        "ratio": round(ratio, 4),
    }


def scan_fp_table(table_name: str, file_path: Path) -> list[dict[str, object]]:
    dataframe = pd.read_excel(file_path)
    measurement_rows = dataframe[dataframe["sampleID"].astype(str).str.contains("-")].copy()
    measurement_rows["base_sample"] = measurement_rows["sampleID"].astype(str).str.replace(
        r"-\d+$",
        "",
        regex=True,
    )

    evidence_rows: list[dict[str, object]] = []
    element_columns = list(measurement_rows.columns[3:-1] if "base_sample" in measurement_rows.columns else measurement_rows.columns[3:])
    if "base_sample" in element_columns:
        element_columns.remove("base_sample")

    for sample_id, group in measurement_rows.groupby("base_sample", sort=False):
        measurement_names = group["sampleID"]
        for element in element_columns:
            evidence = evaluate_positive_group(
                values=group[element],
                measurement_names=measurement_names,
                sample_id=str(sample_id),
                table_name=table_name,
                element=element,
            )
            if evidence is not None:
                evidence_rows.append(evidence)

    return evidence_rows


def extract_main_measurement_groups(columns: Iterable[str]) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = {}
    for column in columns:
        matched = MAIN_MEASUREMENT_PATTERN.match(str(column))
        if not matched:
            continue
        sample_id = matched.group("sample_id")
        groups.setdefault(sample_id, []).append(str(column))
    return groups


def scan_main_table(file_path: Path) -> list[dict[str, object]]:
    dataframe = pd.read_excel(file_path)
    indexed = dataframe.set_index("Element")
    measurement_groups = extract_main_measurement_groups(indexed.columns)

    evidence_rows: list[dict[str, object]] = []
    for sample_id, measurement_columns in measurement_groups.items():
        subset = indexed[measurement_columns]
        for element, row in subset.iterrows():
            evidence = evaluate_positive_group(
                values=row,
                measurement_names=pd.Series(measurement_columns, index=measurement_columns),
                sample_id=sample_id,
                table_name="mainFPs",
                element=str(element),
            )
            if evidence is not None:
                evidence_rows.append(evidence)

    return evidence_rows


def collect_evidence() -> pd.DataFrame:
    evidence_rows: list[dict[str, object]] = []

    for table_name in TABLES_TO_CHECK:
        file_path = TABLE_FILE_MAP[table_name]
        if not file_path.exists():
            raise FileNotFoundError(f"Missing result file: {file_path}")

        if table_name == "mainFPs":
            evidence_rows.extend(scan_main_table(file_path))
        else:
            evidence_rows.extend(scan_fp_table(table_name, file_path))

    evidence_columns = [
        "sampleID",
        "tableName",
        "element",
        "measurementCount",
        "minMeasurement",
        "maxMeasurement",
        "minValue",
        "maxValue",
        "ratio",
    ]

    evidence_df = pd.DataFrame(evidence_rows, columns=evidence_columns)
    if evidence_df.empty:
        return evidence_df

    table_order = {name: index for index, name in enumerate(TABLES_TO_CHECK)}
    sort_keys = evidence_df["sampleID"].astype(str).map(split_sample_for_sort)
    evidence_df["_prefix"] = sort_keys.map(lambda item: item[0])
    evidence_df["_number"] = sort_keys.map(lambda item: item[1])
    evidence_df["_sample_lower"] = sort_keys.map(lambda item: item[2])
    evidence_df["_table_order"] = evidence_df["tableName"].map(table_order)
    evidence_df = evidence_df.sort_values(
        by=["_prefix", "_number", "_sample_lower", "_table_order", "ratio", "element"],
        ascending=[True, True, True, True, False, True],
        kind="stable",
    ).drop(columns=["_prefix", "_number", "_sample_lower", "_table_order"])
    return evidence_df.reset_index(drop=True)


def build_sample_summary(evidence_df: pd.DataFrame) -> pd.DataFrame:
    summary_columns = ["sampleID", "hitCount", "tablesHit", "elementsHit", "maxRatio"]
    if evidence_df.empty:
        return pd.DataFrame(columns=summary_columns)

    summary_rows: list[dict[str, object]] = []
    for sample_id, group in evidence_df.groupby("sampleID", sort=False):
        table_names = sorted(group["tableName"].astype(str).unique(), key=lambda name: TABLES_TO_CHECK.index(name))
        element_names = sorted(group["element"].astype(str).unique())
        summary_rows.append(
            {
                "sampleID": str(sample_id),
                "hitCount": int(len(group)),
                "tablesHit": "; ".join(table_names),
                "elementsHit": "; ".join(element_names),
                "maxRatio": round(float(group["ratio"].max()), 4),
            }
        )

    summary_df = pd.DataFrame(summary_rows, columns=summary_columns)
    return sort_dataframe_by_sample(summary_df)


def write_output(sample_summary_df: pd.DataFrame, evidence_df: pd.DataFrame) -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        sample_summary_df.to_excel(writer, sheet_name="sample_summary", index=False)
        evidence_df.to_excel(writer, sheet_name="evidence_detail", index=False)


def print_summary(sample_summary_df: pd.DataFrame) -> None:
    print(f"Number of outlier samples: {len(sample_summary_df)}")
    if sample_summary_df.empty:
        print("No samples met the thresholds.")
        return

    for _, row in sample_summary_df.iterrows():
        print(f"{row['sampleID']}: maxRatio={row['maxRatio']}")


def main() -> None:
    evidence_df = collect_evidence()
    sample_summary_df = build_sample_summary(evidence_df)
    write_output(sample_summary_df, evidence_df)
    print_summary(sample_summary_df)
    print(f"Wrote screening results: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
