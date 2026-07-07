# FIFA World Cup 2026 Player Performance: Data Cleaning & Exploratory Analysis in Excel

*A case study in diagnosing data quality issues, cleaning a large dataset, and building exploratory PivotTable/PivotChart analysis.*

## Overview

This project takes a raw Kaggle dataset of FIFA World Cup 2026 player records, diagnoses and resolves a significant data quality issue (53,352 exact duplicate rows), and builds an exploratory analysis of player physical attributes using Excel PivotTables and PivotCharts. The goal was to practice a full, realistic data-cleaning workflow — not just apply a tool, but verify the problem, choose the right method, and validate the result before drawing conclusions.

## The Dataset

**Source:** FIFA World Cup 2026 Player Performance dataset (Kaggle), selected-columns version.

- **Columns:** `player_id`, `player_name`, `age`, `nationality`, `team`, `jersey_number`, `position`, `height_cm`, `weight_kg`, `preferred_foot`
- **Raw size:** 54,600 rows across 48 national teams

## The Problem: Duplicate Records

On first inspection, player names and stats appeared to repeat multiple times throughout the dataset. Before removing anything, I confirmed the nature of the duplication rather than assuming it:

- Checked whether repeats were **exact full-row duplicates** (same player, same stats, everything identical) versus legitimate repeat appearances.
- Used **conditional formatting** (Highlight Duplicate Values) to visually confirm the pattern before writing any formulas.
- Built a helper column concatenating all fields into a single string, then used `COUNTIF` to count occurrences of each full-row combination — confirming these were true exact duplicates.

**Finding:** every unique player record was repeated 5 or more times throughout the file, consistent with the same export or merge having been appended to the dataset multiple times.

## Data Cleaning Process

1. Converted the range to an **Excel Table** (`Ctrl+T`) for stable, self-adjusting formulas on a large dataset.
2. Verified the duplication was uniform (a clean repeat) rather than a messier partial overlap, using `COUNTIFS` across a subset of identifying columns.
3. Sorted by player name to cluster duplicate rows together for a final visual check.
4. Applied **Data → Remove Duplicates** with all columns selected, so a row was only removed if every field matched exactly.
5. Verified the result by re-running the duplicate count on the cleaned data — every value returned `1`.

| Metric | Raw Dataset | Cleaned Dataset |
|---|---|---|
| Total rows | 54,600 | 1,248 |
| Duplicate rows removed | — | 53,352 |
| Unique players | 1,248 | 1,248 |
| Teams represented | 48 | 48 |

> **Note:** while cleaning, the workbook briefly became unresponsive due to volatile whole-column formulas (e.g. `COUNTIF(A:A, ...)`) recalculating on every edit across tens of thousands of rows. This was resolved by switching calculation to **Manual** during cleanup and bounding formula ranges to the actual data size — a practical lesson in working with large datasets in Excel.

## Exploratory Analysis: PivotTable & PivotChart

With a clean, deduplicated dataset (1,248 unique players), I built a PivotTable/PivotChart to explore physical attributes by position and preferred foot:

- **Rows:** `position`
- **Columns:** `preferred_foot`
- **Values:** Average of `age`, `height_cm`, `weight_kg`
- **Filters:** `team`, `nationality` — `jersey_number` and `player_name` excluded, as identifiers add no analytical value at the aggregate level

This layout supports comparisons such as *"do goalkeepers differ in average height from midfielders?"* or *"is there a physical difference between left- and right-footed players at the same position?"* — both filterable down to a single team or nationality.

*(Insert PivotChart screenshot here)*

## Key Insights

Averages across the full cleaned dataset (all 48 teams):

| Position | Avg. Age | Avg. Height (cm) | Avg. Weight (kg) |
|---|---|---|---|
| Goalkeeper | 28.7 | 189.1 | 78.7 |
| Defender | 26.7 | 183.5 | 76.4 |
| Midfielder | 25.5 | 178.9 | 75.1 |
| Forward | 25.6 | 178.6 | 74.1 |

- **Goalkeepers** are, on average, the **oldest (28.7 yrs)** and **tallest (189.1 cm)** position — consistent with the position's physical and experience demands.
- **Forwards and midfielders** are the youngest positions on average (~25.5 yrs) and closely matched in height and weight.
- **Preferred foot:** right-footed players outnumber left-footed roughly **3 to 1** (928 vs. 320) across the dataset, in line with general population footedness patterns.
- Height and weight scale predictably with position — Goalkeeper > Defender > Midfielder ≈ Forward — matching expectations from real-world football physiology.

## Tools & Skills Demonstrated

`Excel` · `Data Cleaning` · `COUNTIF/COUNTIFS` · `PivotTables` · `PivotCharts` · `Data Validation` · `Large Dataset Troubleshooting`

- Diagnosing a data quality issue before acting on it
- Formula-based duplicate detection (`COUNTIF`, `COUNTIFS`, `CONCATENATE`)
- Data cleaning with Excel Tables and Remove Duplicates
- Performance troubleshooting on large datasets (manual calculation mode, bounded ranges)
- Multi-dimensional PivotTable/PivotChart design (rows, columns, values, filters)
- Translating pivot output into a written analytical finding

## Reflection

The most valuable part of this project wasn't the cleanup itself — it was slowing down to verify what kind of duplication I was dealing with before deleting anything, and building a habit of checking results (row counts, re-run duplicate checks) rather than trusting a single pass. Next time, I'd reach for Power Query earlier for a dataset this size, since it handles duplicate removal and large-data transforms more efficiently than live worksheet formulas.
