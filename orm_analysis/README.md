# ORM vs. In-House Digital Title Analysis

## Purpose

Evaluate whether digital-title promotion performs better when assigned to Open Road Media (ORM), which proposes retaining 35% of ORM-channel sales, or managed internally through promotion and BISAC changes.

## Cohorts

The SQL assigns each ISBN to one cohort. Cohort labels are administrative labels and are **not assumed to be the actual ORM start date**.

Robert's in-house test is stored under `March2026`, but the effective analysis date is **April 1, 2026**. BISAC work began around March 15, so March is excluded as an incomplete intervention month.

## Verified intervention dates

For ORM cohorts, the effective start is the first accounting period with positive ORM revenue. Monthly activity was checked to confirm that sales continued afterward and were not isolated transactions.

| Cohort | Label month | Verified start |
|---|---|---|
| oct2021 | October 2021 | December 2021 |
| Apr2022 | April 2022 | January 2022 |
| jun2022 | June 2022 | August 2022 |
| jul2022 | July 2022 | September 2022 |
| mar2023 | March 2023 | May 2023 |
| jun2023 | June 2023 | August 2023 |
| feb2025 | February 2025 | April 2025 |
| March2026 | March 2026 | April 2026, managed in-house by Robert |

`Apr2022` is unusual because ORM sales begin before its label. January 2022 contained positive ORM revenue across 36 of the cohort's 41 titles, followed by sustained activity, so January is retained as the observed sales start.

## Current comparison method

- Treat every cohort independently; no cohort's titles or dates affect another cohort.
- For each ORM cohort, compare its first 12 months beginning with the verified ORM start against the same 12 calendar months one year earlier.
- For Robert, use only complete post-intervention months. The current window is April-June 2026 versus April-June 2025.
- Exclude July 2026 for now because customer files are still arriving and the period is incomplete.
- Compare the same ISBN cohort with itself to control for different title mixes.
- Use the matching prior-year months to reduce seasonality effects.

## Overlap and fee treatment

Some cohort titles continue selling through customers other than ORM after ORM begins.

- `Post Gross Revenue` = ORM revenue + non-ORM revenue.
- `Post Net After Fee` = non-ORM revenue + 65% of ORM revenue.
- The 35% fee is applied only to ORM-channel revenue.
- `Gross Lift` compares total post revenue with the prior-year baseline.
- `Net Lift` compares revenue retained after the hypothetical ORM fee with the prior-year baseline.
- Overlap remains visible at title level and should not be credited automatically to ORM.

Because only 65% of ORM revenue is retained, ORM needs approximately 53.85% more gross revenue to equal otherwise-equivalent direct revenue.

## August refresh checklist

1. Confirm that July 2026 is closed and all digital customer files have arrived.
2. Rerun `orm_sales_query.sql` and replace the saved raw CSV and Parquet results.
3. Verify that April 2026 remains Robert's effective start.
4. Add July to Robert's window: April-July 2026 versus April-July 2025.
5. Recheck each ORM cohort's first positive ORM month and confirm sustained activity.
6. Refresh the verified-start audit, independent cohort analysis, title analysis, overlap analysis, category analysis, and top-title concentration.
7. Evaluate both gross lift and net lift after the 35% fee.
8. Call out incomplete data, returns/negative revenue, title concentration, and non-ORM overlap in the conclusions.

## Key files

- `orm_sales_query.sql`: supplied sales query and cohort definitions.
- `orm_sales_raw.parquet` / `orm_sales_raw.csv`: saved query output.
- `orm_sales_analysis.xlsx`: analysis workbook.
- `orm_verified_start_audit.csv`: month-level evidence used to establish ORM starts.
- `orm_independent_cohort_analysis.csv`: latest independent cohort results.

