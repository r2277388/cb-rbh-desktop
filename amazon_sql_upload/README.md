# Amazon SQL Upload Data Pipeline - Process Overview

This process combines, cleans, and standardizes weekly Amazon sales, inventory, traffic, and catalog data for reporting and analysis.

## Steps

1. **File Selection**
   - Automatically finds the latest weekly CSV files for sales, inventory, traffic, and catalog from designated folders.

2. **Data Cleaning**
   - Cleans numeric columns (removes $, commas, fills missing values with zero).
   - Pads ASINs and ISBNs with leading zeroes for consistent formatting.
   - Removes unwanted publishers and availability statuses from catalog data.
   - Drops rows with duplicate or missing ASINs/ISBNs.

3. **Data Merging**
   - Combines all sources into a single DataFrame using ASIN as the key.
   - Ensures each ASIN is unique in the final output.

4. **ISBN Assignment**
   - Assigns ISBNs from catalog data where available.
   - If missing, attempts to assign from EAN, Amazon ISBN, or Model Number if they match a known ISBN.
   - Applies manual ASIN-to-ISBN overrides for special cases.
   - Marks ASINs with no valid ISBN as "NO_ISBN".

5. **Filtering**
   - Removes rows for specific ASINs (as defined in a removal list).
   - Excludes products with titles ending in "anglais".
   - Drops rows where all numeric columns are zero.

6. **Column Selection & Renaming**
   - Keeps only relevant columns: ASIN, ISBN (renamed to External ID), and key numeric fields.
   - Renames columns for clarity (e.g., "Ordered Units" to "Customer Orders").

7. **Output**
   - Saves the final cleaned and merged data to an Excel file, ready for reporting.
   - Provides summary statistics and diagnostics in the console.

## Purpose

This pipeline ensures that Amazon weekly data is accurate, standardized, and ready for business analysis, eliminating duplicates, cleaning values, and providing a unified report for further use.