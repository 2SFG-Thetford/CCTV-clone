# CCTV Sediment Tray Log

This folder contains code (Power Query/Visual Basic/Excel formulae) used in CCTV Sediment Tray Log.

## Explanation

<details>
  <summary>Query.pq</summary>

  + Look in the folder for files
    * Go to the folder:
    `\\2sfg.net\filestore\Thetford Filestore\Technical\02 Audits\Internal Audits\2024-2025\CCTV Audit\CCTV Sediment tray\Testing`
    * Make a list of all files.
  + Keep only Excel files
    * Filter to keep only files that end with `.xlsx` (case-insensitive).
  + Pick the latest file
    * Sort the Excel files by **Date modified**, newest first.
    * Take the first file in the sorted list.
    * If there are no files, give an **error**.
  + Load the workbook
    * Open the Excel file.
    * Take the **first sheet**.
  + Promote headers
    * Make the first row of the sheet the column names.
  + Clean the column names
    * Remove extra spaces from column names.
  + Get existing references from Table1
    * Load `Table1` from the current workbook.
    * Get the list of existing `Ref` values as text.
  + Keep only new references
    * Filter the audit sheet to **exclude rows** where `Ref` already exists in `Table1`.
  + Fix the main Date column
    * If there is a `Date` column:
      * Convert it to proper **date type**.
  + Identify fixed vs extra columns
    * Fixed columns (never collapse): `"Ref","Date","Questionnaire Type","Site","Location","Questionnaire","Auditor","Audit Score","Audit Score Comments"`
    * Extra columns (may have duplicates that need collapsing).
  + Collapse duplicate extra columns
    * For each duplicate column:
      * Combine multiple columns of the same name into **one column**.
      * Keep the **first non-null value** for each row.
  + Remove completely empty columns
    * Delete columns that have **no data at all**.
  + Update Audit Score Comments
    * Create a new `Audit Score Comments` column based on `Audit Score`:
      * `Green` → `Pass`
      * `Amber` → `Caution`
      * `Red` → `Fail`
      * Anything else → keep the original comment
  + Move Audit Score Comments next to Audit Score
    * Make sure `Audit Score Comments` appears **immediately after `Audit Score`**.
  + Sort by Ref
    * Sort the table **ascending by Ref**.
  
  Super simple summary:\
  “Take the latest CCTV sediment tray audit Excel file, remove any rows that already exist in Table1, clean up duplicate columns, convert dates, replace audit score comments with friendly labels, remove empty columns, reorder so comments follow scores, and return a nice sorted table of only new audit records.”
</details>
