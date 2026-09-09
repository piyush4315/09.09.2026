# Sort toggle — how to use

The workbook **Search any withot error.xlsx** can now sort any column
ascending / descending. Two ways to drive it:

## Option A — no macros needed (works everywhere, incl. Excel Online)

1. On **Live Search**, scroll just right of the table to the navy
   **SORT BY** box (cell **AJ2**) and pick a column, e.g. `C - Buyer`.
2. Pick a direction in the **ORDER** box (cell **AL2**):
   ▲ Ascending or ▼ Descending.
3. The sorted header shows ▲/▼, the status bar in row 3 names the sort,
   and rows reorder **inside each group** — every ∑ / TOTAL MATCHED /
   REMAINING / GRAND TOTAL row stays exactly where it was.
4. Choose `(none)` in SORT BY to restore the original order.

Notes:

- Dates (`dd.mm.yyyy`) sort chronologically, numbers numerically,
  text A–Z (case-insensitive), and blanks always go last — just like
  Excel's own sort.
- Sorting never breaks the search: change the search box and the new
  results come out already sorted.
- On **Final Calculation Sheet** the filter arrows on row 3 already
  sort (click arrow → Sort A to Z / Z to A). Its range was narrowed to
  rows 4–40 so the total rows 41–43 can never be dragged into the data.

## Option B — one-click sorting by clicking a header (needs macros)

Excel cannot run "click header = sort" from an `.xlsx` file, so true
click-to-toggle ships as the macro module **`SortToggle.bas`**.
One-minute install (Excel for Windows / Mac; Excel Online cannot run
macros — use Option A there):

1. Open the workbook, then **File → Save As → Excel Macro-Enabled
   Workbook (\*.xlsm)**. (An `.xlsx` cannot store macros.)
2. Show the Developer tab if needed: **File → Options → Customize
   Ribbon → tick Developer**. Then **Developer → Visual Basic**.
3. In the VBA editor: **File → Import File → pick `SortToggle.bas`**.
4. In the Project pane double-click the **Live Search** sheet module
   and paste at the top:

   ```vb
   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
       SortToggle.ToggleLiveSearchSort Target
   End Sub

   Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
       If Not Intersect(Target, Me.Range("A5:AI5")) Is Nothing Then
           SortToggle.ResetLiveSearchSort
           Cancel = True
       End If
   End Sub
   ```

5. Double-click the **Final Calculation Sheet** sheet module and paste:

   ```vb
   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
       SortToggle.ToggleFinalSort Target
   End Sub
   ```

6. **File → Close and Return to Microsoft Excel**, save, and click
   **Enable Content** when Excel asks about macros.

Then:

- **Live Search** — single-click any header in row 5 to sort by it;
  click the same header again to flip ▲/▼; double-click a header to
  reset to the original order.
- **Final Calculation Sheet** — single-click any header in row 3 to
  sort the 37 data rows; click again to flip. Total rows are untouched.

## Troubleshooting

- *"The macro doesn't run"* — the file must be saved as `.xlsm`, and
  macros must be enabled (Enable Content). Macro security lives under
  File → Options → Trust Center.
- *"I clicked a Final Calculation Sheet header and nothing visibly
  changed"* — a filter may be hiding rows; clear it (Data → Clear) and
  click again.
- *Date columns on Final Calculation Sheet* sort as text
  (`dd.mm.yyyy`), because that is how they are stored. On Live Search
  they sort chronologically.
- The sort boxes live in the helper area (columns AJ/AL, rows 1–2)
  and the sort list in AJ5:AJ40 — don't delete those columns.
