# Investigation Report Macro (Excel → Word)

An Excel VBA macro that turns a reviewed transactions workbook into a Word investigation report. Suspicious MSISDNs are marked in red (or pink fill) during the review; the macro does the counting, the tables and the layout.

## What it does

1. Scans every sheet named `Analysis - <Country>` and collects the MSISDNs marked with red font or pink fill.
2. If the country name itself is red on its sheet, the whole corridor counts as suspicious.
3. Asks which countries to include: leave it empty for all, or type numbers like `1,3,5`.
4. Reads the send-money transactions from `Sheet1` and works out, per country, the total and suspicious amount and transaction count.
5. Builds a governorate table (transactions, unique senders, amount), sorted by transactions, with a total row.
6. Writes the Word report: a header table, one section per country, empty *Key Observations* and *Conclusion* to fill in, and an *Action Taken* section that starts with the number of flagged customers.
7. Saves it as `.docx` next to the workbook, adding `_2`, `_3`… instead of overwriting an older report.

## Input layout

- `Sheet1` columns: 10 = Governorate, 12 = Sender MSISDN, 13 = Type (`SEND_MONEY`), 14 = Amount, 16 = Destination Country.
- Cell `A1` of the first `Analysis - …` sheet holds the period, for example `1-Oct-2026 to 31-Oct-2026`. It's used in the report and the file name.

## Setup

1. In Excel press `Alt + F11`, then **File** → **Import File** and choose `InvestigationReport.bas`. Import it into `PERSONAL.XLSB` to have it in every workbook.
2. Set `AML_MANAGER` and `DEPUTY_MANAGER` at the top of the module.
3. Save the workbook, press `Alt + F8` and run `GenerateInvestigationReport`.

Word is driven through late binding, so no extra references are needed.
