# AML Analytics Toolkit (Excel VBA)

Excel VBA tools I built to analyze Western Union send-money activity for AML reviews. They take a raw transaction export, turn it into per-country risk tables, and end with a Word investigation report. The analyst still makes every call; the tools do the counting, matching and first drafts.

Patterns they look for include funnel accounts (many senders, one receiver), transfers sitting at the transfer limit, newly opened mule wallets, agent and recruiter concentration, and the same receiver hidden behind different spellings.

## How the pieces fit

1. **Analyze**: `WUReportAnalysis_Final_Version` reads the export and builds an `Analysis - <COUNTRY>` sheet for each country you pick. Cell `A1` gets the data period.
2. **Review**: the analyst goes through the tables and marks suspicious MSISDNs in red (or pink fill).
3. **Collect**: `Collect_Red_Phones_Correctly` gathers every marked number, with its destination country, on one sheet.
4. **Report**: `GenerateInvestigationReport` turns the marked sheets into a Word investigation report, and `ExportDashboardToWord` fills the *WU Accounts* report template.

`ScreenBeneficiaryNames` is a separate check for junk or random beneficiary names.

## Modules

### WUReportAnalyzer.bas

The core analysis. A settings form asks for the target countries (or `ALL COUNTRIES`), the thresholds, and which optional analyses to run.

- **Primary country per account**: each receiving account is tied to the country where it got the most transfers, so one account isn't split across countries.
- **Account risk**: counts transactions or unique senders per receiving account (your choice). At or above the threshold it's *Suspicious*; from 3 up to the threshold it's *Review*. Each table has totals, the number of countries, and the bank / wallet split.
- **Geography**: countries and governorates split into high-risk (with suspicious activity) and general, with the bank / mobile wallet / cash-pickup mix, plus a nationality overview of the senders.
- **Suspicious senders**: one row per sender with every transfer side by side (amount, account, country, receiver, MTCN, date) and a ready-to-paste Arabic note.
- **Many-to-one**: five or more transfers going to the same Model ID, or to receiver names that sound the same (Metaphone), so spelling variations still group together.
- **Cash-in agents** (optional, from `Sheet2`): links suspicious senders to the agents who funded their wallets, flags agents above a threshold, and shows what share of their customers are in the agent's own city and which countries they send to.
- **Raw data**: the country's transactions with added count columns (receiver name, account, Model ID, mother's name, agents) and a filter row.

### InvestigationReport.bas

Builds the Word report from the sheets the analyst has marked.

- Reads the marked numbers from every `Analysis - <COUNTRY>` sheet (red font or pink fill). A country name written in red marks the whole corridor.
- Lets you pick the countries: leave it empty for all, or type numbers like `1,3,5`.
- Per country: total vs. suspicious amount and transactions, and a governorate table sorted by transactions.
- Drafts the **Key Observations** from the data. A line is written only when its pattern is actually there:
  - the suspicious share of the corridor's volume
  - the top two governorates
  - transfers exactly at, or just below, the transfer limit, with real examples
  - unique customers by nationality
  - receiving channels, with the number of unique bank accounts and wallets
  - one declared purpose covering 60% or more of the transfers
  - wallets opened less than 30 days before the first suspicious transfer, including same-day ones
  - one recruiter behind five or more flagged customers
  - one receiver collecting from three or more senders
  - most amounts in exact multiples of 250,000 IQD
  - senders spread over four or more governorates but behaving the same way
  - flagged customers moving 1.5× or more than a typical sender in the corridor
- Drafts a conclusion, starts *Action Taken* with the flagged-customer count, and saves the `.docx` next to the workbook without overwriting older reports.

The transfer limit and excluded recruiter IDs are constants at the top of the module. Everything it writes is a draft to review before sending.

### DashboardToWord.bas

Fills the *WU Accounts* report from a Word template.

- Reads the headline numbers from the analysis sheet: period, suspicious amount and transactions, unique suspicious senders, suspicious accounts split by bank and wallet, and the top governorate.
- Replaces these tags in the template: `(Today)`, `(From-To)`, `(Total$)`, `(Trx.no.)`, `(TUC)`, `(TB/W)`, `(TB)`, `(TW)`, `(TopG)`, `(TopG$)`.
- Pastes clean copies of the governorate and country tables at `(Governorate Table)` and `(Country Table)`, with unneeded columns removed, headers renamed and colors stripped.
- Saves `WU Accounts <period>.docx` next to the workbook.

### CollectRedPhones.bas

Collects every red or pink MSISDN (column M) from the country sheets into one `Red_Phones` sheet, with its destination country (column R). It reads the displayed color, so conditional formatting counts too, and it accepts any shade of red or pink rather than one exact color code.

### NameScreening.bas

Scores each beneficiary `full_name` for signs of a junk or random entry, and copies the flagged ones to a `Gibberish_Review` sheet. It's a triage aid: unusual real names can land in *Review*, so a person always makes the call.

| Signal | Points |
|---|---|
| Placeholder word (`test`, `n/a`, `unknown`…) or keyboard mash (`asdfgh`, `qwerty`…) | +5 |
| Same character 3 or more times in a row | +4 |
| Digits | +3 |
| Shorter than 3 characters | +3 |
| A word with no vowels | +3 |
| Unexpected symbols | +2 |
| Very low vowel ratio (under 15%) | +2 |
| Five or more consonants in a row | +2 |

A score of 2 or more goes to *Review*, 5 or more to *High risk*. Running it again updates the same columns instead of adding new ones.

## Setup

1. In Excel press `Alt + F11`, then **File** → **Import File** for each file. Import them into `PERSONAL.XLSB` to use them in any workbook.
2. The analyzer uses a settings form (`frmSettings`) and a class (`clsBtn`) that aren't in this repo yet. The other modules work on their own.
3. Run a macro with `Alt + F8`.

Word is driven through late binding, so no extra references are needed.

## Data layout

Transactions export (`Sheet1`, columns A:AB): 1 Actor ID, 3 KYC date, 4 Recruiter, 5 Nationality, 10 Governorate, 12 Sender MSISDN, 13 Type (`SEND_MONEY`), 14 Amount, 15 Receiver name, 16 Destination country, 17 MTCN, 18 Send date, 20 Created at, 22 Purpose, 25 Delivery method, 26 Receiving account, 27 Provider, 28 Model ID. The analyzer writes the receiver-name Metaphone to column 29.

Cash-in data (`Sheet2`): 5 Customer MSISDN, 7 Amount, 10 Agent MSISDN, 17 Agent city.
