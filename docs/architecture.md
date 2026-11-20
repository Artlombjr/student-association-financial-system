# Architecture

The system has four parts that share one Google Sheet: the claim web app, the sheet tools bound to the Sheet, a Python budget pipeline, and a standalone dashboard. All names and data are fictional.

## Components

| Part | Runs on | Role |
|---|---|---|
| Claim web app (`apps-script/claim-web-app`) | Standalone Apps Script project, deployed as a web app | Form, approval chain, PDFs, Drive filing, reminders, writes claims and ledger rows |
| Sheet tools (`apps-script/sheet-tools`) | Apps Script project bound to the claim log Sheet | Budget tracker, manual-expense sidebar, claim-tracking columns, digests and reminders, GL list |
| Budget pipeline (`budget-pipeline`) | Python, run locally | Budget workbook to `TrackerData.gs` |
| Dashboard (`dashboard`) | Static HTML in the browser | Reads a downloaded `.xlsx` export of the Sheet |

The web app and the sheet tools are separate projects because the web app executes as the deploying user for anonymous visitors, while the sheet tools need the Sheet's own UI (menus, sidebars, `onEdit`).

## Claim lifecycle

Statuses are stored in the Claims log, in the order a claim normally moves through them. The code reads them by prefix (`stageOf_`).

1. `Pending VP Finance signature` - set on submission. VP Finance signs on screen and picks a GL code for every line.
2. `Pending President signature` - the packet PDF is rebuilt with the VP signature and the President gets a one-time link.
3. `Pending NU Staff signed PDF` - the President has signed. The VP receives the PDF plus a ready-to-send subject and message, because the university mail server does not accept automated Gmail. The staff signer signs the PDF itself and replies; the VP forwards the reply to the script's Gmail keeping the `(Claim <id>)` tag in the subject.
4. `Approved - signed PDF from NU Staff saved` - `checkStaffSignedPdfs`, a trigger that runs every 10 minutes, finds the reply, saves the signed PDF in the claim folder, labels the thread and emails the PDF to VP Finance, who sends it to Accounts Payable by hand.

Side paths:

- `Returned by President` and `Returned by NU Staff`: the signer returns the claim with a note; it goes back to VP Finance, who can re-send it.
- `Sent back for changes`: VP Finance sends the claim to the student. The email contains an edit link (`?action=edit&id=...&token=...`) that reopens the form pre-filled. Resubmitting keeps the same claim ID and returns the claim to `Pending VP Finance signature`. The claim folder moves to a "Sent back" folder meanwhile.
- Statuses from earlier versions are still understood: `Pending NU Staff signature` (link-based staff signing) and `Pending VP Finance final approval` / `Approved - final PDF sent to VP Finance`.

Each step rotates the claim's token, so an old link stops working once its step is done. Reminders (`sendDailyReminders`, scheduled by `installDailyReminderTrigger`) go to whoever holds the claim after 5, 7 and 14 days at the same step.

Ledger status follows the claim status: pending steps map to `Pending`, approved claims to `Approved`, and sent-back claims to `Void`.

## Expense Ledger data model

One row per expense line. Claims write rows automatically; the VP adds manual rows through the sidebar.

| Column | Meaning |
|---|---|
| Entry ID | Claim line key, or `M-001`... for manual entries |
| Source | Claim or Manual |
| Date | Expense date |
| GL Code | Full GL code assigned at VP signing (or chosen in the sidebar) |
| GL Name, GL Group | Description and department group of the code |
| Charge to Department (optional) | Overrides the default department for a code shared by several budget lines |
| Description, Payee | What was bought and who is paid |
| Amount (incl. tax), GST, HST, Net Expense | Amounts; the tracker counts amounts including tax |
| Status | Pending, Approved, Posted or Void |
| Locked | Yes means the claim system never overwrites the row (it only updates Claim Status and Last Updated) |
| Claim ID, Line # | Link back to the claim |
| Claim Status, Last Updated | Mirrors the Claims log |
| Notes | Free text |

Columns A to T are owned by the claim system and tracker. `ClaimTracking.gs` adds three dropdown columns after them: Submitted (received / submitted to Accounts Payable / sent back), Cheque Received from AP, Cheque Picked Up. Its `onEdit` copies a choice to every row with the same Claim ID.

## How the tracker keys budget lines by GL code

`TrackerData.gs` (generated) holds the budget: departments, GL lines, line items under each GL line, and a GL map. `Tracker.gs` builds these tabs from it: Dashboard, Charts, GL Budget, Line Items, Tracker Ledger, GL Map, Dept Summary and Tracker Notes. A separate Revenue Ledger tab is created once and never deleted by a rebuild.

- Each GL row on a department tab of the budget workbook is a budget line; sub-items beneath it are line items.
- The GL Map decides which budget line a ledger row is charged to. By default a code goes to the line with the largest budget; "Charge to Department" overrides this. A few code-to-line decisions are kept as explicit extra keys.
- Tracker Ledger reads the Expense Ledger live and works out the budget line, department and semester for each row. GL Budget then sums Spent (approved and posted) and Pending per line with SUMIFS and shows remaining, percent used and a status (On track, Watch, Almost out, Over budget, Unbudgeted spend).
- Codes with spending but no budget appear as $0 lines, so spending outside the budget is visible.
- One department is split Fall/Spring, with the current semester shown by default.
- `rebuildBudgetTracker` replaces the tabs for a new budget but puts back line-item budget edits and GL Map changes unless the new budget changed that same line; a copy of the old Line Items is kept.
- Dept Summary ties each department's total back to the budget workbook so a mismatch is obvious.

## Budget pipeline

`parse.py` reads the department tabs of the budget workbook (Excel), normalises GL codes, reads line items and checks totals against the P&L. `build2.py` assembles the tracker model, `gsheet_export.py` writes `TrackerData.gs`, and `gl_master_gsheet.py` writes `GlCodes.gs` (the GL List). Real workbooks are never committed (`*.xlsx` is git-ignored).

## How the tests mock Google services

Apps Script has no local runtime. Each test reads the real `.gs` source as text and runs it in a Node `vm` context in which the Google globals (`SpreadsheetApp`, `DriveApp`, `GmailApp`, `HtmlService`, `ScriptApp`, `PropertiesService` and so on) are small in-memory fakes. The fakes hold sheets as arrays, files as objects and Gmail threads as lists, so a test can submit a claim, walk it through each signer, and then assert on the Sheet rows, emails sent, files saved and status changes.

The tracker test uses a strict spreadsheet mock that only knows real Apps Script method names, so a mistyped call throws. It then checks the tab sizes and every formula written. These tests catch structure and logic errors; they cannot evaluate Google Sheets formulas, so numbers were confirmed in a real Sheet.

Every test is plain Node with built-in modules only. `tests/paths.js` maps file names to locations.
