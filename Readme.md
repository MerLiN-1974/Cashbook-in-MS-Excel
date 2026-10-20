# Construction Accounts Manager

An offline Windows desktop app for recording credit and debit transactions across multiple construction sites. Each site is its own Excel file in a folder you choose. No internet and no Excel installation are needed to run it.

---

## Contents

1. [Status](#status)
2. [Features](#features)
3. [Quick start](#quick-start)
4. [The window](#the-window)
5. [Folder layout](#folder-layout)
6. [Site file format](#site-file-format)
7. [Ordering rules and manual edits in Excel](#ordering-rules-and-manual-edits-in-excel)
8. [What happens when you add a transaction](#what-happens-when-you-add-a-transaction)
9. [Validation rules](#validation-rules)
10. [Backups](#backups)
11. [Settings](#settings)
12. [Errors and how the app responds](#errors-and-how-the-app-responds)
13. [Tech stack](#tech-stack)
14. [Project structure](#project-structure)
15. [Development setup and tests](#development-setup-and-tests)
16. [Building the Windows executable](#building-the-windows-executable)
17. [Limitations](#limitations)
18. [Out of scope](#out-of-scope)

---

## Status

| Part | State |
|---|---|
| Excel logic, ordering, formulas, backups, amount/INR helpers | Written; 28 unit tests pass |
| Window (layout follows the supplied mockup) | Written; 13 tests drive the real window (add entries, back-dated entries, manual Excel edits, locked file, Add Site, calendar, validation) |
| Formulas, date sorting and number formats | Checked by opening generated files in LibreOffice |
| All 41 tests | Run on Linux with a virtual display, not on Windows |
| Windows `.exe` | **Not built yet**; `build.bat` must be run on Windows |
| Not yet verified | Running on Windows, the hidden-folder attribute on Windows, opening files in Microsoft Excel itself, antivirus behaviour of the `.exe`, appearance on high-DPI screens |

---

## Features

- **One Excel file per site**, all in one folder you pick (type the path or Browse).
- **Site Selection** drop-down lists every valid site file. The last entry, **Add Site...**, creates a new file in that folder.
- **Current Balance** card, recalculated from the file every time, so manual edits are picked up.
- **Transaction form:** date, bill date, bill number, credit/debit, amount, description.
- **Calendar pickers** for Date (prefilled with today, editable) and Bill Date (empty, with a clear button).
- **Recent Transactions table** and **Total Credit / Total Debit** footer.
- **Sorted by date automatically**, including rows typed straight into Excel. Rows in the wrong format go to the bottom.
- **Sr No and running balance are Excel formulas.**
- **Backups** after every change, kept 3 days, in a hidden `.backup` folder.
- **INR** with Indian digit grouping (₹1,25,000.00).
- **Remembers** the last folder and site.

---

## Quick start

1. Run `ConstructionAccountsManager.exe` (no installation).
2. Under **Folder Path**, type a path or click **Browse** and choose the folder for the site files (an empty folder is fine).
3. In **Site Selection**, choose **Add Site...**, then enter a name such as `Pune Villa`. This creates `Pune Villa.xlsx`.
4. To record an opening balance, add it as a normal **Credit** transaction.
5. Fill in the form and click **Add Transaction** (or press Enter in any text field).

Close a site's file in Excel before adding transactions to it from the app.

---

## The window

Top to bottom, as in the mockup:

| Area | Contents |
|---|---|
| Folder band | Folder path box, **Browse** button, hint text |
| Site Selection | Drop-down of sites ending in **Add Site...** |
| Current Balance | Balance of the selected site: green if zero or positive, red if negative |
| New Transaction | The form below |
| Recent Transactions | Table of the selected site's rows (latest 100; the title shows the count if there are more) |
| Footer | Total Credit, Total Debit (all rows), app name and version |

**Form fields**

| Field | Default | Required | Notes |
|---|---|---|---|
| Date of Transaction | Today | Yes | Typed as `DD-MM-YYYY` or picked from the calendar button |
| Bill Date | Empty (hint `DD-MM-YYYY`) | No | Calendar button and **×** clear button |
| Bill Number | Empty | No | Any text |
| Credit / Debit | Credit (C) | Yes | |
| Amount | Empty (hint `0.00`) | Yes | `12500`, `12,500.50` and `₹1,25,000` are accepted |
| Description / Purpose | Empty | Yes | For example Cement, Labor, Interior |

**Table columns:** Sr No, Date, Bill No, Bill Date, C/D, Description, Credit (₹), Debit (₹), Balance (₹). Blank cells show `-`. Rows in the wrong format are shown in red. This order is for display only; the Excel file uses the order in [Site file format](#site-file-format).

**Behaviour**

- **Add Transaction** is disabled until a site is selected.
- After a successful add: Amount, Description, Bill Number and Bill Date are cleared, Date returns to today, the Credit/Debit choice is kept, and focus returns to Amount.
- Cancelling **Add Site...** returns to the previously selected site.
- When the window regains focus (for example coming back from Excel) the site is reloaded.
- A message line beside the button reports successes, warnings and errors.

---

## Folder layout

```
Accounts/                       <- the folder chosen in the app
├── .backup/                    <- hidden, created automatically (sorts to the top when visible)
│   ├── Pune Villa_2026-10-17_14-32-05.xlsx
│   └── Baner Plaza_2026-10-16_09-10-41.xlsx
├── Baner Plaza.xlsx            <- one file = one site
└── Pune Villa.xlsx
```

- File name = site name. Renaming a file renames the site; deleting a file removes the site.
- The drop-down lists only `.xlsx` files whose column headings match exactly (below). Other Excel files, Excel's `~$` lock files and folders are ignored.
- A file copied in from elsewhere appears as a site if its headings match.

---

## Site file format

One sheet named **Accounts**. Row 1 holds the headings (frozen); data starts at row 2.

| Col | Heading | Content | Written as |
|---|---|---|---|
| A | Sr No | 1, 2, 3, … | Formula `=ROW()-1` |
| B | Date | Date of transaction | Excel date, `DD-MM-YYYY` |
| C | Description | Purpose | Text |
| D | Credit | Money received; blank on debit rows | Number, ₹ with Indian grouping |
| E | Debit | Money spent; blank on credit rows | Number, ₹ with Indian grouping |
| F | Balance | Running balance | Formula, ₹ |
| G | Bill Date | Optional | Excel date, `DD-MM-YYYY`, or blank |
| H | Bill No | Optional | Text (so `0042` keeps its zeros), or blank |

**Balance formulas**

| Row | Formula |
|---|---|
| 2 (first) | `=N(D2)-N(E2)` |
| n ≥ 3 | `=F{n-1}+N(D{n})-N(E{n})` (balance above + credit − debit) |

`N()` makes a blank cell (or stray text) count as zero.

**Number formats:** Credit and Debit always use Indian grouping. The Balance column uses a format that always shows a minus sign for a negative balance; its digit grouping follows the PC's region setting (Indian grouping on a PC set to English (India), for example ₹4,999,999.50 on a US-region PC). The app window always uses Indian grouping.

**Example**

| Sr No | Date | Description | Credit | Debit | Balance | Bill Date | Bill No |
|---|---|---|---|---|---|---|---|
| 1 | 01-10-2026 | Client advance | ₹50,00,000.00 | | ₹50,00,000.00 | 30-09-2026 | B-17 |
| 2 | 05-10-2026 | Second instalment | ₹2,50,000.00 | | ₹52,50,000.00 | | |
| 3 | 10-10-2026 | Cement | | ₹85,000.00 | ₹51,65,000.00 | | |
| 4 | 10-10-2026 | Labour | | ₹60,000.50 | ₹51,04,999.50 | | |

**Current Balance (shown in the app)** = sum of all Credit − sum of all Debit, read from the file. It does not depend on the Balance column.

---

## Ordering rules and manual edits in Excel

The app keeps every site file in order, whether a row was added by the app or typed in Excel.

**Order**

1. Rows with a readable date come first, **oldest date first**.
2. Rows with the same date keep the order they were entered in.
3. Rows that are **not in the expected format** go to the bottom, in their existing order, and are left exactly as typed.

A row is "not in the expected format" if any of these is true:

| Problem | Example |
|---|---|
| Date missing or not a real date | `someday`, `31-02-2026`, empty |
| Bill Date filled in but not a real date | `n/a` |
| Credit or Debit is not a plain number | `abc`, a formula like `=5+5` |

Readable date forms: a real Excel date, or text as `DD-MM-YYYY`, `DD/MM/YYYY`, `DD.MM.YYYY` or `YYYY-MM-DD`. Text dates are converted to real dates when the file is tidied.

**When the app tidies a file** (when you add a transaction, select the site, or return to the window): if rows are out of date order, a row is missing its Sr No or Balance formula, or there are blank rows in the middle, the app

1. takes a backup of the file as you left it,
2. sorts the rows as above,
3. rewrites every Sr No and Balance formula,
4. shows a message: "Changes made in Excel were sorted by date and renumbered."

If the file is open in Excel at that moment it cannot be rewritten. The window still shows the sorted rows and the correct balance, and the file is tidied the next time it is free.

Rows in the wrong format are counted in the message ("2 rows with a date or amount that could not be read are kept at the bottom") and shown in red. Their numeric amounts still count toward the balance.

| Safe in Excel | Not supported |
|---|---|
| Typing new rows anywhere (date, description, credit or debit; formulas are added for you) | Changing or deleting the column headings (the file stops being listed as a site) |
| Correcting any value | Formulas in Credit or Debit (the row goes to the bottom) |
| Deleting or inserting rows | Your own data in extra columns (it does not move when rows are re-sorted) |
| Renaming or deleting the file | Saving from the app while the file is open in Excel |

---

## What happens when you add a transaction

1. The form is checked; any problem is shown and nothing is written.
2. The app checks that the file is not locked by Excel.
3. It reads all existing rows and adds the new one.
4. Rows are sorted by the rules above.
5. Every Sr No and Balance formula is rewritten.
6. The file is written to a temporary file and then swapped in, so a failure never leaves a half-written file.
7. A backup copy is made and backups older than 3 days are deleted.
8. The table, balance and totals are refreshed from the file.

A back-dated entry therefore renumbers and rebalances every row below it.

---

## Validation rules

**Amount:** greater than zero, at most 2 decimal places, below ₹1,00,00,00,00,00,000. Commas, spaces and ₹ are ignored.

**Date and Bill Date:** `DD-MM-YYYY` (`/` and `.` also accepted when typing). Must be a real date.

**Description:** cannot be empty.

**Site name:** not empty; at most 100 characters; none of `\ / : * ? " < > |`; cannot end with a dot or start with `~$`; not a Windows reserved name (CON, PRN, AUX, NUL, COM1–9, LPT1–9); no existing file with the same name (case-insensitive).

---

## Backups

- A copy of the site file is saved in `.backup` inside the accounts folder after every saved transaction, and before the app rewrites a file you edited by hand.
- Name: `<Site>_YYYY-MM-DD_HH-MM-SS.xlsx`.
- Files in `.backup` older than 3 days (by modified time) are deleted after each save and when a folder is loaded.
- To restore, copy a backup out of `.backup` into the accounts folder under the site's name.
- If a backup cannot be made, a warning is shown; the transaction itself is already saved.

**Why `.backup`, and is it hidden?** On Windows a leading dot does **not** hide a folder (that is a Linux/macOS convention). The dot makes the name sort first in a folder list, so it stays at the top. To actually hide it, the app also gives the folder the Windows *Hidden* attribute when it creates it. To see it, open File Explorer, then View > Show > Hidden items. To keep the folder visible instead, set `HIDE_BACKUP_FOLDER = False` in `backup_manager.py`. (This attribute is not yet tested on Windows.)

---

## Settings

The last folder and site are stored in `%APPDATA%\ConstructionAccountsManager\config.json`. Delete it to reset. If it cannot be written, the app works normally and simply forgets.

---

## Errors and how the app responds

| Situation | Response |
|---|---|
| Folder missing or empty path | Message line shows the problem; site list disabled |
| File open in Excel when adding | Warning to close it and try again; nothing changes; what you typed is kept |
| File open in Excel when reloading | Table shown sorted anyway; message says the file will be tidied once closed |
| Invalid amount or date, empty description | Warning naming the problem; nothing is written |
| Duplicate or invalid site name | Warning; the name is asked for again |
| File missing or headings changed | Message; the file is not touched |
| Rows in the wrong format | Kept at the bottom, shown in red, counted in a message |
| Unexpected error while saving | Error shown; the existing file is left unchanged |
| Unexpected error elsewhere | A message box instead of a crash |

---

## Tech stack

| Part | Choice |
|---|---|
| Language | Python 3.10+ |
| Window | Tkinter / ttk (`clam` theme), `tkcalendar` for the calendar pop-up |
| Excel files | `openpyxl` (no Excel needed) |
| Packaging | PyInstaller, single `.exe` |
| Icons | Drawn in code (no image files); app icon `assets/icon.ico` made with Pillow (only needed to regenerate it) |

---

## Project structure

```
construction-accounts-manager/
├── main.py              # entry point: window size, icon, error box
├── gui.py               # the window and all its behaviour
├── widgets.py           # entry with hint text; date field with calendar pop-up
├── theme.py             # colours and fonts
├── icons.py             # calendar, folder, clear and plus icons
├── excel_handler.py     # site files: create, list, add, tidy, read ledger
├── money.py             # amount parsing and INR formatting
├── backup_manager.py    # .backup folder and 3-day cleanup
├── config.py            # remembered folder and site
├── make_icon.py         # regenerates assets/icon.ico (needs Pillow)
├── requirements.txt     # openpyxl, tkcalendar, pyinstaller
├── build.bat            # builds the .exe on Windows
├── assets/icon.ico
├── tests/
│   ├── test_logic.py        # 28 tests: files, ordering, formulas, backups, amounts
│   └── test_gui_smoke.py    # 13 tests: drives the real window
└── README.md
```

**Main functions in `excel_handler.py`**

| Function | Purpose |
|---|---|
| `list_sites(folder)` | Valid site names in a folder |
| `validate_site_name(name, folder)` | Error message or `None` |
| `create_site(folder, name)` | New empty site file |
| `add_transaction(path, ...)` | Add, sort, rewrite formulas, save; returns the new balance |
| `needs_tidy(path)` / `tidy_site(path, before_write)` | Detect and fix a hand-edited file |
| `read_ledger(path)` | Sorted rows with running balance, totals, count of unreadable rows |
| `get_balance(path)` | Balance recalculated from Credit and Debit |

---

## Development setup and tests

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
python -m unittest discover -s tests -v
```

The window tests need a display and are skipped automatically without one. On Linux without a screen: `xvfb-run -a python -m unittest discover -s tests -v`.

---

## Building the Windows executable

Double-click `build.bat` on a Windows PC with Python 3.10+ installed. It runs:

```bat
pyinstaller --noconfirm --onefile --windowed --name ConstructionAccountsManager ^
  --icon assets\icon.ico --add-data "assets\icon.ico;assets" ^
  --hidden-import babel.numbers --collect-data babel main.py
```

The result is `dist\ConstructionAccountsManager.exe`; copy it to the client's PC.

- `--hidden-import babel.numbers` and `--collect-data babel` are needed because `tkcalendar` depends on Babel.
- Unsigned PyInstaller executables are sometimes flagged by antivirus. Test on the client's PC and add an exception or sign the file if needed.
- Test the `.exe` on a PC without Python installed.

---

## Limitations

- One user at a time per folder; there is no locking between computers.
- Rows are rewritten when a file is tidied, so manual cell colours or comments do not follow rows when they are re-sorted.
- Credit and Debit must be plain numbers.
- No edit or delete of entries inside the app; use Excel (backups cover mistakes for 3 days).
- Dates are `DD-MM-YYYY`.
- The table shows the latest 100 rows (totals and balance always cover all rows).
- Windows only.

---

## Out of scope

Summary across all sites, statement export to PDF, search and filter, password protection, multi-user sync, undo.

---

## License

Private project developed for a client. All rights reserved.
