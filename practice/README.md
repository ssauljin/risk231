# Practice files

Excel workbooks that go with the RISK 231 lecture slides: in-class walkthrough exercises for the VBA lectures, plus the course's assignments and projects.

| File | Type | Goes with | What it is |
|---|---|---|---|
| [`9 _ VBA_Intro_and_Recording_Macros.xlsm`](#9-vba-intro-and-recording-macros) | Walkthrough | Ch. 3 (VBA), §3.1 | Recording macros, buttons, and a loan calculator |
| [`10 _ VBA_Variables_Objects.xlsm`](#10-vba-variables-and-objects) | Walkthrough | Ch. 3 (VBA), §3.2 | `InputBox`, `MsgBox`, variables and objects |
| [`11 _ VBA_If_Statements_Loops_Functions.xlsm`](#11-vba-if-statements-loops-and-functions) | Walkthrough | Ch. 3 (VBA), §3.3 | Sorting students into grade tabs with `If` and loops |
| [`Assignment_2.xlsx`](#assignment-2--age-to-age-factors) | Assignment | Ch. 2 (Excel) | Claims development: age-to-age factors |
| [`Assignment_3.xlsm`](#assignment-3--loan-calculator-macros) | Assignment | Ch. 2 §2.4, Ch. 3 §3.1 | Loan calculator solved with recorded Goal Seek macros |
| [`Project 1.xlsx`](#project-1--sales-data-analysis) | Project | Ch. 2 (Excel) | Term-life sales data: text and date functions, lookups, filtering, sorting, pivot tables |
| [`Project 2.xlsm`](#project-2--premium-tables-with-vba) | Project | Ch. 2–3 | Universal-life illustration model: funding patterns and a looping macro that builds premium tables |

The slides are the `Ch0N_*.pdf` files in the repository root.

## Before you start

- **Use desktop Excel (Windows or Mac).** The `.xlsm` files contain VBA macros, which don't run in Excel for the web, Google Sheets, or LibreOffice.
- **Enable macros.** Excel blocks macros in downloaded files. If you see a security warning, close the file, right-click it → *Properties* → tick **Unblock** (Windows), reopen it, and click **Enable Content**.
- **Show the Developer tab** to open the VBA editor and record macros: *File → Options → Customize Ribbon → Developer* (on Mac: *Excel → Preferences → Ribbon & Toolbar*). See the "Intro to Excel ~ Tabs" slide in Ch. 2.
- **Keep the original.** Work on a copy so that you can start over.
- **Look for the colour-coded cells.** The workbooks use fill colours to mark the cells you should fill in, or where to put a button; each instruction sheet says which colour means what.

## In-class walkthroughs (VBA)

These follow the lecture slides step by step. Each sheet starts with a **Walkthrough Practice** box describing the task, and the workbooks already contain the instructor's example macros (open the VBA editor with **Alt+F11**, or **Opt+F11** on a Mac) so you can compare your code against them.

### 9: VBA intro and recording macros

`9 _ VBA_Intro_and_Recording_Macros.xlsm`

| Sheet | Task |
|---|---|
| `Ex 1` | Record a macro that copies a value from a *source* to a *destination* cell, then write variations of it by hand in the VBE (`Ex1a`, `Ex1b`, `Ex1c`, …). |
| `Ex 2` | Record `Ex_2` to format a premium summary table: underline the headers, apply `$` number format, and bold the totals row. Then read the generated code. |
| `Ex 3` | Record `Ex_3` to wrap an inforce table's `VLOOKUP` in `IFERROR`, fill the formula down, and attach the macro to a button. |
| `Ex 4` | Record `Ex_4` (type "Whoot!" into a cell, then clear it) and step through it line by line with **F8**. |
| `Ex 5` | Loan amortization schedule: record `Ex_5`, which uses **Goal Seek** to find the payment that pays off the loan, and run it from a button. |

### 10: VBA variables and objects

`10 _ VBA_Variables_Objects.xlsm`. Sheet `Examples`:

1. Write a macro that reads two numbers with `Application.InputBox`, adds them, and shows the result in a `MsgBox` as `x + y = z`. Add a button to run it.
2. Write a macro that asks for the user's name, gender (F/M) and age with input boxes, and writes the answers into the small table on the sheet. Add a button for it as well.

The module also contains the `VtypeDemo1`–`VtypeDemo3` examples from the slides on variable types.

### 11: VBA if statements, loops, and functions

`11 _ VBA_If_Statements_Loops_Functions.xlsm`

| Sheet | Task |
|---|---|
| `If Statements` | Given student IDs and scores out of 100, write a macro that copies each student to the tab for their letter grade (A = 90–100, B = 80–89, C = 70–79, D = 60–69, F = 0–59). Run it from a button. |
| `A`, `B`, `C`, `D`, `F` | Output tabs for the `If` version. |
| `Loops` | The same exercise, using a loop over all the students instead of one `If` block per student. |
| `A_` … `F_` | Output tabs for the loop version. |

## Assignments

### Assignment 2: age-to-age factors

`Assignment_2.xlsx`. Sheets: `Assignment 2` (instructions and work area) and `Template` (paid-claims triangle, accident years 1999–2008 by development age 12–120 months).

You build a claims development triangle of **age-to-age factors** (the ratio of successive cumulative claims) and summarise them:

1. Build the column labels `12-24`, `24-36`, …, `108-120` with a formula (hint: `CONCAT`).
2. Make the factor formulas ignore both `0` and `#DIV/0!` (hint: `IF` and `IFERROR`).
3. Fill in the summary rows (simple averages over all years or the latest three years, and volume-weighted means) with formulas that copy across without manual edits (hint: `OFFSET`).

A hidden sheet, `Values`, holds the expected results (values only), so you can check your answers. To see it, right-click any sheet tab → **Unhide**.

Skills: Ch. 2 §2.2–2.5 (`IF`, `IFERROR`, `OFFSET`). `CONCAT` isn't covered in the Excel slides: it joins text, so `=CONCAT(12,"-",24)` gives `12-24`.

### Assignment 3: loan calculator macros

`Assignment_3.xlsm`. Sheets: `Instructions`, `LOAN CALCULATOR`, and `Answer Sheet`.

Redo the loan problems from Assignment 1 by recording macros that run **Goal Seek** on the loan calculator:

- Q1–Q3: solve for the monthly payment with one macro and button.
- Q4: solve for the loan principal, and Q5: solve for the monthly interest rate, each with its own macro and button. You can record them, or write them by copying and editing your first macro's code.

Enter the answers in the `Answer Sheet`.

Skills: Ch. 2 §2.3–2.4 (Goal Seek, loan calculator) and Ch. 3 §3.1 (recording macros, buttons).

## Projects

### Project 1: sales data analysis

`Project 1.xlsx`. Two parts; tabs labelled in **green** belong to Part 1 and tabs labelled in **blue** to Part 2. The first tab, `Project 1`, is the cover sheet.

**Part 1** (`P1 Instructions`, `P1 Sales Extract`, `P1 Agent Table`, `P1 Inforce`): about 4,000 term-life policies sold through June 2017.

1. Convert the issue date to a real Excel date, and assign Series A or B by date.
2. Extract each agent's last name from the full name (`FIND`, `MID`/`RIGHT`).
3. Look up the agent's last name for every inforce policy (`VLOOKUP`), then carry it back to the sales extract (non-inforce policies are flagged).
4. Answer summary questions with filtering and `SUBTOTAL`, list a particular agent's policies that have riders, and re-sort the data by agent and policy number.

**Part 2** (`P2 Instructions`, `P2 Sales Extract`, `P2 Q1`–`P2 Q3`): build pivot tables from the sales extract to summarise total premium by year, premium and face amount by issue-age band, and average premium and face by gender and risk class. Then copy the results into the answer grids with *Paste Special → Values*.

Skills: Ch. 2 §2.2–2.6 (text and date functions, `VLOOKUP`, filtering, sorting, pivot tables).

### Project 2: premium tables with VBA

`Project 2.xlsm`. This extends the universal-life illustration model from the Excel group project.

| Sheet | Purpose |
|---|---|
| `Instructions` | The task and the marking scheme (10 points) |
| `Inputs_Outputs` | Model inputs (gender, risk class, issue age, funding pattern), the premium, and the maximum account value outputs; holds the **Get Premiums** button |
| `Calculations` | Monthly account-value projection to age 99 |
| `COI tables`, `Per K tables` | Cost-of-insurance rates and per-$1,000 charges by gender, class, age and policy year |
| `MN Monthly` … `FS Single-pay` | 12 empty premium tables: 2 genders (M/F) × 2 classes (N = non-smoker, S = smoker) × 3 funding patterns |

Tasks:

1. In `Calculations`, change the premium-paid formula to support three funding patterns through the named range `Pattern`: **Monthly** (level monthly premium), **Ten-pay** (annual premium for the first 10 years), and **Single-pay** (one premium at issue).
2. In `Inputs_Outputs`, add formulas (cells C10:C12) for the maximum account value and the year and month in which it occurs.
3. Write one macro that asks for gender, class, funding pattern, and a low and high issue age with `InputBox`. For each age, it should run Goal Seek so the account value at age 99 is $1, then paste the premium and max-AV results into the matching table. Attach it to **Get Premiums** and fill in all 12 tables (issue ages 18–70).

Skills: Ch. 2 (named ranges, `IF`, Goal Seek) and Ch. 3 §3.2–3.3 (`InputBox`, variables, `If`, loops, selecting sheets).
