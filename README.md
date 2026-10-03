# Assessment Marks Splitter: Excel Plugin

This Excel plugin is built with Python and **xlwings**. You give it a student's **Term Work total**, and it dynamically fills in the **component marks** (Assignment, Test, Practical, …) for you using a simple Excel formula.

Every mark it generates:

- is between **0 and that component's maximum**, and
- adds up **exactly** to the total you entered.

---

## Contents

1. [Requirements](#1-requirements)
2. [One-time installation](#2-one-time-installation)
3. [User manual](#3-user-manual-the-formula-generate_marks)
4. [How the marks are generated](#4-how-the-marks-are-generated)
5. [Keeping the marks fixed](#5-keeping-the-marks-fixed-important)
6. [Troubleshooting](#6-troubleshooting)
7. [FAQ](#7-faq)

---

## 1. Requirements

- **Windows** with **Microsoft Excel** (desktop app; Excel Online is not supported)
- **Python 3.8 or newer**: <https://www.python.org/downloads/>
  (during installation, tick **"Add Python to PATH"**)
- The Python package **xlwings**

---

## 2. One-time installation

You only do this **once per computer**.

### Step 1: Install xlwings

Open **Command Prompt** and run:

```
pip install xlwings
xlwings addin install
```

Restart Excel. A new **xlwings** tab should now appear in the ribbon.

### Step 2: Allow xlwings to talk to Excel

1. In Excel, go to **File → Options → Trust Center → Trust Center Settings…**
2. Click **Macro Settings**.
3. Tick **"Trust access to the VBA project object model"**.
4. Click **OK**, then **OK** again.

> If you skip this, you'll get `Run-time error '1004': Method 'VBProject' of object '_Workbook' failed`.

### Step 3: Point xlwings at the plugin

On the **xlwings** ribbon tab, fill in these boxes:

| Box | What to type |
|---|---|
| **Interpreter** | `python` |
| **PYTHONPATH** | The folder that contains `assessment_tool.py`, e.g. `C:\Users\Ani\OneDrive\Desktop\codes\Chandan sirs -excel plugin` |
| **UDF Modules** | `assessment_tool` (no `.py`) |

These settings apply to **every workbook** on this computer, so the plugin file can stay in its own folder. You don't have to copy it next to each workbook.

The installation is finished.

---

## 3. User manual: the formula `GENERATE_MARKS`

### Syntax

```
=GENERATE_MARKS(total, max_marks_range)
```

| Argument | Meaning |
|---|---|
| `total` | The cell with the student's Term Work total, e.g. `B3` |
| `max_marks_range` | The cells with each component's maximum marks, e.g. `$C$1:$E$1` |

The result **spills** into as many cells as there are max-marks cells. A row of maximums spills across; a column of maximums spills down.

### Using it in a workbook (each new workbook)

1. Save the workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
   *(Import Functions stores a small piece of code inside the workbook. A plain `.xlsx` can't keep that code after you save.)*
2. On the **xlwings** tab, click **Import Functions**.
3. Type the formula in a cell (see the example below).

### Worked example

Suppose your sheet looks like this:

|   | A | B | C | D | E |
|---|---|---|---|---|---|
| **1** | | *(max →)* | **10** | **10** | **5** |
| **2** | Student | Term Work | Comp 1 | Comp 2 | Comp 3 |
| **3** | 1 | 20 | | | |
| **4** | 2 | 20 | | | |
| **5** | 3 | 24 | | | |
| **6** | 4 | 12 | | | |

1. Click cell **C3**.
2. Type:
   ```
   =GENERATE_MARKS(B3, $C$1:$E$1)
   ```
3. Press **Enter**. Cells **C3, D3 and E3** fill in, for example `8  7  5`, which adds up to 20.
4. Drag the small square at the bottom-right corner of **C3** down to row 6. Every student now has marks.

> **Why the `$` signs?** They lock the max-marks row, so it stays on row 1 when you drag the formula down.
> Leave the cells to the right of the formula (D and E here) **empty**. If something is in them, Excel shows `#SPILL!`.

### Formula behaviour

| Situation | What you get |
|---|---|
| Total cell is empty | Blank cells |
| A max-marks cell is empty | A blank in that position (the other components still add up to the total) |
| Total has a half mark, e.g. `17.5` | Marks in steps of 0.5 |
| Total is larger than the sum of maximums | `Error: Total 30 exceeds maximum 25.` |
| Total or max is text / TRUE / FALSE | `Error: Total must be a number, got 'abc'.` |
| Total is negative | `Error: Total cannot be negative.` |
| All max-marks cells are empty | `Error: No maximum marks given.` |

---

## 4. How the marks are generated

- The components are filled in a **random order**. Each mark is chosen at random within a range that guarantees the remaining components can still reach the total, so the marks always add up exactly.
- **Step size** is picked automatically from the total:
  `20` gives whole marks, `17.5` gives half marks, `17.25` gives quarter marks, and so on down to 0.01.
- A mark can never exceed its maximum, even if the maximum is a decimal like `2.5`.

---

## 5. Keeping the marks fixed (important)

The **formula** gives new random marks whenever its input cells change or Excel fully recalculates (for example Ctrl + Alt + F9).

When you're happy with the marks and want to keep them:

1. Select the cells with the marks.
2. Press **Ctrl + C**.
3. Right-click the same cells and choose **Paste Special → Values** (the `123` icon).

The cells now hold plain numbers and won't change again.

---

## 6. Troubleshooting

| Problem | Fix |
|---|---|
| No **xlwings** tab in Excel | Run `xlwings addin install` in Command Prompt, then restart Excel. |
| `Run-time error '1004': Method 'VBProject' ... failed` | Do [Step 2](#step-2-allow-xlwings-to-talk-to-excel): tick *Trust access to the VBA project object model*. |
| `#NAME?` in the cell | The function isn't imported into this workbook yet. Click **Import Functions** on the xlwings tab. |
| `Object required` in the cell | 1) Save the workbook as `.xlsm`. 2) Press **Alt + F11 → Tools → References** and tick **xlwings**. 3) Put `python` in the **Interpreter** box. 4) Click **Restart UDF Server**, then **Import Functions**. 5) Click the formula cell and press **Enter**. |
| `ModuleNotFoundError: No module named 'assessment_tool'` | The **PYTHONPATH** box doesn't point at the folder that contains `assessment_tool.py`, or **UDF Modules** is misspelled. |
| `ModuleNotFoundError: No module named 'Book1'` (or your workbook's name) | The **UDF Modules** box is empty, so xlwings looks for a file named after the workbook. Type `assessment_tool` in it. |
| `#SPILL!` | Clear the cells to the right of (or below) the formula so the result has room. |
| Changed `assessment_tool.py` but Excel still uses the old version | Click **Restart UDF Server**, then **Import Functions**. |
| Marks keep changing | That's expected for formulas. See [Keeping the marks fixed](#5-keeping-the-marks-fixed-important). |

---

## 7. FAQ

**Do I need to copy the `.py` file next to every workbook?**
No. As long as **PYTHONPATH** points at this folder, every workbook can use it.

**Do I have to click "Import Functions" for every workbook?**
Yes, once per workbook. After that, saving as `.xlsm` keeps the function in that workbook.

**Can I have a different number of components per sheet?**
Yes. The formula uses however many max-marks cells you select.

**Can the sheet use columns instead of rows for the max marks?**
Yes. If the max-marks range is a column, the result spills downwards.

**Is anything sent over the internet?**
No. Everything runs locally on your computer.
