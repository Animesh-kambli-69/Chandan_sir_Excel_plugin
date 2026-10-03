"""
assessment_tool.py

Configurable Excel assessment helper.

REAL-ASSESSMENT MODE:
- Read component names and maximum marks from the "Setup" sheet.
- Rebuild/update the "Marks" sheet headers.
- Calculate Term Work from the actual component marks entered by the teacher.

SYNTHETIC/DEMO MODE:
- Generate random component combinations for a TARGET TOTAL for testing/demo data.
- Keep this separate from real assessment records. Do not use synthetic values
  as actual student marks.

Excel requirement:
- Microsoft Excel on Windows
- Python package: xlwings

Install:
    pip install xlwings

Run:
    python assessment_tool.py assessment_tool_template.xlsx
"""

import random
import sys
from pathlib import Path

import xlwings as xw


SETUP_SHEET = "Setup"
MARKS_SHEET = "Marks"
SYNTHETIC_SHEET = "Synthetic Test"

FIRST_CONFIG_ROW = 2
LAST_CONFIG_ROW = 20


def read_config(book):
    """Read enabled component names and maximum marks from Setup."""
    sheet = book.sheets[SETUP_SHEET]

    components = []

    for row in range(FIRST_CONFIG_ROW, LAST_CONFIG_ROW + 1):
        name = sheet.range(f"A{row}").value
        max_marks = sheet.range(f"B{row}").value
        enabled = sheet.range(f"C{row}").value

        # Stop at the first completely empty component row.
        if name in (None, "") and max_marks in (None, "") and enabled in (None, ""):
            continue

        if name in (None, ""):
            raise ValueError(f"Setup row {row}: component name is missing.")

        if enabled is None:
            enabled = "Y"

        enabled = str(enabled).strip().upper()

        if enabled not in {"Y", "N"}:
            raise ValueError(
                f"Setup row {row}: Enabled must be Y or N, not {enabled!r}."
            )

        if enabled == "N":
            continue

        if max_marks is None or max_marks == "":
            raise ValueError(f"Setup row {row}: maximum marks are missing.")

        try:
            max_marks = float(max_marks)
        except (TypeError, ValueError):
            raise ValueError(
                f"Setup row {row}: maximum marks must be numeric."
            )

        if max_marks < 0:
            raise ValueError(
                f"Setup row {row}: maximum marks cannot be negative."
            )

        components.append(
            {
                "name": str(name).strip(),
                "max": max_marks,
            }
        )

    if not components:
        raise ValueError("No enabled components found in Setup.")

    total_name = sheet.range("G2").value or "Term Work"

    return components, str(total_name).strip()


def excel_column_number(n):
    """1-based column number -> Excel letters."""
    letters = ""
    while n:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def build_marks_sheet(book, components, total_name):
    """
    Build the Marks headers and fill component marks based on the entered Term Work.
    """
    sheet = book.sheets[MARKS_SHEET]

    headers = ["Student Name"] + [c["name"] for c in components] + [total_name]
    sheet.range(
        f"A1:{excel_column_number(len(headers))}1"
    ).value = [headers]

    total_col_num = len(headers)
    total_col = excel_column_number(total_col_num)

    # Read Term Work and generate components
    for row in range(2, 502):
        target_val = sheet.range(f"{total_col}{row}").value
        if target_val is not None and isinstance(target_val, (int, float)):
            try:
                target_total = float(target_val)
                values = random_synthetic_breakdown(target_total, components)
                for j, value in enumerate(values, start=2):
                    col = excel_column_number(j)
                    sheet.range(f"{col}{row}").value = value
            except Exception as e:
                print(f"Row {row}: Could not break down total {target_val}. Error: {e}")

    # Basic formatting.
    sheet.range(
        f"A1:{total_col}1"
    ).api.Font.Bold = True
    sheet.range("A:A").column_width = 20
    sheet.range(
        f"B:{total_col}"
    ).column_width = 15

    # Update the maximum total on Setup.
    setup = book.sheets[SETUP_SHEET]
    total_max = sum(c["max"] for c in components)
    setup.range("G3").value = len(components)
    setup.range("G4").value = total_max

    return total_max


def validate_actual_marks(values, components):
    """Validate actual component marks before writing/calculating."""
    if len(values) != len(components):
        raise ValueError("The number of entered marks does not match the setup.")

    for value, component in zip(values, components):
        if value is None or value == "":
            continue

        try:
            value = float(value)
        except (TypeError, ValueError):
            raise ValueError(
                f"{component['name']}: mark must be numeric."
            )

        if value < 0 or value > component["max"]:
            raise ValueError(
                f"{component['name']}: {value} is outside 0-{component['max']}."
            )


def calculate_termwork(values, components):
    """Calculate total from actual component marks."""
    validate_actual_marks(values, components)
    return sum(float(v) for v in values if v not in (None, ""))


def random_synthetic_breakdown(target_total, components, step=1):
    """
    Generate one random valid combination that adds to target_total.

    Every component stays between 0 and its configured maximum.
    This is for synthetic/demo data only.
    """
    if target_total < 0:
        raise ValueError("Target total cannot be negative.")

    max_total = sum(c["max"] for c in components)

    if target_total > max_total:
        raise ValueError(
            f"Target total {target_total} exceeds maximum total {max_total}."
        )

    # Work in integer units so step=1 means integer marks.
    if step <= 0:
        raise ValueError("step must be positive.")

    scaled_target = round(target_total / step)
    scaled_maxima = [round(c["max"] / step) for c in components]

    if abs(scaled_target * step - target_total) > 1e-9:
        raise ValueError(
            f"Target total {target_total} is not compatible with step {step}."
        )

    # Randomized backtracking. The candidate order changes each attempt,
    # so different valid combinations are produced.
    order = list(range(len(components)))
    random.SystemRandom().shuffle(order)

    result = [0] * len(components)

    def search(position, remaining):
        if position == len(order):
            return remaining == 0

        idx = order[position]
        units_left = len(order) - position - 1

        # Bounds contributed by components still after this one.
        remaining_min = 0
        remaining_max = sum(scaled_maxima[order[j]] for j in range(position + 1, len(order)))

        low = max(0, remaining - remaining_max)
        high = min(scaled_maxima[idx], remaining - remaining_min)

        if low > high:
            return False

        choices = list(range(low, high + 1))
        random.SystemRandom().shuffle(choices)

        for choice in choices:
            result[idx] = choice
            if search(position + 1, remaining - choice):
                return True

        result[idx] = 0
        return False

    if not search(0, scaled_target):
        raise ValueError("No valid component combination exists for this total.")

    return [u * step for u in result]


def create_synthetic_demo_sheet(book, components, total_name):
    """
    Create a clearly-labelled synthetic test sheet.
    Column A contains target totals entered for demo/testing.
    Python fills component values that add to that target.
    """
    if SYNTHETIC_SHEET in [s.name for s in book.sheets]:
        sheet = book.sheets[SYNTHETIC_SHEET]
        sheet.clear()
    else:
        sheet = book.sheets.add(SYNTHETIC_SHEET)

    headers = ["Target Total"] + [c["name"] for c in components] + [total_name]
    sheet.range(
        f"A1:{excel_column_number(len(headers))}1"
    ).value = [headers]

    sheet.range("A1").api.Font.Bold = True
    sheet.range("A1").api.Interior.Color = 0xD9EAF7

    # Example target totals for testing; replace them with your own demo values.
    example_totals = [21, 23, 19, 25]

    max_total = sum(c["max"] for c in components)
    valid_totals = [x for x in example_totals if x <= max_total]

    for i, target in enumerate(valid_totals, start=2):
        sheet.range(f"A{i}").value = target

        values = random_synthetic_breakdown(target, components)

        for j, value in enumerate(values, start=2):
            col = excel_column_number(j)
            sheet.range(f"{col}{i}").value = value

        total_col = excel_column_number(len(headers))
        sheet.range(f"{total_col}{i}").value = sum(values)

    sheet.range("A:A").column_width = 16
    sheet.range(
        f"B:{excel_column_number(len(headers))}"
    ).column_width = 16


def main():
    if len(sys.argv) != 2:
        print("Usage: python assessment_tool.py <excel_file>")
        print("Example: python assessment_tool.py assessment_tool_template.xlsx")
        sys.exit(1)

    file_path = Path(sys.argv[1]).resolve()

    if not file_path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    # Connect to the workbook if it's already open, otherwise open it.
    book = xw.Book(str(file_path))

    components, total_name = read_config(book)

    total_max = build_marks_sheet(book, components, total_name)

    # Synthetic sheet is optional and clearly separated from real marks.
    create_synthetic_demo_sheet(book, components, total_name)

    book.save()

    print("\nAssessment workbook updated.")
    print("Components:")
    for c in components:
        print(f"  {c['name']}: {c['max']}")
    print(f"Maximum total: {total_max}")
    print(f"Calculated column: {total_name}")
    print("\nEnter ACTUAL marks in the Marks sheet.")
    print("Synthetic Test is for demo/testing data only.")



if __name__ == "__main__":
    main()
