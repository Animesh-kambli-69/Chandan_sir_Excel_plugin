"""
Assessment Marks Splitter: Excel Plugin

Excel plugin (xlwings) that splits a Term Work total into component marks.
Every component stays between 0 and its maximum, and the components always
add up exactly to the total.

Usage:
Formula (works in any workbook once added to xlwings "UDF Modules"):
    =GENERATE_MARKS(total_cell, max_marks_range)
    e.g. =GENERATE_MARKS(F2, $B$1:$E$1)

Requirement: Windows + Microsoft Excel + `pip install xlwings`
"""

import random

import xlwings as xw

STEPS = (1, 0.5, 0.25, 0.1, 0.01)  # supported mark granularities
_rng = random.SystemRandom()

# --------------------------------------------------------------------------- #
# Core logic
# --------------------------------------------------------------------------- #
def _to_number(value, label):
    """Convert an Excel cell value to float with a readable error."""
    if isinstance(value, bool):
        raise ValueError(f"{label} must be a number, not TRUE/FALSE.")
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{label} must be a number, got {value!r}.") from None


def _pick_step(value):
    """Smallest-precision step that represents `value` exactly (1, 0.5, ...)."""
    for step in STEPS:
        if abs(value / step - round(value / step)) < 1e-9:
            return step
    return STEPS[-1]


def split_total(target_total, maxima):
    """
    Return a random list of marks (same order as `maxima`) that sums exactly
    to `target_total`, with 0 <= mark <= max for every component.
    """
    if not maxima:
        raise ValueError("No maximum marks given.")
    if target_total < 0:
        raise ValueError("Total cannot be negative.")
    if any(m < 0 for m in maxima):
        raise ValueError("Maximum marks cannot be negative.")

    step = _pick_step(target_total)
    target_units = round(target_total / step)
    max_units = [int(m / step + 1e-9) for m in maxima]  # floor, never exceed max

    if target_units > sum(max_units):
        raise ValueError(f"Total {target_total:g} exceeds maximum {sum(maxima):g}.")

    # Fill components in random order. Each pick is bounded so the remaining
    # components can always absorb what is left -> no backtracking needed.
    order = list(range(len(maxima)))
    _rng.shuffle(order)

    result = [0] * len(maxima)
    remaining = target_units
    capacity_left = sum(max_units)

    for idx in order:
        capacity_left -= max_units[idx]
        low = max(0, remaining - capacity_left)
        high = min(max_units[idx], remaining)
        result[idx] = _rng.randint(low, high)
        remaining -= result[idx]

    if step == 1:
        return result
    return [round(u * step, 2) for u in result]


# --------------------------------------------------------------------------- #
# Excel formula:  =GENERATE_MARKS(total, max_marks_range)
# --------------------------------------------------------------------------- #
@xw.func
@xw.arg("max_marks", ndim=2)
def GENERATE_MARKS(target_total, max_marks):
    """
    Split target_total into component marks.
    Usage: =GENERATE_MARKS(total_cell, max_marks_range)
    The result has the same shape as max_marks_range (row or column).
    Blank cells in max_marks_range give blank results.
    """
    rows = len(max_marks)
    cols = len(max_marks[0]) if rows else 0
    flat = [v for row in max_marks for v in row]
    output = [""] * len(flat)

    if target_total not in (None, ""):
        try:
            total = _to_number(target_total, "Total")
            used = [i for i, v in enumerate(flat) if v not in (None, "")]
            maxima = [_to_number(flat[i], "Max marks") for i in used]
            for i, mark in zip(used, split_total(total, maxima)):
                output[i] = mark
        except ValueError as e:
            return f"Error: {e}"

    return [output[r * cols:(r + 1) * cols] for r in range(rows)]


@xw.func
@xw.arg("max_marks", ndim=2)
def GENERATE_SCALED_MARKS(scaled_mark, max_scaled, max_marks):
    """
    Distribute marks proportionally based on a scaled score (e.g., 4 out of 5).
    Usage: =GENERATE_SCALED_MARKS(scaled_mark_cell, max_scaled_cell, max_marks_range)
    Example: =GENERATE_SCALED_MARKS(4, 5, $B$1:$E$1)
    """
    rows = len(max_marks)
    cols = len(max_marks[0]) if rows else 0
    flat = [v for row in max_marks for v in row]
    output = [""] * len(flat)

    if scaled_mark not in (None, "") and max_scaled not in (None, ""):
        try:
            s_mark = _to_number(scaled_mark, "Scaled mark")
            m_scaled = _to_number(max_scaled, "Max scaled")
            
            if m_scaled <= 0:
                return "Error: Max scaled must be > 0"
                
            used = [i for i, v in enumerate(flat) if v not in (None, "")]
            maxima = [_to_number(flat[i], "Max marks") for i in used]
            
            # Calculate the proportional target total and round to the nearest whole number 
            # (or you could round to nearest 0.5 by doing round(target * 2) / 2)
            total_max = sum(maxima)
            target_total = round((s_mark / m_scaled) * total_max)
            
            for i, mark in zip(used, split_total(target_total, maxima)):
                output[i] = mark
        except ValueError as e:
            return f"Error: {e}"

    return [output[r * cols:(r + 1) * cols] for r in range(rows)]
