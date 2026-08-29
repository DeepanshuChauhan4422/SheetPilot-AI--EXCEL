from langchain_core.tools import tool


# ============================================================
# UPDATE CELL
# ============================================================

@tool
def update_cell(
    cell: str,
    value: str
) -> dict:
    """Update a specific Excel cell."""

    return {
        "action": "update_cell",
        "cell": cell,
        "value": value
    }


# ============================================================
# WRITE RANGE
# ============================================================

@tool
def write_range(
    range_address: str,
    values: list
) -> dict:
    """Write a 2D list of values into an Excel range."""

    return {
        "action": "write_range",
        "range_address": range_address,
        "values": values
    }


# ============================================================
# READ RANGE
# ============================================================

@tool
def read_range(
    range_address: str
) -> dict:
    """Read values from an Excel range."""

    return {
        "action": "read_range",
        "range_address": range_address
    }


# ============================================================
# CLEAR RANGE
# ============================================================

@tool
def clear_range(
    range_address: str
) -> dict:
    """Clear the contents of an Excel range."""

    return {
        "action": "clear_range",
        "range_address": range_address
    }


# ============================================================
# FORMAT RANGE
# ============================================================

@tool
def format_range(
    range_address: str,
    format_type: str,
    value: str = ""
) -> dict:
    """
    Format an Excel range.

    Supported format types:
    - bold
    - italic
    - fill_color
    - font_color
    - font_size
    - number_format
    - horizontal_alignment
    - borders
    - autofit_columns
    - autofit_rows
    """

    return {
        "action": "format_range",
        "range_address": range_address,
        "format_type": format_type,
        "value": value
    }


# ============================================================
# ADD FORMULA
# ============================================================

@tool
def add_formula(
    cell: str,
    formula: str
) -> dict:
    """Add an Excel formula to a cell."""

    return {
        "action": "add_formula",
        "cell": cell,
        "formula": formula
    }


# ============================================================
# SORT RANGE
# ============================================================

@tool
def sort_range(
    range_address: str,
    column: str,
    order: str = "descending"
) -> dict:
    """
    Sort an Excel range by a column.

    column can be:
    - Excel column letter such as A, B, C
    - Excel header name such as name, revenue, expenses, profit

    order:
    - ascending
    - descending
    """

    normalized_order = str(order).lower().strip()

    if normalized_order not in [
        "ascending",
        "descending"
    ]:
        normalized_order = "descending"

    return {
        "action": "sort_range",
        "range_address": range_address,
        "column": str(column).strip(),
        "order": normalized_order
    }


# ============================================================
# CREATE CHART
# ============================================================

@tool
def create_chart(
    range_address: str,
    chart_type: str = "ColumnClustered",
    title: str = "Excel Chart"
) -> dict:
    """
    Create an Excel chart from an Excel range.
    """

    return {
        "action": "create_chart",
        "range_address": range_address,
        "chart_type": chart_type,
        "title": title
    }