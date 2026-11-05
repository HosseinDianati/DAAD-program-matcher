"""
Production Excel and Google Sheet generator for DAAD Program Matcher.
13 Standard Columns (Option 2):
1. University
2. Program
3. Admission Chance
4. Tuition fee
5. Semester fee
6. Intake
7. Deadline
8. Requirements
9. Required Documents
10. Application Method
11. Program link
12. Note
13. Apply Link
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Column widths tailored for 13 columns
COLUMN_WIDTHS = {
    "A": 24.0,  # University
    "B": 26.0,  # Program
    "C": 18.0,  # Admission Chance
    "D": 14.0,  # Tuition fee
    "E": 15.0,  # Semester fee
    "F": 12.0,  # Intake
    "G": 18.0,  # Deadline
    "H": 48.0,  # Requirements
    "I": 42.0,  # Required Documents
    "J": 16.0,  # Application Method
    "K": 20.0,  # Program link
    "J": 16.0,  # Application Method
    "L": 20.0,  # Note
    "M": 20.0,  # Apply Link
}

# Fix duplicate key in dictionary
COLUMN_WIDTHS = {
    "A": 24.0,  # University
    "B": 26.0,  # Program
    "C": 18.0,  # Admission Chance
    "D": 14.0,  # Tuition fee
    "E": 15.0,  # Semester fee
    "F": 12.0,  # Intake
    "G": 18.0,  # Deadline
    "H": 48.0,  # Requirements
    "I": 42.0,  # Required Documents
    "J": 16.0,  # Application Method
    "K": 20.0,  # Program link
    "L": 20.0,  # Note
    "M": 20.0,  # Apply Link
}

HEADERS_13 = [
    "University",
    "Program",
    "Admission Chance",
    "Tuition fee",
    "Semester fee",
    "Intake",
    "Deadline",
    "Requirements",
    "Required Documents",
    "Application Method",
    "Program link",
    "Note",
    "Apply Link"
]

def get_chance_fill(chance_str: str):
    """Returns soft color fills for Admission Chance."""
    c = str(chance_str).lower()
    if "high" in c or "safe" in c:
        return PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"), Font(name="Arial", size=10, bold=True, color="155724")
    elif "medium" in c or "target" in c:
        return PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"), Font(name="Arial", size=10, bold=True, color="856404")
    elif "low" in c or "reach" in c:
        return PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"), Font(name="Arial", size=10, bold=True, color="721C24")
    return PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"), Font(name="Arial", size=10)


def create_study_search_workbook(rows: list[dict], title: str = "Germany") -> openpyxl.Workbook:
    """
    Creates an openpyxl Workbook populated with 13-column search results.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Study Programs"

    # Row 1: Title ("Germany")
    title_font = Font(name="Arial", size=12, bold=True)
    title_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    header_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

    ws.cell(row=1, column=1, value=title)
    ws.cell(row=1, column=1).font = title_font
    ws.cell(row=1, column=1).fill = title_fill

    for col in range(2, len(HEADERS_13) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = title_fill

    # Row 2: Headers
    header_font = Font(name="Arial", size=11, bold=True)
    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )

    for col_idx, header in enumerate(HEADERS_13, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Data Rows
    data_font = Font(name="Arial", size=10)
    link_font = Font(name="Arial", size=10, color="0000FF", underline="single")

    for row_idx, row_data in enumerate(rows, start=3):
        for col_idx, header in enumerate(HEADERS_13, start=1):
            val = row_data.get(header, "")
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border

            # Specific column styling
            if header == "Admission Chance":
                fill, font = get_chance_fill(val)
                cell.fill = fill
                cell.font = font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.value = val
            elif header in ["Requirements", "Required Documents"]:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.value = val
            elif header in ["Program link", "Apply Link"]:
                cell.alignment = Alignment(horizontal="left", vertical="top")
                if val and str(val).startswith("http"):
                    cell.hyperlink = str(val)
                    cell.font = link_font
                    cell.value = str(val)
                else:
                    cell.value = val
            elif header in ["Tuition fee", "Semester fee", "Intake", "Deadline", "Application Method"]:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.value = val
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.value = val

    # Set column widths
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # Row heights
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 28
    for r in range(3, len(rows) + 3):
        ws.row_dimensions[r].height = 90  # Room for multi-line requirements and document lists

    return wb


def save_search_to_excel(rows: list[dict], output_path: str, title: str = "Germany"):
    wb = create_study_search_workbook(rows, title)
    wb.save(output_path)
    print(f"Saved {len(rows)} programs to Excel: {output_path}")


def export_to_google_sheet_format(rows: list[dict], sheet_title: str = "DAAD Study Program Matches") -> dict:
    """
    Formats the rows for Google Sheets API / CSV / MCP sync.
    Returns standard 2D matrix structure [ [headers], [row1], [row2], ... ]
    """
    matrix = [HEADERS_13]
    for row in rows:
        matrix.append([str(row.get(h, "")) for h in HEADERS_13])
    return {
        "title": sheet_title,
        "values": matrix,
        "rowCount": len(matrix),
        "columnCount": len(HEADERS_13)
    }
