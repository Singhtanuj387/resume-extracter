import pandas as pd
import io
import re


def _clean_cell(value):
    """
    Clean a cell value for Excel display:
    - Replace literal \\n with actual newlines
    - Collapse multiple blank lines
    - Strip leading/trailing whitespace
    - Return empty string for NaN/None
    """
    if pd.isna(value) or value is None:
        return ""
    s = str(value).strip()
    # Normalize newlines
    s = s.replace("\\n", "\n")
    # Collapse 3+ newlines into 2
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s


def _estimate_row_height(text, col_width=40):
    """Estimate the row height needed based on text length and wrapping."""
    if not text:
        return 20
    lines = text.split('\n')
    total_lines = 0
    for line in lines:
        # Estimate wrapped lines within each line
        wrapped = max(1, len(line) // col_width + 1)
        total_lines += wrapped
    return max(20, min(total_lines * 15, 400))


# Date pattern for detecting title lines (e.g., "October 2025", "Jan 2026")
_DATE_RE = re.compile(
    r'(January|February|March|April|May|June|July|August|September|October|November|December|'
    r'Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}',
    re.IGNORECASE
)

# Columns that contain mixed heading + description content
_RICH_COLUMNS = {
    'Area of Interest / Objective',
    'Education',
    'Experience & Internships',
    'Projects',
    'Awards & Achievements',
    'Skills',
    'Extra Curriculars & Leadership',
}


def _is_heading_line(line):
    """
    Determine if a line is a heading/title vs a bullet-point description.
    Headings typically: contain '|', have dates, or don't start with bullet chars.
    """
    stripped = line.strip()
    if not stripped:
        return False
    # Bullet lines are definitely not headings
    if stripped[0] in ('•', '–', '-', '▪', '▸', '*', '►', '○', '●'):
        return False
    # Lines with pipe separator are titles (e.g., "Backend Dev | Company Oct 2025")
    if '|' in stripped:
        return True
    # Lines with date patterns are titles
    if _DATE_RE.search(stripped):
        return True
    # Lines that start with uppercase and are short-ish are likely titles/labels
    if stripped[0].isupper() and len(stripped) < 120:
        return True
    return False


def _write_sheet(writer, df, sheet_name, workbook):
    """Write a single sheet with professional formatting."""

    # ── Define formats ──────────────────────────────────────────────
    header_fmt = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#2B579A',
        'border': 1,
        'border_color': '#1B3A6B',
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'font_size': 11,
        'font_name': 'Calibri',
    })

    cell_fmt = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'border_color': '#D9D9D9',
        'font_size': 10,
        'font_name': 'Calibri',
    })

    cell_alt_fmt = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'border_color': '#D9D9D9',
        'bg_color': '#F2F7FC',
        'font_size': 10,
        'font_name': 'Calibri',
    })

    name_fmt = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'border_color': '#D9D9D9',
        'bold': True,
        'font_size': 10,
        'font_name': 'Calibri',
    })

    name_alt_fmt = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'border_color': '#D9D9D9',
        'bold': True,
        'bg_color': '#F2F7FC',
        'font_size': 10,
        'font_name': 'Calibri',
    })

    # Bold formats for project titles
    bold_fmt = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'border_color': '#D9D9D9',
        'bold': True,
        'font_size': 10,
        'font_name': 'Calibri',
        'font_color': '#1B3A6B',
    })

    bold_alt_fmt = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'border_color': '#D9D9D9',
        'bold': True,
        'bg_color': '#F2F7FC',
        'font_size': 10,
        'font_name': 'Calibri',
        'font_color': '#1B3A6B',
    })

    # ── Column width config ─────────────────────────────────────────
    col_widths = {
        'S.No': 6,
        'Name': 20,
        'Contact No': 15,
        'Email': 28,
        'Role': 22,
        'Skills': 45,
        'Education': 45,
        'Projects': 55,
        'Experience & Internships': 50,
        'Area of Interest / Objective': 35,
        'Awards & Achievements': 40,
        'Extra Curriculars & Leadership': 40,
        'Registration No': 18,
        'LinkedIn': 25,
        'GitHub': 22,
        'Source File': 25,
    }

    # ── Prepare data ────────────────────────────────────────────────
    # Add serial number column
    display_df = df.copy()
    display_df.insert(0, 'S.No', range(1, len(display_df) + 1))

    # Clean all cell values
    for col in display_df.columns:
        display_df[col] = display_df[col].apply(_clean_cell)

    # ── Write to sheet ──────────────────────────────────────────────
    worksheet = workbook.add_worksheet(sheet_name)

    # Freeze the header row
    worksheet.freeze_panes(1, 0)

    # Write headers
    for col_idx, col_name in enumerate(display_df.columns):
        worksheet.write(0, col_idx, col_name, header_fmt)

    # Set header row height
    worksheet.set_row(0, 30)

    # Write data rows with alternating colors
    for row_idx in range(len(display_df)):
        is_alt = row_idx % 2 == 1
        max_height = 20

        for col_idx, col_name in enumerate(display_df.columns):
            value = display_df.iloc[row_idx, col_idx]
            row_num = row_idx + 1

            # Pick base format
            if col_name == 'Name':
                fmt = name_alt_fmt if is_alt else name_fmt
                worksheet.write(row_num, col_idx, value, fmt)
            elif col_name in _RICH_COLUMNS and value and '\n' in value:
                # Rich string: bold headings, regular descriptions
                base_fmt = cell_alt_fmt if is_alt else cell_fmt
                b_fmt = bold_alt_fmt if is_alt else bold_fmt

                segments = []
                lines = [l for l in value.split('\n') if l.strip()]
                for li, line in enumerate(lines):
                    if li > 0:
                        segments.append(base_fmt)
                        segments.append('\n')
                    if _is_heading_line(line):
                        segments.append(b_fmt)
                    else:
                        segments.append(base_fmt)
                    segments.append(line)

                if len(segments) >= 2:
                    try:
                        worksheet.write_rich_string(row_num, col_idx, *segments, base_fmt)
                    except Exception:
                        worksheet.write(row_num, col_idx, value, base_fmt)
                else:
                    worksheet.write(row_num, col_idx, value, base_fmt)
            else:
                fmt = cell_alt_fmt if is_alt else cell_fmt
                worksheet.write(row_num, col_idx, value, fmt)

            # Estimate row height from longest cell in the row
            w = col_widths.get(col_name, 30)
            h = _estimate_row_height(value, w)
            if h > max_height:
                max_height = h

        worksheet.set_row(row_idx + 1, min(max_height, 200))

    # Set column widths
    for col_idx, col_name in enumerate(display_df.columns):
        width = col_widths.get(col_name, 30)
        worksheet.set_column(col_idx, col_idx, width)

    # Auto-filter on header row
    if len(display_df) > 0:
        worksheet.autofilter(0, 0, len(display_df), len(display_df.columns) - 1)


def generate_excel(parsed_resumes):
    """
    Generates a professionally formatted Excel file from a list of
    parsed resume dictionaries. Each Role gets its own sheet, plus
    an "All Resumes" master sheet.
    """
    if not parsed_resumes:
        return b""

    # Ensure consistent column ordering
    column_order = [
        'Name', 'Contact No', 'Email', 'Role', 'Skills',
        'Education', 'Projects', 'Experience & Internships',
        'Area of Interest / Objective', 'Awards & Achievements',
        'Extra Curriculars & Leadership', 'Registration No',
        'LinkedIn', 'GitHub', 'Source File'
    ]

    df_all = pd.DataFrame(parsed_resumes)

    # Reorder columns (only keep columns that exist)
    ordered_cols = [c for c in column_order if c in df_all.columns]
    # Add any extra columns not in our order list
    extra_cols = [c for c in df_all.columns if c not in column_order]
    df_all = df_all[ordered_cols + extra_cols]

    output = io.BytesIO()

    workbook = None
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # ── Master sheet ────────────────────────────────────────────
        _write_sheet(writer, df_all, "All Resumes", workbook)

        # ── Per-role sheets ─────────────────────────────────────────
        if 'Role' in df_all.columns:
            for role in df_all['Role'].unique():
                role_df = df_all[df_all['Role'] == role].reset_index(drop=True)

                # Sanitize sheet name (Excel max 31 chars, no special chars)
                sname = str(role)[:31]
                for ch in [':', '/', '\\', '?', '*', '[', ']']:
                    sname = sname.replace(ch, '')
                sname = sname.strip() or "Uncategorized"

                _write_sheet(writer, role_df, sname, workbook)

    return output.getvalue()
