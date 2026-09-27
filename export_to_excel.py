"""
Utility script to export project data to Excel
"""

import pandas as pd
import json
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def export_projects_to_excel(json_file: str, output_file: str = "projects.xlsx"):
    """
    Export projects from JSON to formatted Excel file
    
    Args:
        json_file: Path to the projects JSON file
        output_file: Output Excel filename
    """
    # Load data
    with open(json_file, "r", encoding="utf-8") as f:
        projects = json.load(f)

    # Create DataFrame
    df = pd.DataFrame(projects)

    # Create Excel writer
    output_path = Path(output_file)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Projects", index=False)

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Projects"]

        # Define styles
        header_fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        # Format headers
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border

        # Format data cells
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

                # Format currency columns
                if cell.column_letter in ["H", "S"]:
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        cell.number_format = '"$"#,##0.00'

                # Format percentage columns
                if cell.column_letter in ["R", "V"]:
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        cell.number_format = '0"%"'

        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # Freeze header row
        worksheet.freeze_panes = "A2"

    print(f"✅ Excel file created: {output_path}")
    print(f"📊 Total records: {len(projects)}")
    return output_path


if __name__ == "__main__":
    json_file = "projects_sample_data.json"
    output_file = "projects.xlsx"

    try:
        export_projects_to_excel(json_file, output_file)
        print("\n✨ Export completed successfully!")
    except FileNotFoundError:
        print(f"❌ Error: {json_file} not found")
    except Exception as e:
        print(f"❌ Error: {str(e)}")
