"""
Project Management API

A FastAPI application that manages construction projects with filtering,
data export to Excel, and real-time project tracking.
"""

from fastapi import FastAPI, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, RedirectResponse
from pydantic import BaseModel
from typing import Optional, List
import pandas as pd
from pathlib import Path
import json
from datetime import datetime

app = FastAPI(
    title="Project Management API",
    description="API for managing construction projects with filtering and export capabilities",
    version="1.0.0"
)

# Load project data
DATA_FILE = Path(__file__).parent.parent / "projects_sample_data.json"

with open(DATA_FILE, "r", encoding="utf-8") as f:
    projects_data = json.load(f)

# Define Pydantic model for validation
class Project(BaseModel):
    no: Optional[str] = None
    project_number: str
    project_description: str
    zone: str
    capacity: Optional[int] = None
    mh: Optional[int] = None
    por_received_date: Optional[str] = None
    estimated_budget: Optional[float] = None
    por_no: Optional[str] = None
    req_no: Optional[str] = None
    status: Optional[str] = None
    status_detail: Optional[str] = None
    contractor: Optional[str] = None
    contract_sign: Optional[str] = None
    po_no: Optional[str] = None
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    performance_percentage: Optional[int] = None
    payment: Optional[float] = None
    project_status: Optional[str] = None
    payment_status_percent: Optional[int] = None
    payment_status: Optional[str] = None

    class Config:
        from_attributes = True


@app.get("/")
def root():
    return {
        "message": "Welcome to Project Management API",
        "endpoints": {
            "all_projects": "/projects",
            "get_project": "/projects/{project_number}",
            "filter_by_zone": "/projects/filter/zone?zone=NR",
            "filter_by_status": "/projects/filter/status?status=Approved",
            "filter_by_project_status": "/projects/filter/project_status?project_status=Ongoing",
            "export_to_excel": "/projects/export/excel",
            "export_to_csv": "/projects/export/csv",
            "statistics": "/projects/statistics"
        }
    }


@app.get("/projects", response_model=List[Project])
def get_all_projects():
    """Get all projects"""
    return projects_data


@app.get("/projects/{project_number}", response_model=Project)
def get_project_by_number(project_number: str):
    """Get a specific project by project number"""
    for project in projects_data:
        if project["project_number"] == project_number:
            return project
    raise HTTPException(status_code=404, detail=f"Project {project_number} not found")


@app.get("/projects/filter/zone", response_model=List[Project])
def filter_by_zone(zone: str = Query(..., description="Zone to filter by")):
    """Filter projects by zone"""
    filtered = [p for p in projects_data if p.get("zone") == zone]
    if not filtered:
        raise HTTPException(status_code=404, detail=f"No projects found for zone {zone}")
    return filtered


@app.get("/projects/filter/status", response_model=List[Project])
def filter_by_status(status: str = Query(..., description="Status to filter by")):
    """Filter projects by approval status"""
    filtered = [p for p in projects_data if p.get("status") == status]
    if not filtered:
        raise HTTPException(status_code=404, detail=f"No projects found with status {status}")
    return filtered


@app.get("/projects/filter/project_status", response_model=List[Project])
def filter_by_project_status(
    project_status: str = Query(..., description="Project status to filter by (e.g., Ongoing, paused, Completed)")
):
    """Filter projects by project status"""
    filtered = [p for p in projects_data if p.get("project_status") == project_status]
    if not filtered:
        raise HTTPException(status_code=404, detail=f"No projects found with project_status {project_status}")
    return filtered


@app.get("/projects/filter/contractor", response_model=List[Project])
def filter_by_contractor(contractor: str = Query(..., description="Contractor name to filter by")):
    """Filter projects by contractor"""
    filtered = [
        p for p in projects_data
        if p.get("contractor") and contractor.lower() in p.get("contractor", "").lower()
    ]
    if not filtered:
        raise HTTPException(status_code=404, detail=f"No projects found for contractor {contractor}")
    return filtered


@app.get("/projects/export/excel")
def export_to_excel():
    """Export all projects to Excel file"""
    try:
        df = pd.DataFrame(projects_data)

        # Format currency columns
        for col in ["estimated_budget", "payment"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: f"${x:,.2f}" if x is not None else "")

        # Create Excel file
        output_file = Path("projects_export.xlsx")
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Projects", index=False)
            # Auto-adjust column widths
            workbook = writer.book
            worksheet = writer.sheets["Projects"]
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

        return FileResponse(
            output_file,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="projects_export.xlsx"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting to Excel: {str(e)}")


@app.get("/projects/export/csv")
def export_to_csv():
    """Export all projects to CSV file"""
    try:
        df = pd.DataFrame(projects_data)
        output_file = Path("projects_export.csv")
        df.to_csv(output_file, index=False, encoding="utf-8")

        return FileResponse(
            output_file,
            media_type="text/csv",
            filename="projects_export.csv"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting to CSV: {str(e)}")


@app.get("/projects/statistics")
def get_statistics():
    """Get project statistics and analytics"""
    total_projects = len(projects_data)
    total_budget = sum(p.get("estimated_budget") or 0 for p in projects_data)
    total_payment = sum(p.get("payment") or 0 for p in projects_data)

    # Count by status
    status_counts = {}
    for p in projects_data:
        status = p.get("status") or "Unknown"
        status_counts[status] = status_counts.get(status, 0) + 1

    # Count by project status
    project_status_counts = {}
    for p in projects_data:
        ps = p.get("project_status") or "Unknown"
        project_status_counts[ps] = project_status_counts.get(ps, 0) + 1

    # Count by zone
    zone_counts = {}
    for p in projects_data:
        zone = p.get("zone") or "Unknown"
        zone_counts[zone] = zone_counts.get(zone, 0) + 1

    # Average performance percentage
    performance_values = [p.get("performance_percentage") for p in projects_data if p.get("performance_percentage") is not None]
    avg_performance = sum(performance_values) / len(performance_values) if performance_values else 0

    return {
        "total_projects": total_projects,
        "total_estimated_budget": f"${total_budget:,.2f}",
        "total_payment_made": f"${total_payment:,.2f}",
        "average_performance_percentage": round(avg_performance, 2),
        "by_status": status_counts,
        "by_project_status": project_status_counts,
        "by_zone": zone_counts
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
