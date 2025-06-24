from fastapi import APIRouter, Depends
from pydantic import BaseModel
from typing import Optional
from database import get_db_connection
from fastapi.responses import JSONResponse

router = APIRouter()

# Define request schema
class SearchFilters(BaseModel):
    project_key: Optional[str] = None
    status: Optional[str] = None
    issuetype: Optional[str] = None
    assignee: Optional[str] = None
    reporter: Optional[str] = None
    issue_key: Optional[str] = None  # New input

@router.post("/search-issues")
def search_issues(filters: SearchFilters, conn=Depends(get_db_connection)):
    try:
        cursor = conn.cursor()

        # Base query
        base_query = """
            SELECT ji.issuenum, ji.pkey, ji.summary, iss.name AS status, pr.pname AS priority,
                   ji.created, ji.updated, p.pname AS project, it.pname AS issue_type,
                   assignee.display_name AS assignee_name, reporter.display_name AS reporter_name,
                   ji.id AS issue_id
            FROM jiraissue ji
            JOIN project p ON ji.project = p.id
            JOIN issuetype it ON ji.issuetype = it.id
            JOIN issuestatus iss ON ji.issuestatus = iss.id
            LEFT JOIN priority pr ON ji.priority = pr.id
            LEFT JOIN cwd_user assignee ON assignee.lower_user_name = LOWER(ji.assignee)
            LEFT JOIN cwd_user reporter ON reporter.lower_user_name = LOWER(ji.reporter)
            WHERE 1=1
        """

        params = {}

        if filters.project_key:
            base_query += " AND p.pkey = :project_key"
            params["project_key"] = filters.project_key

        if filters.status:
            base_query += " AND iss.name = :status"
            params["status"] = filters.status

        if filters.issuetype:
            base_query += " AND it.pname = :issuetype"
            params["issuetype"] = filters.issuetype

        if filters.assignee:
            base_query += " AND assignee.display_name = :assignee"
            params["assignee"] = filters.assignee

        if filters.reporter:
            base_query += " AND reporter.display_name = :reporter"
            params["reporter"] = filters.reporter

        if filters.issue_key:
            base_query += " AND (p.pkey || '-' || ji.issuenum) = :issue_key"
            params["issue_key"] = filters.issue_key

        cursor.execute(base_query, params)
        rows = cursor.fetchall()

        issues = []
        for row in rows:
            (issuenum, pkey, summary, status, priority, created, updated, project, issue_type,
             assignee_name, reporter_name, issue_id) = row

            issues.append({
                "issuenum": issuenum,
                "pkey": pkey,
                "summary": summary,
                "status": status,
                "priority": priority,
                "created": created.date().isoformat() if created else None,
                "updated": updated.date().isoformat() if updated else None,
                "project": project,
                "issue_type": issue_type,
                "assignee": assignee_name,
                "reporter": reporter_name
            })

        return JSONResponse(content=issues)

    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)
