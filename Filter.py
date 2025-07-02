from fastapi import APIRouter, Depends, Query
from pydantic import BaseModel
from typing import Optional
from database import get_db_connection
from fastapi.responses import JSONResponse

router = APIRouter()

# Request schema
class SearchFilters(BaseModel):
    status: Optional[str] = None
    issuetype: Optional[str] = None
    assignee: Optional[str] = None
    reporter: Optional[str] = None
    issue_key: Optional[str] = None

@router.post("/search-issues")
def search_issues(
    filters: SearchFilters,
    page: int = Query(1, gt=0),
    page_size: int = Query(10, gt=0, le=100),  # max 100 items per page
    conn=Depends(get_db_connection)
):
    try:
        cursor = conn.cursor()

        # Calculate offset
        offset = (page - 1) * page_size

        # Base query with ROW_NUMBER to eliminate duplicates
        base_query = f"""
            SELECT * FROM (
                SELECT ji.issuenum, ji.summary, ji.description, js.pname AS status,
                       jp.pname AS priority, assignee.display_name AS assignee,
                       reporter.display_name AS reporter, ji.created, ji.updated,
                       p.pkey AS project_key, p.pname AS project_name, it.pname AS issue_type,
                       ji.id, ROW_NUMBER() OVER (PARTITION BY ji.id ORDER BY ji.id) AS rn
                FROM jiraissue ji
                JOIN project p ON ji.project = p.id
                LEFT JOIN issuetype it ON ji.issuetype = it.id
                LEFT JOIN issuestatus js ON ji.issuestatus = js.id
                LEFT JOIN priority jp ON ji.priority = jp.id
                LEFT JOIN cwd_user assignee ON assignee.lower_user_name = LOWER(ji.assignee)
                LEFT JOIN cwd_user reporter ON reporter.lower_user_name = LOWER(ji.reporter)
                WHERE 1=1
        """

        params = {}

        if filters.status:
            base_query += " AND js.pname = :status"
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

        # Close the subquery and apply row filtering and pagination
        base_query += f"""
            ) WHERE rn = 1
            OFFSET {offset} ROWS FETCH NEXT {page_size} ROWS ONLY
        """

        cursor.execute(base_query, params)
        rows = cursor.fetchall()

        issues = []
        for row in rows:
            (issuenum, summary, description, status, priority, assignee_name, reporter_name,
             created, updated, project_key, project_name, issue_type, issue_id, rn) = row

            issues.append({
                "issue_key": f"{project_key}-{issuenum}",
                "summary": summary,
                "description": description,
                "status": status,
                "priority": priority,
                "created": created.date().isoformat() if created else None,
                "updated": updated.date().isoformat() if updated else None,
                "project_key": project_key,
                "project_name": project_name,
                "issue_type": issue_type,
                "assignee": assignee_name,
                "reporter": reporter_name
            })

        return JSONResponse(content={
            "page": page,
            "page_size": page_size,
            "results": issues
        })

    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)
