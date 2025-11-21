import os
import io
import httpx
import asyncpg
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from openpyxl import load_workbook

# -----------------------------
# CONFIG
# -----------------------------
TEMPLATE_S3_URL = "https://your-bucket.s3.amazonaws.com/template.xlsx"

DATABASE_URL = "postgresql://postgres:Bcast%40123@164.52.219.253:5432/sp_dev"

TARGET_SHEET_NAME = "ProjectLookup"

# -----------------------------
# FASTAPI SETUP
# -----------------------------
app = FastAPI(title="Excel Project Code Exporter")


class ExportRequest(BaseModel):
    table: str   # accepted but unused (placeholder for future business logic)


@app.post("/export")
async def export(req: ExportRequest):

    try:
        # 1) Fetch Excel template from S3
        with open("/home/developer/Activity_template.xlsx", "rb") as f:
            template_bytes = f.read()
        #async with httpx.AsyncClient(timeout=30.0) as client:
        #/  r = await client.get(TEMPLATE_S3_URL)
        #/ if r.status_code != 200:
        #/      raise HTTPException(
        #/          status_code=500,
        #/          detail="Failed to download Excel template from S3"
        #/      )

        #/template_bytes = r.content

        # 2) Load workbook from memory
        wb_io = io.BytesIO(template_bytes)
        wb = load_workbook(
            wb_io,
            keep_vba=True,  # Preserves VBA and more XML elements
            data_only=False,
            keep_links=True
       )

        # 3) Get specific worksheet
        if TARGET_SHEET_NAME not in wb.sheetnames:
            raise HTTPException(
                status_code=500,
                detail=f"Worksheet '{TARGET_SHEET_NAME}' not found in template"
            )

        ws = wb[TARGET_SHEET_NAME]

        # 4) Run DB query
        conn = await asyncpg.connect(DATABASE_URL)
        try:
            rows = await conn.fetch(
                "SELECT project_code FROM projects WHERE is_deleted = false;"
            )
        finally:
            await conn.close()

        # 5) Write to Excel starting at A1 always
        for i, row in enumerate(rows, start=1):
            ws.cell(row=i, column=1).value = row["project_code"]

        # 6) Save workbook to buffer
        out_io = io.BytesIO()
        wb.save(out_io)
        out_io.seek(0)

        return StreamingResponse(
            out_io,
            media_type=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition": 'attachment; filename="ProjectLookup_filled.xlsx"'
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        print("Unexpected server error:", str(e))
        raise HTTPException(status_code=500, detail="Internal server error")

