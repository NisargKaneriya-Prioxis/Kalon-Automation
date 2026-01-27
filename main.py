
import os
import time
import threading
import webbrowser
from io import BytesIO
from typing import Optional

import uvicorn
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from fastapi.concurrency import run_in_threadpool

# Import your report generator
from app.generator import generate_clawback_report
from app.generator import NoDataInRange 


# FastAPI app
app = FastAPI(
    title="Clawback Report API",
    version="1.0.0",
    description="Upload two Excel files + dates to generate Clawback Report",
)


# -----------------------------
#     HEALTH CHECK ENDPOINT
# -----------------------------
@app.get("/health")
async def health():
    return {"status": "OK"}


# -----------------------------
#     REPORT GENERATION
# -----------------------------
@app.post("/generate-report")
async def generate_report(
    input1: UploadFile = File(..., description="Raw data Excel"),
    input2: UploadFile = File(..., description="Team mapping Excel"),
    start_date: str = Form(...),
    end_date: str = Form(...),
    report_title: Optional[str] = Form(None),
):
    """Upload 2 Excels + date range -> returns final XLSX report"""

    # Validate file types
    for f in (input1, input2):
        if not f.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(400, f"{f.filename} must be an Excel file (.xlsx or .xlsm).")

    # Read bytes
    in1_bytes = await input1.read()
    in2_bytes = await input2.read()

    try:
        excel_bytes, file_name = await run_in_threadpool(
            generate_clawback_report,
            in1_bytes,
            in2_bytes,
            start_date,
            end_date,
            report_title
        )
    except NoDataInRange as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Processing Error: {e}")

    # Send file back
    return StreamingResponse(
        BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'}
    )


# -----------------------------
# OPEN BROWSER WHEN APP STARTS
# -----------------------------
def open_docs():
    time.sleep(1)
    webbrowser.open("http://127.0.0.1:3978/docs")


# -----------------------------
#     DEV MODE RUNNER
# -----------------------------
if __name__ == "__main__":
    threading.Thread(target=open_docs, daemon=True).start()
    uvicorn.run("main:app", host="0.0.0.0", port=3978, reload=True)
