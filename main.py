
import os
import tempfile
import shutil
import uvicorn
import threading
import webbrowser
import time
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import FileResponse

from app.generator import generate_clawback_report, NoDataInRange

app = FastAPI(title="Clawback Report API")

def cleanup_temp_files(file_paths: list):
    """Deletes the temporary files from the server after the response is sent."""
    for path in file_paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
                print(f"Successfully deleted temp file: {path}")
            except Exception as e:
                print(f"Error deleting temp file {path}: {e}")

@app.post("/generate-report")
async def generate_report(
    background_tasks: BackgroundTasks,
    input1: UploadFile = File(...),
    input2: UploadFile = File(...),
    start_date: str = Form(...),
    end_date: str = Form(...),
    report_title: str = Form(None),
):
    # 1. Create unique temporary paths for inputs
    temp_dir = tempfile.gettempdir()
    in1_path = os.path.join(temp_dir, f"upload1_{time.time()}_{input1.filename}")
    in2_path = os.path.join(temp_dir, f"upload2_{time.time()}_{input2.filename}")

    # 2. Save uploaded files to the temp directory
    try:
        with open(in1_path, "wb") as buffer:
            shutil.copyfileobj(input1.file, buffer)
        with open(in2_path, "wb") as buffer:
            shutil.copyfileobj(input2.file, buffer)
            
        # 3. Process the files
        output_path, file_name = generate_clawback_report(
            in1_path, in2_path, start_date, end_date, report_title
        )

        # 4. Add all files to the cleanup queue
        background_tasks.add_task(cleanup_temp_files, [in1_path, in2_path, output_path])

        # 5. Return the file from disk
        return FileResponse(
            path=output_path,
            filename=file_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except NoDataInRange as e:
        cleanup_temp_files([in1_path, in2_path])
        raise HTTPException(422, str(e))
    except Exception as e:
        cleanup_temp_files([in1_path, in2_path])
        raise HTTPException(500, f"Processing Error: {str(e)}")

def open_docs():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:3978/docs")

if __name__ == "__main__":
    threading.Thread(target=open_docs, daemon=True).start()
    uvicorn.run("main:app", host="0.0.0.0", port=3978, reload=True)