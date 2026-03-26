from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import tempfile, json, base64
from pathlib import Path
import badge_maker as bm

app = FastAPI()
templates = Jinja2Templates(directory="templates")
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
SPEC_PATH = Path(__file__).parent / "designs" / "design_spec.json"
SPEC = json.loads(SPEC_PATH.read_text(encoding="utf-8"))


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {
        "request": request,
        "spec": SPEC,
    })


@app.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(await file.read())
    tmp.close()
    result = bm.read_excel(tmp.name)
    if result and isinstance(result[0], dict) and "error" in result[0]:
        return JSONResponse({"error": result[0]["error"]})
    return JSONResponse({"people": result, "tmp_path": tmp.name, "count": len(result)})


@app.post("/upload-logo")
async def upload_logo(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix or ".png"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    data = await file.read()
    tmp.write(data)
    tmp.close()
    b64 = base64.b64encode(data).decode()
    return JSONResponse({"tmp_path": tmp.name, "b64": b64})


@app.post("/generate")
async def generate(
    excel_path:   str  = Form(...),
    logo_path:    str  = Form(""),
    company_name: str  = Form(""),
    badge_type:   str  = Form(...),
    design:       str  = Form(...),
    color:        str  = Form(...),
    size:         str  = Form(...),
    layout:       str  = Form(...),
    event_mode:   str  = Form("false"),
    event_name:   str  = Form(""),
    event_date:   str  = Form(""),
):
    people = bm.read_excel(excel_path)
    if not people or (isinstance(people[0], dict) and "error" in people[0]):
        return JSONResponse({"error": "엑셀 읽기 실패"}, status_code=400)
    em = event_mode.lower() == "true"
    event_info = {"event_name": event_name, "event_date": event_date} if em else None
    result = bm.generate_badges(
        people=people,
        logo_path=logo_path or None,
        company_name=company_name,
        badge_type=badge_type,
        design=design,
        color=color,
        size=size,
        layout=layout,
        event_mode=em,
        event_info=event_info,
        output_dir=str(OUTPUT_DIR),
    )
    if not result:
        return JSONResponse({"error": "생성 실패"}, status_code=500)
    return JSONResponse({"filename": Path(result).name, "count": len(people)})


@app.get("/download/{filename}")
async def download(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        return JSONResponse({"error": "파일 없음"}, status_code=404)
    return FileResponse(
        path, filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


if __name__ == "__main__":
    import uvicorn
    import os
    
    # Railway가 환경변수로 주는 포트를 사용하되, 없으면 8080을 씁니다.
    port = int(os.environ.get("PORT", 8080))
    
    # host를 0.0.0.0으로 해야 외부에서 접속이 가능해집니다!
    uvicorn.run(app, host="0.0.0.0", port=port)
