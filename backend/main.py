from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from pathlib import Path
from tempfile import NamedTemporaryFile
from backend.logic import generate_api_doc

app = FastAPI()

@app.get("/", response_class=HTMLResponse)
def get_index():
    return Path("frontend/index.html").read_text(encoding="utf-8")

@app.post("/upload/")
async def upload_files(
    excel_file: UploadFile = File(...),
    word_template: UploadFile = File(...),
    sql_properties: UploadFile = File(...)
):
    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel, \
         NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word, \
         NamedTemporaryFile(delete=False, suffix=".properties") as tmp_props:

        tmp_excel.write(await excel_file.read())
        tmp_word.write(await word_template.read())
        tmp_props.write(await sql_properties.read())

        output_path = Path(tmp_word.name).with_name("output.docx")

        generate_api_doc(
            excel_path=Path(tmp_excel.name),
            word_template_path=Path(tmp_word.name),
            output_path=output_path,
            sql_properties_path=Path(tmp_props.name)
        )

        return FileResponse(output_path, filename="API_規格書.docx")

