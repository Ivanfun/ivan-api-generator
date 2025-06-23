from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from pathlib import Path
from tempfile import NamedTemporaryFile
from backend.logic import generate_api_doc
import os

app = FastAPI()

@app.get("/", response_class=HTMLResponse)
def get_index():
    """首頁：回傳前端頁面（HTML）"""
    index_file = Path("frontend/index.html")
    if not index_file.exists():
        return HTMLResponse(content="<h1>前端頁面不存在</h1>", status_code=404)
    return index_file.read_text(encoding="utf-8")

def cleanup_files(paths_to_delete):
    """背景資源清理：安全刪除暫存檔案"""
    for path in paths_to_delete:
        if path and path.exists():
            try:
                os.remove(path)
            except OSError:
                pass  # 忽略刪除錯誤（檔案可能已被刪除）

@app.post("/upload/")
async def upload_files(
    background_tasks: BackgroundTasks,
    excel_file: UploadFile = File(...),
    word_template: UploadFile = File(...),
    sql_properties: UploadFile = File(...)
):
    all_paths = []
    # 【變更點 1】新增一個成功旗標，預設為 False
    success = False
    
    try:
        # --- 檔案暫存 ---
        # 這部分邏輯不變
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel, \
             NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word, \
             NamedTemporaryFile(delete=False, suffix=".properties") as tmp_props:

            tmp_excel_path = Path(tmp_excel.name)
            tmp_word_path = Path(tmp_word.name)
            tmp_props_path = Path(tmp_props.name)
            all_paths.extend([tmp_excel_path, tmp_word_path, tmp_props_path])

            tmp_excel.write(await excel_file.read())
            tmp_word.write(await word_template.read())
            tmp_props.write(await sql_properties.read())

        output_path = tmp_word_path.with_name(f"output_{tmp_word_path.stem}.docx")
        all_paths.append(output_path)

        # --- 核心邏輯調用 ---
        generate_api_doc(
            excel_path=tmp_excel_path,
            word_template_path=tmp_word_path,
            output_path=output_path,
            sql_properties_path=tmp_props_path
        )

        # --- 成功路徑 ---
        # 【變更點 2】如果程式能執行到這裡，代表處理成功，將旗標設為 True
        success = True
        
        # 成功時，清理任務必須交給 BackgroundTasks，才能在檔案回傳後執行
        background_tasks.add_task(cleanup_files, all_paths)

        return FileResponse(
            path=output_path,
            filename="API_規格書.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            background=background_tasks
        )

    except ValueError as e:
        # 失敗時，直接拋出例外，讓 finally 去處理清理工作
        raise HTTPException(status_code=400, detail=f"檔案內容或格式錯誤: {e}")
    except Exception as e:
        # 失敗時，直接拋出例外，讓 finally 去處理清理工作
        raise HTTPException(status_code=500, detail=f"伺服器內部錯誤: {e}")
    finally:
        # 【變更點 3】新增 finally 區塊
        # 這裡的邏輯至關重要：只有在處理不成功時，才執行清理。
        # 如果成功，清理工作已交給 background_tasks，此處不可再清理，否則會刪掉正要回傳的檔案。
        if not success:
            cleanup_files(all_paths)
