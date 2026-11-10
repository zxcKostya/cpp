import os
import datetime
from fastapi import FastAPI, UploadFile, File, Request, Depends
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docx import Document
from docx.shared import Inches, Pt
from sqlalchemy import create_engine, Column, Integer, String, DateTime, JSON
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session


DATABASE_URL = "sqlite:///./docx_checker.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class ValidationHistory(Base):
    __tablename__ = "validation_history"
    
    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String, nullable=False)
    timestamp = Column(DateTime, default=datetime.datetime.utcnow)
    status = Column(String, nullable=False)  
    errors = Column(JSON, nullable=False)   

Base.metadata.create_all(bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def check_docx_formatting(file_path: str) -> list:
    violations = []
    try:
        doc = Document(file_path)
    except Exception as e:
        return [f"Не удалось прочитать файл как DOCX: {str(e)}"]

    for i, section in enumerate(doc.sections):
        if section.top_margin < Inches(0.7):
            violations.append(f"Раздел {i+1}: Верхнее поле меньше нормы (должно быть >= 20 мм).")
        if section.bottom_margin < Inches(0.7):
            violations.append(f"Раздел {i+1}: Нижнее поле меньше нормы (должно быть >= 20 мм).")
        if section.left_margin < Inches(1.1):
            violations.append(f"Раздел {i+1}: Левое поле меньше нормы (должно быть >= 30 мм).")
        if section.right_margin < Inches(0.35):
            violations.append(f"Раздел {i+1}: Правое поле меньше нормы (должно быть >= 10 мм).")

    for p_idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
            
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            
            font_name = run.font.name or paragraph.style.font.name
            if font_name and font_name != "Times New Roman":
                violations.append(f"Абзац {p_idx+1}: Использован некорректный шрифт '{font_name}' вместо Times New Roman.")
                break 

                
            font_size = run.font.size or paragraph.style.font.size
            if font_size and font_size != Pt(14):
                size_pt = font_size.pt
                violations.append(f"Абзац {p_idx+1}: Неверный размер шрифта {size_pt}pt (ожидается 14pt).")
                break


    if len(doc.paragraphs) > 0:
        title_text = "".join([p.text for p in doc.paragraphs[:15]]).lower()
        if "министерство" not in title_text and "тусур" not in title_text:
            violations.append("Титульный лист: Не обнаружены обязательные реквизиты организации (ТУСУР / Министерство).")

    return violations


app = FastAPI(title="Автоматизированная проверка отчетов DOCX")


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Проверка DOCX отчетов</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background-color: #f4f6f9; color: #333; }
        .container { max-width: 800px; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h1 { color: #1a5c96; }
        .upload-box { border: 2px dashed #1a5c96; padding: 30px; text-align: center; border-radius: 6px; cursor: pointer; background: #f9fbfd; }
        button { background: #1a5c96; color: white; border: none; padding: 10px 20px; font-size: 16px; border-radius: 4px; cursor: pointer; margin-top: 15px; }
        button:hover { background: #134470; }
        .result { margin-top: 20px; padding: 15px; border-radius: 4px; display: none; }
        .success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .danger { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        ul { padding-left: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Автоматическая проверка оформления отчетов (DOCX)</h1>
        <p>Выберите файл отчета для автоматической проверки полей, шрифтов и структуры на соответствие нормоконтролю.</p>
        
        <form id="uploadForm" enctype="multipart/form-data">
            <div class="upload-box">
                <input type="file" id="fileInput" name="file" accept=".docx" required><br>
            </div>
            <button type="submit">Запустить проверку</button>
        </form>

        <div id="resultBox" class="result"></div>
    </div>

    <script>
        document.getElementById('uploadForm').onsubmit = async (e) => {
            e.preventDefault();
            const fileInput = document.getElementById('fileInput');
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);

            const resultBox = document.getElementById('resultBox');
            resultBox.style.display = 'block';
            resultBox.className = 'result';
            resultBox.innerHTML = '⚡ Обработка и анализ документа...';

            try {
                const response = await fetch('/api/check', {
                    method: 'POST',
                    body: formData
                });
                const data = await response.json();

                if (data.status === "Passed") {
                    resultBox.className = 'result success';
                    resultBox.innerHTML = '<b>Успешно!</b> Ошибок оформления не обнаружено.';
                } else {
                    resultBox.className = 'result danger';
                    let html = '<b>Обнаружены нарушения оформления:</b><ul>';
                    data.violations.forEach(v => { html += `<li>${v}</li>`; });
                    html += '</ul>';
                    resultBox.innerHTML = html;
                }
            } catch (err) {
                resultBox.className = 'result danger';
                resultBox.innerText = 'Ошибка при отправке файла на сервер.';
            }
        };
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def get_gui():
    return HTML_TEMPLATE

@app.post("/api/check")
async def api_check_document(file: UploadFile = File(...), db: Session = Depends(get_db)):

    temp_path = f"temp_{file.filename}"
    with open(temp_path, "wb") as buffer:
        buffer.write(await file.read())
    

    violations = check_docx_formatting(temp_path)
    

    if os.path.exists(temp_path):
        os.remove(temp_path)
        
    status = "Passed" if len(violations) == 0 else "Failed"
    
  
    history_entry = ValidationHistory(
        filename=file.filename,
        status=status,
        errors=violations
    )
    db.add(history_entry)
    db.commit()
    
    return JSONResponse(content={
        "filename": file.filename,
        "status": status,
        "violations": violations,
        "timestamp": datetime.datetime.now().isoformat()
    })

@app.get("/api/history")
def get_history(db: Session = Depends(get_db)):
    history = db.query(ValidationHistory).order_by(ValidationHistory.id.desc()).all()
    return [{
        "id": h.id,
        "filename": h.filename,
        "timestamp": h.timestamp.isoformat(),
        "status": h.status,
        "errors": h.errors
    } for h in history]

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
