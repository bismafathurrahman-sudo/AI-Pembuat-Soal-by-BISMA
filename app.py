
import os
import json
import io
from pathlib import Path

import fitz
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request
from openai import OpenAI

load_dotenv()

app = FastAPI()
templates = Jinja2Templates(directory="templates")

API_KEY = os.getenv("GROQ_API_KEY")
MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-20b")

client = OpenAI(
    api_key=API_KEY,
    base_url="https://api.groq.com/openai/v1"
)

MAX_FILE_MB = int(os.getenv("MAX_FILE_MB", "20"))
MAX_FILE_SIZE = MAX_FILE_MB * 1024 * 1024


def extract_text(filename: str, content: bytes) -> str:
    extension = Path(filename).suffix.lower()

    try:
        if extension == ".txt":
            return content.decode("utf-8", errors="ignore")

        elif extension == ".pdf":
            pdf = fitz.open(stream=content, filetype="pdf")
            return "\n".join(page.get_text() for page in pdf)

        elif extension == ".docx":
            doc = Document(io.BytesIO(content))
            return "\n".join(p.text for p in doc.paragraphs)

        elif extension == ".pptx":
            presentation = Presentation(io.BytesIO(content))
            slides = []

            for number, slide in enumerate(presentation.slides, start=1):
                slide_text = [
                    shape.text
                    for shape in slide.shapes
                    if hasattr(shape, "text") and shape.text.strip()
                ]
                slides.append(
                    f"Slide {number}\n" + "\n".join(slide_text)
                )

            return "\n\n".join(slides)

        elif extension == ".xlsx":
            workbook = load_workbook(
                io.BytesIO(content),
                data_only=True,
                read_only=True
            )

            sheets = []

            for sheet in workbook.worksheets:
                rows = []

                for row in sheet.iter_rows(values_only=True):
                    values = [
                        str(value)
                        for value in row
                        if value is not None
                    ]

                    if values:
                        rows.append(" | ".join(values))

                sheets.append(
                    f"Sheet: {sheet.title}\n" + "\n".join(rows)
                )

            return "\n\n".join(sheets)

        else:
            raise ValueError("Format file tidak didukung.")

    except Exception as error:
        raise ValueError(f"Gagal membaca file: {error}")


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html"
    )


@app.post("/api/generate")
async def generate_questions(
    file: UploadFile = File(...),
    jumlah: int = Form(10),
    kesulitan: str = Form("Sedang"),
    tipe: str = Form("Pilihan Ganda"),
    mode: str = Form("Belajar")
):
    if not client:
        raise HTTPException(
            status_code=500,
            detail="GROQ_API_KEY belum diatur. Periksa file .env."
        )

    if not 5 <= jumlah <= 100:
        raise HTTPException(
            status_code=400,
            detail="Jumlah soal harus antara 5 sampai 100."
        )

    allowed = {".pdf", ".docx", ".pptx", ".xlsx", ".txt"}
    extension = Path(file.filename or "").suffix.lower()

    if extension not in allowed:
        raise HTTPException(
            status_code=400,
            detail="Format file harus PDF, DOCX, PPTX, XLSX, atau TXT."
        )

    content = await file.read()

    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"Ukuran file maksimal {MAX_FILE_MB} MB."
        )

    try:
        material = extract_text(file.filename, content)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error))

    if not material.strip():
        raise HTTPException(
            status_code=400,
            detail="Tidak ditemukan teks dalam file. Pastikan file bukan hasil scan gambar."
        )

    material = material[:60000]

    prompt = f"""
Kamu adalah AI pengajar dan pembuat soal ujian profesional.

Buat {jumlah} soal berdasarkan materi berikut.

TINGKAT KESULITAN: {kesulitan}
JENIS SOAL: {tipe}
MODE: {mode}

Ketentuan:
- Soal harus sesuai isi materi.
- Jangan membuat pertanyaan yang jawabannya tidak ada dasar dalam materi.
- Buat soal bervariasi dan tidak berulang.
- Jika pilihan ganda, berikan 4 opsi A, B, C, D.
- Sertakan jawaban benar dan pembahasan yang jelas.
- Tingkat kesulitan harus sesuai permintaan.
- Gunakan Bahasa Indonesia.

Kembalikan hanya JSON valid dengan struktur:
{{
  "judul": "Judul latihan",
  "questions": [
    {{
      "question": "Pertanyaan",
      "options": {{
        "A": "Pilihan A",
        "B": "Pilihan B",
        "C": "Pilihan C",
        "D": "Pilihan D"
      }},
      "answer": "A",
      "explanation": "Pembahasan jawaban",
      "difficulty": "Sedang"
    }}
  ]
}}

Jika jenis soal bukan pilihan ganda, tetap gunakan struktur options
sebagai objek kosong dan answer sebagai jawaban teks.

MATERI:
{material}
"""

    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {
                    "role": "system",
                    "content": "Kamu adalah asisten pendidikan yang teliti."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            response_format={"type": "json_object"},
            temperature=0.7
        )

        result = json.loads(
            response.choices[0].message.content
        )

        return {
            "success": True,
            "data": result
        }

    except Exception as error:
        print("ERROR GROQ:", repr(error))
        raise HTTPException(
            status_code=500,
            detail=f"Gagal membuat soal: {str(error)}"
        )


@app.post("/api/evaluate")
async def evaluate_answers(payload: dict):
    questions = payload.get("questions", [])
    answers = payload.get("answers", {})

    if not questions:
        raise HTTPException(
            status_code=400,
            detail="Data soal tidak ditemukan."
        )

    correct = 0
    results = []

    for index, question in enumerate(questions):
        user_answer = str(
            answers.get(str(index), "")
        ).strip()

        correct_answer = str(
            question.get("answer", "")
        ).strip()

        is_correct = (
            user_answer.lower() == correct_answer.lower()
        )

        if is_correct:
            correct += 1

        results.append({
            "question": question.get("question", ""),
            "user_answer": user_answer,
            "correct_answer": correct_answer,
            "is_correct": is_correct,
            "explanation": question.get("explanation", "")
        })

    total = len(questions)
    score = round((correct / total) * 100, 2)

    return {
        "success": True,
        "total": total,
        "correct": correct,
        "incorrect": total - correct,
        "score": score,
        "results": results
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
