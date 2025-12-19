from fastapi import FastAPI, Form
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from fastapi.responses import FileResponse
import uuid
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/generate-ppt")
def generate_ppt(
    title: str = Form(...),
    content: str = Form(...)
):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = content

    filename = f"{uuid.uuid4()}.pptx"
    prs.save(filename)

    return FileResponse(
        filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="result.pptx"
    )
