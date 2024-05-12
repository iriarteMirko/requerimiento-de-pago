from docx import Document
from docx.shared import Pt
from src.utils.resource_path import resource_path


def generar_doc(modelo_2, replacements, ruta_doc):
    doc = Document(modelo_2)
    for paragraph in doc.paragraphs:
        for key, attributes in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, attributes["value"])
                run = paragraph.runs[0]
                run.font.name = 'Arial'
                run.font.size = Pt(attributes["font_size"])
                run.bold = attributes.get("bold", False)
    doc.save(ruta_doc)