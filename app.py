import gradio as gr
import pytesseract
from PIL import Image
from docx import Document
import tempfile

# If on Windows, set this path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def image_to_word(image):
    # OCR
    text = pytesseract.image_to_string(Image.open(image))

    # Create Word document
    doc = Document()
    for line in text.split("\n"):
        doc.add_paragraph(line)

    # Save temporary docx
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)

    return tmp.name

iface = gr.Interface(
    fn=image_to_word,
    inputs=gr.Image(type="filepath"),
    outputs=gr.File(label="Download Editable Word File"),
    title="Image to Editable Word Converter",
    description="Upload an image of notes or document and get an editable Word file."
)

iface.launch()
