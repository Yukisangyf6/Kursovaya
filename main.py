import tkinter as tk
from tkinter import ttk, filedialog
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
import fitz
import language_tool_python
class DocumentFormatterApp:
