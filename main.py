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
    # Инициализация приложения
    def __init__(self, master):
        # Установка главного окна
        self.master = master
        master.title("Форматирование документа")

        # Определение стилей для кнопок
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Text.TButton', background='#568fba', padding=2, font=('Google Sans', 14), foreground='white')
        style.map('Text.TButton', background=[('active', '#274873')])

        # Центрирование элементов в окне
        master.columnconfigure(0, weight=1)
        master.columnconfigure(1, weight=1)
        master.columnconfigure(2, weight=1)
# Переменные для хранения путей к файлам
        self.file_path = tk.StringVar()
        self.download_path = tk.StringVar()
