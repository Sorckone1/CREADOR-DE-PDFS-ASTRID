import json
import os
import re
import subprocess
import sys
import unicodedata
from datetime import datetime
from io import BytesIO
from pathlib import Path

import customtkinter as ctk
from PIL import Image, ImageOps, ImageTk
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from tkinter import filedialog, messagebox


# =========================
# APARIENCIA
# =========================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

WINDOW_BG = "#F5F5F7"
CARD_BG = "#FFFFFF"
CARD_BORDER = "#ECECF1"

TEXT_PRIMARY = "#1D1D1F"
TEXT_SECONDARY = "#6E6E73"

BUTTON_COLOR = "#F4C2C2"
BUTTON_HOVER = "#E49AAE"
BUTTON_TEXT = "#FFFFFF"

SOFT_BUTTON_BG = "#FFFFFF"
SOFT_BUTTON_HOVER = "#F8F1F3"
SOFT_BUTTON_TEXT = "#8E5E6C"
SOFT_BUTTON_BORDER = "#E9DDE1"

TITLE_PINK = "#D98CA3"

INPUT_BG = "#FAFAFC"
INPUT_BORDER = "#E6E6EC"

SELECTED_CARD_BORDER = "#D98CA3"
SELECTED_CARD_BG = "#FFF7F8"

PROGRESS_BG = "#ECECF1"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
PDF_EXTENSIONS = {".pdf"}

APP_STATE_FILE = str(Path.home() / ".astrid_lopez_pdf_app_state.json")


# =========================
# FUENTES
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR, "fonts")

MONTSERRAT_REGULAR_PATH = os.path.join(FONT_DIR, "Montserrat-Regular.ttf")
MONTSERRAT_BOLD_PATH = os.path.join(FONT_DIR, "Montserrat-Bold.ttf")

if os.path.exists(MONTSERRAT_REGULAR_PATH):
    ctk.FontManager.load_font(MONTSERRAT_REGULAR_PATH)

if os.path.exists(MONTSERRAT_BOLD_PATH):
    ctk.FontManager.load_font(MONTSERRAT_BOLD_PATH)

FONT_FAMILY = "Montserrat"


# =========================
# UTILIDADES
# =========================
def sanitize_filename(text: str) -> str:
    text = text.strip()
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.upper()
    text = text.replace(" ", "_")
    text = re.sub(r"[^A-Z0-9_\-\.]", "", text)
    text = re.sub(r"_+", "_", text)
    return text or "SIN_NOMBRE"


def parse_date(date_str: str):
    date_str = date_str.strip()
    if not date_str:
        return None

    formats = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"]

    for fmt in formats:
        try:
            d = datetime.strptime(date_str, fmt)
            return d.strftime("%Y-%m-%d")
        except ValueError:
            pass

    return "INVALID_DATE"


def image_to_pdf_bytes(image_path: str) -> BytesIO:
    with Image.open(image_path) as im:
        im = ImageOps.exif_transpose(im)

        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        elif im.mode == "L":
            im = im.convert("RGB")

        w, h = im.size

        img_buffer = BytesIO()
        im.save(img_buffer, format="JPEG", quality=95)
        img_buffer.seek(0)

        pdf_buffer = BytesIO()
        c = canvas.Canvas(pdf_buffer, pagesize=(w, h))
        c.drawImage(ImageReader(img_buffer), 0, 0, width=w, height=h)
        c.showPage()
        c.save()
        pdf_buffer.seek(0)
        return pdf_buffer


def ensure_unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path

    base, ext = os.path.splitext(path)
    counter = 2

    while True:
        new_path = f"{base}_{counter}{ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def is_image(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in IMAGE_EXTENSIONS


def is_pdf(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in PDF_EXTENSIONS


def safe_filename_only(path: str) -> str:
    return os.path.basename(path)


def open_path(path: str):
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def open_folder(path: str):
    folder = path if os.path.isdir(path) else os.path.dirname(path)
    open_path(folder)


def load_app_state() -> dict:
    try:
        if os.path.exists(APP_STATE_FILE):
            with open(APP_STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_app_state(data: dict):
    try:
        with open(APP_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =========================
# APP
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title_font = ctk.CTkFont(family=FONT_FAMILY, size=28, weight="bold")
        self.subtitle_font = ctk.CTkFont(family=FONT_FAMILY, size=12)
        self.section_font = ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold")
        self.label_font = ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold")
        self.button_font = ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold")
        self.primary_button_font = ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold")
        self.small_font = ctk.CTkFont(family=FONT_FAMILY, size=10)
        self.thumb_font = ctk.CTkFont(family=FONT_FAMILY, size=8)
        self.drop_font = ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold")
        self.icon_font = ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold")

        self.title("Astrid Lopez (conversor pdf)")
        self.geometry("740x620")
        self.minsize(640, 560)
        self.configure(fg_color=WINDOW_BG)

        self.files = []
        self.output_dir = ""
        self.preview_images = []
        self.selected_index = None
        self.last_output_file = None
        self.resize_after_id = None

        self.app_state = load_app_state()

        self.build_ui()
        self.load_saved_values()
        self.bind("<Configure>", self.on_resize)

    # =========================
    # UI
    # =========================
    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.grid(row=0, column=0, sticky="nsew", padx=18, pady=16)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(0, weight=0)
        self.main.grid_rowconfigure(1, weight=0)
        self.main.grid_rowconfigure(2, weight=1)
        self.main.grid_rowconfigure(3, weight=0)

        self.center_wrap = ctk.CTkFrame(self.main, fg_color="transparent")
        self.center_wrap.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.center_wrap.grid_columnconfigure(0, weight=1)

        # Header
        self.header = ctk.CTkFrame(self.center_wrap, fg_color="transparent")
        self.header.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        self.title_label = ctk.CTkLabel(
            self.header,
            text="Astrid Lopez",
            text_color=TITLE_PINK,
            font=self.title_font
        )
        self.title_label.pack(anchor="center")

        self.subtitle_label = ctk.CTkLabel(
            self.header,
            text="Conversor PDF",
            text_color=TEXT_SECONDARY,
            font=self.subtitle_font
        )
        self.subtitle_label.pack(anchor="center", pady=(2, 0))

        # Formulario
        self.form_card = ctk.CTkFrame(
            self.center_wrap,
            fg_color=CARD_BG,
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=20
        )
        self.form_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        self.form_inner = ctk.CTkFrame(self.form_card, fg_color="transparent")
        self.form_inner.pack(fill="both", expand=True, padx=14, pady=14)

        self.patient_wrap, self.patient_entry = self.make_field(
            self.form_inner, "Paciente", "Ej. Juan Pérez"
        )
        self.study_wrap, self.study_entry = self.make_field(
            self.form_inner, "Estudio", "Ej. Radiografía"
        )
        self.date_wrap, self.date_entry = self.make_field(
            self.form_inner, "Fecha opcional", "YYYY-MM-DD"
        )

        # Archivos
        self.upload_card = ctk.CTkFrame(
            self.center_wrap,
            fg_color=CARD_BG,
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=20
        )
        self.upload_card.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        self.upload_card.grid_rowconfigure(0, weight=1)
        self.upload_card.grid_columnconfigure(0, weight=1)

        self.upload_inner = ctk.CTkFrame(self.upload_card, fg_color="transparent")
        self.upload_inner.pack(fill="both", expand=True, padx=14, pady=14)

        self.upload_top = ctk.CTkFrame(self.upload_inner, fg_color="transparent")
        self.upload_top.pack(fill="x", pady=(0, 6))

        self.upload_title = ctk.CTkLabel(
            self.upload_top,
            text="Archivos",
            text_color=TEXT_PRIMARY,
            font=self.section_font
        )
        self.upload_title.pack(anchor="w")

        self.files_counter = ctk.CTkLabel(
            self.upload_top,
            text="0 archivos seleccionados",
            text_color=TEXT_SECONDARY,
            font=self.small_font
        )
        self.files_counter.pack(anchor="w", pady=(2, 0))

        self.drop_zone = ctk.CTkFrame(
            self.upload_inner,
            fg_color="#FCFCFE",
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=16,
            height=170
        )
        self.drop_zone.pack(fill="both", expand=True)
        self.drop_zone.pack_propagate(False)

        self.drop_title = ctk.CTkLabel(
            self.drop_zone,
            text="Zona de carga",
            text_color=TEXT_PRIMARY,
            font=self.drop_font
        )
        self.drop_title.pack(pady=(10, 2))

        self.drop_text = ctk.CTkLabel(
            self.drop_zone,
            text="Imágenes, PDFs o carpetas",
            text_color=TEXT_SECONDARY,
            font=self.small_font
        )
        self.drop_text.pack(pady=(0, 5))

        self.preview_scroll = ctk.CTkScrollableFrame(
            self.drop_zone,
            fg_color="transparent",
            height=88,
            orientation="horizontal"
        )
        self.preview_scroll.pack(fill="x", padx=8, pady=(0, 8))

        self.actions_row = ctk.CTkFrame(self.upload_inner, fg_color="transparent")
        self.actions_row.pack(fill="x", pady=(8, 0))

        self.add_images_button = self.make_soft_button(self.actions_row, "Imágenes", self.add_images)
        self.add_pdf_button = self.make_soft_button(self.actions_row, "PDF", self.add_pdf)
        self.add_folder_button = self.make_soft_button(self.actions_row, "Carpeta", self.add_folder)
        self.clear_button = self.make_soft_button(self.actions_row, "Limpiar", self.clear_list)

        self.manage_row = ctk.CTkFrame(self.upload_inner, fg_color="transparent")
        self.manage_row.pack(fill="x", pady=(8, 0))

        self.move_left_button = self.make_soft_button(self.manage_row, "←", self.move_selected_left)
        self.move_right_button = self.make_soft_button(self.manage_row, "→", self.move_selected_right)
        self.remove_button = self.make_soft_button(self.manage_row, "Quitar", self.remove_selected)

        # Salida
        self.bottom_card = ctk.CTkFrame(
            self.center_wrap,
            fg_color=CARD_BG,
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=20
        )
        self.bottom_card.grid(row=3, column=0, sticky="ew")

        self.bottom_inner = ctk.CTkFrame(self.bottom_card, fg_color="transparent")
        self.bottom_inner.pack(fill="both", expand=True, padx=14, pady=14)

        self.output_title = ctk.CTkLabel(
            self.bottom_inner,
            text="Carpeta de salida",
            text_color=TEXT_PRIMARY,
            font=self.label_font
        )
        self.output_title.grid(row=0, column=0, sticky="w", columnspan=2, pady=(0, 5))

        self.output_button = ctk.CTkButton(
            self.bottom_inner,
            text="Seleccionar carpeta",
            command=self.select_output,
            fg_color=SOFT_BUTTON_BG,
            hover_color=SOFT_BUTTON_HOVER,
            text_color=SOFT_BUTTON_TEXT,
            corner_radius=12,
            height=32,
            border_width=1,
            border_color=SOFT_BUTTON_BORDER,
            font=self.button_font,
            width=150
        )
        self.output_button.grid(row=1, column=0, sticky="w")

        self.output_label = ctk.CTkLabel(
            self.bottom_inner,
            text="No seleccionada",
            text_color=TEXT_SECONDARY,
            font=self.small_font,
            wraplength=320,
            justify="left"
        )
        self.output_label.grid(row=1, column=1, sticky="w", padx=(10, 0))

        self.progress = ctk.CTkProgressBar(
            self.bottom_inner,
            progress_color=BUTTON_COLOR,
            fg_color=PROGRESS_BG,
            height=10
        )
        self.progress.set(0)
        self.progress.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 6))

        self.status_label = ctk.CTkLabel(
            self.bottom_inner,
            text="Listo",
            text_color=TEXT_SECONDARY,
            font=self.small_font
        )
        self.status_label.grid(row=3, column=0, sticky="w")

        self.convert_button = ctk.CTkButton(
            self.bottom_inner,
            text="Convertir a PDF",
            command=self.convert,
            fg_color=BUTTON_COLOR,
            hover_color=BUTTON_HOVER,
            text_color=BUTTON_TEXT,
            corner_radius=14,
            height=38,
            font=self.primary_button_font,
            width=180
        )
        self.convert_button.grid(row=3, column=1, sticky="e")

        self.result_row = ctk.CTkFrame(self.bottom_inner, fg_color="transparent")
        self.result_row.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        self.open_pdf_button = self.make_soft_button(self.result_row, "Abrir PDF", self.open_last_pdf)
        self.open_folder_button = self.make_soft_button(self.result_row, "Abrir carpeta", self.open_output_folder)

        self.apply_responsive_layout(self.winfo_width())
        self.update_manage_buttons()

    def make_field(self, parent, label, placeholder):
        wrap = ctk.CTkFrame(parent, fg_color="transparent")

        field_label = ctk.CTkLabel(
            wrap,
            text=label,
            text_color=TEXT_PRIMARY,
            font=self.label_font
        )
        field_label.pack(anchor="w", pady=(0, 4))

        entry = ctk.CTkEntry(
            wrap,
            placeholder_text=placeholder,
            fg_color=INPUT_BG,
            border_color=INPUT_BORDER,
            text_color=TEXT_PRIMARY,
            corner_radius=12,
            height=32,
            font=self.small_font
        )
        entry.pack(fill="x")

        return wrap, entry

    def make_soft_button(self, parent, text, command):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            fg_color=SOFT_BUTTON_BG,
            hover_color=SOFT_BUTTON_HOVER,
            text_color=SOFT_BUTTON_TEXT,
            corner_radius=12,
            height=32,
            border_width=1,
            border_color=SOFT_BUTTON_BORDER,
            font=self.button_font
        )

    # =========================
    # ESTADO GUARDADO
    # =========================
    def load_saved_values(self):
        study = self.app_state.get("last_study", "")
        output_dir = self.app_state.get("last_output_dir", "")
        window_geometry = self.app_state.get("last_geometry")

        if study:
            self.study_entry.insert(0, study)

        if output_dir and os.path.isdir(output_dir):
            self.output_dir = output_dir
            self.output_label.configure(text=self.shorten_path(output_dir))

        if isinstance(window_geometry, str):
            try:
                self.geometry(window_geometry)
            except Exception:
                pass

    def persist_state(self):
        data = {
            "last_study": self.study_entry.get().strip(),
            "last_output_dir": self.output_dir,
            "last_geometry": self.geometry()
        }
        save_app_state(data)

    # =========================
    # RESPONSIVE
    # =========================
    def on_resize(self, event):
        if event.widget != self:
            return

        if self.resize_after_id:
            self.after_cancel(self.resize_after_id)

        self.resize_after_id = self.after(40, self.handle_resize)

    def handle_resize(self):
        width = self.winfo_width()
        self.apply_responsive_layout(width)
        self.persist_state()

    def apply_responsive_layout(self, width: int):
        # Formulario
        self.patient_wrap.grid_forget()
        self.study_wrap.grid_forget()
        self.date_wrap.grid_forget()

        for i in range(3):
            self.form_inner.grid_columnconfigure(i, weight=0)

        if width >= 900:
            self.form_inner.grid_columnconfigure(0, weight=1)
            self.form_inner.grid_columnconfigure(1, weight=1)
            self.form_inner.grid_columnconfigure(2, weight=1)

            self.patient_wrap.grid(row=0, column=0, sticky="ew")
            self.study_wrap.grid(row=0, column=1, sticky="ew", padx=(8, 0))
            self.date_wrap.grid(row=0, column=2, sticky="ew", padx=(8, 0))

        elif width >= 720:
            self.form_inner.grid_columnconfigure(0, weight=1)
            self.form_inner.grid_columnconfigure(1, weight=1)

            self.patient_wrap.grid(row=0, column=0, sticky="ew")
            self.study_wrap.grid(row=0, column=1, sticky="ew", padx=(8, 0))
            self.date_wrap.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        else:
            self.form_inner.grid_columnconfigure(0, weight=1)
            self.patient_wrap.grid(row=0, column=0, sticky="ew")
            self.study_wrap.grid(row=1, column=0, sticky="ew", pady=(8, 0))
            self.date_wrap.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        # Botones principales
        for widget in self.actions_row.winfo_children():
            widget.grid_forget()

        if width >= 760:
            self.actions_row.grid_columnconfigure(0, weight=1)
            self.actions_row.grid_columnconfigure(1, weight=1)
            self.actions_row.grid_columnconfigure(2, weight=1)
            self.actions_row.grid_columnconfigure(3, weight=1)

            self.add_images_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
            self.add_pdf_button.grid(row=0, column=1, sticky="ew", padx=(0, 6))
            self.add_folder_button.grid(row=0, column=2, sticky="ew", padx=(0, 6))
            self.clear_button.grid(row=0, column=3, sticky="ew")
        else:
            self.actions_row.grid_columnconfigure(0, weight=1)
            self.actions_row.grid_columnconfigure(1, weight=1)

            self.add_images_button.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=(0, 6))
            self.add_pdf_button.grid(row=0, column=1, sticky="ew", pady=(0, 6))
            self.add_folder_button.grid(row=1, column=0, sticky="ew", padx=(0, 6))
            self.clear_button.grid(row=1, column=1, sticky="ew")

        # Botones de gestión
        for widget in self.manage_row.winfo_children():
            widget.grid_forget()

        self.manage_row.grid_columnconfigure(0, weight=1)
        self.manage_row.grid_columnconfigure(1, weight=1)
        self.manage_row.grid_columnconfigure(2, weight=1)

        self.move_left_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.move_right_button.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self.remove_button.grid(row=0, column=2, sticky="ew")

        # Bloque final
        self.output_button.grid_forget()
        self.output_label.grid_forget()
        self.status_label.grid_forget()
        self.convert_button.grid_forget()
        self.open_pdf_button.grid_forget()
        self.open_folder_button.grid_forget()

        if width >= 760:
            self.bottom_inner.grid_columnconfigure(1, weight=1)
            self.result_row.grid_columnconfigure(1, weight=1)

            self.output_button.grid(row=1, column=0, sticky="w")
            self.output_label.grid(row=1, column=1, sticky="w", padx=(10, 0))
            self.output_label.configure(wraplength=320)

            self.status_label.grid(row=3, column=0, sticky="w")
            self.convert_button.grid(row=3, column=1, sticky="e")

            self.open_pdf_button.grid(row=0, column=0, sticky="w", padx=(0, 6))
            self.open_folder_button.grid(row=0, column=1, sticky="w")
        else:
            self.bottom_inner.grid_columnconfigure(0, weight=1)
            self.result_row.grid_columnconfigure(0, weight=1)

            self.output_button.grid(row=1, column=0, sticky="ew")
            self.output_label.grid(row=2, column=0, sticky="w", pady=(8, 0))
            self.output_label.configure(wraplength=250)

            self.status_label.grid(row=3, column=0, sticky="w", pady=(8, 0))
            self.convert_button.grid(row=4, column=0, sticky="ew", pady=(8, 0))

            self.open_pdf_button.grid(row=0, column=0, sticky="ew", pady=(0, 6))
            self.open_folder_button.grid(row=1, column=0, sticky="ew")

    # =========================
    # ESTADO UI
    # =========================
    def set_status(self, text: str):
        self.status_label.configure(text=text)
        self.update_idletasks()

    def set_progress(self, value: float):
        self.progress.set(max(0.0, min(1.0, value)))
        self.update_idletasks()

    def shorten_path(self, folder: str) -> str:
        return folder if len(folder) < 42 else "..." + folder[-39:]

    def update_manage_buttons(self):
        has_selection = self.selected_index is not None and 0 <= self.selected_index < len(self.files)

        self.remove_button.configure(state="normal" if has_selection else "disabled")
        self.move_left_button.configure(
            state="normal" if has_selection and self.selected_index > 0 else "disabled"
        )
        self.move_right_button.configure(
            state="normal" if has_selection and self.selected_index < len(self.files) - 1 else "disabled"
        )
        self.open_pdf_button.configure(
            state="normal" if self.last_output_file and os.path.exists(self.last_output_file) else "disabled"
        )
        self.open_folder_button.configure(
            state="normal" if self.output_dir else "disabled"
        )

    # =========================
    # MINIATURAS
    # =========================
    def select_thumbnail(self, index: int):
        self.selected_index = index
        self.refresh_previews()
        self.update_manage_buttons()
        if 0 <= index < len(self.files):
            self.set_status(f"Seleccionado: {safe_filename_only(self.files[index])}")

    def refresh_previews(self):
        for widget in self.preview_scroll.winfo_children():
            widget.destroy()

        self.preview_images = []

        if self.files:
            if self.drop_text.winfo_ismapped():
                self.drop_text.pack_forget()
        else:
            self.selected_index = None
            if not self.drop_text.winfo_ismapped():
                self.drop_text.pack(pady=(0, 6))
            self.update_manage_buttons()
            return

        if self.selected_index is None or self.selected_index >= len(self.files):
            self.selected_index = 0

        max_items = min(len(self.files), 8)

        for i in range(max_items):
            path = self.files[i]
            selected = i == self.selected_index

            card = ctk.CTkFrame(
                self.preview_scroll,
                fg_color=SELECTED_CARD_BG if selected else "#FFFFFF",
                border_width=2 if selected else 1,
                border_color=SELECTED_CARD_BORDER if selected else CARD_BORDER,
                corner_radius=12,
                width=82,
                height=82
            )
            card.pack(side="left", padx=5, pady=2)
            card.pack_propagate(False)

            if is_image(path):
                try:
                    with Image.open(path) as img:
                        img = ImageOps.exif_transpose(img)
                        img.thumbnail((34, 34))
                        tk_img = ImageTk.PhotoImage(img.copy())
                        self.preview_images.append(tk_img)

                    label_img = ctk.CTkLabel(card, text="", image=tk_img)
                    label_img.pack(pady=(6, 2))
                    label_img.bind("<Button-1>", lambda e, idx=i: self.select_thumbnail(idx))
                except Exception:
                    icon = ctk.CTkLabel(
                        card,
                        text="IMG",
                        text_color=TEXT_PRIMARY,
                        font=self.icon_font
                    )
                    icon.pack(pady=(12, 4))
                    icon.bind("<Button-1>", lambda e, idx=i: self.select_thumbnail(idx))
            elif is_pdf(path):
                icon = ctk.CTkLabel(
                    card,
                    text="PDF",
                    text_color=TEXT_PRIMARY,
                    font=self.icon_font
                )
                icon.pack(pady=(12, 4))
                icon.bind("<Button-1>", lambda e, idx=i: self.select_thumbnail(idx))

            name = os.path.basename(path)
            if len(name) > 10:
                name = name[:7] + "..."

            label_name = ctk.CTkLabel(
                card,
                text=name,
                text_color=TEXT_SECONDARY,
                font=self.thumb_font,
                wraplength=64
            )
            label_name.pack(padx=4, pady=(0, 4))
            label_name.bind("<Button-1>", lambda e, idx=i: self.select_thumbnail(idx))
            card.bind("<Button-1>", lambda e, idx=i: self.select_thumbnail(idx))

        remaining = len(self.files) - max_items
        if remaining > 0:
            more = ctk.CTkLabel(
                self.preview_scroll,
                text=f"+{remaining}",
                text_color=TEXT_SECONDARY,
                font=self.small_font
            )
            more.pack(side="left", padx=8, pady=26)

        self.update_manage_buttons()

    def refresh_file_state(self):
        total = len(self.files)
        self.files_counter.configure(
            text="1 archivo seleccionado" if total == 1 else f"{total} archivos seleccionados"
        )
        self.refresh_previews()

    # =========================
    # CARGA
    # =========================
    def add_images(self):
        files = filedialog.askopenfilenames(title="Selecciona imágenes")
        added = 0
        for f in files:
            if is_image(f) and f not in self.files:
                self.files.append(f)
                added += 1

        if added and self.selected_index is None:
            self.selected_index = 0

        self.refresh_file_state()
        self.set_status(f"Se agregaron {added} imagen(es)")

    def add_pdf(self):
        files = filedialog.askopenfilenames(title="Selecciona archivos PDF")
        added = 0
        for f in files:
            if is_pdf(f) and f not in self.files:
                self.files.append(f)
                added += 1

        if added and self.selected_index is None:
            self.selected_index = 0

        self.refresh_file_state()
        self.set_status(f"Se agregaron {added} PDF(s)")

    def add_folder(self):
        folder = filedialog.askdirectory(title="Selecciona una carpeta")
        if not folder:
            return

        added = 0
        for root, _, files in os.walk(folder):
            for file in files:
                path = os.path.join(root, file)
                if (is_image(path) or is_pdf(path)) and path not in self.files:
                    self.files.append(path)
                    added += 1

        self.files.sort(key=lambda x: x.lower())

        if added and self.selected_index is None:
            self.selected_index = 0

        self.refresh_file_state()
        self.set_status(f"Se agregaron {added} archivo(s) desde carpeta")

    def clear_list(self):
        self.files = []
        self.selected_index = None
        self.refresh_file_state()
        self.set_status("Lista limpiada")

    def remove_selected(self):
        if self.selected_index is None or not (0 <= self.selected_index < len(self.files)):
            return

        removed = safe_filename_only(self.files[self.selected_index])
        del self.files[self.selected_index]

        if not self.files:
            self.selected_index = None
        elif self.selected_index >= len(self.files):
            self.selected_index = len(self.files) - 1

        self.refresh_file_state()
        self.set_status(f"Se quitó: {removed}")

    def move_selected_left(self):
        if self.selected_index is None or self.selected_index <= 0:
            return

        idx = self.selected_index
        self.files[idx - 1], self.files[idx] = self.files[idx], self.files[idx - 1]
        self.selected_index -= 1
        self.refresh_file_state()
        self.set_status("Archivo movido a la izquierda")

    def move_selected_right(self):
        if self.selected_index is None or self.selected_index >= len(self.files) - 1:
            return

        idx = self.selected_index
        self.files[idx + 1], self.files[idx] = self.files[idx], self.files[idx + 1]
        self.selected_index += 1
        self.refresh_file_state()
        self.set_status("Archivo movido a la derecha")

    def select_output(self):
        folder = filedialog.askdirectory(title="Selecciona carpeta de salida")
        if folder:
            self.output_dir = folder
            self.output_label.configure(text=self.shorten_path(folder))
            self.set_status("Carpeta de salida seleccionada")
            self.persist_state()
            self.update_manage_buttons()

    # =========================
    # RESULTADO
    # =========================
    def open_last_pdf(self):
        if self.last_output_file and os.path.exists(self.last_output_file):
            open_path(self.last_output_file)

    def open_output_folder(self):
        if self.output_dir and os.path.isdir(self.output_dir):
            open_folder(self.output_dir)

    # =========================
    # CONVERSIÓN
    # =========================
    def convert(self):
        if not self.files:
            messagebox.showerror("Error", "No hay archivos cargados.")
            return

        if not self.output_dir:
            messagebox.showerror("Error", "Selecciona una carpeta de salida.")
            return

        patient = self.patient_entry.get().strip()
        study = self.study_entry.get().strip()
        date_text = self.date_entry.get().strip()

        if not patient or not study:
            messagebox.showerror("Error", "Completa nombre del paciente y tipo de estudio.")
            return

        safe_date = parse_date(date_text)
        if safe_date == "INVALID_DATE":
            messagebox.showerror("Error", "Fecha inválida. Usa YYYY-MM-DD o DD/MM/YYYY.")
            return

        safe_patient = sanitize_filename(patient)
        safe_study = sanitize_filename(study)

        if safe_date:
            output_name = f"{safe_patient}_{safe_study}_{safe_date}.pdf"
        else:
            output_name = f"{safe_patient}_{safe_study}.pdf"

        output_path = os.path.join(self.output_dir, output_name)
        output_path = ensure_unique_path(output_path)

        writer = PdfWriter()
        failed_files = []
        total = len(self.files)

        self.set_status("Construyendo PDF...")
        self.set_progress(0)
        self.last_output_file = None
        self.update_manage_buttons()
        self.persist_state()

        try:
            for idx, f in enumerate(self.files, start=1):
                try:
                    if is_image(f):
                        temp_pdf = image_to_pdf_bytes(f)
                        reader = PdfReader(temp_pdf)
                        for page in reader.pages:
                            writer.add_page(page)

                    elif is_pdf(f):
                        reader = PdfReader(f)
                        for page in reader.pages:
                            writer.add_page(page)

                    else:
                        failed_files.append(f"{safe_filename_only(f)}: formato no compatible")

                except Exception as e:
                    failed_files.append(f"{safe_filename_only(f)}: {str(e)}")

                self.set_progress(idx / total)
                self.set_status(f"Procesando {idx} de {total}...")

            if len(writer.pages) == 0:
                self.set_progress(0)
                messagebox.showerror("Error", "No se pudo crear el PDF porque no hubo archivos válidos.")
                self.set_status("No se generó el PDF")
                return

            with open(output_path, "wb") as out:
                writer.write(out)

            self.last_output_file = output_path
            self.set_progress(1)
            self.set_status("PDF creado correctamente")
            self.update_manage_buttons()

            msg = f"PDF creado correctamente.\n\n{output_path}"
            if failed_files:
                preview_errors = "\n".join(failed_files[:6])
                msg += f"\n\nSe omitieron {len(failed_files)} archivo(s):\n{preview_errors}"
                if len(failed_files) > 6:
                    msg += "\n..."

            messagebox.showinfo("Listo", msg)

        except Exception as e:
            self.set_progress(0)
            self.set_status("Ocurrió un error")
            messagebox.showerror("Error", f"No se pudo crear el PDF.\n\nDetalle: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
