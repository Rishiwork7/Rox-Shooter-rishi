import tkinter as tk
import customtkinter as ctk
import multiprocessing
import threading
import queue
import asyncio
import os
import shutil
import string
import re
import random
import tempfile
from datetime import datetime
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.util import Inches
from PIL import Image as PILImage
import xlsxwriter
import subprocess
import sys
import uuid

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ── Light Theme Color Palette ──
COLOR = {
    "bg":           "#F5F7FA",
    "surface":      "#FFFFFF",
    "surface_alt":  "#EEF1F6",
    "border":       "#DDE2EA",
    "text":         "#1E293B",
    "text_sec":     "#64748B",
    "primary":      "#4F6DF5",
    "primary_hov":  "#3B5BDB",
    "success":      "#22C55E",
    "success_hov":  "#16A34A",
    "danger":       "#EF4444",
    "danger_hov":   "#DC2626",
    "warn":         "#F59E0B",
    "warn_hov":     "#D97706",
    "muted":        "#94A3B8",
    "muted_hov":    "#64748B",
    "accent":       "#8B5CF6",
    "accent_hov":   "#7C3AED",
    "info":         "#0EA5E9",
    "info_hov":     "#0284C7",
    "input_bg":     "#F1F5F9",
    "input_border": "#CBD5E1",
    "log_bg":       "#F8FAFC",
    "log_text":     "#334155",
    "tag_card":     "#F1F5F9",
}

class TextParser:
    def __init__(self):
        pass

    def generate_random_string(self, length, weight_letters=0.75):
        """Generates alphanumeric string weighted towards uppercase letters."""
        num_letters = int(length * weight_letters)
        num_digits = length - num_letters
        
        letters = ''.join(random.choices(string.ascii_uppercase, k=num_letters))
        digits = ''.join(random.choices(string.digits, k=num_digits))
        
        result = list(letters + digits)
        random.shuffle(result)
        return ''.join(result)

    def generate_word(self, length):
        """Generates random uppercase letters."""
        return ''.join(random.choices(string.ascii_uppercase, k=length))

    def generate_invoice(self, length):
        """Generates numbers ending with 1-2 uppercase letters."""
        num_suffix = random.randint(1, 2)
        num_digits = length - num_suffix
        digits = ''.join(random.choices(string.digits, k=max(1, num_digits)))
        letters = ''.join(random.choices(string.ascii_uppercase, k=num_suffix))
        return digits + letters

    def generate_number(self, length):
        """Generates pure random numbers."""
        return ''.join(random.choices(string.digits, k=length))

    def parse(self, text, recipient_email, tfn_number):
        if not text:
            return ""

        def replacer(match):
            tag = match.group(1).lower()
            val = match.group(2)
            length = int(val) if val else 6

            if tag == "random":
                return self.generate_random_string(length)
            elif tag == "word":
                return self.generate_word(length)
            elif tag == "invoice_no":
                return self.generate_invoice(length)
            elif tag in ["rand", "number"]:
                return self.generate_number(length)
            elif tag == "mail":
                return recipient_email
            elif tag == "date":
                return datetime.now().strftime("%d-%m-%Y")
            elif tag == "tfn":
                return tfn_number
            return match.group(0) # Fallback

        # Match only supported tags to avoid swallowing following underscores or text
        supported_tags = ["random", "word", "invoice_no", "rand", "number", "mail", "email", "date", "tfn"]
        pattern = r"\$(" + "|".join(supported_tags) + r")(?:[\(\[])?(\d+)?(?:[\)\]])?"
        text = re.sub(pattern, replacer, text, flags=re.IGNORECASE)

        return text

class Converter:
    def __init__(self, log_callback):
        self.log = log_callback
        # Use system temp directory instead of relative path
        # This ensures compatibility with PyInstaller .exe on Windows
        self.temp_dir = os.path.join(tempfile.gettempdir(), "rox_shooter_attachments")
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        os.makedirs(self.temp_dir, exist_ok=True)

    async def html_to_pdf(self, html_content, filename="attachment.pdf"):
        """Convert HTML to PDF using system-installed Chrome (no Chromium download needed)."""
        path = os.path.join(self.temp_dir, filename)
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(
                    headless=True,
                    channel="chrome"  # Uses system Chrome — no download needed
                )
                page = await browser.new_page()
                await page.set_content(html_content, wait_until="networkidle")
                await page.pdf(path=path, format="A4")
                await browser.close()
            return path
        except Exception as e:
            self.log(f"PDF Conversion Error: {e}")
            return None

    async def html_to_image(self, html_content, filename="attachment.png"):
        """Convert HTML to Image using system-installed Chrome (no Chromium download needed)."""
        path = os.path.join(self.temp_dir, filename)
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(
                    headless=True,
                    channel="chrome"  # Uses system Chrome — no download needed
                )
                page = await browser.new_page()
                await page.set_content(html_content, wait_until="networkidle")
                await page.screenshot(path=path, full_page=True)
                await browser.close()
            return path
        except Exception as e:
            self.log(f"Image Conversion Error: {e}")
            return None


    async def html_to_image_pptx(self, html_content, filename="attachment_img.pptx"):
        """HTML -> Image -> PPTX Slide."""
        path = os.path.join(self.temp_dir, filename)
        unique_id = uuid.uuid4().hex[:8]
        img_name = f"temp_pptx_{unique_id}.png"
        img_path = os.path.join(self.temp_dir, img_name)
        try:
            await self.html_to_image(html_content, img_name)
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(img_path, Inches(0.5), Inches(0.5), width=Inches(9))
            prs.save(path)
            if os.path.exists(img_path): os.remove(img_path)
            return path
        except Exception as e:
            self.log(f"Image then PPTX Error: {e}")
            return None

    async def html_to_image_pdf(self, html_content, filename="attachment_img.pdf"):
        """HTML -> Image -> PDF."""
        path = os.path.join(self.temp_dir, filename)
        unique_id = uuid.uuid4().hex[:8]
        img_name = f"temp_pdf_{unique_id}.png"
        img_path = os.path.join(self.temp_dir, img_name)
        try:
            await self.html_to_image(html_content, img_name)
            img = PILImage.open(img_path)
            pdf_img = img.convert('RGB')
            pdf_img.save(path)
            if os.path.exists(img_path): os.remove(img_path)
            return path
        except Exception as e:
            self.log(f"Image then PDF Error: {e}")
            return None

    async def html_to_image_xls(self, html_content, filename="attachment_img.xlsx"):
        """HTML -> Image -> XLS."""
        path = os.path.join(self.temp_dir, filename)
        unique_id = uuid.uuid4().hex[:8]
        img_name = f"temp_xls_{unique_id}.png"
        img_path = os.path.join(self.temp_dir, img_name)
        try:
            await self.html_to_image(html_content, img_name)
            workbook = xlsxwriter.Workbook(path)
            worksheet = workbook.add_worksheet()
            worksheet.insert_image('B2', img_path, {'x_scale': 0.5, 'y_scale': 0.5})
            workbook.close()
            if os.path.exists(img_path): os.remove(img_path)
            return path
        except Exception as e:
            self.log(f"Image then XLS Error: {e}")
            return None

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Window Configuration ---
        self.title("Rox-Shooter v1.1")
        self.geometry("1340x920")
        self.configure(fg_color=COLOR["bg"])

        # UI Scaling/State
        self.after(0, lambda: self.state('zoomed'))
        
        # --- Internal State & Threading ---
        self.log_queue = queue.Queue()
        self.loop = None
        self.thread = None
        self.is_shooting = False
        self.contexts = {} # {id: {"context": context, "page": page, "frame": frame}}
        self.parser = TextParser()
        self.converter = Converter(self.log)
        
        # --- Background Loop ---
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(target=self.run_background_loop, daemon=True)
        self.thread.start()
        
        # --- UI Construction ---
        self.setup_grid()
        self.create_header_bar()
        self.create_launch_controls()
        self.create_main_container()
        self.create_activity_log()
        
        # Start logging check
        self.after(100, self.process_logs)
        
        # Handle Close
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.log("System Initialized. Ready for Phase 1.")

    def setup_grid(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # Header Bar
        self.grid_rowconfigure(1, weight=0)  # Window Management
        self.grid_rowconfigure(2, weight=1)  # Main Content
        self.grid_rowconfigure(3, weight=0)  # Activity Logs

    def create_header_bar(self):
        """Branded header with Gmail-style logo text and Rox avatar."""
        header = ctk.CTkFrame(self, fg_color=COLOR["surface"], corner_radius=0, height=56)
        header.grid(row=0, column=0, sticky="nsew")
        header.grid_columnconfigure(1, weight=1)

        # Logo text
        logo_frame = ctk.CTkFrame(header, fg_color="transparent")
        logo_frame.pack(side="left", padx=24, pady=12)
        ctk.CTkLabel(logo_frame, text="R", font=("Inter", 22, "bold"), text_color="#4285F4").pack(side="left")
        ctk.CTkLabel(logo_frame, text="OX-", font=("Inter", 22, "bold"), text_color=COLOR["text"]).pack(side="left")
        ctk.CTkLabel(logo_frame, text="SHOOTER", font=("Inter", 22, "bold"), text_color=COLOR["primary"]).pack(side="left")
        ctk.CTkLabel(logo_frame, text="  —  v1.1", font=("Inter", 13), text_color=COLOR["text_sec"]).pack(side="left", padx=(4, 0))

        # Rox avatar circle (top-right)
        avatar_frame = ctk.CTkFrame(header, fg_color="transparent")
        avatar_frame.pack(side="right", padx=24, pady=10)
        avatar = ctk.CTkButton(
            avatar_frame, text="S", width=36, height=36, corner_radius=18,
            fg_color=COLOR["primary"], hover_color=COLOR["primary_hov"],
            text_color="#FFFFFF", font=("Inter", 16, "bold"), state="disabled"
        )
        avatar.pack(side="right")
        ctk.CTkLabel(avatar_frame, text="Shooter", font=("Inter", 12, "bold"), text_color=COLOR["text"]).pack(side="right", padx=(0, 8))

    def create_launch_controls(self):
        # ── Window Management Bar ──
        self.launch_frame = ctk.CTkFrame(self, fg_color=COLOR["surface"], corner_radius=12, border_width=1, border_color=COLOR["border"])
        self.launch_frame.grid(row=1, column=0, padx=24, pady=(12, 8), sticky="nsew")

        # Left group: session controls
        left_group = ctk.CTkFrame(self.launch_frame, fg_color="transparent")
        left_group.pack(side="left", padx=(20, 10), pady=14)

        ctk.CTkLabel(left_group, text="🖥  Window Management", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(left_group, text="Sessions:", font=("Inter", 12), text_color=COLOR["text_sec"]).pack(side="left", padx=(0, 6))

        self.entry_num_windows = ctk.CTkEntry(left_group, width=56, height=34, corner_radius=8,
                                               fg_color=COLOR["input_bg"], border_color=COLOR["input_border"],
                                               text_color=COLOR["text"], placeholder_text="1", font=("Inter", 13))
        self.entry_num_windows.insert(0, "1")
        self.entry_num_windows.pack(side="left", padx=(0, 10))

        self.btn_launch = ctk.CTkButton(left_group, text="Initialize Session(s)", height=34, corner_radius=8,
                                         fg_color=COLOR["primary"], hover_color=COLOR["primary_hov"],
                                         text_color="#FFFFFF", font=("Inter", 12, "bold"), command=self.on_launch_windows)
        self.btn_launch.pack(side="left", padx=(0, 8))

        self.btn_terminate = ctk.CTkButton(left_group, text="Force Close All", height=34, corner_radius=8,
                                            fg_color=COLOR["danger"], hover_color=COLOR["danger_hov"],
                                            text_color="#FFFFFF", font=("Inter", 12, "bold"), command=self.on_terminate_all)
        self.btn_terminate.pack(side="left", padx=(0, 8))

        # Right group: utility actions
        right_group = ctk.CTkFrame(self.launch_frame, fg_color="transparent")
        right_group.pack(side="right", padx=(10, 20), pady=14)

        self.btn_reset = ctk.CTkButton(right_group, text="Full System Reset", width=130, height=34, corner_radius=8,
                                        fg_color=COLOR["surface_alt"], hover_color=COLOR["border"],
                                        text_color=COLOR["text"], border_width=1, border_color=COLOR["border"],
                                        font=("Inter", 12, "bold"), command=self.on_reset)
        self.btn_reset.pack(side="right", padx=(8, 0))

        self.btn_clear_log = ctk.CTkButton(right_group, text="Erase Logs", width=100, height=34, corner_radius=8,
                                            fg_color=COLOR["surface_alt"], hover_color=COLOR["border"],
                                            text_color=COLOR["text"], border_width=1, border_color=COLOR["border"],
                                            font=("Inter", 12, "bold"), command=self.on_clear_log)
        self.btn_clear_log.pack(side="right")

    def create_main_container(self):
        # Main layout: Left (Active Sessions), Right (Settings Tabs)
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=2, column=0, padx=24, pady=8, sticky="nsew")

        self.main_container.grid_columnconfigure(0, weight=1)  # Active Sessions
        self.main_container.grid_columnconfigure(1, weight=2)  # Tabs
        self.main_container.grid_rowconfigure(0, weight=1)

        # ── Active Sessions Panel ──
        self.active_windows_frame = ctk.CTkScrollableFrame(
            self.main_container, label_text="⚡  Active Sessions",
            label_font=("Inter", 14, "bold"),
            fg_color=COLOR["surface"], corner_radius=12,
            border_width=1, border_color=COLOR["border"],
            label_fg_color=COLOR["surface"]
        )
        self.active_windows_frame.grid(row=0, column=0, padx=(0, 12), sticky="nsew")

        # ── Settings Tabs ──
        self.tabview = ctk.CTkTabview(
            self.main_container, corner_radius=12,
            fg_color=COLOR["surface"], border_width=1, border_color=COLOR["border"],
            segmented_button_fg_color=COLOR["surface"], # White bar creates visible gaps
            segmented_button_selected_color=COLOR["surface"],
            segmented_button_unselected_color=COLOR["surface_alt"],
            segmented_button_selected_hover_color=COLOR["surface"],
            segmented_button_unselected_hover_color="#CBD5E1"
        )
        # Fix tab colors and spacing appearance
        try:
            self.tabview._segmented_button.configure(
                font=("Inter", 12, "bold"),
                corner_radius=6, # Slightly rounded box
                text_color="#000000",
                unselected_color=COLOR["surface_alt"],
                selected_color=COLOR["surface"]
            )
        except:
            pass
        self.tabview.grid(row=0, column=1, sticky="nsew")

        self.tab_data = self.tabview.add("Target Contacts")
        self.tab_subject_body = self.tabview.add("Subject & Body")
        self.tab_content = self.tabview.add("Content Options")
        self.tab_settings = self.tabview.add("Advanced Params")
        self.tab_shooter = self.tabview.add("Execute Send")
        self.tab_tags = self.tabview.add("System Tags")

        self.setup_tabs()

    def setup_tabs(self):
        # --- Tab 1: Data ---
        self.data_frame = ctk.CTkFrame(self.tab_data, fg_color="transparent")
        self.data_frame.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(self.data_frame, text="Target Contacts", font=("Inter", 14, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 8))

        btn_row = ctk.CTkFrame(self.data_frame, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 10))

        self.btn_load_data = ctk.CTkButton(
            btn_row, text="📂  Load CSV / Excel / TXT", height=36, corner_radius=8,
            fg_color=COLOR["info"], hover_color=COLOR["info_hov"],
            text_color="#FFFFFF", font=("Inter", 12, "bold"),
            command=self.handle_load_file
        )
        self.btn_load_data.pack(side="left")

        self.text_emails = ctk.CTkTextbox(
            self.data_frame, height=300, font=("Consolas", 12),
            fg_color=COLOR["input_bg"], text_color=COLOR["text"],
            border_width=1, border_color=COLOR["input_border"], corner_radius=8
        )
        self.text_emails.pack(fill="both", expand=True, pady=(0, 10))
        self.text_emails.bind("<KeyRelease>", self.update_email_counter)

        self.lbl_email_counter = ctk.CTkLabel(
            self.data_frame, text="Total Emails: 0",
            font=("Inter", 14, "bold"), text_color=COLOR["primary"]
        )
        self.lbl_email_counter.pack(pady=(0, 4))

        # --- Tab 2: Subject & Body ---
        sb_frame = ctk.CTkFrame(self.tab_subject_body, fg_color="transparent")
        sb_frame.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(sb_frame, text="Shooter Subject", font=("Inter", 14, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 6))
        self.entry_subject = ctk.CTkEntry(
            sb_frame, placeholder_text="e.g. $invoice_no7 — Important Notice", font=("Inter", 13), height=38,
            corner_radius=8, fg_color=COLOR["input_bg"], border_color=COLOR["input_border"],
            text_color=COLOR["text"]
        )
        self.entry_subject.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(sb_frame, text="Main Message Body", font=("Inter", 14, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 6))
        self.text_body = ctk.CTkTextbox(
            sb_frame, height=300, font=("Inter", 13),
            fg_color=COLOR["input_bg"], text_color=COLOR["text"],
            border_width=1, border_color=COLOR["input_border"], corner_radius=8
        )
        self.text_body.pack(fill="both", expand=True)
        
        # --- Tab 3: Content ---
        self.content_frame = ctk.CTkFrame(self.tab_content, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=24, pady=20)

        # Row 1: Conversion and Filename Mode
        row1 = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 12))

        # Left: Conversion
        conv_col = ctk.CTkFrame(row1, fg_color="transparent")
        conv_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(conv_col, text="Conversion Type", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 4))
        self.conv_options = [
            "None", "HTML to raw PDF", "HTML to Image", "HTML to Image then PDF",
            "HTML to Image then PPTX", "HTML to Image then XLS"
        ]
        self.dropdown_conversion = ctk.CTkComboBox(
            conv_col, values=self.conv_options, width=260, height=36, corner_radius=8,
            fg_color=COLOR["input_bg"], border_color=COLOR["input_border"],
            text_color=COLOR["text"], button_color=COLOR["primary"],
            button_hover_color=COLOR["primary_hov"], dropdown_fg_color=COLOR["surface"],
            dropdown_text_color=COLOR["text"], dropdown_hover_color=COLOR["surface_alt"]
        )
        self.dropdown_conversion.set("HTML to raw PDF")
        self.dropdown_conversion.pack(anchor="w", pady=(0, 4))

        # Right: Filename Mode
        file_col = ctk.CTkFrame(row1, fg_color="transparent")
        file_col.pack(side="left", fill="x", expand=True, padx=(24, 0))
        ctk.CTkLabel(file_col, text="Filename Mode", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 4))
        self.file_mode = ctk.CTkSegmentedButton(
            file_col, values=["Random", "Custom"], command=self.toggle_filename_entry,
            selected_color=COLOR["primary"], selected_hover_color=COLOR["primary_hov"],
            unselected_color=COLOR["surface_alt"], unselected_hover_color=COLOR["border"],
            text_color=COLOR["text"], text_color_disabled=COLOR["muted"],
            font=("Inter", 12, "bold"), corner_radius=8
        )
        self.file_mode.set("Random")
        self.file_mode.pack(anchor="w", pady=(0, 4))

        # Row 2: Custom Filename Entry
        self.filename_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.filename_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(self.filename_frame, text="Custom Filename (Supports Tags):", font=("Inter", 12), text_color=COLOR["text_sec"]).pack(anchor="w")
        self.entry_filename = ctk.CTkEntry(
            self.filename_frame, placeholder_text="e.g. Invoice_$invoice_no7", width=400, height=36,
            corner_radius=8, fg_color=COLOR["input_bg"], border_color=COLOR["input_border"],
            text_color=COLOR["text"]
        )
        self.entry_filename.pack(anchor="w", pady=5)
        self.filename_frame.pack_forget()  # Hidden by default

        # Row 3: HTML Content
        ctk.CTkLabel(self.content_frame, text="HTML Content", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(4, 6))
        self.text_html = ctk.CTkTextbox(
            self.content_frame, font=("Consolas", 12),
            fg_color=COLOR["input_bg"], text_color=COLOR["text"],
            border_width=1, border_color=COLOR["input_border"], corner_radius=8
        )
        self.text_html.pack(fill="both", expand=True, pady=(0, 12))

        preview_row = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        preview_row.pack(fill="x", pady=(0, 4))

        self.btn_preview = ctk.CTkButton(
            preview_row, text="👁  Preview Selected Format", height=36, corner_radius=8,
            fg_color=COLOR["accent"], hover_color=COLOR["accent_hov"],
            text_color="#FFFFFF", font=("Inter", 12, "bold"), command=self.on_preview_attachment
        )
        self.btn_preview.pack(side="left", padx=(0, 10))

        self.btn_preview_all = ctk.CTkButton(
            preview_row, text="📁  Preview All Formats", height=36, corner_radius=8,
            fg_color=COLOR["info"], hover_color=COLOR["info_hov"],
            text_color="#FFFFFF", font=("Inter", 12, "bold"), command=self.on_preview_all
        )
        self.btn_preview_all.pack(side="left")
        
        # --- Tab 4: Settings ---
        settings_frame = ctk.CTkFrame(self.tab_settings, fg_color="transparent")
        settings_frame.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(settings_frame, text="Advanced Parameters", font=("Inter", 16, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 16))

        # Delay card
        delay_card = ctk.CTkFrame(settings_frame, fg_color=COLOR["surface_alt"], corner_radius=10, border_width=1, border_color=COLOR["border"])
        delay_card.pack(fill="x", pady=(0, 12))
        delay_inner = ctk.CTkFrame(delay_card, fg_color="transparent")
        delay_inner.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(delay_inner, text="Delay Between Emails (seconds)", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(side="left")
        self.entry_delay = ctk.CTkEntry(
            delay_inner, width=80, height=36, corner_radius=8,
            fg_color=COLOR["input_bg"], border_color=COLOR["input_border"], text_color=COLOR["text"], font=("Inter", 13)
        )
        self.entry_delay.insert(0, "2")
        self.entry_delay.pack(side="right")
        ctk.CTkLabel(settings_frame, text="⚠  Increasing delay helps avoid Gmail spam detection.", font=("Inter", 11), text_color=COLOR["muted"]).pack(anchor="w", pady=(0, 16))

        # TFN card
        tfn_card = ctk.CTkFrame(settings_frame, fg_color=COLOR["surface_alt"], corner_radius=10, border_width=1, border_color=COLOR["border"])
        tfn_card.pack(fill="x", pady=(0, 8))
        tfn_inner = ctk.CTkFrame(tfn_card, fg_color="transparent")
        tfn_inner.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(tfn_inner, text="Toll-Free Number (TFN)", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 6))
        self.entry_tfn = ctk.CTkEntry(
            tfn_inner, placeholder_text="e.g. +1-800-XXX-XXXX", width=300, height=36,
            corner_radius=8, fg_color=COLOR["input_bg"], border_color=COLOR["input_border"], text_color=COLOR["text"]
        )
        self.entry_tfn.pack(anchor="w")

        # --- Tab 5: Shooter ---
        shooter_frame = ctk.CTkFrame(self.tab_shooter, fg_color="transparent")
        shooter_frame.pack(fill="both", expand=True)

        shooter_card = ctk.CTkFrame(shooter_frame, fg_color=COLOR["surface_alt"], corner_radius=16, border_width=1, border_color=COLOR["border"])
        shooter_card.pack(expand=True, padx=60, pady=60)

        ctk.CTkLabel(shooter_card, text="🚀", font=("Inter", 48)).pack(pady=(30, 10))
        ctk.CTkLabel(shooter_card, text="Ready to send?", font=("Inter", 14), text_color=COLOR["text_sec"]).pack(pady=(0, 16))

        self.btn_start_shooting = ctk.CTkButton(
            shooter_card, text="START SENDING", height=56, width=280,
            corner_radius=28, font=("Inter", 16, "bold"),
            fg_color=COLOR["primary"], hover_color=COLOR["primary_hov"],
            text_color="#FFFFFF", command=self.on_start_shooting
        )
        self.btn_start_shooting.pack(pady=(0, 30))

        # --- Tab 6: Tags ---
        self.setup_tags_tab()

    def setup_tags_tab(self):
        container = ctk.CTkScrollableFrame(self.tab_tags, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(container, text="Available System Tags", font=("Inter", 16, "bold"), text_color=COLOR["text"]).pack(anchor="w", pady=(0, 4))
        ctk.CTkLabel(container, text="Use these tags in Subject, Body, HTML Content, or Custom Filenames.\nThey are replaced with dynamic data during shooting.",
                      font=("Inter", 12), justify="left", wraplength=550, text_color=COLOR["text_sec"]).pack(anchor="w", pady=(0, 16))

        tags_info = [
            ("$random(length)", "Generates a mix of uppercase letters and numbers.\nExample: $random(8) → A1B2C3D4", COLOR["info"]),
            ("$word(length)", "Generates random uppercase letters only.\nExample: $word(6) → KIMOXQ", COLOR["accent"]),
            ("$invoice_no(length)", "Generates numbers ending with 1-2 letters.\nExample: $invoice_no(7) → 12345AB", COLOR["success"]),
            ("$rand(length)", "Generates pure random numbers.\nExample: $rand(6) → 982374", COLOR["warn"]),
            ("$mail", "Replaces with the recipient's email address.", "#EC4899"),
            ("$date", "Replaces with current date (DD-MM-YYYY).", COLOR["muted"]),
            ("$tfn", "Replaces with the Toll-Free Number from Settings.", COLOR["primary"]),
        ]

        for tag, desc, color in tags_info:
            frame = ctk.CTkFrame(container, fg_color=COLOR["tag_card"], corner_radius=10, border_width=1, border_color=COLOR["border"])
            frame.pack(fill="x", pady=4)

            lbl_tag = ctk.CTkLabel(frame, text=tag, font=("Consolas", 13, "bold"), text_color=color, width=160)
            lbl_tag.pack(side="left", padx=16, pady=12)

            lbl_desc = ctk.CTkLabel(frame, text=desc, font=("Inter", 12), justify="left", text_color=COLOR["text"])
            lbl_desc.pack(side="left", padx=10, pady=12, fill="x", expand=True)

        ctk.CTkLabel(container, text="Note: (length) is optional and defaults to 6 if not specified.", font=("Inter", 11, "italic"), text_color=COLOR["muted"]).pack(anchor="w", pady=(10, 0))


    def create_activity_log(self):
        # ── Activity Logs ──
        self.log_frame = ctk.CTkFrame(self, fg_color=COLOR["surface"], corner_radius=12, border_width=1, border_color=COLOR["border"], height=160)
        self.log_frame.grid(row=3, column=0, padx=24, pady=(8, 20), sticky="nsew")

        log_header = ctk.CTkFrame(self.log_frame, fg_color="transparent")
        log_header.pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkLabel(log_header, text="✨  Activity Logs", font=("Inter", 13, "bold"), text_color=COLOR["text"]).pack(side="left")

        self.log_textbox = ctk.CTkTextbox(
            self.log_frame, font=("Consolas", 11.5), state="disabled",
            fg_color=COLOR["log_bg"], text_color=COLOR["log_text"],
            corner_radius=8, border_width=1, border_color=COLOR["border"]
        )
        self.log_textbox.pack(fill="both", expand=True, padx=16, pady=(8, 14))

    # --- UI Logic & Handlers ---
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {message}\n")

    def process_logs(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", msg)
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
        self.after(100, self.process_logs)

    def update_email_counter(self, event=None):
        content = self.text_emails.get("1.0", "end-1c").strip()
        emails = [e for e in content.split("\n") if e.strip()]
        self.lbl_email_counter.configure(text=f"Total Emails: {len(emails)}")

    def on_launch_windows(self):
        try:
            count_str = self.entry_num_windows.get()
            count = int(count_str) if count_str.isdigit() else 1
            self.run_coro(self.launch_and_arrange(count))
        except Exception as e:
            self.log(f"Launch Error: {e}")

    async def launch_and_arrange(self, count):
        """Launch all sessions sequentially, then tile them split-screen."""
        for i in range(1, count + 1):
            if i not in self.contexts:
                await self.launch_browser_task(i)

        # Silently arrange all active windows in split-screen
        await self.arrange_windows_split()

    async def arrange_windows_split(self):
        """Tile all active browser windows in a grid across the screen using CDP."""
        import math
        active_ids = sorted(self.contexts.keys())
        n = len(active_ids)
        if n == 0:
            return

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        cols = math.ceil(math.sqrt(n))
        rows = math.ceil(n / cols)
        win_w = screen_w // cols
        win_h = screen_h // rows

        for idx, wid in enumerate(active_ids):
            x = (idx % cols) * win_w
            y = (idx // cols) * win_h
            try:
                page = self.contexts[wid]["page"]
                cdp = await page.context.new_cdp_session(page)
                win_info = await cdp.send("Browser.getWindowForTarget")
                await cdp.send("Browser.setWindowBounds", {
                    "windowId": win_info["windowId"],
                    "bounds": {"left": x, "top": y, "width": win_w, "height": win_h, "windowState": "normal"}
                })
            except:
                pass

    async def launch_browser_task(self, window_id):
        try:
            self.log(f"Initializing Session {window_id}...")
            profile_path = os.path.abspath(f"profiles/profile_{window_id}")
            os.makedirs(profile_path, exist_ok=True)
            
            # Force remove lock file if it exists (prevents ProcessSingleton errors)
            lock_file = os.path.join(profile_path, "SingletonLock")
            if os.path.exists(lock_file):
                try:
                    os.remove(lock_file)
                except:
                    pass
            
            # Start Playwright for this context
            pw = await async_playwright().start()
            context = await pw.chromium.launch_persistent_context(
                user_data_dir=profile_path,
                channel="chrome",
                headless=False,
                args=["--incognito", "--no-sandbox", "--disable-blink-features=AutomationControlled"]
            )
            
            # Initial page: reuse default tab if exists, else create new
            if context.pages:
                page = context.pages[0]
            else:
                page = await context.new_page()
            
            # Store in app state
            self.contexts[window_id] = {
                "pw": pw,
                "context": context,
                "page": page
            }
            
            # Update UI (thread-safe via after)
            self.after(0, lambda: self.add_window_row(window_id))
            self.log(f"Session {window_id} ready.")
            
        except Exception as e:
            self.log(f"Session {window_id} failed: {e}")

    def add_window_row(self, window_id):
        # Create a styled row card in the scrollable frame
        row = ctk.CTkFrame(self.active_windows_frame, fg_color=COLOR["surface_alt"], corner_radius=8, border_width=1, border_color=COLOR["border"])
        row.pack(fill="x", padx=8, pady=4)

        ctk.CTkLabel(row, text=f"⚡ Active Session {window_id}", width=140, font=("Inter", 12, "bold"), text_color=COLOR["text"]).pack(side="left", padx=12, pady=10)

        btn_gmail = ctk.CTkButton(
            row, text="Open Browser", width=110, height=30, corner_radius=6,
            fg_color=COLOR["info"], hover_color=COLOR["info_hov"], text_color="#FFFFFF",
            font=("Inter", 11, "bold"),
            command=lambda: self.run_coro(self.open_gmail_task(window_id))
        )
        btn_gmail.pack(side="left", padx=4)

        btn_close = ctk.CTkButton(
            row, text="End", width=60, height=30, corner_radius=6,
            fg_color=COLOR["surface"], hover_color=COLOR["border"], text_color=COLOR["danger"],
            border_width=1, border_color=COLOR["danger"],
            font=("Inter", 11, "bold"),
            command=lambda: self.run_coro(self.close_window_task(window_id))
        )
        btn_close.pack(side="right", padx=12)

        # Save the row frame for deletion later
        self.contexts[window_id]["ui_row"] = row

    async def open_gmail_task(self, window_id):
        if window_id in self.contexts:
            page = self.contexts[window_id]["page"]
            self.log(f"Navigating Window {window_id} to Gmail...")
            await page.goto("https://mail.google.com")
            await page.bring_to_front()

    async def close_window_task(self, window_id):
        if window_id in self.contexts:
            # Get reference but don't pop yet to keep tracking
            data = self.contexts.get(window_id)
            if not data: return
            
            self.log(f"Closing Window {window_id}...")
            try:
                if "page" in data:
                    await data["page"].close()
                if "context" in data:
                    await data["context"].close()
                if "pw" in data:
                    await data["pw"].stop()
                self.log(f"Window {window_id} closed successfully.")
            except Exception as e:
                self.log(f"Error closing Window {window_id}: {e}")
            finally:
                # Remove from tracking after attempt
                if window_id in self.contexts:
                    self.contexts.pop(window_id)
            
            # Update UI
            if "ui_row" in data:
                try:
                    self.after(0, data["ui_row"].destroy)
                except:
                    pass

    def on_terminate_all(self):
        self.log("Terminating all windows...")
        ids = list(self.contexts.keys())
        for window_id in ids:
            self.run_coro(self.close_window_task(window_id))

    def on_clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def on_reset(self):
        self.entry_num_windows.delete(0, "end")
        self.entry_num_windows.insert(0, "1")
        self.text_emails.delete("1.0", "end")
        self.entry_subject.delete(0, "end")
        self.text_body.delete("1.0", "end")
        self.text_html.delete("1.0", "end")
        self.entry_tfn.delete(0, "end")
        self.update_email_counter()
        self.log("Application state reset.")

    def handle_load_file(self):
        from tkinter import filedialog
        import pandas as pd
        import re
        
        file_path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xlsx *.xls *.txt")])
        if not file_path:
            return
            
        try:
            self.log(f"Loading data from {os.path.basename(file_path)}...")
            if file_path.endswith(".csv"):
                df = pd.read_csv(file_path)
                text = df.to_string()
            elif file_path.endswith((".xlsx", ".xls")):
                df = pd.read_excel(file_path)
                text = df.to_string()
            else:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()
            
            # Extract emails using regex
            emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
            unique_emails = list(dict.fromkeys(emails)) # Remove duplicates while preserving order
            
            self.text_emails.delete("1.0", "end")
            self.text_emails.insert("1.0", "\n".join(unique_emails))
            self.update_email_counter()
            self.log(f"Successfully imported {len(unique_emails)} unique emails.")
            
        except Exception as e:
            self.log(f"Error loading file: {e}")

    def toggle_filename_entry(self, value):
        if value == "Custom":
            self.filename_frame.pack(after=self.file_mode.master, fill="x", pady=(0, 10))
        else:
            self.filename_frame.pack_forget()

    def on_preview_attachment(self):
        self.run_coro(self.preview_attachment_task())

    async def preview_attachment_task(self):
        html_template = self.text_html.get("1.0", "end-1c")
        if not html_template.strip():
            self.log("Error: No HTML content to preview.")
            return

        conversion_type = self.dropdown_conversion.get()
        if conversion_type == "None":
            self.log("Error: Conversion type is 'None'. Choose a type to preview.")
            return

        self.log(f"Generating preview for {conversion_type}...")
        
        # Parse with dummy data
        tfn = self.entry_tfn.get() or "1-800-PREVIEW"
        parsed_html = self.parser.parse(html_template, "preview@example.com", tfn)
        
        # Extension
        ext = ".pdf"
        if "Image" in conversion_type and "PDF" not in conversion_type and "PPTX" not in conversion_type and "XLS" not in conversion_type:
            ext = ".png"
        elif "PPTX" in conversion_type:
            ext = ".pptx"
        elif "XLS" in conversion_type:
            ext = ".xlsx"
            
        filename = f"preview_test{ext}"
        path = None
        
        try:
            if conversion_type == "HTML to raw PDF":
                path = await self.converter.html_to_pdf(parsed_html, filename)
            elif conversion_type == "HTML to Image":
                path = await self.converter.html_to_image(parsed_html, filename)
            elif conversion_type == "HTML to Image then PDF":
                path = await self.converter.html_to_image_pdf(parsed_html, filename)

            elif conversion_type == "HTML to Image then PPTX":
                path = await self.converter.html_to_image_pptx(parsed_html, filename)
            elif conversion_type == "HTML to Image then XLS":
                path = await self.converter.html_to_image_xls(parsed_html, filename)
            
            if path and os.path.exists(path):
                self.log(f"Preview generated: {path}")
                # Open the file based on OS
                if sys.platform == "darwin":
                    subprocess.run(["open", path])
                elif sys.platform == "win32":
                    os.startfile(path)
                else:
                    try:
                        subprocess.run(["xdg-open", path])
                    except:
                        self.log(f"File saved at: {path} (Could not open automatically)")
            else:
                self.log("Error: Failed to generate preview file.")
        except Exception as e:
            self.log(f"Preview Error: {e}")

    def on_preview_all(self):
        self.run_coro(self.preview_all_task())

    async def preview_all_task(self):
        html_template = self.text_html.get("1.0", "end-1c")
        if not html_template.strip():
            self.log("Error: No HTML content to preview.")
            return

        self.log("Generating previews for ALL formats. This may take a moment...")
        
        # Parse with dummy data
        tfn = self.entry_tfn.get() or "1-800-PREVIEW"
        parsed_html = self.parser.parse(html_template, "preview@example.com", tfn)
        
        tasks = [
            ("PDF", self.converter.html_to_pdf(parsed_html, "preview_all_pdf.pdf")),
            ("Image", self.converter.html_to_image(parsed_html, "preview_all_img.png")),
            ("Image then PDF", self.converter.html_to_image_pdf(parsed_html, "preview_all_img_pdf.pdf")),

            ("Image then PPTX", self.converter.html_to_image_pptx(parsed_html, "preview_all_img_pptx.pptx")),
            ("Image then XLS", self.converter.html_to_image_xls(parsed_html, "preview_all_img_xls.xlsx"))
        ]
        
        for name, coro in tasks:
            try:
                self.log(f"Generating {name}...")
                path = await coro
                if path and os.path.exists(path):
                    self.log(f"{name} generated: {path}")
                    # Open the file
                    if sys.platform == "darwin":
                        subprocess.run(["open", path])
                    elif sys.platform == "win32":
                        os.startfile(path)
                    else:
                        subprocess.run(["xdg-open", path], capture_output=True)
                else:
                    self.log(f"Error: Failed to generate {name}.")
            except Exception as e:
                self.log(f"Error generating {name}: {e}")
        
        self.log("All previews generated.")

    def on_start_shooting(self):
        if self.is_shooting:
            self.is_shooting = False
            self.btn_start_shooting.configure(text="START SENDING", fg_color=COLOR["primary"], hover_color=COLOR["primary_hov"])
            self.log("Sending PAUSED. Waiting for current tasks to finish...")
        else:
            self.is_shooting = True
            self.btn_start_shooting.configure(text="STOP SENDING", fg_color=COLOR["danger"], hover_color=COLOR["danger_hov"])
            self.run_coro(self.shooter_engine_task())

    async def shooter_engine_task(self):
        self.log("Initializing Shooter Engine...")
        
        # 1. Get Data
        raw_emails = self.text_emails.get("1.0", "end-1c").strip().split("\n")
        emails = [e.strip() for e in raw_emails if e.strip()]
        
        if not emails:
            self.log("Error: No recipient emails found.")
            self.on_start_shooting()
            return

        active_ids = list(self.contexts.keys())
        if not active_ids:
            self.log("Error: No active windows found. Launch windows first.")
            self.on_start_shooting()
            return

        subject_template = self.entry_subject.get()
        body_template = self.text_body.get("1.0", "end-1c")
        html_template = self.text_html.get("1.0", "end-1c")
        tfn = self.entry_tfn.get()
        conversion_type = self.dropdown_conversion.get()
        file_mode = self.file_mode.get()
        custom_filename_template = self.entry_filename.get()
        delay_sec = float(self.entry_delay.get() or 2)

        self.log(f"Starting shoot for {len(emails)} emails using {len(active_ids)} windows...")

        # 2. Round-Robin Loop
        for i, recipient in enumerate(emails):
            if not self.is_shooting:
                break
            
            window_id = active_ids[i % len(active_ids)]

            # 3. Parse Content
            parsed_subject = self.parser.parse(subject_template, recipient, tfn)
            parsed_body = self.parser.parse(body_template, recipient, tfn)
            parsed_html = self.parser.parse(html_template, recipient, tfn)

            # 4. Handle Conversion
            attachment_path = None
            if parsed_html.strip() and conversion_type != "None":
                # Determine Filename
                if file_mode == "Custom" and custom_filename_template:
                    # For filename, user wants only the username part of the email (remove everything after @)
                    recipient_user = recipient.split("@")[0] if "@" in recipient else recipient
                    base_name = self.parser.parse(custom_filename_template, recipient_user, tfn)
                else:
                    base_name = f"attachment_{self.parser.generate_random_string(6)}"
                
                # Add correct extension
                ext = ".pdf"
                if "Image" in conversion_type and "PDF" not in conversion_type and "PPTX" not in conversion_type and "XLS" not in conversion_type:
                    ext = ".png"
                elif "PPTX" in conversion_type:
                    ext = ".pptx"
                elif "XLS" in conversion_type:
                    ext = ".xlsx"
                
                final_filename = f"{base_name}{ext}"

                if conversion_type == "HTML to raw PDF":
                    attachment_path = await self.converter.html_to_pdf(parsed_html, final_filename)
                elif conversion_type == "HTML to Image":
                    attachment_path = await self.converter.html_to_image(parsed_html, final_filename)
                elif conversion_type == "HTML to Image then PDF":
                    attachment_path = await self.converter.html_to_image_pdf(parsed_html, final_filename)

                elif conversion_type == "HTML to Image then PPTX":
                    attachment_path = await self.converter.html_to_image_pptx(parsed_html, final_filename)
                elif conversion_type == "HTML to Image then XLS":
                    attachment_path = await self.converter.html_to_image_xls(parsed_html, final_filename)

            # 5. Execute Automation
            success = await self.automate_gmail_send(window_id, recipient, parsed_subject, parsed_body, attachment_path)
            
            if success:
                self.log(f"Successfully sent to {recipient}")
            else:
                self.log(f"Failed to send to {recipient}")

            if i < len(emails) - 1:
                await asyncio.sleep(delay_sec)

        self.log("Shooting session COMPLETED.")
        if self.is_shooting:
            self.on_start_shooting() # Toggle UI back

    async def automate_gmail_send(self, window_id, recipient, subject, body, attachment_path):
        if window_id not in self.contexts:
            return False
        
        page = self.contexts[window_id]["page"]
        
        try:
            # Check if we are on Gmail
            if "mail.google.com" not in page.url:
                await page.goto("https://mail.google.com")
                await asyncio.sleep(3)

            # Click Compose
            await page.get_by_role("button", name="Compose").click(timeout=15000)
            await asyncio.sleep(random.uniform(1.5, 3.0))

            # Fill To
            self.log(f"[S{window_id}] Sending to {recipient}...")
            to_input = page.locator('div[aria-label="To"] input, input[aria-label="To"]').first
            await to_input.click()
            await to_input.fill(recipient)
            await page.keyboard.press("Enter")
            await asyncio.sleep(0.5)

            # Fill Subject
            subject_input = page.locator('input[name="subjectbox"], input[placeholder="Subject"]').first
            await subject_input.fill(subject)
            await asyncio.sleep(0.5)

            # Fill Body
            body_input = page.locator('div[role="textbox"][aria-label="Message Body"], div.editable[contenteditable="true"]').first
            await body_input.click()
            await body_input.fill(body)
            await asyncio.sleep(0.5)

            # Upload Attachment
            if attachment_path and os.path.exists(attachment_path):
                
                upload_success = False
                try:
                    # Method 1: Using expect_file_chooser (More robust for modern Gmail)
                    async with page.expect_file_chooser() as fc_info:
                        # Click the "Attach files" button (paperclip)
                        # We try multiple selectors for the attach button
                        attach_btn = page.locator('div[command="Files"][aria-label*="Attach files"], div[aria-label*="Attach files"]').first
                        await attach_btn.click(timeout=5000)
                    file_chooser = await fc_info.value
                    await file_chooser.set_files(attachment_path)
                    upload_success = True
                except Exception as e:
                    try:
                        # Method 2: Direct set_input_files on hidden input
                        file_input = page.locator('input[type="file"][name="Filedata"], input[type="file"]').last
                        await file_input.set_input_files(attachment_path)
                        upload_success = True
                    except Exception as e2:
                        self.log(f"[W{window_id}] Method 2 (Direct Input) failed: {e2}")

                if upload_success:
                    await asyncio.sleep(0.5) # Minimal buffer for Gmail UI

            # Send
            try:
                # Use a more robust selector for the Send button
                # Gmail's send button usually contains "Send" in aria-label
                send_btn = page.locator('div[role="button"][aria-label*="Send"], div[role="button"]:has-text("Send")').first
                await send_btn.click(timeout=7000)
            except Exception as e:
                # Fallback: Gmail shortcut for Send is Ctrl+Enter (or Cmd+Enter on Mac)
                # We try both to be sure
                await page.keyboard.press("Control+Enter")
                await asyncio.sleep(0.5)
                # If still not sent, try Meta+Enter for Mac
                await page.keyboard.press("Meta+Enter")
            
            # Optional: Wait for "Message sent" confirmation
                pass
            except:
                pass
                
            return True

        except Exception as e:
            self.log(f"[W{window_id}] Automation Error: {e}")
            # Try to recover by closing compose if stuck
            try:
                await page.keyboard.press("Escape")
            except:
                pass
            return False

    # --- Threading & Async Bridges ---
    def run_background_loop(self):
        """Runs the asyncio event loop in a separate thread."""
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def run_coro(self, coro):
        """Safely submits a coroutine to the background loop."""
        return asyncio.run_coroutine_threadsafe(coro, self.loop)

    def on_closing(self):
        """Handles application shutdown."""
        self.is_blasting = False
        self.log("Shutting down... Terminating browsers.")
        
        # 1. Start termination for all contexts
        ids = list(self.contexts.keys())
        for window_id in ids:
            self.run_coro(self.close_window_task(window_id))
        
        # 2. Wait longer for browsers to finish cleanup before killing the loop
        # Increased to 3 seconds for better reliability with real Chrome
        self.after(3000, self.final_cleanup)

    def final_cleanup(self):
        try:
            self.loop.call_soon_threadsafe(self.loop.stop)
            temp_attachments = os.path.join(tempfile.gettempdir(), "gmail_mailer_attachments")
            if os.path.exists(temp_attachments):
                shutil.rmtree(temp_attachments)
        except:
            pass
        self.destroy()

if __name__ == "__main__":
    # Essential for PyInstaller standalone executables
    multiprocessing.freeze_support()
    
    # No browser download needed — we use system Chrome (channel="chrome")
    # Chrome must be installed on the target machine (it almost always is)
        
    app = App()
    app.mainloop()
