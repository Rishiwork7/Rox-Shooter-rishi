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
from datetime import datetime
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.util import Inches
from PIL import Image as PILImage
import xlsxwriter
import subprocess
import sys

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

def install_browsers():
    """Ensure playwright browsers are installed without recursive calls."""
    if getattr(sys, 'frozen', False):
        # In a frozen app, we don't use sys.executable to install
        # Browsers should be installed during setup or bundled
        return
    try:
        import subprocess
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], capture_output=True)
    except:
        pass

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
        supported_tags = ["random", "word", "invoice_no", "rand", "number", "mail", "date", "tfn"]
        pattern = r"\$(" + "|".join(supported_tags) + r")(?:[\(\[])?(\d+)?(?:[\)\]])?"
        text = re.sub(pattern, replacer, text, flags=re.IGNORECASE)

        return text

class Converter:
    def __init__(self, log_callback):
        self.log = log_callback
        self.temp_dir = os.path.abspath("temp_attachments")
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        os.makedirs(self.temp_dir, exist_ok=True)

    async def html_to_pdf(self, html_content, filename="attachment.pdf"):
        path = os.path.join(self.temp_dir, filename)
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True)
                page = await browser.new_page()
                await page.set_content(html_content)
                await page.pdf(path=path, format="A4")
                await browser.close()
            return path
        except Exception as e:
            self.log(f"PDF Conversion Error: {e}")
            return None

    async def html_to_image(self, html_content, filename="attachment.png"):
        path = os.path.join(self.temp_dir, filename)
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True)
                page = await browser.new_page()
                await page.set_content(html_content)
                await page.screenshot(path=path, full_page=True)
                await browser.close()
            return path
        except Exception as e:
            self.log(f"Image Conversion Error: {e}")
            return None

    async def html_to_pptx(self, html_content, filename="attachment.pptx"):
        """Direct text conversion: Adds cleaned HTML text to a PPTX slide."""
        path = os.path.join(self.temp_dir, filename)
        try:
            # Better text cleaning
            clean_text = re.sub('<br\s*/?>', '\n', html_content, flags=re.IGNORECASE)
            clean_text = re.sub('<[^<]+?>', '', clean_text).strip()
            
            prs = Presentation()
            slide_layout = prs.slide_layouts[1] # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            
            title = slide.shapes.title
            title.text = "Business Proposal" # Professional default
            
            content = slide.placeholders[1]
            content.text = clean_text[:2000] # Increased limit
            
            prs.save(path)
            # Small delay to ensure OS file system flushes
            await asyncio.sleep(0.5)
            return path
        except Exception as e:
            self.log(f"Direct PPTX Error: {e}")
            return None

    async def html_to_image_pptx(self, html_content, filename="attachment_img.pptx"):
        """HTML -> Image -> PPTX Slide."""
        path = os.path.join(self.temp_dir, filename)
        img_path = os.path.join(self.temp_dir, "temp_pptx_slide.png")
        try:
            await self.html_to_image(html_content, "temp_pptx_slide.png")
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
        img_path = os.path.join(self.temp_dir, "temp_pdf_wrap.png")
        try:
            await self.html_to_image(html_content, "temp_pdf_wrap.png")
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
        img_path = os.path.join(self.temp_dir, "temp_xls_insert.png")
        try:
            await self.html_to_image(html_content, "temp_xls_insert.png")
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
        self.title("Elite Gmail Blaster | Automation Suite")
        self.geometry("1200x850")
        
        # UI Scaling/State
        self.after(0, lambda: self.state('zoomed'))
        
        # --- Internal State & Threading ---
        self.log_queue = queue.Queue()
        self.loop = None
        self.thread = None
        self.is_blasting = False
        self.contexts = {} # {id: {"context": context, "page": page, "frame": frame}}
        self.parser = TextParser()
        self.converter = Converter(self.log)
        
        # --- Background Loop ---
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(target=self.run_background_loop, daemon=True)
        self.thread.start()
        
        # --- UI Construction ---
        self.setup_grid()
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
        self.grid_rowconfigure(0, weight=0) # Launch Controls
        self.grid_rowconfigure(1, weight=1) # Main Content (Active Windows + Tabs)
        self.grid_rowconfigure(2, weight=0) # Activity Log

    def create_launch_controls(self):
        # Step 2: "Launch Controls" frame
        self.launch_frame = ctk.CTkFrame(self, height=80)
        self.launch_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="nsew")
        
        ctk.CTkLabel(self.launch_frame, text="Number of Windows:", font=("Inter", 13, "bold")).pack(side="left", padx=(20, 5), pady=20)
        
        self.entry_num_windows = ctk.CTkEntry(self.launch_frame, width=60, placeholder_text="1")
        self.entry_num_windows.insert(0, "1")
        self.entry_num_windows.pack(side="left", padx=5, pady=20)
        
        self.btn_launch = ctk.CTkButton(self.launch_frame, text="Launch Windows", fg_color="#28a745", hover_color="#218838", font=("Inter", 13, "bold"), command=self.on_launch_windows)
        self.btn_launch.pack(side="left", padx=10, pady=20)
        
        self.btn_terminate = ctk.CTkButton(self.launch_frame, text="Terminate All", fg_color="#dc3545", hover_color="#c82333", font=("Inter", 13, "bold"), command=self.on_terminate_all)
        self.btn_terminate.pack(side="left", padx=5, pady=20)
        
        self.btn_clear_log = ctk.CTkButton(self.launch_frame, text="Clear Log", fg_color="#6c757d", hover_color="#5a6268", font=("Inter", 13, "bold"), command=self.on_clear_log)
        self.btn_clear_log.pack(side="left", padx=5, pady=20)
        
        self.btn_reset = ctk.CTkButton(self.launch_frame, text="Reset", fg_color="#17a2b8", hover_color="#138496", font=("Inter", 13, "bold"), command=self.on_reset)
        self.btn_reset.pack(side="left", padx=5, pady=20)

    def create_main_container(self):
        # Main layout: Left (Active Windows), Right (Configuration Tabs)
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        
        self.main_container.grid_columnconfigure(0, weight=1) # Active Windows
        self.main_container.grid_columnconfigure(1, weight=2) # Tabs
        self.main_container.grid_rowconfigure(0, weight=1)
        
        # Step 3: "Active Windows" frame
        self.active_windows_frame = ctk.CTkScrollableFrame(self.main_container, label_text="Active Windows Manager", label_font=("Inter", 14, "bold"))
        self.active_windows_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        
        # Step 4: CTkTabview for Configuration
        self.tabview = ctk.CTkTabview(self.main_container)
        self.tabview.grid(row=0, column=1, sticky="nsew")
        
        self.tab_data = self.tabview.add("Data")
        self.tab_subject_body = self.tabview.add("Subject & Body")
        self.tab_content = self.tabview.add("Content")
        self.tab_settings = self.tabview.add("Settings")
        self.tab_blaster = self.tabview.add("Blaster")
        
        self.setup_tabs()

    def setup_tabs(self):
        # --- Tab 1: Data ---
        self.data_frame = ctk.CTkFrame(self.tab_data, fg_color="transparent")
        self.data_frame.pack(fill="both", expand=True, padx=20, pady=15)
        
        ctk.CTkLabel(self.data_frame, text="Recipient Emails:", font=("Inter", 13, "bold")).pack(anchor="w", pady=(0, 5))
        
        btn_row = ctk.CTkFrame(self.data_frame, fg_color="transparent")
        btn_row.pack(fill="x", pady=5)
        
        self.btn_load_data = ctk.CTkButton(btn_row, text="📁 Load CSV / Excel / TXT", fg_color="#17a2b8", hover_color="#138496", 
                                            command=self.handle_load_file)
        self.btn_load_data.pack(side="left", padx=0)
        
        self.text_emails = ctk.CTkTextbox(self.data_frame, height=300, font=("Consolas", 12))
        self.text_emails.pack(fill="both", expand=True, pady=10)
        self.text_emails.bind("<KeyRelease>", self.update_email_counter)
        
        self.lbl_email_counter = ctk.CTkLabel(self.data_frame, text="Total Emails: 0", font=("Inter", 14, "bold"), text_color="#17a2b8")
        self.lbl_email_counter.pack(pady=5)
        
        # --- Tab 2: Subject & Body ---
        ctk.CTkLabel(self.tab_subject_body, text="Email Subject:", font=("Inter", 13, "bold")).pack(anchor="w", padx=20, pady=(15, 5))
        self.entry_subject = ctk.CTkEntry(self.tab_subject_body, placeholder_text="Enter subject here...", font=("Inter", 13))
        self.entry_subject.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(self.tab_subject_body, text="Email Body:", font=("Inter", 13, "bold")).pack(anchor="w", padx=20, pady=(15, 5))
        self.text_body = ctk.CTkTextbox(self.tab_subject_body, height=300, font=("Inter", 13))
        self.text_body.pack(fill="both", expand=True, padx=20, pady=5)
        
        # --- Tab 3: Content ---
        self.content_frame = ctk.CTkFrame(self.tab_content, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=15)

        # Row 1: Conversion and Filename Mode
        row1 = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 10))

        # Left: Conversion
        conv_col = ctk.CTkFrame(row1, fg_color="transparent")
        conv_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(conv_col, text="Conversion Type:", font=("Inter", 13, "bold")).pack(anchor="w")
        self.conv_options = [
            "None", "HTML to raw PDF", "HTML to Image", "HTML to Image then PDF", 
            "HTML to PPTX (Direct)", "HTML to Image then PPTX", "HTML to Image then XLS"
        ]
        self.dropdown_conversion = ctk.CTkComboBox(conv_col, values=self.conv_options, width=250)
        self.dropdown_conversion.set("HTML to raw PDF")
        self.dropdown_conversion.pack(anchor="w", pady=5)

        # Right: Filename Mode
        file_col = ctk.CTkFrame(row1, fg_color="transparent")
        file_col.pack(side="left", fill="x", expand=True, padx=(20, 0))
        ctk.CTkLabel(file_col, text="Filename Mode:", font=("Inter", 13, "bold")).pack(anchor="w")
        self.file_mode = ctk.CTkSegmentedButton(file_col, values=["Random", "Custom"], command=self.toggle_filename_entry)
        self.file_mode.set("Random")
        self.file_mode.pack(anchor="w", pady=5)

        # Row 2: Custom Filename Entry
        self.filename_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.filename_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(self.filename_frame, text="Custom Filename (Supports Tags):", font=("Inter", 12)).pack(anchor="w")
        self.entry_filename = ctk.CTkEntry(self.filename_frame, placeholder_text="e.g. Invoice_$invoice_no7", width=400)
        self.entry_filename.pack(anchor="w", pady=5)
        self.filename_frame.pack_forget() # Hidden by default

        # Row 3: HTML Content
        ctk.CTkLabel(self.content_frame, text="HTML Content:", font=("Inter", 13, "bold")).pack(anchor="w", pady=(5, 5))
        self.text_html = ctk.CTkTextbox(self.content_frame, font=("Consolas", 12))
        self.text_html.pack(fill="both", expand=True, pady=(0, 10))
        
        # --- Tab 4: Settings ---
        ctk.CTkLabel(self.tab_settings, text="Configuration Settings", font=("Inter", 15, "bold")).pack(pady=15)
        
        delay_frame = ctk.CTkFrame(self.tab_settings, fg_color="transparent")
        delay_frame.pack(fill="x", padx=40, pady=10)
        
        ctk.CTkLabel(delay_frame, text="Delay Between Emails (seconds):", font=("Inter", 13)).pack(side="left")
        self.entry_delay = ctk.CTkEntry(delay_frame, width=100)
        self.entry_delay.insert(0, "2")
        self.entry_delay.pack(side="left", padx=10)
        
        ctk.CTkLabel(self.tab_settings, text="* Increasing delay helps avoid Gmail spam detection.", font=("Inter", 11), text_color="gray").pack(pady=5)
        
        ctk.CTkLabel(self.tab_settings, text="Toll-Free Number (TFN):", font=("Inter", 13)).pack(anchor="w", padx=40, pady=(10, 0))
        self.entry_tfn = ctk.CTkEntry(self.tab_settings, placeholder_text="e.g. +1-800-XXX-XXXX", width=300)
        self.entry_tfn.pack(anchor="w", padx=40, pady=5)
        
        # --- Tab 5: Blaster ---
        self.btn_start_blasting = ctk.CTkButton(self.tab_blaster, text="START BLASTING", height=100, font=("Inter", 24, "bold"), 
                                                fg_color="#fd7e14", hover_color="#e8590c", command=self.on_start_blasting)
        self.btn_start_blasting.pack(expand=True, padx=50, pady=50)

    def create_activity_log(self):
        # Step 5: Activity Log
        self.log_frame = ctk.CTkFrame(self, height=150)
        self.log_frame.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="nsew")
        
        self.log_textbox = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), state="disabled", text_color="#00ff00")
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=10)

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
            for i in range(1, count + 1):
                if i not in self.contexts:
                    self.run_coro(self.launch_browser_task(i))
                else:
                    self.log(f"Window {i} is already active.")
        except Exception as e:
            self.log(f"Launch Error: {e}")

    async def launch_browser_task(self, window_id):
        try:
            self.log(f"Launching Window {window_id} with persistent profile...")
            profile_path = os.path.abspath(f"profiles/profile_{window_id}")
            os.makedirs(profile_path, exist_ok=True)
            
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
            self.log(f"Window {window_id} Ready.")
            
        except Exception as e:
            self.log(f"Error launching Window {window_id}: {e}")

    def add_window_row(self, window_id):
        # Create a row in the scrollable frame
        row = ctk.CTkFrame(self.active_windows_frame)
        row.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(row, text=f"Profile {window_id}", width=100, font=("Inter", 12, "bold")).pack(side="left", padx=10)
        
        btn_gmail = ctk.CTkButton(row, text="Open Gmail", width=100, fg_color="#17a2b8", 
                                  command=lambda: self.run_coro(self.open_gmail_task(window_id)))
        btn_gmail.pack(side="left", padx=5)
        
        btn_close = ctk.CTkButton(row, text="Close", width=60, fg_color="#6c757d", 
                                  command=lambda: self.run_coro(self.close_window_task(window_id)))
        btn_close.pack(side="right", padx=10)
        
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

    def on_start_blasting(self):
        if self.is_blasting:
            self.is_blasting = False
            self.btn_start_blasting.configure(text="START BLASTING", fg_color="#fd7e14")
            self.log("Blasting PAUSED. Waiting for current tasks to finish...")
        else:
            self.is_blasting = True
            self.btn_start_blasting.configure(text="STOP BLASTING", fg_color="#dc3545")
            self.run_coro(self.blaster_engine_task())

    async def blaster_engine_task(self):
        self.log("Initializing Blaster Engine...")
        
        # 1. Get Data
        raw_emails = self.text_emails.get("1.0", "end-1c").strip().split("\n")
        emails = [e.strip() for e in raw_emails if e.strip()]
        
        if not emails:
            self.log("Error: No recipient emails found.")
            self.on_start_blasting()
            return

        active_ids = list(self.contexts.keys())
        if not active_ids:
            self.log("Error: No active windows found. Launch windows first.")
            self.on_start_blasting()
            return

        subject_template = self.entry_subject.get()
        body_template = self.text_body.get("1.0", "end-1c")
        html_template = self.text_html.get("1.0", "end-1c")
        tfn = self.entry_tfn.get()
        conversion_type = self.dropdown_conversion.get()
        file_mode = self.file_mode.get()
        custom_filename_template = self.entry_filename.get()
        delay_sec = float(self.entry_delay.get() or 2)

        self.log(f"Starting blast for {len(emails)} emails using {len(active_ids)} windows...")

        # 2. Round-Robin Loop
        for i, recipient in enumerate(emails):
            if not self.is_blasting:
                break
            
            window_id = active_ids[i % len(active_ids)]
            self.log(f"Processing {recipient} via Window {window_id}...")

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
                elif conversion_type == "HTML to PPTX (Direct)":
                    attachment_path = await self.converter.html_to_pptx(parsed_html, final_filename)
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

            # 6. User-defined delay between emails
            if i < len(emails) - 1:
                self.log(f"Waiting {delay_sec} seconds before next email...")
                await asyncio.sleep(delay_sec)

        self.log("Blasting session COMPLETED.")
        if self.is_blasting:
            self.on_start_blasting() # Toggle UI back

    async def automate_gmail_send(self, window_id, recipient, subject, body, attachment_path):
        if window_id not in self.contexts:
            return False
        
        page = self.contexts[window_id]["page"]
        
        try:
            # Check if we are on Gmail
            if "mail.google.com" not in page.url:
                await page.goto("https://mail.google.com")
                await asyncio.sleep(3)

            # Step 3: STRICT LOCATOR RULES
            # Click Compose
            self.log(f"[W{window_id}] Clicking Compose...")
            await page.get_by_role("button", name="Compose").click(timeout=15000)
            await asyncio.sleep(random.uniform(1.5, 3.0))

            # Fill To
            self.log(f"[W{window_id}] Filling recipient...")
            to_input = page.locator('div[aria-label="To"] input, input[aria-label="To"]').first
            await to_input.click()
            await to_input.fill(recipient)
            await page.keyboard.press("Enter")
            await asyncio.sleep(0.5)

            # Fill Subject
            self.log(f"[W{window_id}] Filling subject...")
            subject_input = page.locator('input[name="subjectbox"], input[placeholder="Subject"]').first
            await subject_input.fill(subject)
            await asyncio.sleep(0.5)

            # Fill Body
            self.log(f"[W{window_id}] Filling body...")
            body_input = page.locator('div[role="textbox"][aria-label="Message Body"], div.editable[contenteditable="true"]').first
            await body_input.click()
            await body_input.fill(body)
            await asyncio.sleep(0.5)

            # Upload Attachment
            if attachment_path and os.path.exists(attachment_path):
                self.log(f"[W{window_id}] Uploading attachment: {os.path.basename(attachment_path)}...")
                # More robust way to find the file input in Gmail's compose
                file_input = page.locator('input[type="file"][name="Filedata"], input[type="file"]').last
                await file_input.set_input_files(attachment_path)
                
                # Wait for attachment to appear in UI (chip)
                try:
                    await page.wait_for_selector('div[role="link"][aria-label*="Attachment"]', timeout=10000)
                    self.log(f"[W{window_id}] Attachment uploaded successfully.")
                except:
                    self.log(f"[W{window_id}] Warning: Attachment chip not detected, but continuing...")
                
                await asyncio.sleep(3.0) # Buffer for large files

            # Send
            self.log(f"[W{window_id}] Clicking Send...")
            # Use a more specific selector for the primary Send button
            await page.locator('div[role="button"][aria-label^="Send"]').first.click()
            await asyncio.sleep(1.0) # Ensure send triggers
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
            if os.path.exists("temp_attachments"):
                shutil.rmtree("temp_attachments")
        except:
            pass
        self.destroy()

if __name__ == "__main__":
    # Essential for PyInstaller standalone executables
    multiprocessing.freeze_support()
    
    # Only try to install if NOT frozen (i.e., running from source)
    # For frozen apps, browsers should be bundled or pre-installed
    if not getattr(sys, 'frozen', False):
        threading.Thread(target=install_browsers, daemon=True).start()
        
    app = App()
    app.mainloop()
