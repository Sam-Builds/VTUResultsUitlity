import os
import sys
import time
import tempfile
import base64
import threading
import ctypes
from ctypes import wintypes
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import ssl
os.environ['WDM_SSL_VERIFY'] = '0'
ssl._create_default_https_context = ssl._create_unverified_context
reader = None
USE_GPU = False

def get_desktop_dir():
    if os.name == "nt":
        try:
            FOLDERID_Desktop = ctypes.c_char_p(
                b"{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
            )
            path_ptr = ctypes.c_wchar_p()
            shell32 = ctypes.windll.shell32
            ole32 = ctypes.windll.ole32

            class GUID(ctypes.Structure):
                _fields_ = [
                    ("Data1", ctypes.c_ulong),
                    ("Data2", ctypes.c_ushort),
                    ("Data3", ctypes.c_ushort),
                    ("Data4", ctypes.c_ubyte * 8),
                ]

            guid = GUID()
            ole32.IIDFromString(wintypes.LPCOLESTR("{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"), ctypes.byref(guid))
            hr = shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(path_ptr))
            if hr == 0 and path_ptr.value:
                desktop = path_ptr.value
                ctypes.windll.ole32.CoTaskMemFree(path_ptr)
                if os.path.isdir(desktop):
                    return desktop
        except Exception:
            pass

    candidates = [
        os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
        os.path.join(os.environ.get("OneDrive", ""), "Desktop"),
        os.path.join(os.path.expanduser("~"), "Desktop"),
    ]
    for candidate in candidates:
        if candidate and os.path.isdir(candidate):
            return candidate

    return os.path.expanduser("~")


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def apply_window_icon(window: tk.Misc):
    try:
        ico_path = resource_path("combined.ico")
        png_path = resource_path("combined.png")

        if os.path.exists(ico_path):
            try:
                window.iconbitmap(ico_path)
            except tk.TclError:
                pass

        if os.path.exists(png_path):
            try:
                icon_img = tk.PhotoImage(file=png_path)
                window._combined_icon_ref = icon_img
                window.iconphoto(True, icon_img)
            except tk.TclError:
                pass
    except Exception:
        pass

def initialize_ocr():
    global reader, USE_GPU
    try:
        import torch
        import easyocr
        
        USE_GPU = torch.cuda.is_available()
        
        local_model_path = resource_path('model')
        
        reader = easyocr.Reader(
            ['en'], 
            gpu=USE_GPU, 
            model_storage_directory=local_model_path, 
            download_enabled=False 
        )
        print("OCR Engine Ready.")
    except Exception as e:
        import traceback
        with open("ocr_crash.txt", "w") as f:
            f.write(traceback.format_exc())
        print(f"OCR Engine failed to initialize: {e}")
        reader = "ERROR"

def solve_captcha(img_path):
    import cv2
    global reader
    if reader is None:
        initialize_ocr()
        
    img = cv2.imread(img_path)
    if img is None: return ""
        
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    adjusted = cv2.convertScaleAbs(gray, alpha=0.3, beta=33)
    _, cleaned = cv2.threshold(adjusted, 65, 255, cv2.THRESH_BINARY)

    results = reader.readtext(
        cleaned, detail=0, 
        allowlist='abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        paragraph=False, min_size=20, contrast_ths=0.1, slope_ths=0.5,     
        adjust_contrast=0.5, text_threshold=0.8, link_threshold=0.1, add_margin=0.1        
    )
    
    prediction = "".join(results).replace(" ", "").strip()
    if len(prediction) > 6:
        prediction = prediction[:6]
    return prediction


def create_temp_captcha_path(usn, attempt):
    # Use a unique path in the OS temp folder so packaged installs never read/write bundled files.
    safe_usn = "".join(ch for ch in str(usn) if ch.isalnum()) or "usn"
    stamp = int(time.time() * 1000)
    filename = f"vtu_captcha_{safe_usn}_{attempt}_{stamp}.png"
    return os.path.join(tempfile.gettempdir(), filename)

def save_as_pdf(driver, usn, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    print_options = {
        'landscape': False,
        'displayHeaderFooter': False,
        'printBackground': True,
        'preferCSSPageSize': True,
    }
    pdf_data = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    
    file_path = os.path.join(output_dir, f"{usn}.pdf")
    with open(file_path, "wb") as f:
        f.write(base64.b64decode(pdf_data['data']))
    return file_path


def write_scrape_run_report(
    output_dir,
    mode,
    target_url,
    requested_usns,
    downloaded_usns,
    no_result_usns,
    error_map,
    was_cancelled,
):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(output_dir, f"scrape_run_report_{stamp}.txt")

    total_requested = len(requested_usns)
    total_downloaded = len(downloaded_usns)
    total_no_result = len(no_result_usns)
    total_errors = len(error_map)

    lines = [
        "VTU SCRAPER RUN REPORT",
        "=" * 60,
        f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Mode: {mode}",
        f"Target URL: {target_url}",
        f"Output Folder: {output_dir}",
        f"Cancelled: {'Yes' if was_cancelled else 'No'}",
        "",
        "Summary",
        "-" * 60,
        f"Requested USNs: {total_requested}",
        f"Downloaded PDFs: {total_downloaded}",
        f"No Results Found: {total_no_result}",
        f"Errored USNs: {total_errors}",
        "",
    ]

    if downloaded_usns:
        lines.append("Downloaded USNs")
        lines.append("-" * 60)
        lines.extend(downloaded_usns)
        lines.append("")

    if no_result_usns:
        lines.append("USNs With No Results")
        lines.append("-" * 60)
        lines.extend(no_result_usns)
        lines.append("")

    if error_map:
        lines.append("USNs With Errors")
        lines.append("-" * 60)
        for usn in sorted(error_map):
            lines.append(f"{usn}: {error_map[usn]}")
        lines.append("")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return report_path

def scraper_worker(usn_list, mode, target_url, output_dir, gui_app):
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import NoAlertPresentException
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.edge.service import Service as EdgeService
    from webdriver_manager.microsoft import EdgeChromiumDriverManager

    if mode == "Auto":
        if reader is None:
            gui_app.update_status("Loading AI Engine...")
            while reader is None:
                time.sleep(0.5)
        
        if reader == "ERROR":
            gui_app.update_log("OCR Engine failed to initialize! Check missing PyInstaller imports (e.g. torch, cv2).")
            gui_app.update_status("AI Engine Error.")
            gui_app.root.after(0, gui_app.reset_gui_state)
            return

    gui_app.update_status("Setting up browser...")
    try:
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
    except Exception as e:
        gui_app.update_log(f"Browser launch failed: {e}")
        gui_app.update_status("Browser Error.")
        gui_app.root.after(0, gui_app.reset_gui_state)
        return
        
    downloaded_count = 0
    downloaded_usns = []
    no_result_usns = []
    error_map = {}

    for usn in usn_list:
        if gui_app.cancel_flag:
            break 

        gui_app.update_status(f"Processing: {usn}")
        attempts = 1
        success = False
        
        while not success:
            if gui_app.cancel_flag:
                break 

            gui_app.update_log(f"[{usn}] Attempt {attempts}...")
            driver.get(target_url)
            
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'captcha')]"))
                )
                
                usn_box = driver.find_element(By.NAME, "lns") 
                captcha_box = driver.find_element(By.NAME, "captchacode")
                submit_btn = driver.find_element(By.ID, "submit")
                captcha_img = driver.find_element(By.XPATH, "//img[contains(@src, 'captcha')]")
                
                captcha_path = create_temp_captcha_path(usn, attempts)
                try:
                    captcha_img.screenshot(captcha_path)

                    if mode == "Auto":
                        predicted_code = solve_captcha(captcha_path)
                        gui_app.update_log(f"[{usn}] Automatic Guess: {predicted_code}")
                    else:
                        predicted_code = gui_app.request_manual_captcha(captcha_path)
                        if predicted_code is None:
                            break
                finally:
                    try:
                        if os.path.exists(captcha_path):
                            os.remove(captcha_path)
                    except OSError:
                        pass

                usn_box.clear()
                usn_box.send_keys(usn)
                captcha_box.clear()
                captcha_box.send_keys(predicted_code)
                submit_btn.click()
                time.sleep(1.5)
                
                try:
                    alert = driver.switch_to.alert
                    alert_text = alert.text
                    alert.accept()
                    
                    if "Invalid" in alert_text or "captcha" in alert_text.lower():
                        gui_app.update_log(f"[{usn}] Captcha Failed. Retrying...")
                        attempts += 1
                        continue
                    elif "found" in alert_text.lower() or "exist" in alert_text.lower():
                        gui_app.update_log(f"[{usn}] Server says: No Results Found. Skipping.")
                        no_result_usns.append(usn)
                        success = True 
                        continue
                except NoAlertPresentException:
                    pass 
                
                if "Semester" in driver.page_source or "phide" in driver.page_source:
                    gui_app.update_log(f"[{usn}] Login Success! Saving PDF...")
                    saved_path = save_as_pdf(driver, usn, output_dir)
                    gui_app.update_log(f"[{usn}] Saved: {saved_path}")
                    downloaded_count += 1
                    downloaded_usns.append(usn)
                    success = True
                else:
                    gui_app.update_log(f"[{usn}] Unrecognized page state. Retrying...")
                    attempts += 1

            except Exception as e:
                gui_app.update_log(f"Error on {usn}: {str(e)}")
                error_map[usn] = str(e)
                time.sleep(2)
                attempts += 1

    report_path = None
    try:
        report_path = write_scrape_run_report(
            output_dir=output_dir,
            mode=mode,
            target_url=target_url,
            requested_usns=usn_list,
            downloaded_usns=downloaded_usns,
            no_result_usns=no_result_usns,
            error_map=error_map,
            was_cancelled=gui_app.cancel_flag,
        )
        gui_app.update_log(f"Run report saved: {report_path}")
    except Exception as report_err:
        gui_app.update_log(f"Could not write run report: {report_err}")

    if gui_app.cancel_flag:
        gui_app.update_log("Scraping Terminated by User.")
        gui_app.update_status("Cancelled.")
        if report_path:
            gui_app.update_log(f"Cancellation report: {report_path}")
    else:
        gui_app.update_status("All USNs processed!")
        gui_app.update_log(
            f"Scraping complete. Downloaded {downloaded_count}/{len(usn_list)} PDFs. "
            f"Check the '{output_dir}' folder."
        )
        if report_path:
            gui_app.update_log(f"Summary report: {report_path}")

        all_downloaded = downloaded_count == len(usn_list)
        if all_downloaded and hasattr(gui_app, 'on_complete_callback') and gui_app.on_complete_callback:
            gui_app.root.after(0, lambda: gui_app.on_complete_callback(output_dir))
        elif all_downloaded:
            messagebox.showinfo("Done", f"Scraping complete. Check the '{output_dir}' folder.")
        else:
            messagebox.showwarning(
                "Scraping Finished",
                f"Downloaded {downloaded_count} of {len(usn_list)} requested PDFs.\n\n"
                f"Output folder:\n{output_dir}"
            )
    
    driver.quit()
    try:
        gui_app.root.after(0, gui_app.reset_gui_state)
    except tk.TclError:
        pass

class VTUScraperGUI(tk.Frame):
    def __init__(self, parent, on_complete_callback=None, on_cancel_callback=None):
        super().__init__(parent, bg="#f5f7fa")
        self.parent = parent
        self.root = parent.winfo_toplevel()
        apply_window_icon(self.root)
        self.on_complete_callback = on_complete_callback
        self.on_cancel_callback = on_cancel_callback
        
        self.BG = "#f5f7fa"
        self.CARD = "#ffffff"
        self.ACCENT = "#1a73e8"
        self.ACCENT_DK = "#1558b0"
        self.ACCENT_PALE = "#e8f0fe"
        self.T_DARK = "#202124"
        self.T_MID = "#5f6368"
        self.T_LIGHT = "#9aa0a6"
        self.BORDER_CLR = "#dadce0"
        self.OK_CLR = "#1e8e3e"
        self.ERR_CLR = "#d93025"
        
        self.F_TITLE = ("Segoe UI", 14, "bold")
        self.F_HEAD = ("Segoe UI", 12, "bold")
        self.F_BODY = ("Segoe UI", 11)
        self.F_SMALL = ("Segoe UI", 10)
        self.F_LOG = ("Consolas", 10)
        
        self.cancel_flag = False
        self.is_scraping = False
        self.waiting_for_manual = False
        self.manual_captcha_result = None
        self.manual_popup = None
        
        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg=self.ACCENT, height=50)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(
            header, 
            text="VTU Auto-Scraper",
            font=self.F_TITLE,
            bg=self.ACCENT,
            fg="white"
        ).pack(side="left", padx=20, pady=12)
        
        main_host = tk.Frame(self, bg=self.BG)
        main_host.pack(fill="both", expand=True, padx=14, pady=10)

        self.main_canvas = tk.Canvas(main_host, bg=self.BG, highlightthickness=0, bd=0)
        self.main_scroll = ttk.Scrollbar(main_host, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scroll.set)
        self.main_scroll.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)

        main = tk.Frame(self.main_canvas, bg=self.BG)
        self.main_inner = main
        self.main_canvas_window = self.main_canvas.create_window((0, 0), window=main, anchor="nw")

        def _sync_scroll_region(_event=None):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
            self._refresh_scraper_scrollbar()

        def _sync_main_width(event):
            self.main_canvas.itemconfigure(self.main_canvas_window, width=event.width)
            _sync_scroll_region()

        main.bind("<Configure>", _sync_scroll_region)
        self.main_canvas.bind("<Configure>", _sync_main_width)

        def _on_mousewheel(event):
            self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.main_canvas.bind("<Enter>", lambda _e: self.main_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.main_canvas.bind("<Leave>", lambda _e: self.main_canvas.unbind_all("<MouseWheel>"))

        url_frame = tk.Frame(main, bg=self.CARD, highlightbackground=self.BORDER_CLR, highlightthickness=1)
        url_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            url_frame,
            text="Target Website URL",
            font=self.F_HEAD,
            bg=self.CARD,
            fg=self.T_DARK,
            anchor="w"
        ).pack(anchor="w", padx=12, pady=(8, 4))

        self.url_entry = ttk.Entry(url_frame, font=self.F_BODY)
        self.url_entry.insert(0, "https://results.vtu.ac.in/D25J26Ecbcs/index.php")
        self.url_entry.pack(fill="x", padx=12, pady=(0, 10))

        range_frame = tk.Frame(main, bg=self.CARD, highlightbackground=self.BORDER_CLR, highlightthickness=1)
        range_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            range_frame,
            text="USN Target Range",
            font=self.F_HEAD,
            bg=self.CARD,
            fg=self.T_DARK,
            anchor="w"
        ).pack(anchor="w", padx=12, pady=(8, 6))

        prefix_row = tk.Frame(range_frame, bg=self.CARD)
        prefix_row.pack(fill="x", padx=12, pady=(0, 8))

        tk.Label(
            prefix_row,
            text="Prefix (e.g., 4DM25CD):",
            font=self.F_SMALL,
            bg=self.CARD,
            fg=self.T_MID,
            width=18,
            anchor="w"
        ).pack(side="left")

        self.prefix_entry = ttk.Entry(prefix_row, font=self.F_BODY, width=20)
        self.prefix_entry.insert(0, "4DM25CD")
        self.prefix_entry.pack(side="left", padx=(8, 0))

        start_row = tk.Frame(range_frame, bg=self.CARD)
        start_row.pack(fill="x", padx=12, pady=(0, 8))

        tk.Label(
            start_row,
            text="Start Number (e.g., 1):",
            font=self.F_SMALL,
            bg=self.CARD,
            fg=self.T_MID,
            width=18,
            anchor="w"
        ).pack(side="left")

        self.start_entry = ttk.Entry(start_row, font=self.F_BODY, width=20)
        self.start_entry.insert(0, "1")
        self.start_entry.pack(side="left", padx=(8, 0))

        end_row = tk.Frame(range_frame, bg=self.CARD)
        end_row.pack(fill="x", padx=12, pady=(0, 10))

        tk.Label(
            end_row,
            text="End Number (e.g., 50):",
            font=self.F_SMALL,
            bg=self.CARD,
            fg=self.T_MID,
            width=18,
            anchor="w"
        ).pack(side="left")

        self.end_entry = ttk.Entry(end_row, font=self.F_BODY, width=20)
        self.end_entry.insert(0, "10")
        self.end_entry.pack(side="left", padx=(8, 0))

        mode_frame = tk.Frame(main, bg=self.CARD, highlightbackground=self.BORDER_CLR, highlightthickness=1)
        mode_frame.pack(fill="x", pady=(0, 16))

        tk.Label(
            mode_frame,
            text="Captcha Mode",
            font=self.F_HEAD,
            bg=self.CARD,
            fg=self.T_DARK,
            anchor="w"
        ).pack(anchor="w", padx=12, pady=(8, 6))

        mode_options = tk.Frame(mode_frame, bg=self.CARD)
        mode_options.pack(anchor="w", padx=12, pady=(0, 10))

        self.mode_var = tk.StringVar(master=self, value="Auto")
        auto_frame = tk.Frame(mode_options, bg=self.CARD)
        auto_frame.pack(side="left", padx=(0, 20))

        auto_radio = tk.Radiobutton(
            auto_frame,
            text="Auto-Solve",
            variable=self.mode_var,
            value="Auto",
            font=self.F_BODY,
            bg=self.CARD,
            fg=self.T_DARK,
            activebackground=self.CARD,
            selectcolor=self.CARD,
            highlightbackground=self.BORDER_CLR,
            indicatoron=1
        )
        auto_radio.pack(side="left")

        manual_frame = tk.Frame(mode_options, bg=self.CARD)
        manual_frame.pack(side="left")

        manual_radio = tk.Radiobutton(
            manual_frame,
            text="Manual Entry",
            variable=self.mode_var,
            value="MANUAL",
            font=self.F_BODY,
            bg=self.CARD,
            fg=self.T_DARK,
            activebackground=self.CARD,
            selectcolor=self.CARD,
            highlightbackground=self.BORDER_CLR,
            indicatoron=1
        )
        manual_radio.pack(side="left")
        btn_frame = tk.Frame(main, bg=self.BG)
        btn_frame.pack(fill="x", pady=(0, 12))

        self.start_btn = tk.Button(
            btn_frame,
            text="START SCRAPING",
            font=("Segoe UI", 11, "bold"),
            bg=self.ACCENT,
            fg="white",
            activebackground=self.ACCENT_DK,
            activeforeground="white",
            bd=0,
            padx=20,
            pady=10,
            relief="flat",
            cursor="hand2",
            command=self.start_scraping
        )
        self.start_btn.pack(side="left", expand=True, fill="x", padx=(0, 6))

        self.cancel_btn = tk.Button(
            btn_frame,
            text="CANCEL",
            font=("Segoe UI", 11, "bold"),
            bg=self.T_LIGHT,
            fg="white",
            activebackground=self.ERR_CLR,
            activeforeground="white",
            bd=0,
            padx=20,
            pady=10,
            relief="flat",
            cursor="hand2",
            state="disabled",
            command=self.trigger_cancel
        )
        self.cancel_btn.pack(side="right", expand=True, fill="x", padx=(6, 0))

        self.status_lbl = tk.Label(
            main,
            text="Status: Ready",
            font=self.F_BODY,
            bg=self.BG,
            fg=self.T_MID,
            anchor="w"
        )
        self.status_lbl.pack(fill="x", pady=(0, 8))

        log_outer = tk.Frame(
            main,
            bg=self.CARD,
            highlightbackground=self.BORDER_CLR,
            highlightthickness=1
        )
        log_outer.pack(fill="both", expand=True)

        self.log_area = tk.Text(
            log_outer,
            height=10,
            font=self.F_LOG,
            bg=self.CARD,
            fg=self.T_DARK,
            bd=0,
            wrap="word",
            state="disabled",
            cursor="arrow",
            selectbackground=self.ACCENT_PALE
        )
        scrollbar = ttk.Scrollbar(log_outer, command=self.log_area.yview)
        self.log_area.configure(yscrollcommand=scrollbar.set)

        self.log_area.pack(side="left", fill="both", expand=True, padx=8, pady=6)
        scrollbar.pack(side="right", fill="y")

        self.log_area.tag_configure("ok", foreground=self.OK_CLR)
        self.log_area.tag_configure("err", foreground=self.ERR_CLR)
        self.log_area.tag_configure("warn", foreground="#f9ab00")
        self.log_area.tag_configure("info", foreground=self.T_MID)

        self._refresh_scraper_scrollbar()
        self.root.after(120, self._refresh_scraper_scrollbar)

    def _refresh_scraper_scrollbar(self):
        try:
            self.root.update_idletasks()
            content_h = self.main_inner.winfo_reqheight()
            canvas_h = self.main_canvas.winfo_height()
            if content_h <= max(1, canvas_h):
                if self.main_scroll.winfo_ismapped():
                    self.main_scroll.pack_forget()
            else:
                if not self.main_scroll.winfo_ismapped():
                    self.main_scroll.pack(side="right", fill="y")
        except tk.TclError:
            return

    def update_status(self, text):
        try:
            self.status_lbl.config(text=f"Status: {text}")
            self.root.update_idletasks()
        except tk.TclError:
            return

    def update_log(self, text):
        try:
            self.log_area.configure(state="normal")
        except tk.TclError:
            return

        if "Error" in text or "Failed" in text:
            tag = "err"
        elif "Success" in text or "Saved" in text:
            tag = "ok"
        else:
            tag = "info"
        
        self.log_area.insert("end", text + "\n", tag)
        self.log_area.see("end")
        self.log_area.configure(state="disabled")
        try:
            self.root.update_idletasks()
        except tk.TclError:
            return

    def start_scraping(self):
        if self.is_scraping:
            return

        mode = self.mode_var.get()
        print(f"Selected mode: {mode}")
        target_url = self.url_entry.get().strip()
        if not target_url:
            messagebox.showerror("Error", "Please enter a valid Target Website URL.")
            return
        prefix = self.prefix_entry.get().strip().upper()
        if not prefix:
            messagebox.showerror("Error", "Please enter a valid USN prefix.")
            return

        invalid_chars = set('\\/:*?"<>|')
        if any(ch in invalid_chars for ch in prefix):
            messagebox.showerror("Error", "Prefix contains invalid filename characters for Windows.")
            return

        try:
            start_num = int(self.start_entry.get().strip())
            end_num = int(self.end_entry.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Start and End must be numbers.")
            return

        if start_num <= 0 or end_num <= 0:
            messagebox.showerror("Error", "Start and End must be positive numbers.")
            return

        if end_num < start_num:
            messagebox.showerror("Error", "End Number must be greater than or equal to Start Number.")
            return

        if (end_num - start_num + 1) > 500:
            messagebox.showwarning(
                "Large Batch",
                "You are attempting to scrape more than 500 USNs in one run. Consider splitting into smaller batches.",
            )

        self.prefix_entry.delete(0, tk.END)
        self.prefix_entry.insert(0, prefix)

        usn_list = [f"{prefix}{str(i).zfill(3)}" for i in range(start_num, end_num + 1)]
        mode = self.mode_var.get()
        self.cancel_flag = False
        self.is_scraping = True
        self.start_btn.config(state="disabled", bg=self.T_LIGHT)
        self.cancel_btn.config(state="normal", bg=self.ERR_CLR)
        
        self.log_area.configure(state="normal")
        self.log_area.delete(1.0, tk.END)
        self.log_area.configure(state="disabled")
        
        
        # Calculate desktop output directory based on prefix
        desktop_dir = get_desktop_dir()
        output_dir = os.path.join(desktop_dir, prefix)
        os.makedirs(output_dir, exist_ok=True)

        threading.Thread(target=scraper_worker, args=(usn_list, mode, target_url, output_dir, self), daemon=True).start()

    def trigger_cancel(self):
        if not self.is_scraping:
            return
        self.cancel_flag = True
        self.update_status("Cancelling...")
        self.cancel_btn.config(state="disabled", bg=self.T_LIGHT)
        if self.on_cancel_callback:
            try:
                self.root.after(0, self.on_cancel_callback)
            except tk.TclError:
                pass
        if self.waiting_for_manual:
            self.waiting_for_manual = False

    def reset_gui_state(self):
        self.is_scraping = False
        try:
            if hasattr(self, "start_btn") and self.start_btn.winfo_exists():
                self.start_btn.config(state="normal", bg=self.ACCENT)
            if hasattr(self, "cancel_btn") and self.cancel_btn.winfo_exists():
                self.cancel_btn.config(state="disabled", bg=self.T_LIGHT)
            if self.manual_popup is not None and self.manual_popup.winfo_exists():
                self.manual_popup.destroy()
                self.manual_popup = None
        except tk.TclError:
            return

    def request_manual_captcha(self, img_path):
        self.manual_captcha_result = None
        self.waiting_for_manual = True
        self.root.after(0, lambda: self._update_or_create_popup(img_path))
        while self.waiting_for_manual:
            if self.cancel_flag: return None
            time.sleep(0.1)
        return self.manual_captcha_result




    def _update_or_create_popup(self, img_path):
        img = Image.open(img_path)
        img = img.resize((150, 50), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        
        if self.manual_popup is None or not self.manual_popup.winfo_exists():
            self.manual_popup = tk.Toplevel(self.root)
            self.manual_popup.title("Manual Captcha")
            self.manual_popup.configure(bg=self.CARD)
            self.manual_popup.resizable(False, False)
            self.manual_popup.attributes("-topmost", True)
            apply_window_icon(self.manual_popup)

            
            self.manual_popup.update_idletasks()
            width, height = 300, 250
            x = (self.manual_popup.winfo_screenwidth() - width) // 2
            y = (self.manual_popup.winfo_screenheight() - height) // 2
            self.manual_popup.geometry(f"{width}x{height}+{x}+{y}")
            

            tk.Label(
                self.manual_popup,
                text="Enter Captcha Code",
                font=self.F_HEAD,
                bg=self.CARD,
                fg=self.T_DARK
            ).pack(pady=(15, 10))
            
            self.manual_img_lbl = tk.Label(self.manual_popup, bg=self.CARD, image=photo)
            self.manual_img_lbl.image = photo
            self.manual_img_lbl.pack(pady=5)
            
            self.manual_entry = ttk.Entry(
                self.manual_popup,
                font=("Segoe UI", 14),
                justify="center",
                width=10
            )
            self.manual_entry.pack(pady=10)
            
            def on_submit(event=None):
                self.manual_captcha_result = self.manual_entry.get().strip()
                self.waiting_for_manual = False
            
            self.manual_popup.bind('<Return>', on_submit)
            
            submit_btn = tk.Button(
                self.manual_popup,
                text="Submit (Enter)",
                font=self.F_BODY,
                bg=self.ACCENT,
                fg="white",
                activebackground=self.ACCENT_DK,
                activeforeground="white",
                bd=0,
                padx=20,
                pady=6,
                relief="flat",
                cursor="hand2",
                command=on_submit
            )
            submit_btn.pack(pady=5)
            
            def on_close():
                self.trigger_cancel()
                self.manual_popup.destroy()
            
            self.manual_popup.protocol("WM_DELETE_WINDOW", on_close)
            
        else:
            self.manual_img_lbl.config(image=photo)
            self.manual_img_lbl.image = photo
            self.manual_entry.delete(0, tk.END)
            self.manual_popup.lift()
        
        self.manual_entry.focus_set()


if __name__ == "__main__":
    root = tk.Tk()
    apply_window_icon(root)
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    win_w = max(620, min(760, int(screen_w * 0.42)))
    win_h = max(560, min(700, int(screen_h * 0.76)))
    pos_x = max(30, (screen_w - win_w) // 2)
    pos_y = max(20, (screen_h - win_h) // 2)
    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    root.withdraw()
    splash = tk.Toplevel(root)
    splash.title("Starting")
    splash.geometry("300x120")
    splash.overrideredirect(True)
    
    splash.configure(bg="#f5f7fa")
    
    screen_width = splash.winfo_screenwidth()
    screen_height = splash.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (120 // 2)
    splash.geometry(f"300x120+{x}+{y}")

    tk.Label(
        splash, 
        text="VTU AUTO-SCRAPER", 
        font=("Segoe UI", 14, "bold"),
        bg="#f5f7fa",
        fg="#1a73e8"
    ).pack(pady=20)
    
    tk.Label(
        splash, 
        text="Initializing components...", 
        font=("Segoe UI", 10),
        bg="#f5f7fa",
        fg="#5f6368"
    ).pack()
    
    splash.update()

    app = VTUScraperGUI(root)
    app.pack(fill="both", expand=True)

    threading.Thread(target=initialize_ocr, daemon=True).start()

    splash.after(800, lambda: [splash.destroy(), root.deiconify()])
    
    root.mainloop()