import customtkinter as ctk
import os
from PIL import Image

COLOR_PASS = "#10b981"
COLOR_FAIL = "#ef4444"
COLOR_WARN = "#f59e0b"
COLOR_ACCENT = "#2563eb" 
COLOR_BG_DEEP = "#0f172a" 
COLOR_BG_LIGHT = "#1e293b"

def load_logo(size=(50, 50)):
    """Load the Rijksoverheid logo from the app directory."""
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    try:
        return ctk.CTkImage(light_image=Image.open(logo_path), dark_image=Image.open(logo_path), size=size)
    except Exception:
        return None

class LoadingFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, fg_color=COLOR_BG_DEEP)
        self.controller = controller
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Logo linksboven
        logo_img = load_logo()
        if logo_img:
            self.logo_lbl = ctk.CTkLabel(self, image=logo_img, text="")
            self.logo_lbl.grid(row=0, column=0, padx=15, pady=10, sticky="nw")
        
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.grid(row=1, column=0)
        
        self.title_lbl = ctk.CTkLabel(self.container, text="Bezig met scannen...", font=("Segoe UI Black", 24), text_color="white")
        self.title_lbl.pack(pady=20)
        
        self.status_lbl = ctk.CTkLabel(self.container, text="Voorbereiden...", font=("Segoe UI", 16), text_color="#e2e8f0")
        self.status_lbl.pack(pady=10)
        
        self.progress = ctk.DoubleVar(value=0)
        self.progressbar = ctk.CTkProgressBar(self.container, width=400, variable=self.progress, progress_color=COLOR_ACCENT)
        self.progressbar.pack(pady=20)
        
        self.time_lbl = ctk.CTkLabel(self.container, text="Resterende tijd berekenen...", font=("Consolas", 12), text_color="#888888")
        self.time_lbl.pack(pady=(0, 10))
        
        self.btn_cancel = ctk.CTkButton(self.container, text="✖ Annuleren", font=("Segoe UI", 14, "bold"), fg_color="#7f1d1d", hover_color="#991b1b", command=self.cancel_scan)
        self.btn_cancel.pack(pady=10)

    def reset(self):
        import time
        self.progress.set(0)
        self.start_time = time.time()
        self.title_lbl.configure(text="Bezig met scannen...")
        self.status_lbl.configure(text="Voorbereiden...")
        self.time_lbl.configure(text="Resterende tijd berekenen...")
        
    def cancel_scan(self):
        self.title_lbl.configure(text="Bezig met annuleren...")
        self.controller.cancel_analysis()
        
    def update_progress(self, val, status_text=None):
        import time
        if hasattr(val, "get"):
            val = val.get()
        self.progress.set(val)
        if status_text:
            self.status_lbl.configure(text=status_text)
            
        if hasattr(self, 'start_time') and val > 0.05:
            elapsed = time.time() - self.start_time
            if val > 0:
                total_est = elapsed / val
                rem = max(0, total_est - elapsed)
                m, s = divmod(int(rem), 60)
                self.time_lbl.configure(text=f"Geschatte resterende tijd: {m}m {s}s")
        else:
            self.time_lbl.configure(text="Resterende tijd berekenen...")
            
        self.update_idletasks()

import customtkinter as ctk
import threading
import queue
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD

# Engines
from centrale_engine import Centrale_Engine
from compliance_engine import ComplianceEngine
from KwaliteitEngine import QualityEngine

# Frames
from home import AuthFrame
from home import HomeFrame

from dashboard import DashboardFrame

COLOR_PASS = "#10b981"
COLOR_FAIL = "#ef4444"
COLOR_WARN = "#f59e0b"
COLOR_ACCENT = "#2563eb" 
COLOR_BG_DEEP = "#0f172a" 
COLOR_BG_LIGHT = "#1e293b"

import sys
import os

def get_app_dir():
    if getattr(sys, 'frozen', False): return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

theme_path = os.path.join(get_app_dir(), "theme", "modern_corporate.json")
try:
    if os.path.exists(theme_path): ctk.set_default_color_theme(theme_path)
    else: ctk.set_default_color_theme("blue")
except Exception:
    ctk.set_default_color_theme("blue")

ctk.set_appearance_mode("Dark")

class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

import centrale_engine as database

class ComplianceApp(TkinterDnD_CTk):
    def __init__(self):
        super().__init__()
        database.init_db()
        self.title("Document Scanner")
        self.geometry("1100x750") 
        self.minsize(950, 650)
        self.configure(fg_color=COLOR_BG_DEEP) 
        
        # Laad het logo één keer voor alle frames
        self.logo_image = load_logo()
        
        self.current_user = None
        self.q = queue.Queue()
        self.is_analyzing = False 
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Initalize all frames
        self.frames = {}
        for F in (AuthFrame, HomeFrame, LoadingFrame, DashboardFrame):
            frame_name = F.__name__
            frame = F(parent=self, controller=self)
            self.frames[frame_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
            
        self.show_frame("AuthFrame")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def set_user(self, user):
        self.current_user = user

    def show_frame(self, frame_name):
        frame = self.frames[frame_name]
        frame.tkraise()
        # Optionally trigger an update method if frame has one
        if hasattr(frame, "update_ui_for_user"):
            frame.update_ui_for_user()

    def start_analysis(self, afdeling, local_paths, sharepoint_sites, active_comp, active_qual):
        if self.is_analyzing: return
        self.is_analyzing = True
        
        self.current_afdeling = afdeling
        
        # Laat loading screen zien
        self.show_frame("LoadingFrame")
        self.frames["LoadingFrame"].reset()

        active_engines = []
        if active_comp:
            active_engines.append(ComplianceEngine(active_comp))
        if active_qual:
            active_engines.append(QualityEngine(active_qual))
            
        self.stop_event = threading.Event()
        # Pass current_user to the engine
        scanner = Centrale_Engine(local_paths, sharepoint_sites, active_engines, self.stop_event)
        # Sla op welke modules we gebruiken om later dashboard op te splitsen
        self.last_run_comp_modules = active_comp
        self.last_run_qual_modules = active_qual
        threading.Thread(target=scanner.process, args=(self.q,), daemon=True).start()
        self.check_queue()

    def check_queue(self):
        if not self.winfo_exists(): return 
        
        try:
            while True:
                msg_type, data = self.q.get_nowait()
                if msg_type == "progress": 
                    self.frames["LoadingFrame"].update_progress(data)
                elif msg_type == "status":
                    self.frames["LoadingFrame"].update_progress(self.frames["LoadingFrame"].progress.get(), data)
                elif msg_type == "error":
                    messagebox.showerror("QA Systeemfout", data)
                    self.is_analyzing = False
                    self.show_frame("HomeFrame")
                    return
                elif msg_type == "canceled":
                    self.is_analyzing = False
                    self.show_frame("HomeFrame")
                    return
                elif msg_type == "done":
                    self.is_analyzing = False
                    
                    # Sla scan op in database
                    if self.current_user and hasattr(self.current_user, 'id'):
                        import centrale_engine as database
                        from datetime import datetime
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        database.save_scan(self.current_user.id, self.current_afdeling, timestamp, data.get("results", []))

                    # Load data into dashboard and show
                    self.frames["DashboardFrame"].load_data(data, self.current_afdeling)
                    self.show_frame("DashboardFrame")
                    return 
        except queue.Empty: pass 
        
        if self.is_analyzing:
            self.after(100, self.check_queue)
            
    def cancel_analysis(self):
        if self.is_analyzing and hasattr(self, 'stop_event'):
            self.stop_event.set()

    def on_closing(self):
        self.destroy() 

if __name__ == "__main__":
    app = ComplianceApp()
    app.mainloop()
