# ==============================================================================
# ana.py — Applicatie-entrypoint: ComplianceApp en LoadingFrame
# ==============================================================================

# --- Stdlib imports ---
import os
import queue
import sys
import threading
import time
from datetime import datetime

# --- Third-party imports ---
import customtkinter as ctk
from PIL import Image
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD

# --- Local imports ---
import centrale_engine as database
# Heavy engines worden pas geladen bij start_analysis om RAM/CPU te sparen
from dashboard import DashboardFrame
from home import AuthFrame, HomeFrame


# ==============================================================================
# KLEUR-CONSTANTEN (Gold/Blue theme)
# ==============================================================================

COLOR_BG_DEEP    = "#001538"   # Diepdonkerblauw — hoofdachtergrond
COLOR_BG_LIGHT   = "#1a2b4b"   # Licht donkerblauw — cards / rijen
COLOR_ACCENT     = "#cf9d1f"   # Goud — knoppen, titels, actieve accenten
COLOR_ACCENT_HOVER = "#b0841a" # Donkerder goud — hover-state
COLOR_PASS       = "#10b981"   # Groen — compliant
COLOR_FAIL       = "#ef4444"   # Rood — fout / kritiek
COLOR_WARN       = "#f59e0b"   # Oranje — waarschuwing


# ==============================================================================
# THEMA & WEERGAVE
# ==============================================================================

def get_app_dir():
    """Geeft de map terug waar de applicatie draait (ook bij frozen/exe)."""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _apply_theme():
    """Laad het gold/blue CustomTkinter thema. Valt terug op 'blue' bij fout."""
    theme_path = os.path.join(get_app_dir(), "gold_blue_theme.json")
    try:
        if os.path.exists(theme_path):
            ctk.set_default_color_theme(theme_path)
        else:
            ctk.set_default_color_theme("blue")
    except Exception:
        ctk.set_default_color_theme("blue")


_apply_theme()
ctk.set_appearance_mode("Dark")


# ==============================================================================
# HELPER FUNCTIES
# ==============================================================================

def load_logo(size=(50, 50)):
    """Laad het Rijksoverheid-logo vanuit de app-map."""
    logo_path = os.path.join(get_app_dir(), "logo.png")
    try:
        return ctk.CTkImage(
            light_image=Image.open(logo_path),
            dark_image=Image.open(logo_path),
            size=size
        )
    except Exception:
        return None


# ==============================================================================
# LOADING FRAME
# ==============================================================================

class LoadingFrame(ctk.CTkFrame):
    """Scan-voortgangsscherm met progressbar, status en annuleerknop."""

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

        # Gecentreerde container
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.grid(row=1, column=0)

        self.title_lbl = ctk.CTkLabel(
            self.container, text="Bezig met scannen...",
            font=("Segoe UI Black", 24), text_color="white"
        )
        self.title_lbl.pack(pady=20)

        self.status_lbl = ctk.CTkLabel(
            self.container, text="Voorbereiden...",
            font=("Segoe UI", 16), text_color="#e2e8f0"
        )
        self.status_lbl.pack(pady=10)

        self.progress = ctk.DoubleVar(value=0)
        self.progressbar = ctk.CTkProgressBar(
            self.container, width=400,
            variable=self.progress, progress_color=COLOR_ACCENT
        )
        self.progressbar.pack(pady=20)

        self.time_lbl = ctk.CTkLabel(
            self.container, text="Resterende tijd berekenen...",
            font=("Consolas", 12), text_color="#888888"
        )
        self.time_lbl.pack(pady=(0, 10))

        self.btn_cancel = ctk.CTkButton(
            self.container, text="✖ Annuleren",
            font=("Segoe UI", 14, "bold"),
            fg_color="#7f1d1d", hover_color="#991b1b",
            command=self.cancel_scan
        )
        self.btn_cancel.pack(pady=10)

    def reset(self):
        """Reset de voortgangsbalk en labels voor een nieuwe scan."""
        self.progress.set(0)
        self.start_time = time.time()
        self.title_lbl.configure(text="Bezig met scannen...")
        self.status_lbl.configure(text="Voorbereiden...")
        self.time_lbl.configure(text="Resterende tijd berekenen...")

    def cancel_scan(self):
        self.title_lbl.configure(text="Bezig met annuleren...")
        self.controller.cancel_analysis()

    def update_progress(self, val, status_text=None):
        """Update de voortgangsbalk en de geschatte resterende tijd."""
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


# ==============================================================================
# TKINTERDND + CUSTOMTKINTER WRAPPER
# ==============================================================================

class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    """Combineert CustomTkinter met drag-and-drop ondersteuning."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


# ==============================================================================
# HOOFD APPLICATIE
# ==============================================================================

class ComplianceApp(TkinterDnD_CTk):
    """Hoofdvenster — beheert frame-navigatie, scan-orchestratie en database."""

    def __init__(self):
        super().__init__()
        database.init_db()

        self.title("Document Scanner")
        self.geometry("450x550")
        self.minsize(450, 550)
        self.configure(fg_color=COLOR_BG_DEEP)

        # State
        self.current_user = None
        self.is_analyzing = False
        self.logo_image = load_logo()

        # Inactivity Tracking
        self._last_activity_time = time.time()
        self._inactivity_timeout = 60  # seconden
        self._warning_duration = 10    # seconden
        self._warning_popup = None
        self._warning_lbl = None
        
        self.bind_all("<Any-KeyPress>", self._update_activity)
        self.bind_all("<Any-ButtonPress>", self._update_activity)
        self.bind_all("<Motion>", self._update_activity)

        # Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Frames aanmaken en registreren
        self.frames = {}
        for F in (AuthFrame, HomeFrame, LoadingFrame, DashboardFrame):
            frame = F(parent=self, controller=self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("AuthFrame")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self._check_inactivity()

    # --------------------------------------------------------------------------
    # Gebruikersbeheer
    # --------------------------------------------------------------------------

    def set_user(self, user):
        self.current_user = user
        if user is not None:
            self._update_activity()

    # --------------------------------------------------------------------------
    # Inactivity Management
    # --------------------------------------------------------------------------

    def _update_activity(self, event=None):
        self._last_activity_time = time.time()
        if self._warning_popup is not None and self._warning_popup.winfo_exists():
            self._warning_popup.destroy()
            self._warning_popup = None

    def _check_inactivity(self):
        if self.current_user is not None:
            elapsed = time.time() - self._last_activity_time
            if elapsed >= self._inactivity_timeout:
                self._auto_logout()
            elif elapsed >= (self._inactivity_timeout - self._warning_duration):
                self._show_warning_countdown(self._inactivity_timeout - elapsed)
        
        self.after(1000, self._check_inactivity)

    def _show_warning_countdown(self, remaining):
        remaining_int = max(0, int(remaining))
        if self._warning_popup is None or not self._warning_popup.winfo_exists():
            self._warning_popup = ctk.CTkToplevel(self)
            self._warning_popup.title("Inactiviteit Waarschuwing")
            self._warning_popup.geometry("350x200")
            self._warning_popup.transient(self)
            self._warning_popup.attributes("-topmost", True)
            self._warning_popup.grab_set()
            self._warning_popup.focus_force()
            
            lbl = ctk.CTkLabel(
                self._warning_popup, 
                text="Je wordt automatisch uitgelogd\nwegens inactiviteit in:", 
                font=("Segoe UI", 16)
            )
            lbl.pack(pady=(30, 20))
            
            self._warning_lbl = ctk.CTkLabel(
                self._warning_popup, 
                text=str(remaining_int), 
                font=("Segoe UI Black", 40), 
                text_color=COLOR_FAIL
            )
            self._warning_lbl.pack()
        else:
            if self._warning_lbl and self._warning_lbl.winfo_exists():
                self._warning_lbl.configure(text=str(remaining_int))

    def _auto_logout(self):
        if self._warning_popup is not None and self._warning_popup.winfo_exists():
            self._warning_popup.destroy()
            self._warning_popup = None
            
        self.set_user(None)
        if self.is_analyzing:
            self.cancel_analysis()

        # Update auth frame to clear inputs / state properly
        if "AuthFrame" in self.frames:
            self.frames["AuthFrame"]._reset_ui()
            
        self.show_frame("AuthFrame")
        
        # Show a dialog after switching frame
        messagebox.showinfo("Automatisch uitgelogd", "Je bent automatisch uitgelogd wegens inactiviteit.")

    # --------------------------------------------------------------------------
    # Frame-navigatie
    # --------------------------------------------------------------------------

    def show_frame(self, frame_name):
        """Breng het opgegeven frame naar de voorgrond en pas venstergrootte aan."""
        if frame_name == "AuthFrame":
            self.geometry("450x550")
            self.minsize(450, 550)
        else:
            self.geometry("1100x750")
            self.minsize(950, 650)

        frame = self.frames[frame_name]
        frame.tkraise()
        if hasattr(frame, "update_ui_for_user"):
            frame.update_ui_for_user()

    # --------------------------------------------------------------------------
    # Scan-beheer
    # --------------------------------------------------------------------------

    def start_analysis(self, afdeling, local_paths, sharepoint_sites, active_comp, active_qual):
        """Start een nieuwe scan in een achtergrond-thread."""
        if self.is_analyzing:
            return
        self.is_analyzing = True
        self.current_afdeling = afdeling

        self.show_frame("LoadingFrame")
        self.frames["LoadingFrame"].reset()

        # Lazy loading van de engines: dit triggert parsing libraries (docx/PyPDF)
        # pas op het moment dat de gebruiker 'Start Analyse' klikt. Dit voorkomt
        # dat de app bij het inlogscherm al veel rekenkracht of RAM vereist.
        from centrale_engine import CentraleEngine
        from compliance_engine import ComplianceEngine
        from kwaliteit_engine import QualityEngine

        active_engines = []
        if active_comp:
            active_engines.append(ComplianceEngine(active_comp))
        if active_qual:
            active_engines.append(QualityEngine(active_qual))

        self.last_run_comp_modules = active_comp
        self.last_run_qual_modules = active_qual

        self.q = queue.Queue()
        self.stop_event = threading.Event()
        scanner = CentraleEngine(local_paths, sharepoint_sites, active_engines, self.stop_event)
        threading.Thread(target=scanner.process, args=(self.q,), daemon=True).start()
        self.check_queue()

    def check_queue(self):
        """Verwerk berichten uit de scan-queue en update de UI."""
        if not self.winfo_exists():
            return

        try:
            while True:
                msg_type, data = self.q.get_nowait()

                if msg_type == "progress":
                    self.frames["LoadingFrame"].update_progress(data)

                elif msg_type == "status":
                    self.frames["LoadingFrame"].update_progress(
                        self.frames["LoadingFrame"].progress.get(), data
                    )

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
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        database.save_scan(
                            self.current_user.id,
                            self.current_afdeling,
                            timestamp,
                            data.get("results", [])
                        )

                    self.frames["DashboardFrame"].load_data(data, self.current_afdeling)
                    self.show_frame("DashboardFrame")
                    return

        except queue.Empty:
            pass

        if self.is_analyzing:
            self.after(100, self.check_queue)

    def cancel_analysis(self):
        if self.is_analyzing and hasattr(self, 'stop_event'):
            self.stop_event.set()

    def on_closing(self):
        self.destroy()


# ==============================================================================
# ENTRYPOINT
# ==============================================================================

if __name__ == "__main__":
    if "--reset-admin" in sys.argv:
        try:
            database.reset_admin()
            print("\n✔️  SUCCES: Admin wachtwoord is succesvol gereset naar 'admin' en lockouts zijn verwijderd.\n")
        except Exception as e:
            print(f"\n❌  FOUT bij het resetten van admin wachtwoord: {e}\n")
        sys.exit(0)

    app = ComplianceApp()
    app.mainloop()
