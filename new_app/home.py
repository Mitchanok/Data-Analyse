# ==============================================================================
# home.py — AuthFrame (login) en HomeFrame (scan-configuratie)
# ==============================================================================

# --- Stdlib imports ---
import os

# --- Third-party imports ---
import customtkinter as ctk
from PIL import Image
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES

# --- Local imports ---
from centrale_engine import User


# ==============================================================================
# KLEUR-CONSTANTEN (Gold/Blue theme)
# ==============================================================================

COLOR_BG_DEEP    = "#001538"   # Diepdonkerblauw — hoofdachtergrond
COLOR_BG_LIGHT   = "#1a2b4b"   # Licht donkerblauw — cards / invoervelden
COLOR_ACCENT     = "#cf9d1f"   # Goud — knoppen, titels, activeaccenten
COLOR_ACCENT_HOVER = "#b0841a" # Donkerder goud — hover-state
COLOR_PASS       = "#10b981"   # Groen — compliant
COLOR_FAIL       = "#ef4444"   # Rood — fout / kritiek
COLOR_WARN       = "#f59e0b"   # Oranje — waarschuwing


# ==============================================================================
# HELPER FUNCTIES
# ==============================================================================

def _load_logo(size=(50, 50)):
    """Laad het Rijksoverheid-logo vanuit de app-map."""
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    try:
        return ctk.CTkImage(
            light_image=Image.open(logo_path),
            dark_image=Image.open(logo_path),
            size=size
        )
    except Exception:
        return None



# ==============================================================================
# AUTH FRAME — Keuzescherm & Admin Login
# ==============================================================================

class AuthFrame(ctk.CTkFrame):
    """Startscherm met keuze tussen medewerker en beheerder."""

    _C_PANEL   = "#0d1b2e"   
    _C_INPUT   = "#162033"   
    _C_BORDER  = "#1e3a5f"   
    _C_ACCENT  = "#2a5298"   
    _C_HOVER   = "#1e3f7a"   
    _C_ERROR   = "#b91c1c"   
    _C_WARN    = "#92400e"   
    _C_TEXT    = "#e2e8f0"   
    _C_MUTED   = "#64748b"   

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color=COLOR_BG_DEEP)
        self.controller = controller
        self._username = ""
        self._countdown_job = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_ui()

    def _build_ui(self):
        self.card = ctk.CTkFrame(
            self, corner_radius=4, fg_color=self._C_PANEL,
            border_width=1, border_color=self._C_BORDER,
            width=360, height=480
        )
        self.card.grid(row=0, column=0)
        self.card.grid_propagate(False)
        self.card.grid_columnconfigure(0, weight=1)

        # Koppen
        ctk.CTkLabel(self.card, text="DOCUMENT SCANNER", font=("Segoe UI", 10, "bold"), text_color=self._C_MUTED).grid(row=0, column=0, pady=(40, 0))
        ctk.CTkLabel(self.card, text="Welkom", font=("Segoe UI Light", 26), text_color=self._C_TEXT).grid(row=1, column=0, pady=(2, 0))
        ctk.CTkFrame(self.card, height=1, fg_color=self._C_BORDER).grid(row=2, column=0, sticky="ew", padx=35, pady=(15, 15))

        self.lbl_fout = ctk.CTkLabel(self.card, text="", font=("Segoe UI", 11), text_color=self._C_ERROR, wraplength=290)
        self.lbl_fout.grid(row=3, column=0, padx=35, sticky="ew")

        # ==========================================================
        # Frame 1: Keuzescherm (Knoppen)
        # ==========================================================
        self.frame_keuze = ctk.CTkFrame(self.card, fg_color="transparent")
        self.frame_keuze.grid(row=4, column=0, padx=35, sticky="ew")

        ctk.CTkButton(
            self.frame_keuze, text="📄 Start Scan (Medewerker)",
            font=("Segoe UI", 14, "bold"), height=55, corner_radius=3,
            fg_color=self._C_ACCENT, hover_color=self._C_HOVER, text_color="#ffffff",
            command=self._login_als_gast
        ).pack(fill="x", pady=(10, 20))

        ctk.CTkButton(
            self.frame_keuze, text="🔐 Inloggen als Beheerder",
            font=("Segoe UI", 12), height=44, corner_radius=3,
            fg_color="transparent", hover_color=self._C_INPUT, text_color=self._C_MUTED,
            border_width=1, border_color=self._C_BORDER,
            command=self._toon_admin_login
        ).pack(fill="x")

        # ==========================================================
        # Frame 2: Admin Login
        # ==========================================================
        self.frame_admin = ctk.CTkFrame(self.card, fg_color="transparent")

        ctk.CTkLabel(self.frame_admin, text="Gebruikersnaam", font=("Segoe UI", 11), text_color=self._C_MUTED, anchor="w").pack(fill="x")
        self.entry_username = ctk.CTkEntry(
            self.frame_admin, height=40, font=("Segoe UI", 14),
            fg_color=self._C_INPUT, border_color=self._C_BORDER, border_width=1,
            text_color=self._C_TEXT, placeholder_text="Gebruikersnaam"
        )
        self.entry_username.pack(fill="x", pady=(2, 10))

        ctk.CTkLabel(self.frame_admin, text="Wachtwoord", font=("Segoe UI", 11), text_color=self._C_MUTED, anchor="w").pack(fill="x")
        self.entry_password = ctk.CTkEntry(
            self.frame_admin, height=40, font=("Segoe UI", 14), show="●",
            fg_color=self._C_INPUT, border_color=self._C_BORDER, border_width=1,
            text_color=self._C_TEXT, placeholder_text="Wachtwoord"
        )
        self.entry_password.pack(fill="x", pady=(2, 10))
        self.entry_password.bind("<Return>", lambda e: self._login_admin())

        self.btn_login = ctk.CTkButton(
            self.frame_admin, text="Inloggen", font=("Segoe UI", 13, "bold"),
            height=44, fg_color=self._C_ACCENT, hover_color=self._C_HOVER,
            command=self._login_admin
        )
        self.btn_login.pack(fill="x", pady=(5, 5))

        self.btn_forgot = ctk.CTkButton(
            self.frame_admin, text="Wachtwoord vergeten?", font=("Segoe UI", 11),
            height=28, fg_color="transparent", hover_color=self._C_INPUT,
            text_color=self._C_MUTED, command=self._wachtwoord_vergeten
        )
        self.btn_forgot.pack(fill="x")

        self.lbl_pogingen = ctk.CTkLabel(self.frame_admin, text="", font=("Segoe UI", 10), text_color=self._C_MUTED, anchor="w")
        self.lbl_pogingen.pack(fill="x", pady=(2, 0))

        self.btn_terug = ctk.CTkButton(
            self.frame_admin, text="← Terug", font=("Segoe UI", 11),
            height=32, fg_color="transparent", hover_color=self._C_INPUT,
            text_color=self._C_MUTED, command=self._terug_naar_keuze
        )
        self.btn_terug.pack(fill="x", pady=(5, 0))

    # --- Flow Logic ---
    def _login_als_gast(self):
        self._stop_countdown()
        from centrale_engine import User
        user = User("Gast", is_admin=False)
        user.id = 0
        self.controller.set_user(user)
        self.controller.show_frame("HomeFrame")
        self._reset_ui()

    def _toon_admin_login(self):
        self.lbl_fout.configure(text="")
        self.frame_keuze.grid_remove()
        self.frame_admin.grid(row=4, column=0, padx=35, sticky="ew")
        self.entry_username.focus()

    def _terug_naar_keuze(self):
        self._stop_countdown()
        self.lbl_fout.configure(text="")
        self.entry_username.delete(0, "end")
        self.entry_password.delete(0, "end")
        self.frame_admin.grid_remove()
        self.frame_keuze.grid(row=4, column=0, padx=35, sticky="ew")

    # --- Wachtwoord Vergeten ---
    def _wachtwoord_vergeten(self):
        from tkinter import messagebox
        username = self.entry_username.get().strip()
        if not username:
            self._toon_fout("Vul eerst je 'Gebruikersnaam' in.")
            return
            
        import centrale_engine as database
        if not database.user_exists(username):
            self._toon_fout("Gebruiker niet gevonden in de database.")
            return
            
        succes, msg = database.send_recovery_email(username, "mickstruijs@gmail.com")
        if succes:
            messagebox.showinfo("Email Verzonden", msg)
            self._toon_fout("Herstel-email gestuurd (check spam/junk folder).")
        else:
            messagebox.showerror("Fout bij verzenden", "SMTP Configuratie is niet voltooid. Zie Terminal voor evt. wachtwoordherstel/errors.")
            print(f"\n🚨 WACHTWOORD IS HERSTELD MAAR DE MAIL MISLUKTE:\n{msg}\n\n")

    # --- Admin Authenticatie ---
    def _login_admin(self):
        import centrale_engine as database
        username = self.entry_username.get().strip()
        password = self.entry_password.get()

        if not username or not password:
            self._toon_fout("Voer zowel gebruikersnaam als wachtwoord in.")
            return

        self._username = username
        status = database.get_lockout_status(self._username)
        if status["locked"] or status["rate_limited"]:
            self._start_lockout_countdown(status["seconds_remaining"], locked=status["locked"])
            return

        db_user = database.verify_user(self._username, password)
        if db_user:
            self._stop_countdown()
            from centrale_engine import User
            user = User(db_user["username"], db_user["is_admin"])
            user.id = db_user["id"]
            self.controller.set_user(user)
            self.controller.show_frame("HomeFrame")
            self._reset_ui()
        else:
            new_status = database.get_lockout_status(self._username)
            if new_status["locked"] or new_status["rate_limited"]:
                self._start_lockout_countdown(new_status["seconds_remaining"], locked=new_status["locked"])
            else:
                self._toon_fout("Ongeldige inloggegevens.")
                self._update_pogingen_label(new_status["attempt_count"])
                self.entry_password.delete(0, "end")
                self.entry_password.focus()
                
    # --- Utilities ---
    def _start_lockout_countdown(self, seconds: int, locked: bool):
        self._stop_countdown()
        self.entry_username.configure(state="disabled")
        self.entry_password.configure(state="disabled")
        self.btn_login.configure(state="disabled")

        kleur, prefix = (self._C_ERROR, "🔒 Account geblokkeerd") if locked else (self._C_WARN, "⏳ Even wachten")

        def _tick(remaining):
            if not self.winfo_exists(): return
            if remaining <= 0:
                self.entry_username.configure(state="normal")
                self.entry_password.configure(state="normal")
                self.btn_login.configure(state="normal")
                self.lbl_fout.configure(text="")
                self._update_pogingen_label(__import__("centrale_engine").get_lockout_status(self._username)["attempt_count"])
                return

            self.lbl_fout.configure(text=f"{prefix} over {remaining}s", text_color=kleur)
            self._countdown_job = self.after(1000, _tick, remaining - 1)

        _tick(seconds)

    def _stop_countdown(self):
        if self._countdown_job:
            self.after_cancel(self._countdown_job)
            self._countdown_job = None

    def _toon_fout(self, tekst: str):
        self.lbl_fout.configure(text=tekst, text_color=self._C_ERROR)

    def _update_pogingen_label(self, count: int):
        import centrale_engine as database
        if count == 0:
            self.lbl_pogingen.configure(text="")
            return
        resterend = database.MAX_ATTEMPTS_LOCKOUT - count
        if resterend > 0:
            self.lbl_pogingen.configure(text=f"Nog {resterend} poging(en).", text_color=self._C_WARN)
        else:
            self.lbl_pogingen.configure(text="")

    def _reset_ui(self):
        self._stop_countdown()
        self.lbl_fout.configure(text="")
        self.entry_username.delete(0, "end")
        self.entry_password.delete(0, "end")
        self._terug_naar_keuze()


# ==============================================================================
# HOME FRAME — Scan-configuratie
# ==============================================================================



# Tooltip helper — positioneert op basis van muispositie
class Tooltip:
    _tw = None   # gedeeld Toplevel (OS window)
    _lbl = None

    def __init__(self, widget, text: str):
        self._widget = widget
        self._text = text
        self._after_id = None
        self._mx = 0
        self._my = 0
        widget.bind("<Enter>",      self._on_enter)
        widget.bind("<Motion>",     self._on_motion)
        widget.bind("<Leave>",      self._hide)
        widget.bind("<ButtonPress>",self._hide)

    def _on_enter(self, event):
        self._mx = event.x_root
        self._my = event.y_root
        self._cancel()
        self._after_id = self._widget.after(400, self._show)

    def _on_motion(self, event):
        self._mx = event.x_root
        self._my = event.y_root
        if Tooltip._tw and Tooltip._tw.winfo_ismapped():
            self._place()

    def _hide(self, event=None):
        self._cancel()
        if Tooltip._tw:
            try:
                Tooltip._tw.withdraw()
            except Exception:
                pass

    def _cancel(self):
        if self._after_id:
            try:
                self._widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        import tkinter as tk
        if Tooltip._tw is None or not Tooltip._tw.winfo_exists():
            Tooltip._tw = tk.Toplevel(self._widget)
            Tooltip._tw.wm_overrideredirect(True)
            Tooltip._tw.wm_attributes("-topmost", True)
            try:
                Tooltip._tw.tk.call("::tk::unsupported::MacWindowStyle", "style", Tooltip._tw._w, "help", "noActivates")
            except Exception:
                pass
            
            # Subtiele gouden rand toevoegen om het echt te laten opvallen
            frame = ctk.CTkFrame(Tooltip._tw, border_width=1, border_color="#cf9d1f", fg_color="#1a2d4a", corner_radius=0)
            frame.pack(fill="both", expand=True)

            Tooltip._lbl = ctk.CTkLabel(
                frame, text="", font=("Segoe UI", 13),
                text_color="#e2e8f0", padx=15, pady=10,
                wraplength=280, justify="left"
            )
            Tooltip._lbl.pack()
        
        try:
            Tooltip._lbl.configure(text=self._text)
            Tooltip._tw.update_idletasks() # Ensure dimensions are calculated
            Tooltip._tw.deiconify()
            self._place()
        except Exception:
            pass

    def _place(self):
        if not Tooltip._tw:
            return
        
        # Plaats iets rechts en onder de cursor
        x = self._mx + 15
        y = self._my + 15
        
        # Voorkom dat de tekst van het scherm valt (heel simpele clamp)
        screen_w = self._widget.winfo_screenwidth()
        tip_w = Tooltip._tw.winfo_reqwidth()
        if x + tip_w > screen_w:
            x = self._mx - tip_w - 5

        try:
            Tooltip._tw.wm_geometry(f"+{x}+{y}")
        except Exception:
            pass





class AccordionWidget(ctk.CTkFrame):
    """Uitklapbaar frame per dimensie met vinkjes, alles-toggle en tooltip."""

    # Beschrijvingen per dimensie (voor tooltip)
    _DESCRIPTIONS = {
        "1. Nauwkeurigheid (Accuracy)": "Komt de opgeslagen waarde overeen met de realiteit?\nControleert of de auteur-eigenschappen van bestanden ingevuld zijn.",
        "2. Volledigheid (Completeness)": "Ontbreken er cruciale onderdelen?\nCheckt lege bestanden (<1KB) en documenten zonder Titel-metadata.",
        "3. Consistentie (Consistency)": "Is de informatie hetzelfde in de hele organisatie?\nControleert of de bestandsextensie klopt met het interne formaat.",
        "4. Tijdigheid (Timeliness)": "Is de data recent genoeg?\nSignaleert bestanden die meer dan 3 of 5 jaar niet zijn gewijzigd.",
        "5. Validiteit (Validity)": "Voldoet de data aan de technische syntax?\nCheckt op vergrendelde/corrupte bestanden en illegale tekens in namen.",
        "6. Uniciteit (Uniqueness)": "Wordt elk item slechts één keer geregistreerd?\nVindt exacte dubbele kopieën van bestanden op de schijf of SharePoint.",
        "7. Integriteit (Integrity)": "Zijn de koppelingen tussen databronnen correct?\nVindt dode snelkoppelingen (.lnk/.url) die verwijzen naar verwijderde locaties.",
    }

    def __init__(self, master, title, rule_vars: dict, **kwargs):
        """
        rule_vars: dict van {regel_naam: ctk.BooleanVar}
        """
        super().__init__(master, fg_color="#101e35", corner_radius=8,
                         border_width=1, border_color="#1e3a5f", **kwargs)
        self._title = title
        self._rule_vars = rule_vars
        self.is_open = False

        # ---------- Header ----------
        self.header = ctk.CTkFrame(self, fg_color="#162440", corner_radius=6)
        self.header.pack(fill="x", padx=4, pady=(4, 0))
        self.header.grid_columnconfigure(1, weight=1)
        self.header.grid_columnconfigure(2, minsize=24)  # ⓘ kolom
        self.header.grid_columnconfigure(3, minsize=78)  # toggle-knop kolom

        # Pijltje
        self.lbl_arrow = ctk.CTkLabel(
            self.header, text="►", font=("Segoe UI", 11),
            text_color="#cf9d1f", width=20
        )
        self.lbl_arrow.grid(row=0, column=0, padx=(8, 2), pady=8)

        # Titel
        self.lbl_title = ctk.CTkLabel(
            self.header, text=title, font=("Segoe UI", 12, "bold"),
            text_color="#e2e8f0", anchor="w"
        )
        self.lbl_title.grid(row=0, column=1, sticky="ew", pady=8)

        # Tooltip-trigger: apart ⓘ icoontje naast de titel (klik toggle staat hier los van)
        tip_text = self._DESCRIPTIONS.get(title, "")
        if tip_text:
            lbl_info = ctk.CTkLabel(
                self.header, text="ⓘ", font=("Segoe UI", 12),
                text_color="#4a6fa5", width=20, cursor="question_arrow"
            )
            lbl_info.grid(row=0, column=2, padx=(0, 4), pady=8)
            Tooltip(lbl_info, tip_text)

        # Alles AAN/UIT knop (rechts in header)
        self.btn_toggle_all = ctk.CTkButton(
            self.header, text="Alles uit",
            font=("Segoe UI", 10), width=70, height=24,
            fg_color="#0f2540", hover_color="#1e3a5f",
            border_width=1, border_color="#334155",
            text_color="#94a3b8", corner_radius=4,
            command=self._toggle_all
        )
        self.btn_toggle_all.grid(row=0, column=3, padx=8, pady=6)
        Tooltip(self.btn_toggle_all, "Zet alle regels in deze dimensie tegelijk aan of uit.")

        # Klik op header → toggle uitklappen
        for w in (self.header, self.lbl_arrow, self.lbl_title):
            w.bind("<Button-1>", self.toggle)
            w.configure(cursor="hand2")

        # ---------- Inhoud (verborgen) ----------
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Initieel: zijn alle regels aan?
        self._update_toggle_label()

    def toggle(self, event=None):
        if self.is_open:
            self.content_frame.pack_forget()
            self.lbl_arrow.configure(text="►")
            self.is_open = False
        else:
            self.content_frame.pack(fill="x", padx=12, pady=(2, 8))
            self.lbl_arrow.configure(text="▼")
            self.is_open = True

    def _toggle_all(self):
        """Zet alle regels in één klap aan of uit."""
        # Als alles aan is → alles uit. Anders → alles aan.
        all_on = all(v.get() for v in self._rule_vars.values())
        new_val = not all_on
        for v in self._rule_vars.values():
            v.set(new_val)
        self._update_toggle_label()

    def _update_toggle_label(self):
        all_on = all(v.get() for v in self._rule_vars.values())
        if all_on:
            self.btn_toggle_all.configure(text="Alles uit", text_color="#94a3b8")
        else:
            self.btn_toggle_all.configure(text="Alles aan", text_color="#cf9d1f")


class HomeFrame(ctk.CTkFrame):
    """Hoofdscherm voor het instellen en starten van een scan."""

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="transparent")
        self.controller = controller

        self.selected_local_paths = set()
        self.selected_sharepoint_sites = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_frame()

    def update_ui_for_user(self):
        """Toon of verberg de afdeling-selector op basis van admin-rol."""
        user = self.controller.current_user
        if user and hasattr(user, 'is_admin') and user.is_admin:
            self.afdeling_frame.pack(fill="x", pady=(0, 25), ipadx=15, ipady=10, before=self.lbl_modules)
        else:
            self.afdeling_frame.pack_forget()

    # --------------------------------------------------------------------------
    # Sidebar
    # --------------------------------------------------------------------------

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=COLOR_BG_DEEP)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(4, weight=1)

        # Logo + titel bovenaan sidebar
        logo_title_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_title_frame.grid(row=0, column=0, padx=20, pady=(20, 20))

        logo_img = _load_logo(size=(45, 45))
        if logo_img:
            ctk.CTkLabel(logo_title_frame, image=logo_img, text="").pack(pady=(5, 5))

        self.logo_label = ctk.CTkLabel(
            logo_title_frame, text="DOCUMENT\nSCANNER",
            font=("Segoe UI Black", 24), text_color=COLOR_ACCENT
        )
        self.logo_label.pack()

        self.lbl_input = ctk.CTkLabel(
            self.sidebar, text="1. Selecteer Bronnen",
            font=("Segoe UI", 14, "bold"), text_color="white"
        )
        self.lbl_input.grid(row=1, column=0, padx=25, pady=(0, 10), sticky="w")

        self.btn_folder = ctk.CTkButton(
            self.sidebar, text="📁 + Lokaal Bestand/Map",
            font=("Segoe UI", 13, "bold"), height=40,
            command=self.open_drag_drop_window
        )
        self.btn_folder.grid(row=2, column=0, padx=25, pady=5, sticky="ew")

        self.btn_sp = ctk.CTkButton(
            self.sidebar, text="🌐 + SharePoint Site",
            font=("Segoe UI", 13, "bold"), height=40,
            command=self.add_sharepoint
        )
        self.btn_sp.grid(row=3, column=0, padx=25, pady=5, sticky="ew")

        self.source_list_frame = ctk.CTkScrollableFrame(
            self.sidebar, fg_color=COLOR_BG_LIGHT, corner_radius=10
        )
        self.source_list_frame.grid(row=4, column=0, padx=20, pady=15, sticky="nsew")
        self.update_source_list()

        self.btn_clear = ctk.CTkButton(
            self.sidebar, text="🗑️ Selectie Wissen",
            font=("Segoe UI", 13, "bold"), height=40,
            fg_color="#7f1d1d", hover_color="#991b1b", text_color="white",
            command=self.clear_selection
        )
        self.btn_clear.grid(row=5, column=0, padx=25, pady=(0, 15), sticky="ew")

        self.btn_logout = ctk.CTkButton(
            self.sidebar, text="🚪 Uitloggen",
            font=("Segoe UI", 13, "bold"), height=40,
            fg_color="transparent", border_width=1,
            border_color="#e2e8f0", text_color="#e2e8f0",
            hover_color=COLOR_BG_LIGHT,
            command=self.logout
        )
        self.btn_logout.grid(row=6, column=0, padx=25, pady=(0, 25), sticky="ew")

    def logout(self):
        self.controller.set_user(None)
        self.controller.show_frame("AuthFrame")

    # --------------------------------------------------------------------------
    # Hoofdpaneel
    # --------------------------------------------------------------------------

    def _build_main_frame(self):
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=40, pady=35)

        self.header_titel = ctk.CTkLabel(
            self.main_container, text="Scan Configuratie",
            font=("Segoe UI Black", 32), text_color="white"
        )
        self.header_titel.pack(anchor="w", pady=(0, 5))

        self.header_sub = ctk.CTkLabel(
            self.main_container,
            text="Stel je scan in volgens de actuele normen en kwaliteitseisen.",
            font=("Segoe UI", 15), text_color="#e2e8f0"
        )
        self.header_sub.pack(anchor="w", pady=(0, 25))

        # Afdeling-selector (alleen zichtbaar voor admins via update_ui_for_user)
        self.afdeling_frame = ctk.CTkFrame(
            self.main_container, fg_color=COLOR_BG_LIGHT, corner_radius=10
        )
        ctk.CTkLabel(
            self.afdeling_frame, text="Afdeling:",
            font=("Segoe UI", 14, "bold"), text_color=COLOR_ACCENT
        ).pack(side="left", padx=(10, 15))

        self.afdeling_var = ctk.StringVar(value="Alle Afdelingen")
        self.opt_afdeling = ctk.CTkOptionMenu(
            self.afdeling_frame,
            values=[
                "Alle Afdelingen",
                "J1: Personeel & Organisatie",
                "J2: Inlichtingen & Veiligheid",
                "J3: Operatiën (Current Ops)",
                "J5: Plannen (Plans)",
                "J6: Verbindings- en Informatiesystemen (CIS)",
                "J9: Civiel-Militaire Samenwerking (CIMIC)",
                "JMED (Medical)",
                "J-Legal (Juridische Zaken)",
                "IMO (Informatie Management Office)",
                "Public Affairs (Communicatie)"
            ],
            variable=self.afdeling_var,
            font=("Segoe UI", 14), width=300
        )
        self.opt_afdeling.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # Modules en Config Vars initializeren (backend datastructuur intact laten)
        self.comp_modules = {
            "Naamgeving":    ctk.BooleanVar(value=True),
            "Metadata":      ctk.BooleanVar(value=True),
            "Rubricering":   ctk.BooleanVar(value=True),
            "Bewaartermijn": ctk.BooleanVar(value=True),
        }
        self.qual_modules = {
            "Auteurs-validatie": ctk.BooleanVar(value=True),
            "Bestandsbody Check": ctk.BooleanVar(value=True),
            "Metadata (Titel) Check": ctk.BooleanVar(value=True),
            "Extensie-correlatie": ctk.BooleanVar(value=True),
            "Actualiteits-norm": ctk.BooleanVar(value=True),
            "Leesbaarheids-garantie": ctk.BooleanVar(value=True),
            "Geen vreemde tekens": ctk.BooleanVar(value=True),
            "Exacte duplicatie": ctk.BooleanVar(value=True),
            "Dode Snelkoppelingen": ctk.BooleanVar(value=True),
        }

        # Unified Governance Filter UI
        self.lbl_modules = ctk.CTkLabel(
            self.main_container, text="2. Unified Governance Filter",
            font=("Segoe UI", 18, "bold"), text_color=COLOR_ACCENT
        )
        self.lbl_modules.pack(anchor="w", pady=(0, 10))

        # Master Controls
        master_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        master_frame.pack(fill="x", pady=(0, 10))

        def toggle_comp():
            state = not all(v.get() for v in self.comp_modules.values())
            for v in self.comp_modules.values(): v.set(state)
            
        def toggle_qual():
            state = not all(v.get() for v in self.qual_modules.values())
            for v in self.qual_modules.values(): v.set(state)

        def reset_all():
            for v in list(self.comp_modules.values()) + list(self.qual_modules.values()):
                v.set(True)

        ctk.CTkButton(master_frame, text="🛡️ Compliance Toggle", fg_color=COLOR_ACCENT, text_color="#18181b", hover_color=COLOR_ACCENT_HOVER, font=("Segoe UI", 13, "bold"), command=toggle_comp).pack(side="left", padx=(0, 10))
        ctk.CTkButton(master_frame, text="✨ Kwaliteit Toggle", fg_color=COLOR_ACCENT, text_color="#18181b", hover_color=COLOR_ACCENT_HOVER, font=("Segoe UI", 13, "bold"), command=toggle_qual).pack(side="left", padx=(0, 10))
        ctk.CTkButton(master_frame, text="↺ Reset", fg_color="transparent", border_width=1, border_color="#334155", text_color="#94a3b8", hover_color="#162440", command=reset_all).pack(side="left")

        # Container for side-by-side lists
        rules_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        rules_container.pack(fill="both", expand=True, pady=(0, 20))
        rules_container.grid_columnconfigure(0, weight=1)
        rules_container.grid_columnconfigure(1, weight=1)

        # Left: Compliance
        comp_scroll = ctk.CTkScrollableFrame(rules_container, fg_color=COLOR_BG_DEEP, corner_radius=15, border_width=1, border_color=COLOR_ACCENT, height=280)
        comp_scroll.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # Right: Quality
        qual_scroll = ctk.CTkScrollableFrame(rules_container, fg_color=COLOR_BG_DEEP, corner_radius=15, border_width=1, border_color=COLOR_ACCENT, height=280)
        qual_scroll.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        rule_tooltips = {
            "Naamgeving": "De bestandslocatie of naam moet voldoen aan de vastgestelde kaders.",
            "Metadata": "Bestanden moeten bepaalde basismetadata bevatten.",
            "Rubricering": "Data moet de juiste classificatie of rubricering dragen.",
            "Bewaartermijn": "Data mag niet te lang bewaard worden (bewaartermijnbeleid).",
            "Auteurs-validatie": "Controleert of de auteur-eigenschappen van bestanden ingevuld zijn.",
            "Bestandsbody Check": "Checkt of het bestand extreem klein of leeg is (<1KB).",
            "Metadata (Titel) Check": "Controleert op bestanden zonder correcte Titel-metadata.",
            "Extensie-correlatie": "Valideert of de bestandsextensie klopt met de effectieve encoding.",
            "Actualiteits-norm": "Signaleert bestanden die buitengewoon lang niet zijn gewijzigd.",
            "Leesbaarheids-garantie": "Spoort vergrendelde of niet-leessbare (corrupte) bestanden op.",
            "Geen vreemde tekens": "Controleert bestands- of mapnamen op niet-toegestane tekens.",
            "Exacte duplicatie": "Vindt dubbele bestanden op basis van een exacte inhoud-match (hash).",
            "Dode Snelkoppelingen": "Detecteert .lnk of .url bestanden waarvan de doellocatie niet meer bestaat."
        }

        # Populate Compliance (Left)
        for name, var in self.comp_modules.items():
            row = ctk.CTkFrame(comp_scroll, fg_color="#14253e", corner_radius=6)
            row.pack(fill="x", padx=5, pady=3)
            row.grid_columnconfigure(2, weight=1)
            
            cb = ctk.CTkCheckBox(row, text="", variable=var, width=30, checkmark_color="#cf9d1f", border_color="#3b82f6")
            cb.grid(row=0, column=0, padx=(5, 5), pady=8)
            
            lbl_icon = ctk.CTkLabel(row, text="🛡️", font=("Segoe UI", 16))
            lbl_icon.grid(row=0, column=1, padx=(0, 5))
            
            lbl_name = ctk.CTkLabel(row, text=name, font=("Segoe UI", 12, "bold"), text_color="#e2e8f0", anchor="w")
            lbl_name.grid(row=0, column=2, sticky="w")
            
            lbl_info = ctk.CTkLabel(row, text="ⓘ", font=("Segoe UI", 12, "bold"), text_color="#64748b", cursor="question_arrow")
            lbl_info.grid(row=0, column=3, padx=5)
            Tooltip(lbl_info, rule_tooltips.get(name, "Geen details beschikbaar."))

        # Populate Quality (Right)
        for name, var in self.qual_modules.items():
            row = ctk.CTkFrame(qual_scroll, fg_color="#0e1b2f", corner_radius=6)
            row.pack(fill="x", padx=5, pady=3)
            row.grid_columnconfigure(2, weight=1)
            
            cb = ctk.CTkCheckBox(row, text="", variable=var, width=30, checkmark_color="#cf9d1f", border_color="#3b82f6")
            cb.grid(row=0, column=0, padx=(5, 5), pady=8)
            
            lbl_icon = ctk.CTkLabel(row, text="✨", font=("Segoe UI", 16))
            lbl_icon.grid(row=0, column=1, padx=(0, 5))
            
            lbl_name = ctk.CTkLabel(row, text=name, font=("Segoe UI", 12, "bold"), text_color="#e2e8f0", anchor="w")
            lbl_name.grid(row=0, column=2, sticky="w")
            
            lbl_info = ctk.CTkLabel(row, text="ⓘ", font=("Segoe UI", 12, "bold"), text_color="#64748b", cursor="question_arrow")
            lbl_info.grid(row=0, column=3, padx=5)
            Tooltip(lbl_info, rule_tooltips.get(name, "Geen details beschikbaar."))



        # Start-knop
        self.btn_analyze = ctk.CTkButton(
            self.main_container, text="▶ START ANALYSE",
            font=("Segoe UI Black", 18), height=60, corner_radius=10,
            text_color="#18181b", command=self.start_analysis_request
        )
        self.btn_analyze.pack(fill="x", pady=(10, 0))

    # --------------------------------------------------------------------------
    # Bronnenlijst
    # --------------------------------------------------------------------------

    def update_source_list(self):
        if not self.winfo_exists():
            return

        for widget in self.source_list_frame.winfo_children():
            widget.destroy()

        if not self.selected_local_paths and not self.selected_sharepoint_sites:
            ctk.CTkLabel(
                self.source_list_frame,
                text="Nog geen bronnen geselecteerd.",
                text_color="#e2e8f0", font=("Segoe UI", 12, "italic")
            ).pack(pady=20)
            return

        for path in self.selected_local_paths:
            map_naam = os.path.basename(path) or path
            icoon = "📄" if os.path.isfile(path) else "📁"
            ctk.CTkLabel(
                self.source_list_frame,
                text=f"{icoon} {map_naam}",
                font=("Segoe UI", 13), anchor="w"
            ).pack(fill="x", pady=5, padx=5)

        for sp in self.selected_sharepoint_sites:
            sp_naam = sp['url'].replace("https://", "").split("/")[0]
            ctk.CTkLabel(
                self.source_list_frame,
                text=f"🌐 {sp_naam}",
                font=("Segoe UI", 13), anchor="w"
            ).pack(fill="x", pady=5, padx=5)

    # --------------------------------------------------------------------------
    # Bestand/map toevoegen
    # --------------------------------------------------------------------------

    def open_drag_drop_window(self):
        drop_win = ctk.CTkToplevel(self)
        drop_win.title("Voeg Map of Bestand toe")
        drop_win.geometry("480x320")
        drop_win.transient(self.controller)
        drop_win.grab_set()
        drop_win.focus_force()

        drop_frame = ctk.CTkFrame(
            drop_win, corner_radius=15,
            border_width=2, border_color=COLOR_ACCENT,
            fg_color=COLOR_BG_DEEP
        )
        drop_frame.pack(fill="both", expand=True, padx=25, pady=25)

        lbl = ctk.CTkLabel(
            drop_frame,
            text="📥 Sleep bestanden of mappen hierheen\n\nOf kies via de verkenner:",
            font=("Segoe UI", 15, "bold")
        )
        lbl.pack(pady=(45, 25))

        btn_frame = ctk.CTkFrame(drop_frame, fg_color="transparent")
        btn_frame.pack()

        ctk.CTkButton(
            btn_frame, text="📁 Kies Map",
            font=("Segoe UI", 12, "bold"), width=130,
            command=lambda: self.browse_folder(drop_win)
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            btn_frame, text="📄 Kies Bestand",
            font=("Segoe UI", 12, "bold"), width=130,
            command=lambda: self.browse_file(drop_win)
        ).pack(side="left", padx=10)

        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', lambda e: self.handle_drop(e, drop_win))
        lbl.drop_target_register(DND_FILES)
        lbl.dnd_bind('<<Drop>>', lambda e: self.handle_drop(e, drop_win))

    def handle_drop(self, event, window):
        paths = self.controller.tk.splitlist(event.data)
        for p in paths:
            self.selected_local_paths.add(p)
        self.update_source_list()
        window.destroy()

    def browse_folder(self, window):
        folder = filedialog.askdirectory(parent=window)
        if folder:
            self.selected_local_paths.add(folder)
            self.update_source_list()
            window.destroy()

    def browse_file(self, window):
        files = filedialog.askopenfilenames(parent=window)
        if files:
            for f in files:
                self.selected_local_paths.add(f)
            self.update_source_list()
            window.destroy()

    # --------------------------------------------------------------------------
    # SharePoint toevoegen
    # --------------------------------------------------------------------------

    def add_sharepoint(self):
        sp_window = ctk.CTkToplevel(self)
        sp_window.title("SharePoint Connectie")
        sp_window.geometry("420x250")
        sp_window.transient(self.controller)
        sp_window.grab_set()
        sp_window.focus_force()

        ctk.CTkLabel(
            sp_window, text="SharePoint Link",
            font=("Segoe UI", 16, "bold"), text_color=COLOR_ACCENT
        ).pack(pady=(20, 10))

        entry_url = ctk.CTkEntry(sp_window, width=320)
        entry_url.pack(pady=5)

        def save_sp():
            url = entry_url.get().strip()
            if url:
                self.selected_sharepoint_sites.append({"url": url, "library": "Documents"})
                self.update_source_list()
                sp_window.destroy()

        ctk.CTkButton(
            sp_window, text="Verbinden",
            font=("Segoe UI", 13, "bold"), text_color="#18181b",
            command=save_sp
        ).pack(pady=20)

    # --------------------------------------------------------------------------
    # Selectie beheren & scan starten
    # --------------------------------------------------------------------------

    def clear_selection(self):
        self.selected_local_paths.clear()
        self.selected_sharepoint_sites.clear()
        self.update_source_list()

    def start_analysis_request(self):
        if not self.selected_local_paths and not self.selected_sharepoint_sites:
            messagebox.showwarning("Data Fout", "Selecteer minimaal één bron om te scannen.")
            return

        active_comp = [key for key, var in self.comp_modules.items() if var.get()]
        active_qual = [key for key, var in self.qual_modules.items() if var.get()]

        if not active_comp and not active_qual:
            messagebox.showwarning("Configuratie Fout", "Selecteer minimaal één module of kwaliteitsregel.")
            return

        user = self.controller.current_user
        afdeling = (
            self.afdeling_var.get()
            if user and hasattr(user, 'is_admin') and user.is_admin
            else "Standaard"
        )

        self.controller.start_analysis(
            afdeling=afdeling,
            local_paths=list(self.selected_local_paths),
            sharepoint_sites=self.selected_sharepoint_sites,
            active_comp=active_comp,
            active_qual=active_qual
        )
