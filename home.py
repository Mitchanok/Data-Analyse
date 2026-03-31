import customtkinter as ctk
import os
from PIL import Image
from centrale_engine import User

COLOR_PASS = "#10b981"
COLOR_FAIL = "#ef4444"
COLOR_WARN = "#f59e0b"
COLOR_ACCENT = "#2563eb" 
COLOR_BG_DEEP = "#0f172a" 
COLOR_BG_LIGHT = "#1e293b"

def _load_logo(size=(50, 50)):
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    try:
        return ctk.CTkImage(light_image=Image.open(logo_path), dark_image=Image.open(logo_path), size=size)
    except Exception:
        return None

class AuthFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, fg_color=COLOR_BG_DEEP)
        self.controller = controller
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Logo linksboven
        logo_img = _load_logo()
        if logo_img:
            self.logo_lbl = ctk.CTkLabel(self, image=logo_img, text="")
            self.logo_lbl.place(x=15, y=10)
        
        self.login_box = ctk.CTkFrame(self, corner_radius=15, fg_color=COLOR_BG_LIGHT, width=400, height=500)
        self.login_box.grid(row=0, column=0)
        self.login_box.grid_propagate(False)
        
        ctk.CTkLabel(self.login_box, text="Welkom bij", font=("Segoe UI", 16)).pack(pady=(40, 5))
        ctk.CTkLabel(self.login_box, text="DOCUMENT SCANNER", font=("Segoe UI Black", 24), text_color=COLOR_ACCENT).pack(pady=(0, 40))
        
        self.username_entry = ctk.CTkEntry(self.login_box, placeholder_text="Gebruikersnaam", width=300, height=45, font=("Segoe UI", 14))
        self.username_entry.pack(pady=10)
        
        self.password_entry = ctk.CTkEntry(self.login_box, placeholder_text="Wachtwoord", show="*", width=300, height=45, font=("Segoe UI", 14))
        self.password_entry.pack(pady=10)
        
        self.login_btn = ctk.CTkButton(self.login_box, text="Inloggen", font=("Segoe UI Black", 16), width=300, height=50, command=self.login)
        self.login_btn.pack(pady=(20, 10))
        
        self.guest_btn = ctk.CTkButton(self.login_box, text="Doorgaan als Gast", font=("Segoe UI", 14), fg_color="transparent", border_width=1, border_color=COLOR_ACCENT, text_color=COLOR_ACCENT, width=300, height=40, command=self.login_guest)
        self.guest_btn.pack(pady=(0, 20))

    def login_guest(self):
        user = User("Gast Gebruiker", False)
        user.id = 0
        self.controller.set_user(user)
        self.controller.show_frame("HomeFrame")
        self.password_entry.delete(0, 'end')

    def login(self):
        import centrale_engine as database
        from tkinter import messagebox
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username or not password:
            messagebox.showerror("Fout", "Vul beide velden in.")
            return

        db_user = database.verify_user(username, password)
        if db_user:
            user = User(db_user["username"], db_user["is_admin"])
            user.id = db_user["id"]
            self.controller.set_user(user)
            self.controller.show_frame("HomeFrame")
            self.password_entry.delete(0, 'end')
        else:
            messagebox.showerror("Fout", "Ongeldige inloggegevens. Standaard admin is admin/admin.")

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from tkinterdnd2 import DND_FILES

COLOR_PASS = "#10b981"
COLOR_FAIL = "#ef4444"
COLOR_WARN = "#f59e0b"
COLOR_ACCENT = "#2563eb" 
COLOR_BG_DEEP = "#0f172a" 
COLOR_BG_LIGHT = "#1e293b" 

class HomeFrame(ctk.CTkFrame):
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
        user = self.controller.current_user
        if user and hasattr(user, 'is_admin') and user.is_admin:
            self.afdeling_frame.pack(fill="x", pady=(0, 25), ipadx=15, ipady=10, before=self.lbl_modules)
        else:
            self.afdeling_frame.pack_forget()

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=COLOR_BG_DEEP)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(4, weight=1) 
        
        # Logo + Titel container
        logo_title_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_title_frame.grid(row=0, column=0, padx=20, pady=(20, 20))
        
        logo_img = _load_logo(size=(45, 45))
        if logo_img:
            ctk.CTkLabel(logo_title_frame, image=logo_img, text="").pack(pady=(5, 5))
        
        self.logo_label = ctk.CTkLabel(logo_title_frame, text="DOCUMENT\nSCANNER", font=("Segoe UI Black", 24), text_color=COLOR_ACCENT)
        self.logo_label.pack()
        
        self.lbl_input = ctk.CTkLabel(self.sidebar, text="1. Selecteer Bronnen", font=("Segoe UI", 14, "bold"), text_color="white")
        self.lbl_input.grid(row=1, column=0, padx=25, pady=(0, 10), sticky="w")
        
        self.btn_folder = ctk.CTkButton(self.sidebar, text="📁 + Lokaal Bestand/Map", font=("Segoe UI", 13, "bold"), height=40, command=self.open_drag_drop_window)
        self.btn_folder.grid(row=2, column=0, padx=25, pady=5, sticky="ew")
        
        self.btn_sp = ctk.CTkButton(self.sidebar, text="🌐 + SharePoint Site", font=("Segoe UI", 13, "bold"), height=40, command=self.add_sharepoint)
        self.btn_sp.grid(row=3, column=0, padx=25, pady=5, sticky="ew")

        self.source_list_frame = ctk.CTkScrollableFrame(self.sidebar, fg_color=COLOR_BG_LIGHT, corner_radius=10)
        self.source_list_frame.grid(row=4, column=0, padx=20, pady=15, sticky="nsew")
        self.update_source_list()
        
        self.btn_clear = ctk.CTkButton(self.sidebar, text="🗑️ Selectie Wissen", font=("Segoe UI", 13, "bold"), height=40, fg_color="#7f1d1d", hover_color="#991b1b", text_color="white", command=self.clear_selection)
        self.btn_clear.grid(row=5, column=0, padx=25, pady=(0, 15), sticky="ew")

        self.btn_logout = ctk.CTkButton(self.sidebar, text="🚪 Uitloggen", font=("Segoe UI", 13, "bold"), height=40, fg_color="transparent", border_width=1, border_color="#e2e8f0", text_color="#e2e8f0", hover_color="#1a2b4b", command=self.logout)
        self.btn_logout.grid(row=6, column=0, padx=25, pady=(0, 25), sticky="ew")

    def logout(self):
        self.controller.set_user(None)
        self.controller.show_frame("AuthFrame")

    def _build_main_frame(self):
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=40, pady=35)
        
        self.header_titel = ctk.CTkLabel(self.main_container, text="Scan Configuratie", font=("Segoe UI Black", 32), text_color="white")
        self.header_titel.pack(anchor="w", pady=(0, 5))
        
        self.header_sub = ctk.CTkLabel(self.main_container, text="Stel je scan in volgens de actuele normen en kwaliteitseisen.", font=("Segoe UI", 15), text_color="#e2e8f0")
        self.header_sub.pack(anchor="w", pady=(0, 25))

        # Afdeling Frame (ipv Projectnaam)
        self.afdeling_frame = ctk.CTkFrame(self.main_container, fg_color=COLOR_BG_LIGHT, corner_radius=10)
        # Pak wordt geregeld in update_ui_for_user()
        
        ctk.CTkLabel(self.afdeling_frame, text="Afdeling:", font=("Segoe UI", 14, "bold"), text_color=COLOR_ACCENT).pack(side="left", padx=(10, 15))
        self.afdeling_var = ctk.StringVar(value="Alle Afdelingen")
        self.opt_afdeling = ctk.CTkOptionMenu(self.afdeling_frame, values=["Alle Afdelingen", "HR", "Finance", "IT", "Directie"], variable=self.afdeling_var, font=("Segoe UI", 14), width=300)
        self.opt_afdeling.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # Engines / Modules
        self.lbl_modules = ctk.CTkLabel(self.main_container, text="2. Actieve Engines & Modules", font=("Segoe UI", 18, "bold"), text_color=COLOR_ACCENT)
        self.lbl_modules.pack(anchor="w", pady=(0, 10))

        self.engines_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.engines_frame.pack(fill="x", pady=(0, 30))

        self.comp_modules = {
            "Naamgeving": ctk.BooleanVar(value=True),
            "Metadata": ctk.BooleanVar(value=True),
            "Rubricering": ctk.BooleanVar(value=True),
            "Bewaartermijn": ctk.BooleanVar(value=True)
        }
        self.qual_modules = {
            "Bestandsgrootte": ctk.BooleanVar(value=True),
            "Actualiteit": ctk.BooleanVar(value=True),
            "Leesbaarheid": ctk.BooleanVar(value=True),
            "Volledigheid": ctk.BooleanVar(value=True)
        }

        # Compliance Frame
        self.comp_frame = ctk.CTkFrame(self.engines_frame, fg_color=COLOR_BG_DEEP, corner_radius=15, border_width=1, border_color=COLOR_ACCENT)
        self.comp_frame.pack(side="left", fill="both", expand=True, padx=(0, 10), ipadx=20, ipady=15)
        ctk.CTkLabel(self.comp_frame, text="🛡️ Compliance Engine", font=("Segoe UI", 16, "bold"), text_color="white").pack(anchor="w", pady=(0, 10))
        for naam, var in self.comp_modules.items():
            ctk.CTkCheckBox(self.comp_frame, text=naam, variable=var, font=("Segoe UI", 15)).pack(anchor="w", pady=8, padx=10)

        # Quality Frame
        self.qual_frame = ctk.CTkFrame(self.engines_frame, fg_color=COLOR_BG_DEEP, corner_radius=15, border_width=1, border_color=COLOR_ACCENT)
        self.qual_frame.pack(side="left", fill="both", expand=True, padx=(10, 0), ipadx=20, ipady=15)
        ctk.CTkLabel(self.qual_frame, text="✨ Data Quality Engine", font=("Segoe UI", 16, "bold"), text_color="white").pack(anchor="w", pady=(0, 10))
        for naam, var in self.qual_modules.items():
            ctk.CTkCheckBox(self.qual_frame, text=naam, variable=var, font=("Segoe UI", 15)).pack(anchor="w", pady=8, padx=10)

        self.btn_analyze = ctk.CTkButton(self.main_container, text="▶ START ANALYSE", font=("Segoe UI Black", 18), height=60, corner_radius=10, text_color="#18181b")
        self.btn_analyze.configure(command=self.start_analysis_request)
        self.btn_analyze.pack(fill="x", pady=(10, 0))

    def update_source_list(self):
        if not self.winfo_exists(): return
        
        for widget in self.source_list_frame.winfo_children():
            widget.destroy()

        if not self.selected_local_paths and not self.selected_sharepoint_sites:
            ctk.CTkLabel(self.source_list_frame, text="Nog geen bronnen geselecteerd.", text_color="#e2e8f0", font=("Segoe UI", 12, "italic")).pack(pady=20)
            return

        for path in self.selected_local_paths:
            map_naam = os.path.basename(path) or path
            icoon = "📄" if os.path.isfile(path) else "📁"
            ctk.CTkLabel(self.source_list_frame, text=f"{icoon} {map_naam}", font=("Segoe UI", 13), anchor="w").pack(fill="x", pady=5, padx=5)

        for sp in self.selected_sharepoint_sites:
            sp_naam = sp['url'].replace("https://", "").split("/")[0] 
            ctk.CTkLabel(self.source_list_frame, text=f"🌐 {sp_naam}", font=("Segoe UI", 13), anchor="w").pack(fill="x", pady=5, padx=5)

    def open_drag_drop_window(self):
        drop_win = ctk.CTkToplevel(self)
        drop_win.title("Voeg Map of Bestand toe")
        drop_win.geometry("480x320")
        
        drop_win.transient(self.controller) 
        drop_win.grab_set() 
        drop_win.focus_force()
        
        drop_frame = ctk.CTkFrame(drop_win, corner_radius=15, border_width=2, border_color=COLOR_ACCENT, fg_color=COLOR_BG_DEEP)
        drop_frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        lbl = ctk.CTkLabel(drop_frame, text="📥 Sleep bestanden of mappen hierheen\n\nOf kies via de verkenner:", font=("Segoe UI", 15, "bold"))
        lbl.pack(pady=(45, 25))
        
        btn_frame = ctk.CTkFrame(drop_frame, fg_color="transparent")
        btn_frame.pack()
        
        ctk.CTkButton(btn_frame, text="📁 Kies Map", font=("Segoe UI", 12, "bold"), width=130, command=lambda: self.browse_folder(drop_win)).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="📄 Kies Bestand", font=("Segoe UI", 12, "bold"), width=130, command=lambda: self.browse_file(drop_win)).pack(side="left", padx=10)
        
        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', lambda e: self.handle_drop(e, drop_win))
        lbl.drop_target_register(DND_FILES)
        lbl.dnd_bind('<<Drop>>', lambda e: self.handle_drop(e, drop_win))

    def handle_drop(self, event, window):
        paths = self.controller.tk.splitlist(event.data)
        for p in paths: self.selected_local_paths.add(p)
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
            for f in files: self.selected_local_paths.add(f)
            self.update_source_list()
            window.destroy()

    def add_sharepoint(self):
        sp_window = ctk.CTkToplevel(self)
        sp_window.title("SharePoint Connectie")
        sp_window.geometry("420x250")
        
        sp_window.transient(self.controller)
        sp_window.grab_set()
        sp_window.focus_force()
        
        ctk.CTkLabel(sp_window, text="SharePoint Link", font=("Segoe UI", 16, "bold"), text_color=COLOR_ACCENT).pack(pady=(20, 10))
        entry_url = ctk.CTkEntry(sp_window, width=320)
        entry_url.pack(pady=5)
        
        def save_sp():
            url = entry_url.get().strip()
            if url:
                self.selected_sharepoint_sites.append({"url": url, "library": "Documents"}) 
                self.update_source_list()
                sp_window.destroy()
        
        ctk.CTkButton(sp_window, text="Verbinden", font=("Segoe UI", 13, "bold"), text_color="#18181b", command=save_sp).pack(pady=20)

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
        afdeling = self.afdeling_var.get() if user and hasattr(user, 'is_admin') and user.is_admin else "Standaard"
        
        self.controller.start_analysis(
            afdeling=afdeling,
            local_paths=list(self.selected_local_paths),
            sharepoint_sites=self.selected_sharepoint_sites,
            active_comp=active_comp,
            active_qual=active_qual
        )
