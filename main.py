import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
import csv
import os
from datetime import datetime
from tkinterdnd2 import TkinterDnD, DND_FILES 
from compliance_engine import ComplianceEngine

# WCAG-Compliant kleuren (Geïntegreerd met jouw thema)
COLOR_PASS = "#388E3C"
COLOR_FAIL = "#D32F2F"
COLOR_WARN = "#F57C00"
COLOR_ACCENT = "#cf9d1f" 
COLOR_BG_DEEP = "#001538" 
COLOR_BG_LIGHT = "#1a2b4b" 

try:
    ctk.set_default_color_theme("theme/gold_blue_theme.json")
except FileNotFoundError:
    ctk.set_default_color_theme("blue") 

ctk.set_appearance_mode("Dark")

class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class ComplianceApp(TkinterDnD_CTk):
    def __init__(self):
        super().__init__()
        self.title("Compliance Analyzer Pro V1.0 - Enterprise Edition")
        self.geometry("1100x750") 
        self.minsize(950, 650)
        self.configure(fg_color=COLOR_BG_DEEP) 
        
        self.selected_local_paths = set()
        self.selected_sharepoint_sites = [] 
        self.q = queue.Queue()
        self.analysis_data = None
        self.is_analyzing = False 
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_frame()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.destroy() 

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=COLOR_BG_DEEP)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(4, weight=1) 

        self.logo_label = ctk.CTkLabel(self.sidebar, text="COMPLIANCE\nANALYZER", font=("Segoe UI Black", 24), text_color=COLOR_ACCENT)
        self.logo_label.grid(row=0, column=0, padx=20, pady=(35, 30))
        
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
        self.btn_clear.grid(row=5, column=0, padx=25, pady=(0, 25), sticky="ew")

    def _build_main_frame(self):
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=40, pady=35)
        
        self.header_titel = ctk.CTkLabel(self.main_frame, text="Scan Configuratie", font=("Segoe UI Black", 32), text_color="white")
        self.header_titel.pack(anchor="w", pady=(0, 5))
        
        self.header_sub = ctk.CTkLabel(self.main_frame, text="Stel je compliance scan in volgens de actuele normen.", font=("Segoe UI", 15), text_color="#e2e8f0")
        self.header_sub.pack(anchor="w", pady=(0, 25))

        self.project_frame = ctk.CTkFrame(self.main_frame, fg_color=COLOR_BG_LIGHT, corner_radius=10)
        self.project_frame.pack(fill="x", pady=(0, 25), ipadx=15, ipady=10)
        
        ctk.CTkLabel(self.project_frame, text="Projectnaam:", font=("Segoe UI", 14, "bold"), text_color=COLOR_ACCENT).pack(side="left", padx=(10, 15))
        self.entry_project = ctk.CTkEntry(self.project_frame, font=("Segoe UI", 14), placeholder_text="Geef test naam", width=400)
        self.entry_project.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.lbl_modules = ctk.CTkLabel(self.main_frame, text="2. Actieve Compliance Modules", font=("Segoe UI", 18, "bold"), text_color=COLOR_ACCENT)
        self.lbl_modules.pack(anchor="w", pady=(0, 10))

        self.check_frame = ctk.CTkFrame(self.main_frame, fg_color=COLOR_BG_DEEP, corner_radius=15, border_width=1, border_color=COLOR_ACCENT)
        self.check_frame.pack(fill="x", pady=(0, 30), ipadx=20, ipady=15)

        self.modules = {
            "Naamgeving": ctk.BooleanVar(value=True),
            "Metadata": ctk.BooleanVar(value=True),
            "Rubricering": ctk.BooleanVar(value=True),
            "Bewaartermijn": ctk.BooleanVar(value=True)
        }

        idx = 0
        for naam, var in self.modules.items():
            cb = ctk.CTkCheckBox(self.check_frame, text=naam, variable=var, font=("Segoe UI", 15), checkbox_width=24, checkbox_height=24)
            cb.grid(row=idx // 2, column=idx % 2, padx=40, pady=15, sticky="w")
            idx += 1
            
        self.check_frame.grid_columnconfigure(0, weight=1)
        self.check_frame.grid_columnconfigure(1, weight=1)

        self.btn_analyze = ctk.CTkButton(self.main_frame, text="▶ START ANALYSE", font=("Segoe UI Black", 18), height=60, corner_radius=10, text_color="#18181b")
        self.btn_analyze.configure(command=self.start_analysis)
        self.btn_analyze.pack(fill="x", pady=(10, 0))

        self.progress = ctk.CTkProgressBar(self.main_frame, height=12)
        self.progress.set(0)

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
        
        drop_win.transient(self) 
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
        paths = self.tk.splitlist(event.data)
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
        
        sp_window.transient(self)
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

    def start_analysis(self):
        if self.is_analyzing: return
        
        if not self.selected_local_paths and not self.selected_sharepoint_sites:
            messagebox.showwarning("Data Fout", "Selecteer minimaal één bron om te scannen.")
            return
        
        active_modules = [key for key, var in self.modules.items() if var.get()]
        if not active_modules:
            messagebox.showwarning("Configuratie Fout", "Selecteer minimaal één module.")
            return

        self.is_analyzing = True
        self.btn_analyze.configure(state="disabled", text="⏳ ANALYSEREN...")
        self.progress.pack(fill="x", pady=20)
        self.progress.set(0)
        
        engine = ComplianceEngine(list(self.selected_local_paths), self.selected_sharepoint_sites, active_modules)
        threading.Thread(target=engine.process, args=(self.q,), daemon=True).start()
        self.check_queue()

    def check_queue(self):
        if not self.winfo_exists(): return 
        
        try:
            while True:
                msg_type, data = self.q.get_nowait()
                if msg_type == "progress": 
                    self.progress.set(data)
                elif msg_type == "error":
                    messagebox.showerror("QA Systeemfout", data)
                    self.reset_ui()
                    return
                elif msg_type == "done":
                    self.analysis_data = data
                    self.show_dashboard()
                    return 
        except queue.Empty: pass 
        
        if self.is_analyzing:
            self.after(100, self.check_queue)

    def reset_ui(self):
        if not self.winfo_exists(): return
        self.is_analyzing = False
        self.progress.pack_forget()
        self.btn_analyze.configure(state="normal", text="▶ START ANALYSE")

    def _calc_average(self, domain_dict):
        if not domain_dict: return -1
        
        valid_scores = []
        for scores in domain_dict.values():
            if not scores: continue
            
            # Zet de scores (die mogelijk tekst of procenten zijn) veilig om naar echte getallen
            numeric_scores = []
            for s in scores:
                if isinstance(s, str):
                    # Haal eventuele % tekens weg voor de zekerheid
                    s = s.replace('%', '').strip()
                try:
                    numeric_scores.append(float(s))
                except ValueError:
                    pass # Negeer foute/lege waardes zoals "N/A"
            
            if numeric_scores:
                valid_scores.append(sum(numeric_scores) / len(numeric_scores))
                
        return sum(valid_scores) / len(valid_scores) if valid_scores else -1

    def show_dashboard(self):
        self.reset_ui()
        if not self.winfo_exists(): return
        
        project_naam = self.entry_project.get().strip() or "Naamloze_Scan"
        
        local_avg = self._calc_average(self.analysis_data.get("domain_scores_local", {}))
        sp_avg = self._calc_average(self.analysis_data.get("domain_scores_sp", {}))
        
        totaal_scores = [s for s in [local_avg, sp_avg] if s != -1]
        totaal_avg = sum(totaal_scores) / len(totaal_scores) if totaal_scores else -1
        
        dash = ctk.CTkToplevel(self)
        dash.title(f"Compliance Rapport - {project_naam}")
        dash.geometry("950x850")
        dash.configure(fg_color=COLOR_BG_DEEP) 
        
        dash.transient(self)
        dash.grab_set()
        dash.focus_force()
        
        head = ctk.CTkFrame(dash, height=100, corner_radius=0, fg_color=COLOR_BG_DEEP)
        head.pack(fill="x")
        ctk.CTkLabel(head, text=f"Rapport: {project_naam}", font=("Segoe UI", 18, "bold"), text_color="white").pack(pady=20)

        nav_frame = ctk.CTkFrame(dash, fg_color="transparent")
        nav_frame.pack(fill="x", pady=10, padx=20)

        self.view_total = ctk.CTkFrame(dash, fg_color="transparent")
        self.view_sp = ctk.CTkFrame(dash, fg_color="transparent")
        self.view_local = ctk.CTkFrame(dash, fg_color="transparent")

        def switch_tab(btn_active, view_active):
            for btn in [btn_total, btn_sp, btn_local]:
                btn.configure(fg_color=COLOR_BG_LIGHT, text_color="white")
            btn_active.configure(fg_color=COLOR_ACCENT, text_color="black")
            
            self.view_total.pack_forget()
            self.view_sp.pack_forget()
            self.view_local.pack_forget()
            view_active.pack(fill="both", expand=True, padx=20, pady=5)

        btn_total = ctk.CTkButton(nav_frame, text="Totaal Overzicht", font=("Segoe UI", 14, "bold"), command=lambda: switch_tab(btn_total, self.view_total))
        btn_total.pack(side="left", expand=True, padx=5)

        btn_sp = ctk.CTkButton(nav_frame, text="SharePoint", font=("Segoe UI", 14, "bold"), command=lambda: switch_tab(btn_sp, self.view_sp))
        btn_sp.pack(side="left", expand=True, padx=5)

        btn_local = ctk.CTkButton(nav_frame, text="Lokale Schijf", font=("Segoe UI", 14, "bold"), command=lambda: switch_tab(btn_local, self.view_local))
        btn_local.pack(side="left", expand=True, padx=5)

        risico_bestanden = [res["Naam"] for res in self.analysis_data.get("results", []) if "🚨 KRITIEK" in res.get("Reden", "")]
        
        if risico_bestanden:
            alert_frame = ctk.CTkFrame(dash, fg_color="#7f1d1d", corner_radius=10, border_width=2, border_color="#ef4444")
            alert_frame.pack(fill="x", padx=20, pady=(15, 0))
            
            ctk.CTkLabel(alert_frame, text="🚨 KRITIEKE BEVEILIGINGSWAARSCHUWING 🚨", font=("Segoe UI Black", 18), text_color="white").pack(pady=(15, 5))
            ctk.CTkLabel(alert_frame, text="De volgende schadelijke bestanden zijn gevonden en blokkeren compliance. DIRECT VERWIJDEREN:", font=("Segoe UI", 14, "bold"), text_color="#fca5a5").pack(pady=0)
            
            toon_bestanden = risico_bestanden[:5]
            files_text = "\n".join([f"• {f}" for f in toon_bestanden])
            if len(risico_bestanden) > 5:
                files_text += f"\n... en nog {len(risico_bestanden) - 5} andere(n). Zie CSV export."
                
            ctk.CTkLabel(alert_frame, text=files_text, font=("Consolas", 14), text_color="white", justify="left").pack(pady=(10, 15), padx=20)

        all_keys = set(self.analysis_data["domain_scores_local"].keys()) | set(self.analysis_data["domain_scores_sp"].keys())
        combined_domains = {mod: [] for mod in all_keys}
        
        for mod in all_keys:
            combined_domains[mod].extend(self.analysis_data["domain_scores_local"].get(mod, []))
            combined_domains[mod].extend(self.analysis_data["domain_scores_sp"].get(mod, []))

        self._build_tab_content(self.view_total, totaal_avg, combined_domains)
        self._build_tab_content(self.view_sp, sp_avg, self.analysis_data.get("domain_scores_sp", {}))
        self._build_tab_content(self.view_local, local_avg, self.analysis_data.get("domain_scores_local", {}))

        switch_tab(btn_total, self.view_total)

        ctk.CTkButton(dash, text="📥 Exporteer Data (CSV)", font=("Segoe UI Black", 14), text_color="#18181b", height=50, command=self.export_to_csv).pack(fill="x", padx=20, pady=20)

    def _get_module_reasons(self, module_name):
        reasons_found = set()
        keyword = module_name.split()[0] 
        
        for res in self.analysis_data.get("results", []):
            reden_string = res.get("Reden", "")
            if keyword in reden_string:
                parts = reden_string.split(" | ")
                for part in parts:
                    if keyword in part:
                        reasons_found.add(part)
        return list(reasons_found)

    def _build_tab_content(self, parent_frame, avg_score, domain_dict):
        if avg_score == -1:
            ctk.CTkLabel(parent_frame, text="Geen documenten geanalyseerd in deze bron.", font=("Segoe UI", 16, "italic"), text_color="#e2e8f0").pack(pady=50)
            return

        kleur_score = COLOR_PASS if avg_score >= 75 else (COLOR_WARN if avg_score >= 50 else COLOR_FAIL)
        ctk.CTkLabel(parent_frame, text=f"Score: {avg_score:.1f}%", font=("Segoe UI Black", 50), text_color=kleur_score).pack(pady=(5, 5))
        
        body = ctk.CTkScrollableFrame(parent_frame, fg_color="transparent")
        body.pack(fill="both", expand=True, pady=5)
        
        for mod, scores in domain_dict.items():
            if not scores: continue 
            
            # --- NIEUWE VEILIGE BEREKENING PER MODULE ---
            numeric_scores = []
            for s in scores:
                try:
                    # Forceer naar tekst, strip de % en probeer er een getal van te maken
                    val = str(s).replace('%', '').strip()
                    numeric_scores.append(float(val))
                except ValueError:
                    pass # Negeer foute/lege waardes zoals "N/A"
            
            # Als er na het filteren geen geldige getallen over zijn, sla deze module over
            if not numeric_scores: continue
            
            mod_avg = sum(numeric_scores) / len(numeric_scores)
            # --------------------------------------------
            
            container = ctk.CTkFrame(body, fg_color="transparent")
            container.pack(fill="x", pady=5)
            
            row = ctk.CTkFrame(container, fg_color=COLOR_BG_DEEP, corner_radius=8, border_width=1, border_color=COLOR_ACCENT)
            row.pack(fill="x")
            
            ctk.CTkLabel(row, text=mod, font=("Segoe UI", 14, "bold"), width=150, anchor="w").pack(side="left", padx=15, pady=12)
            
            bar_color = COLOR_PASS if mod_avg >= 70 else (COLOR_WARN if mod_avg >= 50 else COLOR_FAIL)
            bar = ctk.CTkProgressBar(row, progress_color=bar_color, fg_color=COLOR_BG_LIGHT, height=12)
            bar.set(mod_avg / 100)
            bar.pack(side="left", fill="x", expand=True, padx=15)
            
            ctk.CTkLabel(row, text=f"{mod_avg:.1f}%", font=("Segoe UI", 14, "bold"), width=60).pack(side="left", padx=5)

            specific_reasons = self._get_module_reasons(mod)
            
            detail_frame = ctk.CTkFrame(container, fg_color=COLOR_BG_LIGHT, corner_radius=5)
            
            if mod_avg == 100:
                detail_text = "✔️ Geen compliance fouten gevonden in deze module voor de gescande bestanden."
                text_color = COLOR_PASS
            elif not specific_reasons:
                detail_text = "⚠️ Bestanden faalden op dit onderdeel, mogelijk omdat de hoofdlocatie al foutief is (Locatie Beleid)."
                text_color = COLOR_WARN
            else:
                detail_text = "Gevonden fouten in gescande bestanden:\n\n• " + "\n• ".join(specific_reasons)
                text_color = "#e2e8f0" 

            ctk.CTkLabel(detail_frame, text=detail_text, font=("Segoe UI", 13), text_color=text_color, justify="left", anchor="w", wraplength=700).pack(fill="x", padx=20, pady=15)

            def make_toggle_func(df, b):
                def toggle():
                    if df.winfo_ismapped():
                        df.pack_forget()
                        b.configure(text="▼")
                    else:
                        df.pack(fill="x", pady=(2,0))
                        b.configure(text="▲")
                return toggle

            btn_toggle = ctk.CTkButton(row, text="▼", width=35, height=35, fg_color="transparent", hover_color=COLOR_BG_LIGHT, text_color=COLOR_ACCENT, font=("Segoe UI Black", 14))
            btn_toggle.configure(command=make_toggle_func(detail_frame, btn_toggle))
            btn_toggle.pack(side="right", padx=10)

    def export_to_csv(self):
        """Exporteert de data naar CSV."""
        results = self.analysis_data.get("results", [])
        if not results:
            messagebox.showinfo("Export", "Geen data om te exporteren.")
            return

        # Vraag de gebruiker waar hij het bestand wil opslaan
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV bestanden", "*.csv"), ("Alle bestanden", "*.*")],
            title="Sla het Compliance Rapport op"
        )

        if not filepath:
            return # Gebruiker heeft op 'Annuleren' geklikt

        # 1. Genereer de datumstempel van dit exacte moment
        nu_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 2. Injecteer de Datum in elke rij van de resultaten
        for row in results:
            row['ScanDatum'] = nu_str

        # 3. Zet de ScanDatum helemaal vooraan in de kolom-headers
        headers = ['ScanDatum'] + [k for k in results[0].keys() if k != 'ScanDatum']

        # 4. Schrijf het bestand weg (met puntkomma als scheidingsteken voor Europese Excel/Power BI)
        try:
            with open(filepath, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=headers, delimiter=';')
                writer.writeheader()
                writer.writerows(results)
            messagebox.showinfo("Export Succesvol", f"Rapport is veilig opgeslagen als CSV")
        except PermissionError:
            messagebox.showerror("Export Fout", "Kan het bestand niet overschrijven. Staat het toevallig nog open in Excel of Power BI?")
        except Exception as e:
            messagebox.showerror("Export Fout", f"Er is een fout opgetreden:\n{e}")
            
if __name__ == "__main__":
    app = ComplianceApp()
    app.mainloop()