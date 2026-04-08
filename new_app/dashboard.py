# ==============================================================================
# dashboard.py — DashboardFrame: rapport, tabs, bestandenlijst en CSV-export
# ==============================================================================

# --- Stdlib imports ---
import csv
import os
from datetime import datetime

# --- Third-party imports ---
import customtkinter as ctk
from PIL import Image
from tkinter import filedialog, messagebox


# ==============================================================================
# KLEUR-CONSTANTEN (Gold/Blue theme)
# ==============================================================================

COLOR_BG_DEEP    = "#001538"   # Diepdonkerblauw — hoofdachtergrond
COLOR_BG_LIGHT   = "#1a2b4b"   # Licht donkerblauw — cards / rijen
COLOR_ACCENT     = "#cf9d1f"   # Goud — knoppen, titels, actieve tab
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
# DASHBOARD FRAME
# ==============================================================================

class DashboardFrame(ctk.CTkFrame):
    """Rapportage-scherm met tabs voor totaal, SharePoint, lokaal en bestandenlijst."""

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color=COLOR_BG_DEEP)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self._build_header()
        self._build_nav()
        self._build_content_area()
        self._build_export_bar()

        # State
        self.analysis_data = None
        self.full_dataset = []
        self.results_filtered = []
        self.filtered_search_results = []
        self.current_page = 0
        self.items_per_page = 50

    # --------------------------------------------------------------------------
    # Layout-opbouw
    # --------------------------------------------------------------------------

    def _build_header(self):
        """Bouw de header-balk met logo, titel en navigatieknoppen."""
        self.head = ctk.CTkFrame(self, height=80, corner_radius=0, fg_color=COLOR_BG_DEEP)
        self.head.grid(row=0, column=0, sticky="ew")

        logo_img = _load_logo(size=(45, 45))
        if logo_img:
            ctk.CTkLabel(self.head, image=logo_img, text="").pack(
                side="left", padx=(15, 5), pady=10
            )

        title_frame = ctk.CTkFrame(self.head, fg_color="transparent")
        title_frame.pack(side="left", padx=5, pady=10)

        self.lbl_title = ctk.CTkLabel(
            title_frame, text="Rapport",
            font=("Segoe UI", 24, "bold"), text_color="white"
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = ctk.CTkLabel(
            title_frame, text="Wachten op data...",
            font=("Segoe UI", 13), text_color="#94a3b8"
        )
        self.lbl_subtitle.pack(anchor="w")

        self.btn_logout = ctk.CTkButton(
            self.head, text="🚪 Uitloggen",
            font=("Segoe UI", 13, "bold"), width=130, height=36,
            fg_color=COLOR_ACCENT, text_color="black", hover_color=COLOR_ACCENT_HOVER,
            command=self.logout
        )
        self.btn_logout.pack(side="right", padx=10, pady=22)

        self.btn_new_scan = ctk.CTkButton(
            self.head, text="🔄 Nieuwe Scan",
            font=("Segoe UI", 13, "bold"), width=130, height=36,
            fg_color=COLOR_ACCENT, text_color="black", hover_color=COLOR_ACCENT_HOVER,
            command=self.new_scan
        )
        self.btn_new_scan.pack(side="right", padx=(20, 10), pady=22)

    def _build_nav(self):
        """Bouw de tab-navigatiebalk."""
        self.nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.nav_frame.grid(row=1, column=0, sticky="ew", pady=10, padx=20)

        self.view_total = ctk.CTkFrame(self.content_area if hasattr(self, 'content_area') else self, fg_color="transparent")
        self.view_sp    = ctk.CTkFrame(self.content_area if hasattr(self, 'content_area') else self, fg_color="transparent")
        self.view_local = ctk.CTkFrame(self.content_area if hasattr(self, 'content_area') else self, fg_color="transparent")
        self.view_files = ctk.CTkFrame(self.content_area if hasattr(self, 'content_area') else self, fg_color="transparent")

        self.btn_total = ctk.CTkButton(
            self.nav_frame, text="Totaal Overzicht",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.switch_tab(self.btn_total, self.view_total)
        )
        self.btn_total.pack(side="left", expand=True, padx=5)

        self.btn_sp = ctk.CTkButton(
            self.nav_frame, text="SharePoint",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.switch_tab(self.btn_sp, self.view_sp)
        )
        self.btn_sp.pack(side="left", expand=True, padx=5)

        self.btn_local = ctk.CTkButton(
            self.nav_frame, text="Lokale Schijf",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.switch_tab(self.btn_local, self.view_local)
        )
        self.btn_local.pack(side="left", expand=True, padx=5)

        self.btn_files = ctk.CTkButton(
            self.nav_frame, text="🔍 Bestandenlijst",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.switch_tab(self.btn_files, self.view_files)
        )
        self.btn_files.pack(side="left", expand=True, padx=5)

        self.filter_var = ctk.StringVar(value="Alle Afdelingen")
        self.opt_filter = ctk.CTkOptionMenu(
            self.nav_frame,
            values=["Alle Afdelingen", "HR", "Finance", "IT", "Directie"],
            variable=self.filter_var,
            font=("Segoe UI", 14)
        )

    def _build_content_area(self):
        """Bouw het centrale inhoudsgebied."""
        self.content_area = ctk.CTkFrame(self, fg_color="transparent")
        self.content_area.grid(row=2, column=0, sticky="nsew", padx=20, pady=5)

    def _build_export_bar(self):
        """Bouw de exportbalk onderaan."""
        self.export_frame = ctk.CTkFrame(self, fg_color="transparent", height=70)
        self.export_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=10)

        self.btn_export = ctk.CTkButton(
            self.export_frame, text="📥 Exporteer Data (CSV)",
            font=("Segoe UI Black", 14), text_color="white",
            height=50, command=self.export_to_csv
        )
        self.btn_export.pack(fill="x", pady=10)

    # --------------------------------------------------------------------------
    # Data laden
    # --------------------------------------------------------------------------

    def load_data(self, data, afdeling):
        """Laad analyseresultaten in het dashboard en bouw het juiste scherm op."""
        self.analysis_data = data
        self.current_afdeling = afdeling
        self.lbl_title.configure(text=f"Rapport: {afdeling}")

        base_results = self.analysis_data.get("results", [])
        self.full_dataset = base_results

        critical_count = len([r for r in base_results if "🚨 KRITIEK" in r.get("Reden", "")])
        self.lbl_subtitle.configure(
            text=f"{len(base_results)} bestanden geanalyseerd  |  {critical_count} kritieke problemen"
        )

        is_admin = self.controller.current_user and self.controller.current_user.is_admin

        for widget in self.content_area.winfo_children():
            widget.destroy()

        if is_admin:
            self._build_admin_view(base_results, afdeling)
        else:
            self._build_user_view(base_results)

    def _build_admin_view(self, base_results, afdeling):
        """Beheerdersweergave met PowerBI-downloadknop."""
        self.nav_frame.grid_remove()
        self.export_frame.grid_remove()

        admin_view = ctk.CTkFrame(self.content_area, fg_color="transparent")
        admin_view.pack(fill="both", expand=True, pady=50)

        ctk.CTkLabel(
            admin_view, text="🛡️ Administrator Scan Voltooid",
            font=("Segoe UI Black", 30), text_color=COLOR_ACCENT
        ).pack(pady=(20, 10))

        info_text = (
            f"De scan voor afdeling '{afdeling}' is succesvol afgerond.\n\n"
            f"In totaal zijn er {len(base_results)} afwijkende bestanden geïdentificeerd.\n"
            "Omdat administrator-scans enorme hoeveelheden data bevatten, is de\n"
            "visuele in-app tabel uitgeschakeld voor optimale prestaties.\n\n"
            "Klik op de knop hieronder om de dataset te genereren voor uw PowerBI Dashboard."
        )
        ctk.CTkLabel(
            admin_view, text=info_text,
            font=("Segoe UI", 16), text_color="#e2e8f0"
        ).pack(pady=20)

        ctk.CTkButton(
            admin_view, text="📊 Download PowerBI Dataset (.csv)",
            font=("Segoe UI Black", 18),
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER,
            text_color="black", height=60, width=450,
            command=self.export_to_csv
        ).pack(pady=30)

        ctk.CTkButton(
            admin_view, text="👀 Bekijk Bestandenlijst (In-App)",
            font=("Segoe UI", 14), fg_color="transparent",
            border_width=1, border_color="#334155",
            hover_color=COLOR_BG_LIGHT, height=40, width=300,
            command=self.show_admin_file_list
        ).pack(pady=10)

    def _build_user_view(self, base_results):
        """Gebruikersweergave met tabbladen en top-50 bestanden."""
        self.nav_frame.grid()
        self.export_frame.grid_remove()
        self.results_filtered = base_results[:50]

        self.view_total = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.view_sp    = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.view_local = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.view_files = ctk.CTkFrame(self.content_area, fg_color="transparent")

        risico_bestanden = [
            res["Naam"] for res in self.results_filtered
            if "🚨 KRITIEK" in res.get("Reden", "")
        ]
        if risico_bestanden:
            self._build_alert(self.view_total, risico_bestanden)

        all_keys = (
            set(self.analysis_data["domain_scores_local"].keys())
            | set(self.analysis_data["domain_scores_sp"].keys())
        )
        combined_domains = {mod: [] for mod in all_keys}
        for mod in all_keys:
            combined_domains[mod].extend(self.analysis_data["domain_scores_local"].get(mod, []))
            combined_domains[mod].extend(self.analysis_data["domain_scores_sp"].get(mod, []))

        self._build_tab_content(self.view_total, combined_domains)
        self._build_tab_content(self.view_sp, self.analysis_data.get("domain_scores_sp", {}))
        self._build_tab_content(self.view_local, self.analysis_data.get("domain_scores_local", {}))
        self._build_files_tab()

        self.switch_tab(self.btn_total, self.view_total)

        if len(base_results) > 50:
            ctk.CTkLabel(
                self.view_total,
                text="Let op: Als 'User' ziet u hier uw top 50 prioriteitsbestanden.",
                text_color=COLOR_WARN, font=("Segoe UI", 13, "bold")
            ).pack(pady=10)

    def show_admin_file_list(self):
        """Toon de volledige bestandenlijst voor de beheerder."""
        for widget in self.content_area.winfo_children():
            widget.destroy()

        self.nav_frame.grid()
        self.export_frame.grid()

        base_results = self.analysis_data.get("results", [])
        self.results_filtered = base_results

        self.view_total = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.view_sp    = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.view_local = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.view_files = ctk.CTkFrame(self.content_area, fg_color="transparent")

        risico_bestanden = [
            res["Naam"] for res in self.results_filtered
            if "🚨 KRITIEK" in res.get("Reden", "")
        ]
        if risico_bestanden:
            self._build_alert(self.view_total, risico_bestanden)

        all_keys = (
            set(self.analysis_data["domain_scores_local"].keys())
            | set(self.analysis_data["domain_scores_sp"].keys())
        )
        combined_domains = {mod: [] for mod in all_keys}
        for mod in all_keys:
            combined_domains[mod].extend(self.analysis_data["domain_scores_local"].get(mod, []))
            combined_domains[mod].extend(self.analysis_data["domain_scores_sp"].get(mod, []))

        self._build_tab_content(self.view_total, combined_domains)
        self._build_tab_content(self.view_sp, self.analysis_data.get("domain_scores_sp", {}))
        self._build_tab_content(self.view_local, self.analysis_data.get("domain_scores_local", {}))
        self._build_files_tab()

        self.switch_tab(self.btn_files, self.view_files)

        lbl_warn = ctk.CTkLabel(
            self.view_files,
            text="⚠️ Beheerdersmodus: Alle bestanden zijn ingeladen. Navigeren kan vertraagd zijn.",
            text_color=COLOR_WARN, font=("Segoe UI", 13, "bold")
        )
        lbl_warn.pack(before=self.view_files.winfo_children()[0], pady=10)

    # --------------------------------------------------------------------------
    # Navigatie & tabs
    # --------------------------------------------------------------------------

    def new_scan(self):
        self.controller.show_frame("HomeFrame")

    def logout(self):
        self.controller.set_user(None)
        self.controller.show_frame("AuthFrame")

    def switch_tab(self, btn_active, view_active):
        """Activeer een tabblad en verberg de overige."""
        for btn in [self.btn_total, self.btn_sp, self.btn_local, self.btn_files]:
            btn.configure(fg_color=COLOR_BG_LIGHT, text_color="white")
        btn_active.configure(fg_color=COLOR_ACCENT, text_color="black")

        self.view_total.pack_forget()
        self.view_sp.pack_forget()
        self.view_local.pack_forget()
        self.view_files.pack_forget()
        view_active.pack(fill="both", expand=True)

    # --------------------------------------------------------------------------
    # Tab-inhoud builders
    # --------------------------------------------------------------------------

    def _build_alert(self, parent_frame, risico_bestanden):
        """Toon een rode waarschuwingsbanner voor kritieke bestanden."""
        alert_frame = ctk.CTkFrame(
            parent_frame, fg_color="#7f1d1d", corner_radius=10,
            border_width=2, border_color="#ef4444"
        )
        alert_frame.pack(fill="x", padx=20, pady=(15, 0))

        ctk.CTkLabel(
            alert_frame, text="🚨 KRITIEKE BEVEILIGINGSWAARSCHUWING 🚨",
            font=("Segoe UI Black", 18), text_color="white"
        ).pack(pady=(15, 5))

        ctk.CTkLabel(
            alert_frame,
            text="De volgende schadelijke bestanden zijn gevonden en blokkeren compliance. DIRECT VERWIJDEREN:",
            font=("Segoe UI", 14, "bold"), text_color="#fca5a5"
        ).pack(pady=0)

        toon_bestanden = risico_bestanden[:5]
        files_text = "\n".join([f"• {f}" for f in toon_bestanden])
        if len(risico_bestanden) > 5:
            files_text += f"\n... en nog {len(risico_bestanden) - 5} andere(n)."

        ctk.CTkLabel(
            alert_frame, text=files_text,
            font=("Consolas", 14), text_color="white", justify="left"
        ).pack(pady=(10, 15), padx=20)

    def _calc_average(self, domain_dict):
        """Bereken het gemiddelde van alle scores in een domein-dict."""
        if not domain_dict:
            return -1
        valid_scores = []
        for scores in domain_dict.values():
            if not scores:
                continue
            numeric_scores = []
            for s in scores:
                if isinstance(s, str):
                    s = s.replace('%', '').strip()
                try:
                    numeric_scores.append(float(s))
                except ValueError:
                    pass
            if numeric_scores:
                valid_scores.append(sum(numeric_scores) / len(numeric_scores))
        return sum(valid_scores) / len(valid_scores) if valid_scores else -1

    def _get_module_reasons(self, module_name):
        """Haal alle unieke redenen op die betrekking hebben op een bepaalde module."""
        reasons_found = set()
        keyword = module_name.split()[0]
        for res in self.results_filtered:
            reden_string = res.get("Reden", "")
            if keyword in reden_string:
                for part in reden_string.split(" | "):
                    if keyword in part:
                        reasons_found.add(part)
        return list(reasons_found)

    def _build_tab_content(self, parent_frame, domain_dict):
        """Bouw de score-boxes en regellenlijst voor een tabblad."""
        active_comp = getattr(self.controller, 'last_run_comp_modules', [])
        active_qual = getattr(self.controller, 'last_run_qual_modules', [])
        central_comp_keys = ["Security (Risico's)", "Data Duplicatie", "Locatie Beleid"]

        comp_keys = active_comp + central_comp_keys
        qual_keys = active_qual

        comp_dict = {k: v for k, v in domain_dict.items() if k in comp_keys}
        qual_dict = {k: v for k, v in domain_dict.items() if k in qual_keys}

        comp_avg = self._calc_average(comp_dict)
        qual_avg = self._calc_average(qual_dict)

        if comp_avg == -1 and qual_avg == -1:
            ctk.CTkLabel(
                parent_frame,
                text="Geen documenten geanalyseerd in deze bron.",
                font=("Segoe UI", 16, "italic"), text_color="#e2e8f0"
            ).pack(pady=50)
            return

        # Score-boxes bovenaan
        header_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=10)

        def _build_score_box(parent, title, avg):
            box = ctk.CTkFrame(
                parent, fg_color=COLOR_BG_LIGHT, corner_radius=6,
                border_width=1, border_color="#334155"
            )
            box.pack(side="left", fill="both", expand=True, padx=10)
            ctk.CTkLabel(
                box, text=title, font=("Segoe UI", 16, "bold"), text_color="#e2e8f0"
            ).pack(pady=(20, 5))
            if avg == -1:
                ctk.CTkLabel(
                    box, text="N/A", font=("Segoe UI Black", 42), text_color="#475569"
                ).pack(pady=(0, 20))
            else:
                kleur = COLOR_PASS if avg >= 75 else (COLOR_WARN if avg >= 50 else COLOR_FAIL)
                ctk.CTkLabel(
                    box, text=f"{avg:.1f}%", font=("Segoe UI Black", 42), text_color=kleur
                ).pack(pady=(0, 20))

        if active_comp or central_comp_keys:
            _build_score_box(header_frame, "🛡️ Compliance Score", comp_avg)
        if active_qual:
            _build_score_box(header_frame, "✨ Data Quality Score", qual_avg)

        # Regellenlijst
        body = ctk.CTkScrollableFrame(parent_frame, fg_color="transparent")
        body.pack(fill="both", expand=True, pady=10)

        def _build_rules_list(title, rules_dict):
            if not rules_dict:
                return
            if not any(len(v) > 0 for v in rules_dict.values()):
                return

            ctk.CTkLabel(
                body, text=title,
                font=("Segoe UI", 20, "bold"), text_color=COLOR_ACCENT
            ).pack(anchor="w", padx=10, pady=(20, 10))

            for mod, scores in rules_dict.items():
                if not scores:
                    continue
                numeric = []
                for s in scores:
                    try:
                        numeric.append(float(str(s).replace('%', '').strip()))
                    except ValueError:
                        pass
                if not numeric:
                    continue

                mod_avg = sum(numeric) / len(numeric)
                light = "🟢" if mod_avg >= 75 else ("🟡" if mod_avg >= 50 else "🔴")

                row = ctk.CTkFrame(
                    body, fg_color=COLOR_BG_LIGHT, corner_radius=6,
                    border_width=1, border_color="#334155"
                )
                row.pack(fill="x", padx=10, pady=5)

                ctk.CTkLabel(
                    row, text=f"{light}  {mod}  •  {mod_avg:.1f}%",
                    font=("Segoe UI", 15, "bold"), text_color="white"
                ).pack(anchor="w", padx=20, pady=(12, 5))

                specific_reasons = self._get_module_reasons(mod)
                if mod_avg == 100:
                    text_reasons = "✓ Geen regelfouten gedetecteerd in bestanden voor dit onderdeel."
                    c = COLOR_PASS
                elif not specific_reasons:
                    text_reasons = "⚠️ Omlaag gebracht door afgeleide fouten (zoals verkeerde datalocatie)."
                    c = COLOR_WARN
                else:
                    text_reasons = "Gevonden fouten in gescande bestanden:\n" + "\n".join([f"• {r}" for r in specific_reasons])
                    c = "#e2e8f0"

                ctk.CTkLabel(
                    row, text=text_reasons,
                    font=("Consolas", 13), text_color=c, justify="left"
                ).pack(anchor="w", padx=40, pady=(0, 15))

        if active_comp or central_comp_keys:
            _build_rules_list("Regels: Compliance", comp_dict)
        if active_qual:
            _build_rules_list("Regels: Data Quality", qual_dict)

    def _build_files_tab(self):
        """Bouw het bestandenlijst-tabblad met zoekbalk en paginering."""
        search_frame = ctk.CTkFrame(self.view_files, fg_color="transparent")
        search_frame.pack(fill="x", pady=10)

        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            search_frame, textvariable=self.search_var,
            placeholder_text="Zoek op bestandsnaam...",
            font=("Segoe UI", 14), width=300
        )
        self.search_entry.pack(side="left", padx=(0, 10))

        btn_search = ctk.CTkButton(
            search_frame, text="Zoeken", command=self.apply_search
        )
        btn_search.pack(side="left")
        self.search_entry.bind("<Return>", lambda e: self.apply_search())

        header = ctk.CTkFrame(self.view_files, fg_color=COLOR_BG_DEEP)
        header.pack(fill="x", pady=5)
        ctk.CTkLabel(
            header, text="Bestand & Reden",
            font=("Segoe UI", 14, "bold")
        ).pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkLabel(
            header, text="Locatie",
            font=("Segoe UI", 14, "bold")
        ).pack(side="right", padx=10)

        self.file_list_container = ctk.CTkScrollableFrame(
            self.view_files, fg_color=COLOR_BG_LIGHT
        )
        self.file_list_container.pack(fill="both", expand=True, pady=5)

        page_frame = ctk.CTkFrame(self.view_files, fg_color="transparent")
        page_frame.pack(fill="x", pady=10)

        self.btn_prev = ctk.CTkButton(
            page_frame, text="< Vorige", command=self.prev_page, width=100
        )
        self.btn_prev.pack(side="left", padx=10)

        self.lbl_page = ctk.CTkLabel(
            page_frame, text="Pagina 1 van 1", font=("Segoe UI", 14)
        )
        self.lbl_page.pack(side="left", expand=True)

        self.btn_next = ctk.CTkButton(
            page_frame, text="Volgende >", command=self.next_page, width=100
        )
        self.btn_next.pack(side="right", padx=10)

        self.apply_search()

    # --------------------------------------------------------------------------
    # Zoeken & paginering
    # --------------------------------------------------------------------------

    def apply_search(self):
        query = self.search_var.get().lower()
        if not query:
            self.filtered_search_results = self.results_filtered
        else:
            self.filtered_search_results = [
                r for r in self.results_filtered
                if query in r.get("Naam", "").lower() or query in r.get("Reden", "").lower()
            ]
        self.current_page = 0
        self.render_file_page()

    def next_page(self):
        max_page = max(0, (len(self.filtered_search_results) - 1) // self.items_per_page)
        if self.current_page < max_page:
            self.current_page += 1
            self.render_file_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_file_page()

    def render_file_page(self):
        """Render de huidige pagina van de bestandenlijst."""
        for widget in self.file_list_container.winfo_children():
            widget.destroy()

        start_idx = self.current_page * self.items_per_page
        end_idx = start_idx + self.items_per_page
        page_items = self.filtered_search_results[start_idx:end_idx]

        for item in page_items:
            row = ctk.CTkFrame(self.file_list_container, fg_color=COLOR_BG_DEEP)
            row.pack(fill="x", pady=5, padx=5)

            score_color = COLOR_PASS if "100%" in item.get("Score_Totaal", "") else COLOR_WARN
            if "0%" in item.get("Score_Totaal", ""):
                score_color = COLOR_FAIL

            title_text = f"{item.get('Naam', '')}  [{item.get('Score_Totaal', '')}]"
            ctk.CTkLabel(
                row, text=title_text,
                font=("Segoe UI", 14, "bold"), text_color=score_color
            ).pack(anchor="w", padx=10, pady=(5, 5))

            for r in item.get('Reden', '').split(" | "):
                if r.strip():
                    if "KRITIEK" in r:
                        r_color = "#fca5a5"
                        icon = "🚨"
                    else:
                        r_color = "#e2e8f0"
                        icon = "⚠️"
                    ctk.CTkLabel(
                        row, text=f"{icon} {r.strip()}",
                        font=("Consolas", 13), text_color=r_color,
                        justify="left", wraplength=700
                    ).pack(anchor="w", padx=25, pady=(0, 2))

            ctk.CTkLabel(row, text="", height=2).pack()

            pad = item.get('Pad', '')
            parts = pad.replace('\\', '/').split('/')
            kort_pad = ".../" + "/".join(parts[-3:]) if len(parts) > 3 else pad
            ctk.CTkLabel(
                row, text=kort_pad,
                font=("Segoe UI", 11, "italic"), text_color="#888888"
            ).place(relx=1.0, rely=0.05, anchor="ne", x=-10, y=10)

        max_page = max(1, (len(self.filtered_search_results) + self.items_per_page - 1) // self.items_per_page)
        self.lbl_page.configure(text=f"Pagina {self.current_page + 1} van {max_page}")

    # --------------------------------------------------------------------------
    # CSV-export
    # --------------------------------------------------------------------------

    def export_to_csv(self):
        """Exporteer de volledige dataset als PowerBI-klaar CSV-bestand."""
        if not self.full_dataset:
            messagebox.showinfo("Export", "Geen data om te exporteren.")
            return

        documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        afdeling_safe = self.current_afdeling.replace(" ", "_")
        filename = f"PowerBI_Dataset_{afdeling_safe}_{timestamp}.csv"
        filepath = os.path.join(documents_path, filename)

        nu_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        headers = [
            'ScanDatum', 'Afdeling', 'Type', 'Bron', 'Risico_Niveau',
            'Bestandsnaam', 'Pad', 'Score_Totaal', 'Aantal_Overtredingen',
            'Compliance_Security', 'Compliance_Duplicatie', 'Compliance_Locatie',
            'Quality_Bestandsgrootte', 'Quality_Actualiteit', 'Quality_Volledigheid',
            'Quality_Leesbaarheid', 'Details_Reden'
        ]

        export_data = []
        for row in self.full_dataset:
            reden_str = row.get('Reden', '')
            is_kritiek = "KRITIEK" in reden_str
            is_warn = (
                "⚠️" in reden_str or "Locatie:" in reden_str
                or "Duplicatie:" in reden_str or "Waarschuwing" in reden_str
            )
            risico_niveau = "Kritiek" if is_kritiek else ("Waarschuwing" if is_warn else "Veilig")
            bron = "SharePoint" if "SP" in row.get('Mode', '') else "Lokaal"
            aantal_overtredingen = (
                0 if not reden_str or "Compliant" in reden_str
                else len(reden_str.split(" | "))
            )

            export_data.append({
                'ScanDatum': nu_str,
                'Afdeling': self.current_afdeling,
                'Type': row.get('Type', 'Onbekend'),
                'Bron': bron,
                'Risico_Niveau': risico_niveau,
                'Bestandsnaam': row.get('Naam', ''),
                'Pad': row.get('Pad', ''),
                'Score_Totaal': row.get('Score_Totaal', ''),
                'Aantal_Overtredingen': aantal_overtredingen,
                'Compliance_Security': row.get("Security (Risico's)", "100%"),
                'Compliance_Duplicatie': row.get("Data Duplicatie", "100%"),
                'Compliance_Locatie': row.get("Locatie Beleid", "100%"),
                'Quality_Bestandsgrootte': row.get("Bestandsgrootte", "N/A"),
                'Quality_Actualiteit': row.get("Actualiteit", "N/A"),
                'Quality_Volledigheid': row.get("Volledigheid", "N/A"),
                'Quality_Leesbaarheid': row.get("Leesbaarheid", "N/A"),
                'Details_Reden': reden_str.replace('\n', ' ').strip()
            })

        try:
            with open(filepath, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=headers, delimiter=';', extrasaction='ignore')
                writer.writeheader()
                writer.writerows(export_data)
            messagebox.showinfo(
                "Export Succesvol",
                f"PowerBI Dataset is succesvol opgeslagen in:\n{filepath}\n\n"
                "Deze is ingericht met vaste filters en direct klaar voor inlezen!"
            )
        except PermissionError:
            messagebox.showerror("Export Fout", "Kan het bestand niet overschrijven. Is het geopend in Excel of PowerBI?")
        except Exception as e:
            messagebox.showerror("Export Fout", f"Er is een fout opgetreden:\n{e}")
