import customtkinter as ctk

# Kleuren overgenomen uit je hoofdprogramma voor een consistente stijl
COLOR_PASS = "#388E3C"
COLOR_FAIL = "#D32F2F"
COLOR_WARN = "#F57C00"
COLOR_ACCENT = "#cf9d1f" 
COLOR_BG_DEEP = "#001538" 
COLOR_BG_LIGHT = "#1a2b4b" 

class DashboardWindow(ctk.CTkToplevel):
    def __init__(self, master, project_naam, analysis_data, export_callback):
        super().__init__(master)
        self.title(f"Compliance Rapport - {project_naam}")
        self.geometry("1000x800")
        self.configure(fg_color=COLOR_BG_DEEP)
        
        # Zorg dat het venster netjes de focus pakt over het hoofdscherm heen
        self.transient(master)
        self.grab_set()
        self.focus_force()

        self.analysis_data = analysis_data
        self.export_callback = export_callback
        self.results = self.analysis_data.get("results", [])

        self.bouw_ui(project_naam)

    def bouw_ui(self, project_naam):
        # --- HEADER ---
        header_frame = ctk.CTkFrame(self, fg_color=COLOR_BG_DEEP, corner_radius=0)
        header_frame.pack(fill="x", pady=(20, 10))
        
        ctk.CTkLabel(header_frame, text=f"📊 Actie Dashboard: {project_naam}", 
                     font=("Segoe UI Black", 28), text_color=COLOR_ACCENT).pack()

        # --- STATISTIEKEN (KPI's) ---
        totaal_bestanden = len(self.results)
        fout_bestanden = sum(1 for b in self.results if b.get('Score_Totaal') != '100%')
        goed_bestanden = totaal_bestanden - fout_bestanden
        percentage_goed = (goed_bestanden / totaal_bestanden * 100) if totaal_bestanden > 0 else 0

        kpi_frame = ctk.CTkFrame(self, fg_color="transparent")
        kpi_frame.pack(fill="x", padx=40, pady=10)

        self.maak_kpi_kaart(kpi_frame, "Totaal Gescand", str(totaal_bestanden), "white", "left")
        self.maak_kpi_kaart(kpi_frame, "Compliance Score", f"{percentage_goed:.1f}%", COLOR_PASS if percentage_goed > 80 else COLOR_WARN, "left")
        self.maak_kpi_kaart(kpi_frame, "Vereisen Actie", str(fout_bestanden), COLOR_FAIL if fout_bestanden > 0 else COLOR_PASS, "right")

        # --- SCROLLBARE LIJST MET FOUTEN ---
        ctk.CTkLabel(self, text="⚠️ Aandachtspunten & Oplossingen", font=("Segoe UI", 18, "bold"), text_color="white").pack(anchor="w", padx=40, pady=(20, 5))
        
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color=COLOR_BG_LIGHT, corner_radius=10)
        self.scroll_frame.pack(fill="both", expand=True, padx=40, pady=5)

        if fout_bestanden == 0:
            ctk.CTkLabel(self.scroll_frame, text="🎉 Geweldig! Alle bestanden voldoen aan de NAVO-normen.", 
                         font=("Segoe UI", 16), text_color=COLOR_PASS).pack(pady=40)
        else:
            self.vul_fouten_lijst()

        # --- FOOTER (Knoppen) ---
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=40, pady=20)
        
        btn_export = ctk.CTkButton(footer_frame, text="📥 Exporteer naar CSV", font=("Segoe UI Black", 14), 
                                   fg_color=COLOR_ACCENT, text_color="black", hover_color="#b0841a",
                                   height=45, command=self.export_callback)
        btn_export.pack(side="right")
        
        btn_sluit = ctk.CTkButton(footer_frame, text="Sluit Dashboard", font=("Segoe UI", 14), 
                                  fg_color="transparent", border_width=2, border_color="gray", 
                                  height=45, command=self.destroy)
        btn_sluit.pack(side="left")

    def maak_kpi_kaart(self, parent, titel, waarde, kleur, uitlijning):
        kaart = ctk.CTkFrame(parent, fg_color=COLOR_BG_LIGHT, corner_radius=10, border_width=1, border_color=COLOR_ACCENT)
        kaart.pack(side=uitlijning, expand=True, fill="x", padx=10)
        
        ctk.CTkLabel(kaart, text=titel, font=("Segoe UI", 14), text_color="#e2e8f0").pack(pady=(15, 0))
        ctk.CTkLabel(kaart, text=waarde, font=("Segoe UI Black", 32), text_color=kleur).pack(pady=(0, 15))

    def vul_fouten_lijst(self):
        for bestand in self.results:
            if bestand.get('Score_Totaal') == '100%':
                continue # Sla perfecte bestanden over

            card = ctk.CTkFrame(self.scroll_frame, fg_color=COLOR_BG_DEEP, corner_radius=8)
            card.pack(fill="x", pady=8, padx=10, ipadx=10, ipady=10)

            # Bestandsnaam
            ctk.CTkLabel(card, text=f"📄 {bestand.get('Naam', 'Onbekend')}", font=("Segoe UI", 15, "bold"), text_color=COLOR_ACCENT).pack(anchor="w")

            # Foutmelding
            reden = bestand.get('Reden', 'Onbekende fout')
            ctk.CTkLabel(card, text=f"❌ Probleem: {reden}", font=("Segoe UI", 13), text_color="white", justify="left", wraplength=750).pack(anchor="w", pady=(5, 0))

            # Genereer slimme suggestie
            reden_str = str(reden).lower()
            suggestie = "Controleer het bestand handmatig op de compliance regels."
            if "naamgeving" in reden_str: 
                suggestie = "Hernoem het bestand naar: YYYYMMDD_Rubricering_Afdeling_Onderwerp_Versie."
            elif "metadata" in reden_str or "auteur" in reden_str: 
                suggestie = "Open het bestand, ga naar Eigenschappen en vul de 'Auteur' in."
            elif "locatie" in reden_str: 
                suggestie = "Verplaats dit bestand naar SharePoint (het hoort niet op de lokale schijf)."
            elif "schadelijk" in reden_str or "kritiek" in reden_str: 
                suggestie = "🚨 DIRECT VERWIJDEREN! Dit is een schadelijk of ongeoorloofd bestandstype."

            # Suggestie tekst
            ctk.CTkLabel(card, text=f"💡 Oplossing: {suggestie}", text_color=COLOR_PASS, font=("Segoe UI", 13, "bold"), justify="left", wraplength=750).pack(anchor="w", pady=(5, 0))