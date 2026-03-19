import os
import io
import time
from office365.sharepoint.client_context import ClientContext
from requests_negotiate_sspi import HttpNegotiateAuth

class CentraleEngine:
    def __init__(self, local_paths, sharepoint_sites, active_engines):
        self.local_paths = local_paths
        self.sharepoint_sites = sharepoint_sites
        self.active_engines = active_engines
        
        self.results = []
        self.file_registry = {} 
        self.sp_bibliotheken_tracker = {}

        # --- DATA QUALITY & SECURITY BASELINES (Verplaatst vanuit Compliance Engine) ---
        self.ALLOWED_SP_EXTS = {'.docx', '.xlsx', '.pptx', '.pdf', '.txt'}
        self.RISKY_EXTS = {'.exe', '.bat', '.msi', '.ps1', '.vbs', '.cmd', '.sh', '.scr'}
        self.FORBIDDEN_CHARS = set('/\\:*?"<>| !+@')
        
        # Centrale harde regels die ALTIJD gelden
        self.base_domains = ["Security (Risico's)", "Data Duplicatie", "Locatie Beleid"]
        
        # Verzamel alle te scoren domeinen (Base + Engines)
        self.all_domains = list(self.base_domains)
        for engine in self.active_engines:
            self.all_domains.extend(engine.domains)
            
        self.domain_scores_local = {mod: [] for mod in self.all_domains}
        self.domain_scores_sp = {mod: [] for mod in self.all_domains}

        self.EXCEPTIONS_FOLDERS = ["werkomgeving", "concepten", "wip"]

    def process(self, q):
        all_items = []
        
        # 1. LOKALE SCAN
        for path in self.local_paths:
            root_src = os.path.abspath(path)
            if os.path.isdir(root_src):
                for root, dirs, files in os.walk(root_src):
                    for f in files:
                        p = os.path.join(root, f)
                        try:
                            is_werkomgeving = any(exc in p.lower() for exc in self.EXCEPTIONS_FOLDERS)
                            item = {
                                "mode": "local", "path": p, "name": f, 
                                "size": os.path.getsize(p), "root_source": root_src,
                                "in_werkomgeving": is_werkomgeving,
                                "extension": os.path.splitext(f)[1].lower()
                            }
                            all_items.append(item)
                            self._register_file(f, "local")
                        except OSError: pass
                        
        # 2. SHAREPOINT SCAN 
        for sp in self.sharepoint_sites:
            site_url = sp["url"]
            self.sp_bibliotheken_tracker[site_url] = {"Open Bibliotheek": 0, "Gesloten Bibliotheek": 0, "Foutieve Bieb": 0}
            
            try:
                ctx = ClientContext(site_url).with_credentials(HttpNegotiateAuth())
                lists = ctx.web.lists
                ctx.load(lists)
                ctx.execute_query()
                
                for library in lists:
                    if library.base_template == 101 and not library.hidden:
                        self._validate_sp_library_name(site_url, library.title)
                        try:
                            self._walk_sp_recursive(ctx, library.root_folder, library.title, site_url, all_items)
                        except Exception as lib_err:
                            self.results.append({
                                "Type": "SP Structuur", "Naam": "Toegangsfout", "Pad": f"{site_url}/{library.title}", 
                                "Mode": "SP", "Score_Totaal": "0%", "Reden": f"Kan SP-bibliotheek niet scannen: {str(lib_err)}"
                            })
            except Exception as e:
                q.put(("error", f"Kritieke SP Connectiefout op {site_url}: {str(e)}"))
                return

        if not all_items:
            q.put(("error", "Geen bestanden gevonden om te scannen."))
            return

        # 3. ORCHESTRATIE & ANALYSE
        total_items = len(all_items)
        for index, item in enumerate(all_items):
            self._analyze_item(item)
            q.put(("progress", (index + 1) / total_items))
            
        self._rapporteer_sp_bibliotheken()
        
        q.put(("done", {
            "results": self.results, 
            "domain_scores_local": self.domain_scores_local,
            "domain_scores_sp": self.domain_scores_sp
        }))

    def _walk_sp_recursive(self, ctx, folder, current_path, site_url, all_items):
        try:
            ctx.load(folder, ["Folders", "Files"])
            ctx.execute_query()
            
            for f in folder.files:
                try:
                    ctx.load(f, ["Name", "ServerRelativeUrl", "Length", "TimeCreated", "TimeLastModified"])
                    ctx.execute_query()
                    
                    file_path = f"SP: {current_path}/{f.name}"
                    is_werkomgeving = any(exc in file_path.lower() for exc in self.EXCEPTIONS_FOLDERS)
                    
                    item = {
                        "mode": "sp", "path": file_path, "name": f.name, 
                        "size": int(f.length), "sp_url": f.serverRelativeUrl, 
                        "time_created": f.timeCreated, "time_modified": f.timeLastModified, 
                        "ctx": ctx, "root_source": site_url,
                        "in_werkomgeving": is_werkomgeving,
                        "extension": os.path.splitext(f.name)[1].lower()
                    }
                    all_items.append(item)
                    self._register_file(f.name, "sp")
                except Exception: continue
                    
            for sub_folder in folder.folders:
                if sub_folder.name not in ["Forms", "_t", "_w", "Templates"]:
                    self._walk_sp_recursive(ctx, sub_folder, f"{current_path}/{sub_folder.name}", site_url, all_items)
        except Exception: pass 

    def _analyze_item(self, item):
        file_stream = self._get_file_stream(item)
        
        # -- DATA QUALITY GATEKEEPER: Bereken hier de fundamentele eigenschappen --
        filename = item["name"]
        extension = item["extension"]
        mode = item["mode"]
        is_duplicate = len(self.file_registry.get(filename.lower(), set())) > 1
        
        # Injecteer bevindingen in het item voor de sub-engines
        item["is_duplicate"] = is_duplicate
        item["has_forbidden_chars"] = any(c in filename for c in self.FORBIDDEN_CHARS)
        item["is_readable_doc"] = extension in self.ALLOWED_SP_EXTS
        
        # Startscores voor de centrale regels
        all_scores = {"Security (Risico's)": 100, "Locatie Beleid": 100, "Data Duplicatie": 100}
        all_reasons = []

        # 1. SECURITY CHECK
        if extension in self.RISKY_EXTS:
            all_scores["Security (Risico's)"] = 0
            all_reasons.append("🚨 KRITIEK: Schadelijk bestand.")

        # 2. LOCATIE CHECK
        if mode == "sp" and extension not in self.ALLOWED_SP_EXTS:
            all_scores["Locatie Beleid"] = 0
            all_reasons.append(f"Locatie: Extensie {extension} mag niet op SP.")
        elif mode == "local":
            is_large_file = item["size"] >= (2 * 1024 * 1024 * 1024)
            if extension in self.ALLOWED_SP_EXTS and not is_large_file:
                all_scores["Locatie Beleid"] = 0
                all_reasons.append("Locatie: Bestand kan op SP en hoort niet lokaal.")

        # 3. DUPLICATIE CHECK
        if is_duplicate:
            all_scores["Data Duplicatie"] = 0
            all_reasons.append("Duplicatie: Bestand bestaat lokaal én op SP.")
        elif all_scores["Locatie Beleid"] == 0:
            all_scores["Data Duplicatie"] = 0
            all_reasons.append("Duplicatie: Faalt door onjuiste basislocatie.")

        # Sla de centrale scores direct op voor het dashboard
        for domein in self.base_domains:
            if item["mode"] == "sp": self.domain_scores_sp[domein].append(all_scores[domein])
            else: self.domain_scores_local[domein].append(all_scores[domein])

        # -- ROEP DE OVERIGE ENGINES AAN (Bijv. ComplianceEngine) --
        for engine in self.active_engines:
            try:
                # De engine ontvangt nu een verrijkt 'item' object
                engine_data = engine.analyze(item, file_stream)
                for domein, score in engine_data["scores"].items():
                    all_scores[domein] = score
                    if item["mode"] == "sp": self.domain_scores_sp[domein].append(score)
                    else: self.domain_scores_local[domein].append(score)
                all_reasons.extend(engine_data["reasons"])
            except Exception as e:
                all_reasons.append(f"Engine Fout ({engine.__class__.__name__}): {str(e)}")

        # Bereken Totalen en bouw output dictionary
        item_result = {
            "Type": "Bestand", "Naam": item["name"], 
            "Pad": item["path"], "Mode": item["mode"].upper()
        }
        
        active_vals = [v for k, v in all_scores.items() if isinstance(v, int)]
        item_result["Score_Totaal"] = f"{int(sum(active_vals) / len(active_vals))}%" if active_vals else "0%"
        
        for dom in self.all_domains:
            val = all_scores.get(dom, "N/A")
            item_result[dom] = f"{val}%" if isinstance(val, int) else val
            
        item_result["Reden"] = " | ".join(all_reasons) if all_reasons else "Volledig Compliant"
        self.results.append(item_result)
        
        if file_stream:
            try: file_stream.close()
            except: pass

    # Helpers voor stream, map controle, en registratie
    def _get_file_stream(self, item):
        try:
            if item["mode"] == "local": return open(item["path"], "rb")
            elif item["mode"] == "sp": return io.BytesIO(item["ctx"].web.get_file_by_server_relative_url(item["sp_url"]).read())
        except Exception: return None

    def _register_file(self, filename, mode):
        name_lower = filename.lower()
        if name_lower not in self.file_registry: self.file_registry[name_lower] = set()
        self.file_registry[name_lower].add(mode)

    def _validate_sp_library_name(self, site_url, lib_name):
        if lib_name == "Open Bibliotheek": self.sp_bibliotheken_tracker[site_url]["Open Bibliotheek"] += 1
        elif lib_name == "Gesloten Bibliotheek": self.sp_bibliotheken_tracker[site_url]["Gesloten Bibliotheek"] += 1
        elif "bibliotheek" in lib_name.lower() or "bieb" in lib_name.lower(): self.sp_bibliotheken_tracker[site_url]["Foutieve Bieb"] += 1

    def _rapporteer_sp_bibliotheken(self):
        for site, counts in self.sp_bibliotheken_tracker.items():
            if counts["Open Bibliotheek"] > 1 or counts["Gesloten Bibliotheek"] > 1 or counts["Foutieve Bieb"] > 0:
                self.results.append({
                    "Type": "SP Structuur", "Naam": "Bibliotheek Fout", "Pad": site, "Mode": "SP",
                    "Score_Totaal": "0%", "Reden": f"FOUT: Verkeerde bibliotheek formatie. Open: {counts['Open Bibliotheek']}, Gesloten: {counts['Gesloten Bibliotheek']}, Invalide: {counts['Foutieve Bieb']}."
                })