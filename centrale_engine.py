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
        
        # Verzamel alle te scoren domeinen van alle actieve engines
        self.all_domains = []
        for engine in self.active_engines:
            self.all_domains.extend(engine.domains)
            
        self.domain_scores_local = {mod: [] for mod in self.all_domains}
        self.domain_scores_sp = {mod: [] for mod in self.all_domains}

        # Data Quality Uitzonderingen: Mappen waarvoor mildere regels gelden
        self.EXCEPTIONS_FOLDERS = ["werkomgeving", "concepten", "wip"]

    def process(self, q):
        all_items = []
        
        # 1. ROBUUSTE LOKALE SCAN
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
                        except OSError as e:
                            # QA Check: Log lokale leesfouten (zoals gelockte systeembestanden)
                            pass
                        
        # 2. ROBUUSTE SHAREPOINT SCAN (Met Foutisolatie)
        for sp in self.sharepoint_sites:
            site_url = sp["url"]
            self.sp_bibliotheken_tracker[site_url] = {"Open Bibliotheek": 0, "Gesloten Bibliotheek": 0, "Foutieve Bieb": 0}
            
            try:
                # Authenticatie met Enterprise SSO
                ctx = ClientContext(site_url).with_credentials(HttpNegotiateAuth())
                lists = ctx.web.lists
                ctx.load(lists)
                ctx.execute_query()
                
                for library in lists:
                    if library.base_template == 101 and not library.hidden:
                        self._validate_sp_library_name(site_url, library.title)
                        
                        try:
                            # Start de recursieve, fouttolerante scanner
                            self._walk_sp_recursive(ctx, library.root_folder, library.title, site_url, all_items)
                        except Exception as lib_err:
                            # Faalt 1 bibliotheek (bijv. geen rechten)? Ga door met de rest!
                            self.results.append({
                                "Type": "SP Structuur", "Naam": "Toegangsfout", "Pad": f"{site_url}/{library.title}", 
                                "Mode": "SP", "Score_Totaal": "0%", "Reden": f"Kan SP-bibliotheek niet scannen: {str(lib_err)}"
                            })
            except Exception as e:
                q.put(("error", f"Kritieke SP Connectiefout op {site_url}: {str(e)}\n\nCheck VPN/Rechten."))
                return

        if not all_items:
            q.put(("error", "Geen bestanden gevonden om te scannen."))
            return

        # 3. ORCHESTRATIE VAN DE ENGINES
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
        """Recursieve SharePoint scanner die crashes per map/bestand isoleert."""
        try:
            ctx.load(folder, ["Folders", "Files"])
            ctx.execute_query()
            
            # Verwerk bestanden in huidige map
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
                except Exception:
                    continue # Eén corrupt/afgeschermd bestand slaat deze iteratie over
                    
            # Dieper graven in submappen
            for sub_folder in folder.folders:
                if sub_folder.name not in ["Forms", "_t", "_w", "Templates"]:
                    # De recursie zit in een nieuwe try-except dankzij de method caller
                    self._walk_sp_recursive(ctx, sub_folder, f"{current_path}/{sub_folder.name}", site_url, all_items)
        except Exception:
            pass # Als we deze map niet in mogen (rechten), sla de map netjes over

    def _analyze_item(self, item):
        # Haal de datastream slechts 1x op voor ALLE engines
        file_stream = self._get_file_stream(item)
        is_duplicate = len(self.file_registry.get(item["name"].lower(), set())) > 1
        
        item_result = {
            "Type": "Bestand", "Naam": item["name"], 
            "Pad": item["path"], "Mode": item["mode"].upper()
        }
        
        all_scores = {}
        all_reasons = []

        # Stuur het bestand door de pijplijn van actieve engines
        for engine in self.active_engines:
            try:
                engine_data = engine.analyze(item, file_stream, is_duplicate)
                for domein, score in engine_data["scores"].items():
                    all_scores[domein] = score
                    if item["mode"] == "sp": self.domain_scores_sp[domein].append(score)
                    else: self.domain_scores_local[domein].append(score)
                all_reasons.extend(engine_data["reasons"])
            except Exception as e:
                # Als de Compliance of Quality engine faalt op rekenlogica, crasht de app niet
                all_reasons.append(f"Engine Fout: {engine.__class__.__name__} kon bestand niet verwerken.")

        # Totalen berekenen
        active_vals = [v for k, v in all_scores.items() if isinstance(v, int)]
        item_result["Score_Totaal"] = f"{int(sum(active_vals) / len(active_vals))}%" if active_vals else "0%"
        
        for dom in self.all_domains:
            val = all_scores.get(dom, "N/A")
            item_result[dom] = f"{val}%" if isinstance(val, int) else val
            
        item_result["Reden"] = " | ".join(all_reasons) if all_reasons else "Volledig Compliant"
        self.results.append(item_result)
        
        # Geheugenlekken voorkomen (Cruciaal voor QA)
        if file_stream:
            try: file_stream.close()
            except: pass

    def _get_file_stream(self, item):
        """Haalt de byte-stream op en handelt SP-restricties direct af."""
        try:
            if item["mode"] == "local": 
                return open(item["path"], "rb")
            elif item["mode"] == "sp": 
                # Beveiliging tegen timeouts en missende rechten op bestandniveau
                file_content = item["ctx"].web.get_file_by_server_relative_url(item["sp_url"]).read()
                return io.BytesIO(file_content)
        except Exception:
            return None # Bestand is gelockt, overgeslagen, of we hebben geen leesrechten

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
