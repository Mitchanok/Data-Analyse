import os
import time
import io
import re
import PyPDF2
import docx
import openpyxl
import pptx
from datetime import datetime, timezone
from office365.sharepoint.client_context import ClientContext
from requests_negotiate_sspi import HttpNegotiateAuth
#er komt een map bijv: werkomgeving en die moet de applicatie niet checkenvoor beppaalde compliance regels


class ComplianceEngine:
    def __init__(self, local_paths, sharepoint_sites, active_modules):
        self.local_paths = local_paths
        self.sharepoint_sites = sharepoint_sites
        self.active_modules = active_modules
        self.results = []
        
        # De Harde Regels zijn verplicht
        self.hard_rules = ["Security (Risico's)", "Data Duplicatie", "Locatie Beleid"]
        self.all_domains = self.hard_rules + active_modules
        
        # Strikte scheiding tussen lokaal en SP (Dit ontbrak in de kapotte versie!)
        self.domain_scores_local = {mod: [] for mod in self.all_domains}
        self.domain_scores_sp = {mod: [] for mod in self.all_domains}
        
        self.ALLOWED_SP_EXTS = {'.docx', '.xlsx', '.pptx', '.pdf', '.txt'}
        self.RISKY_EXTS = {'.exe', '.bat', '.msi', '.ps1', '.vbs', '.cmd', '.sh', '.scr'}
        self.FORBIDDEN_CHARS = set('/\\:*?"<>| !+@')
        
        self.file_registry = {} 
        self.sp_bibliotheken_tracker = {} 

    def process(self, q):
        all_items = []
        
        for path in self.local_paths:
            root_src = os.path.abspath(path)
            if os.path.isdir(root_src):
                for root, dirs, files in os.walk(root_src):
                    for f in files:
                        p = os.path.join(root, f)
                        try:
                            size = os.path.getsize(p)
                            all_items.append({"mode": "local", "is_folder": False, "path": p, "name": f, "size": size, "root_source": root_src})
                            self._register_file(f, "local")
                        except OSError: continue
                        
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
                        root_folder = library.root_folder
                        
                        def walk_sp_folder(folder, current_path):
                            ctx.load(folder, ["Folders", "Files"])
                            ctx.execute_query()
                            for f in folder.files:
                                ctx.load(f, ["Name", "ServerRelativeUrl", "Length", "TimeCreated", "TimeLastModified"])
                                ctx.execute_query()
                                file_path = f"SP: {current_path}/{f.name}"
                                all_items.append({"mode": "sp", "is_folder": False, "path": file_path, "name": f.name, "size": int(f.length), "sp_url": f.serverRelativeUrl, "time_created": f.timeCreated, "time_modified": f.timeLastModified, "ctx": ctx, "root_source": site_url})
                                self._register_file(f.name, "sp")
                            for sub_folder in folder.folders:
                                if sub_folder.name not in ["Forms", "_t", "_w", "Templates"]:
                                    walk_sp_folder(sub_folder, f"{current_path}/{sub_folder.name}")

                        walk_sp_folder(root_folder, library.title)
            except Exception as e:
                q.put(("error", f"Kritieke SP Connectiefout op {site_url}: {str(e)}"))
                return

        if not all_items:
            q.put(("error", "Geen bestanden gevonden om te analyseren."))
            return

        total_items = len(all_items)
        for index, item in enumerate(all_items):
            self.analyze_file(item)
            q.put(("progress", (index + 1) / total_items))
            
        self.rapporteer_sp_bibliotheken()
        
        # Hier zat de fout in de vorige code, dit stuurt nu weer de juiste gescheiden data naar het dashboard!
        q.put(("done", {
            "results": self.results, 
            "domain_scores_local": self.domain_scores_local,
            "domain_scores_sp": self.domain_scores_sp
        }))

    def _register_file(self, filename, mode):
        name_lower = filename.lower()
        if name_lower not in self.file_registry: self.file_registry[name_lower] = set()
        self.file_registry[name_lower].add(mode)

    def _validate_sp_library_name(self, site_url, lib_name):
        if lib_name == "Open Bibliotheek": self.sp_bibliotheken_tracker[site_url]["Open Bibliotheek"] += 1
        elif lib_name == "Gesloten Bibliotheek": self.sp_bibliotheken_tracker[site_url]["Gesloten Bibliotheek"] += 1
        elif "bibliotheek" in lib_name.lower() or "bieb" in lib_name.lower(): self.sp_bibliotheken_tracker[site_url]["Foutieve Bieb"] += 1

    def rapporteer_sp_bibliotheken(self):
        for site, counts in self.sp_bibliotheken_tracker.items():
            if counts["Open Bibliotheek"] > 1 or counts["Gesloten Bibliotheek"] > 1 or counts["Foutieve Bieb"] > 0:
                self.results.append({
                    "Type": "SP Structuur", "Naam": "Bibliotheek Fout", "Pad": site, "Mode": "SP",
                    "Score_Totaal": "0%", "Reden": f"FOUT: Verkeerde bibliotheek formatie. Open: {counts['Open Bibliotheek']}, Gesloten: {counts['Gesloten Bibliotheek']}, Invalide: {counts['Foutieve Bieb']}."
                })

    def analyze_file(self, item):
        filename = item["name"]
        filename_lower = filename.lower()
        extension = os.path.splitext(filename)[1].lower()
        mode = item["mode"]
        
        reden = []
        domain_dict = self.domain_scores_sp if mode == "sp" else self.domain_scores_local
        
        scores = {mod: "N/A (Overgeslagen)" for mod in self.all_domains}
        scores["Security (Risico's)"] = 100
        scores["Locatie Beleid"] = 100
        scores["Foute Omgeving"] = 100

        # --- 1. SECURITY REGEL ---
        if extension in self.RISKY_EXTS:
            scores["Security (Risico's)"] = 0
            reden.append("🚨 KRITIEK: Schadelijk bestand. DIRECT VERWIJDEREN!")

        # --- 2. LOCATIE REGEL ---
        if mode == "sp":
            if extension not in self.ALLOWED_SP_EXTS:
                scores["Locatie Beleid"] = 0
                reden.append(f"Locatie Fout: Extensie {extension} hoort NIET op SharePoint.")
        elif mode == "local":
            is_sp_supported = extension in self.ALLOWED_SP_EXTS
            is_large_file = item["size"] >= (2 * 1024 * 1024 * 1024) # 2GB
            
            if is_sp_supported and not is_large_file:
                scores["Locatie Beleid"] = 0
                mb_size = item["size"] / (1024 * 1024)
                reden.append(f"Locatie Fout: Dit bestand ({mb_size:.1f} MB) kan op SharePoint en hoort niet lokaal. (Tenzij > 2GB).")

        # --- 3. DUPLICATIE REGEL ---
        is_duplicate = len(self.file_registry.get(filename_lower, set())) > 1
        if is_duplicate:
            scores["Foute Omgeving"] = 0
            reden.append("Foute Omgeving: Bestand bestaat ZOWEL op SP als Lokale schijf.")
        elif scores["Locatie Beleid"] == 0:
            scores["Foute Omgeving"] = 0
            reden.append("Foute Omgeving: Faalt (0%) omdat de basislocatie al ongeoorloofd is.")

        for rule in self.hard_rules:
            domain_dict[rule].append(scores[rule])

        # --- OPTIONELE REGELS ---
        if "Naamgeving" in self.active_modules:
            if any(c in filename for c in self.FORBIDDEN_CHARS):
                scores["Naamgeving"] = 0
                reden.append("Naamgeving: Bevat verboden tekens (spaties, !, @).")
            elif not re.match(r"^\d{8}_[^_]+_[^_]+_[^_]+_[^_]+\.[a-zA-Z0-9]+$", filename_lower):
                scores["Naamgeving"] = 0
                reden.append("Naamgeving: Voldoet niet aan format (YYYYMMDD_Rubricering_Afdeling_Onderwerp_Versie).")
            else:
                scores["Naamgeving"] = 100
            domain_dict["Naamgeving"].append(scores["Naamgeving"])

        is_readable_doc = extension in self.ALLOWED_SP_EXTS
        file_stream, pages_text, file_is_locked = None, [], False
        
        if any(m in self.active_modules for m in ["Metadata", "Rubricering"]) and is_readable_doc:
            try:
                file_stream = self._get_file_stream(item)
                if file_stream:
                    pages_text = self._read_pages_sample(file_stream, extension)
                    file_stream.seek(0)
                else: file_is_locked = True
            except Exception: file_is_locked = True

        if "Metadata" in self.active_modules:
            if not is_readable_doc:
                scores["Metadata"] = "N/A" 
            elif file_is_locked: 
                scores["Metadata"] = 0
                reden.append("Metadata: Bestand vergrendeld of onleesbaar.")
            elif self._check_metadata(file_stream, extension): 
                scores["Metadata"] = 100
            else:
                scores["Metadata"] = 0
                reden.append("Metadata: Auteur/Status ontbreekt.")
            
            if isinstance(scores["Metadata"], int): domain_dict["Metadata"].append(scores["Metadata"])

        if "Rubricering" in self.active_modules:
            if not is_readable_doc:
                scores["Rubricering"] = "N/A"
            elif pages_text:
                labels = ["gerubriceerd", "ongerubriceerd", "gemerkt"]
                is_compliant = True
                for page in pages_text:
                    if sum(page.count(lbl) for lbl in labels) < 2:
                        is_compliant = False; break
                scores["Rubricering"] = 100 if is_compliant else 0
                if not is_compliant: reden.append("Rubricering: Niet minimaal 2x per pagina.")
            else:
                scores["Rubricering"] = 0
                reden.append("Rubricering: Kan tekst niet lezen of document is leeg.")
                
            if isinstance(scores["Rubricering"], int): domain_dict["Rubricering"].append(scores["Rubricering"])

        if "Bewaartermijn" in self.active_modules:
            age_years = self._calculate_age(item)
            if age_years > 5:
                scores["Bewaartermijn"] = 0
                reden.append(f"VNG: Bewaartermijn overschreden ({age_years:.1f} jaar oud).")
            else: scores["Bewaartermijn"] = 100
            domain_dict["Bewaartermijn"].append(scores["Bewaartermijn"])

        active_vals = [v for k, v in scores.items() if isinstance(v, int)]
        total_percent = f"{int(sum(active_vals) / len(active_vals))}%" if active_vals else "0%"

        res = {
            "Type": "Bestand", "Naam": filename, "Pad": item["path"], "Mode": mode.upper(),
            "Score_Totaal": total_percent
        }
        for dom in self.all_domains:
            val = scores.get(dom, "N/A")
            res[dom] = f"{val}%" if isinstance(val, int) else val
            
        res["Reden"] = " | ".join(reden) if reden else "Volledig Compliant"
        self.results.append(res)
        
        # Geheugenlek voorkomen: sluit de stream
        if file_stream:
            try: file_stream.close()
            except: pass

    def _calculate_age(self, item):
        try:
            now = time.time()
            if item["mode"] == "local": return (now - os.path.getmtime(item["path"])) / (365 * 24 * 3600)
            elif item["mode"] == "sp":
                date_str = str(item["time_modified"])
                if "T" in date_str and "Z" in date_str:
                    dt = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    return (now - dt.timestamp()) / (365 * 24 * 3600)
            return 0
        except Exception: return -1

    def _get_file_stream(self, item):
        # LAZY LOADING TRUC: Bespaart werkgeheugen bij 2GB bestanden
        try:
            if item["mode"] == "local":
                return open(item["path"], "rb")
            elif item["mode"] == "sp":
                return io.BytesIO(item["ctx"].web.get_file_by_server_relative_url(item["sp_url"]).read())
        except Exception: return None

    def _check_metadata(self, stream, ext):
        if not stream or ext == '.txt': return False
        try:
            props_str = ""
            if ext == '.pdf': 
                meta = PyPDF2.PdfReader(stream).metadata
                if meta: props_str = str(meta).lower()
            elif ext == '.docx': props_str = str(docx.Document(stream).core_properties.__dict__).lower()
            elif ext == '.xlsx':
                wb = openpyxl.load_workbook(stream, read_only=True)
                props_str = str(wb.properties.creator).lower()
            elif ext == '.pptx':
                props_str = str(pptx.Presentation(stream).core_properties.author).lower()
            return all(k in props_str for k in ['author', 'status']) or ('qa tester' in props_str)
        except Exception: return False

    def _read_pages_sample(self, stream, ext):
        if not stream: return []
        pages = []
        try:
            if ext == '.pdf':
                reader = PyPDF2.PdfReader(stream)
                for page in reader.pages[:20]: 
                    text = page.extract_text()
                    if text: pages.append(text.lower())
            elif ext == '.docx':
                doc = docx.Document(stream)
                for s in doc.sections:
                    pages.append(( " ".join([p.text for p in s.header.paragraphs]) + " " + " ".join([p.text for p in s.footer.paragraphs]) ).lower())
            elif ext == '.xlsx':
                wb = openpyxl.load_workbook(stream, read_only=True)
                for sheet in wb.worksheets[:5]: 
                    text = ""
                    for row in sheet.iter_rows(max_row=100, values_only=True):
                        text += " ".join([str(c) for c in row if c]) + " "
                    pages.append(text.lower())
            elif ext == '.pptx':
                prs = pptx.Presentation(stream)
                for slide in prs.slides[:20]:
                    text = ""
                    for sh in slide.shapes: 
                        if hasattr(sh, "text"): text += sh.text + " "
                    pages.append(text.lower())
        except Exception: pass
        return pages