import os
import io
import time
import logging
import re
import csv
import requests
import urllib3
from datetime import datetime
from typing import List, Dict, Any, Optional
from requests_negotiate_sspi import HttpNegotiateAuth

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class CentraleEngine:
    def __init__(self, local_paths: List[str], sharepoint_sites: List[Dict[str, str]], active_engines: List[Any]):
        self.local_paths = local_paths
        self.sharepoint_sites = sharepoint_sites
        self.active_engines = active_engines
        
        self.results = []
        self.file_registry = {} 
        self.sp_bibliotheken_tracker = {}

        # --- DATA QUALITY & SECURITY BASELINES ---
        self.ALLOWED_SP_EXTS = {'.docx', '.xlsx', '.pptx', '.pdf', '.txt'}
        self.RISKY_EXTS = {'.exe', '.bat', '.msi', '.ps1', '.vbs', '.cmd', '.sh', '.scr'}
        self.FORBIDDEN_CHARS = set('/\\:*?"<>| !+@')
        
        self.base_domains = ["Security (Risico's)", "Data Duplicatie", "Locatie Beleid"]
        self.all_domains = list(self.base_domains)
        for engine in self.active_engines:
            if hasattr(engine, 'domains'):
                self.all_domains.extend(engine.domains)
            
        self.domain_scores_local = {mod: [] for mod in self.all_domains}
        self.domain_scores_sp = {mod: [] for mod in self.all_domains}
        self.EXCEPTIONS_FOLDERS = ["werkomgeving", "concepten", "wip"]

    def _sanitize_url(self, url: str) -> str:
        url = url.strip()
        url = re.split(r'/(SitePages|Shared Documents|Forms|Lists)/', url, flags=re.IGNORECASE)[0]
        url = re.sub(r'/[^/]+\.aspx$', '', url, flags=re.IGNORECASE)
        return url.rstrip('/')

    def process(self, q: Any) -> None:
        os.environ['trust_env'] = '0'
        all_items = []
        
        for path in self.local_paths:
            self._scan_local(path, all_items)

        for sp in self.sharepoint_sites:
            site_url = self._sanitize_url(sp.get("url", ""))
            self.sp_bibliotheken_tracker[site_url] = {"Open Bibliotheek": 0, "Gesloten Bibliotheek": 0, "Foutieve Bieb": 0}
            
            try:
                session = requests.Session()
                session.auth = HttpNegotiateAuth()
                session.verify = False 
                session.trust_env = False
                session.headers.update({'Accept': 'application/json;odata=verbose'})

                lists_api = f"{site_url}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$expand=RootFolder"
                response = session.get(lists_api)
                
                if response.status_code == 401:
                    raise Exception("Windows Login geweigerd. SSPI faalde (401).")
                elif response.status_code != 200:
                    raise Exception(f"Fout bij ophalen lijsten: HTTP {response.status_code} - {response.text}")

                libraries = response.json()['d']['results']
                logging.info(f"VDI Seamless SSO Succesvol! {len(libraries)} bibliotheken gevonden op {site_url}")

                for library in libraries:
                    lib_title = library.get('Title', 'Onbekende Bieb')
                    root_folder_url = library['RootFolder']['ServerRelativeUrl']
                    self._validate_sp_library_name(site_url, lib_title)
                    self._walk_sp_recursive_native(session, root_folder_url, lib_title, site_url, all_items)
                        
            except Exception as e:
                error_msg = f"Kritieke SP Netwerkfout op {site_url}: {str(e)}"
                logging.error(error_msg)
                q.put(("error", error_msg))
                return

        # --- ORCHESTRATIE & ANALYSE ---
        total_items = len(all_items)
        if total_items == 0 and not self.results:
            q.put(("error", "Geen bestanden gevonden om te analyseren."))
            return

        for index, item in enumerate(all_items):
            self._analyze_item(item)
            q.put(("progress", (index + 1) / total_items))
            
        self._rapporteer_sp_bibliotheken()
        
        # 🚨 AUTOMATISCHE CSV EXPORT AANROEPEN VOORAFGAAND AAN AFRONDING
        self._auto_export_csv()
        
        q.put(("done", {
            "results": self.results, 
            "domain_scores_local": self.domain_scores_local,
            "domain_scores_sp": self.domain_scores_sp
        }))

    def _auto_export_csv(self):
        """Maakt automatisch een CSV dump van self.results na de scan."""
        if not self.results:
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Scan_Rapport_{timestamp}.csv"
        scan_datum_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        risky_chars = ('=', '+', '-', '@', '\t', '\r')
        sanitized_results = []
        
        for row in self.results:
            sanitized_row = {'ScanDatum': scan_datum_str}
            for k, v in row.items():
                str_val = str(v)
                if str_val.startswith(risky_chars):
                    str_val = "'" + str_val
                sanitized_row[k] = str_val
            sanitized_results.append(sanitized_row)

        headers = ['ScanDatum'] + [k for k in self.results[0].keys() if k != 'ScanDatum']
        
        try:
            with open(filename, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=headers, delimiter=';')
                writer.writeheader()
                writer.writerows(sanitized_results)
            logging.info(f"Automatische CSV export geslaagd: {filename}")
        except Exception as e:
            logging.error(f"Automatische CSV export gefaald: {e}")

    def _walk_sp_recursive_native(self, session: requests.Session, folder_url: str, current_path: str, site_url: str, all_items: list) -> None:
        try:
            safe_folder_url = folder_url.replace("'", "''")
            files_api = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{safe_folder_url}')/Files"
            r_files = session.get(files_api)
            
            if r_files.status_code == 200:
                files_data = r_files.json()['d']['results']
                for f in files_data:
                    name = f.get('Name', '')
                    file_path = f"SP: {current_path}/{name}"
                    
                    item = {
                        "mode": "sp", "path": file_path, "name": name, 
                        "size": int(f.get('Length', 0)), "sp_url": f.get('ServerRelativeUrl', ''), 
                        "time_created": f.get('TimeCreated', ''), "time_modified": f.get('TimeLastModified', ''), 
                        "session": session, "root_source": site_url,
                        "in_werkomgeving": any(exc in file_path.lower() for exc in self.EXCEPTIONS_FOLDERS),
                        "extension": os.path.splitext(name)[1].lower()
                    }
                    all_items.append(item)
                    self._register_file(name, "sp")

            folders_api = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{safe_folder_url}')/Folders"
            r_folders = session.get(folders_api)
            
            if r_folders.status_code == 200:
                folders_data = r_folders.json()['d']['results']
                
                if len(folders_data) == 0 and ('r_files' in locals() and r_files.status_code == 200 and len(r_files.json()['d']['results']) == 0):
                    self._add_result({
                        "Type": "Structuur", "Naam": current_path.split('/')[-1], "Pad": f"SP: {current_path}", 
                        "Mode": "SP", "Score_Totaal": "0%", "Reden": "Data Vervuiling: Lege SharePoint map."
                    })

                for sub in folders_data:
                    sub_name = sub.get('Name', '')
                    if sub_name not in ["Forms", "_t", "_w", "Templates"]:
                        self._walk_sp_recursive_native(session, sub.get('ServerRelativeUrl', ''), f"{current_path}/{sub_name}", site_url, all_items)
        except Exception as e:
            logging.warning(f"Scan-onderbreking in SP map {current_path}: {e}")

    def _get_file_stream(self, item: Dict[str, Any]) -> Optional[io.BytesIO]:
        try:
            if item["mode"] == "local":
                return open(item["path"], "rb")
            else:
                safe_sp_url = item["sp_url"].replace("'", "''")
                download_api = f"{item['root_source']}/_api/web/GetFileByServerRelativeUrl('{safe_sp_url}')/$value"
                response = item["session"].get(download_api)
                if response.status_code == 200:
                    return io.BytesIO(response.content)
                else:
                    return None
        except Exception as e:
            return None

    def _analyze_item(self, item: Dict[str, Any]) -> None:
        file_stream = None
        try:
            file_stream = self._get_file_stream(item)
            
            filename = item["name"]
            extension = item["extension"]
            mode = item["mode"]
            is_duplicate = len(self.file_registry.get(filename.lower(), set())) > 1
            
            item["is_duplicate"] = is_duplicate
            item["has_forbidden_chars"] = any(c in filename for c in self.FORBIDDEN_CHARS)
            item["is_readable_doc"] = extension in self.ALLOWED_SP_EXTS
            
            all_scores = {"Security (Risico's)": 100, "Locatie Beleid": 100, "Data Duplicatie": 100}
            all_reasons = []

            if extension in self.RISKY_EXTS:
                all_scores["Security (Risico's)"] = 0
                all_reasons.append("🚨 KRITIEK: Schadelijk bestandstype.")

            if mode == "sp" and extension not in self.ALLOWED_SP_EXTS:
                all_scores["Locatie Beleid"] = 0
                all_reasons.append(f"Locatie: {extension} niet toegestaan op SharePoint.")
            elif mode == "local":
                if extension in self.ALLOWED_SP_EXTS and item["size"] < (2 * 1024**3):
                    all_scores["Locatie Beleid"] = 0
                    all_reasons.append("Locatie: Dit bestand hoort op SharePoint.")

            if is_duplicate:
                all_scores["Data Duplicatie"] = 0
                all_reasons.append("Duplicatie: Bestand staat zowel lokaal als op SP.")

            for engine in self.active_engines:
                try:
                    engine_data = engine.analyze(item, file_stream)
                    for dom, score in engine_data.get("scores", {}).items():
                        all_scores[dom] = score
                    all_reasons.extend(engine_data.get("reasons", []))
                except Exception as e:
                    all_reasons.append(f"Engine Crash ({engine.__class__.__name__}): {e}")

            res = {
                "Type": "Bestand", "Naam": filename, "Pad": item["path"], "Mode": mode.upper()
            }
            
            vals = [v for v in all_scores.values() if isinstance(v, int)]
            res["Score_Totaal"] = f"{int(sum(vals)/len(vals))}%" if vals else "0%"
            
            for dom in self.all_domains:
                score = all_scores.get(dom, "N/A")
                res[dom] = f"{score}%" if isinstance(score, int) else score
            
            res["Reden"] = " | ".join(all_reasons) if all_reasons else "Volledig Compliant"
            
            self._add_result(res)

            target_dict = self.domain_scores_sp if mode == "sp" else self.domain_scores_local
            for dom, score in all_scores.items():
                if dom in target_dict: target_dict[dom].append(score)

        finally:
            if file_stream:
                try: file_stream.close()
                except: pass

    def _add_result(self, result: Dict[str, str]) -> None:
        """Slaat het resultaat exclusief op in werkgeheugen."""
        self.results.append(result)

    def _scan_local(self, path: str, all_items: list) -> None:
        root_src = os.path.abspath(path)
        if not os.path.isdir(root_src): return
        for root, dirs, files in os.walk(root_src):
            if not dirs and not files:
                self._add_result({"Type": "Structuur", "Naam": os.path.basename(root), "Pad": root, "Mode": "LOCAL", "Score_Totaal": "0%", "Reden": "Lege map."})
                continue
            for f in files:
                p = os.path.join(root, f)
                all_items.append({
                    "mode": "local", "path": p, "name": f, "size": os.path.getsize(p), "root_source": root_src,
                    "in_werkomgeving": any(exc in p.lower() for exc in self.EXCEPTIONS_FOLDERS),
                    "extension": os.path.splitext(f)[1].lower()
                })
                self._register_file(f, "local")

    def _register_file(self, filename: str, mode: str) -> None:
        name_lower = filename.lower()
        if name_lower not in self.file_registry: self.file_registry[name_lower] = set()
        self.file_registry[name_lower].add(mode)

    def _validate_sp_library_name(self, site_url: str, lib_name: str) -> None:
        if lib_name == "Open Bibliotheek": self.sp_bibliotheken_tracker[site_url]["Open Bibliotheek"] += 1
        elif lib_name == "Gesloten Bibliotheek": self.sp_bibliotheken_tracker[site_url]["Gesloten Bibliotheek"] += 1
        elif "bibliotheek" in lib_name.lower() or "bieb" in lib_name.lower(): self.sp_bibliotheken_tracker[site_url]["Foutieve Bieb"] += 1

    def _rapporteer_sp_bibliotheken(self) -> None:
        for site, counts in self.sp_bibliotheken_tracker.items():
            if counts["Open Bibliotheek"] > 1 or counts["Gesloten Bibliotheek"] > 1 or counts["Foutieve Bieb"] > 0:
                self._add_result({"Type": "SP Structuur", "Naam": "Bieb Fout", "Pad": site, "Mode": "SP", "Score_Totaal": "0%", "Reden": "Invalide bieb formatie."})