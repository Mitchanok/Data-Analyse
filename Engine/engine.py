class User:
    def __init__(self, username, is_admin=False):
        self.username = username
        self.is_admin = is_admin

import sqlite3
import os
import hashlib

import sys

def get_app_dir():
    if getattr(sys, 'frozen', False): return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def get_data_dir():
    appdata = os.getenv('APPDATA')
    if appdata:
        p = os.path.join(appdata, 'DocumentScanner')
        os.makedirs(p, exist_ok=True)
        return p
    return get_app_dir()

DB_PATH = os.path.join(get_data_dir(), "app_data.db")

def _hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Users table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    
    # Scans table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            scan_datum TEXT NOT NULL,
            afdeling TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')
    
    # Scan Results table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scan_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_id INTEGER,
            locatie TEXT,
            bestandsnaam TEXT,
            score TEXT,
            reden TEXT,
            FOREIGN KEY (scan_id) REFERENCES scans (id)
        )
    ''')
    
    # Create default admin if no users exist
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        create_user("admin", "admin", "Admin")
        
    conn.commit()
    conn.close()

def create_user(username, password, role="User"):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", 
                       (username, _hash_password(password), role))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def verify_user(username, password):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role FROM users WHERE username=? AND password_hash=?", 
                   (username, _hash_password(password)))
    user = cursor.fetchone()
    conn.close()
    if user:
        return {"id": user[0], "username": user[1], "role": user[2], "is_admin": user[2] == "Admin"}
    return None

def save_scan(user_id, afdeling, scan_datum, results):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO scans (user_id, scan_datum, afdeling) VALUES (?, ?, ?)", 
                   (user_id, scan_datum, afdeling))
    scan_id = cursor.lastrowid
    
    for r in results:
        cursor.execute("INSERT INTO scan_results (scan_id, locatie, bestandsnaam, score, reden) VALUES (?, ?, ?, ?, ?)",
                       (scan_id, r.get("Pad", ""), r.get("Naam", ""), str(r.get("Score_Totaal", "")), r.get("Reden", "")))
    
    conn.commit()
    conn.close()

import time
from datetime import datetime, timezone
import os

class QualityEngine:
    def __init__(self, active_modules):
        self.domains = active_modules

    def analyze(self, item, stream):
        scores = {mod: "N/A" for mod in self.domains}
        reden = []
        
        # 1. Bestandsgrootte Check
        if "Bestandsgrootte" in self.domains:
            size_b = item.get("size", 0)
            if size_b == 0:
                scores["Bestandsgrootte"] = 0
                reden.append("Grootte: Leeg bestand (0 bytes).")
            elif size_b > (2 * 1024 * 1024 * 1024): # 2 GB
                scores["Bestandsgrootte"] = 0
                reden.append("Grootte: Bestand is groter dan 2GB.")
            else:
                scores["Bestandsgrootte"] = 100

        # 2. Actualiteit Check (Ouder dan 3 jaar = aftrek)
        if "Actualiteit" in self.domains:
            age_years = self._calculate_age(item)
            if age_years > 5:
                scores["Actualiteit"] = 0
                reden.append(f"Actualiteit: Zeer oud archiefbestand ({age_years:.1f} jr).")
            elif age_years > 3:
                scores["Actualiteit"] = 50
                reden.append(f"Actualiteit: Bestand ouder dan 3 jaar ({age_years:.1f} jr).")
            elif age_years >= 0:
                scores["Actualiteit"] = 100
            else:
                scores["Actualiteit"] = 0
                reden.append("Actualiteit: Onbekende wijzigingsdatum.")
                
        # 3. Leesbaarheid & Volledigheid
        # Basis check: Extensie en of we een stream hebben.
        if "Leesbaarheid" in self.domains:
            if stream is None and item.get("size", 0) > 0:
                scores["Leesbaarheid"] = 0
                reden.append("Leesbaarheid: Bestand is vergrendeld of onleesbaar.")
            else:
                scores["Leesbaarheid"] = 100
                
        if "Volledigheid" in self.domains:
            # We assume fullness if size > 1KB and it has an extension.
            if item.get("extension", "") == "":
                scores["Volledigheid"] = 0
                reden.append("Volledigheid: Geen bestandsextensie.")
            elif item.get("size", 0) < 1024:
                scores["Volledigheid"] = 50
                reden.append("Volledigheid: Zeer weinig inhoud (<1KB).")
            else:
                scores["Volledigheid"] = 100

        return {
            "scores": scores,
            "reasons": reden
        }
        
    def _calculate_age(self, item):
        try:
            now = time.time()
            if item["mode"] == "local": 
                return (now - item.get("time_modified", os.path.getmtime(item["path"]))) / (365 * 24 * 3600)
            elif item["mode"] == "sp":
                date_str = str(item.get("time_modified", ""))
                if "T" in date_str and "Z" in date_str:
                    dt = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    return (now - dt.timestamp()) / (365 * 24 * 3600)
            return 0
        except Exception: 
            return -1

import os
import io
import time
import re
from office365.sharepoint.client_context import ClientContext
from requests_negotiate_sspi import HttpNegotiateAuth

class CentraleEngine:
    def __init__(self, local_paths, sharepoint_sites, active_engines, stop_event=None):
        self.local_paths = local_paths
        self.sharepoint_sites = sharepoint_sites
        self.active_engines = active_engines
        self.stop_event = stop_event
        
        self.results = []
        self.file_registry_name = {} 
        self.file_registry_content = {}
        self.sp_bibliotheken_tracker = {}

        # --- DATA QUALITY & SECURITY BASELINES ---
        self.ALLOWED_SP_EXTS = {'.docx', '.xlsx', '.pptx', '.pdf', '.txt'}
        self.RISKY_EXTS = {'.exe', '.bat', '.msi', '.ps1', '.vbs', '.cmd', '.sh', '.scr'}
        self.FORBIDDEN_CHARS = set('/\\:*?"<>| !+@')
        
        self.base_domains = ["Security (Risico's)", "Data Duplicatie", "Locatie Beleid"]
        
        self.all_domains = list(self.base_domains)
        for engine in self.active_engines:
            self.all_domains.extend(engine.domains)
            
        self.domain_scores_local = {mod: [] for mod in self.all_domains}
        self.domain_scores_sp = {mod: [] for mod in self.all_domains}

        self.EXCEPTIONS_FOLDERS = ["werkomgeving", "concepten", "wip"]

    def process(self, q):
        all_items = []
        
        # 1. LOKALE SCAN (Met controle op lege mappen)
        for path in self.local_paths:
            if self.stop_event and self.stop_event.is_set():
                q.put(("canceled", "Geannuleerd tijdens lokale opsomming."))
                return

            root_src = os.path.abspath(path)
            if os.path.isdir(root_src):
                for root, dirs, files in os.walk(root_src):
                    if self.stop_event and self.stop_event.is_set():
                        q.put(("canceled", "Geannuleerd tijdens opsomming."))
                        return
                    # DATA QUALITY CHECK: Is de map volledig leeg?
                    if not dirs and not files:
                        map_naam = os.path.basename(root) or root
                        self.results.append({
                            "Type": "Structuur", "Naam": map_naam, "Pad": root, 
                            "Mode": "LOCAL", "Score_Totaal": "0%", 
                            "Reden": "Data Vervuiling: Lege map gedetecteerd. Verwijder deze om overzicht te behouden."
                        })
                        continue

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
                            self._register_file(item)
                        except OSError: pass
                        
        # 2. SHAREPOINT SCAN 
        for sp in self.sharepoint_sites:
            if self.stop_event and self.stop_event.is_set():
                q.put(("canceled", "Geannuleerd tijdens SharePoint connectie."))
                return

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

        if not all_items and not self.results:
            q.put(("error", "Geen bestanden of structuren gevonden om te scannen."))
            return

        # 3. ORCHESTRATIE & ANALYSE
        total_items = len(all_items)
        for index, item in enumerate(all_items):
            if self.stop_event and self.stop_event.is_set():
                q.put(("canceled", "Geannuleerd tijdens bestandsanalyse."))
                return
            
            self._analyze_item(item)
            if total_items > 0:
                q.put(("progress", (index + 1) / total_items))
            
        self._rapporteer_sp_bibliotheken()
        
        q.put(("done", {
            "results": self.results, 
            "domain_scores_local": self.domain_scores_local,
            "domain_scores_sp": self.domain_scores_sp
        }))

    def _walk_sp_recursive(self, ctx, folder, current_path, site_url, all_items):
        if self.stop_event and self.stop_event.is_set(): return
        
        try:
            ctx.load(folder, ["Folders", "Files"])
            ctx.execute_query()
            
            # DATA QUALITY CHECK: Is de SharePoint map volledig leeg?
            if len(folder.files) == 0 and len(folder.folders) == 0:
                map_naam = current_path.split('/')[-1] if '/' in current_path else current_path
                self.results.append({
                    "Type": "Structuur", "Naam": map_naam, "Pad": f"SP: {current_path}", 
                    "Mode": "SP", "Score_Totaal": "0%", 
                    "Reden": "Data Vervuiling: Lege map gedetecteerd op SharePoint. Ruim deze op."
                })
                return # Niets meer te doen in deze map

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
                    self._register_file(item)
                except Exception: continue
                    
            for sub_folder in folder.folders:
                if sub_folder.name not in ["Forms", "_t", "_w", "Templates"]:
                    self._walk_sp_recursive(ctx, sub_folder, f"{current_path}/{sub_folder.name}", site_url, all_items)
        except Exception: pass 

    def _analyze_item(self, item):
        file_stream = self._get_file_stream(item)
        
        filename = item["name"]
        base_name = self._get_base_filename(filename)
        extension = item["extension"]
        mode = item["mode"]
        size_key = f"{item.get('size', 0)}_{extension}"
        
        name_dups = [loc for loc in self.file_registry_name.get(base_name, []) if loc != item["path"]]
        content_dups = [loc for loc in self.file_registry_content.get(size_key, []) if loc != item["path"]] if item.get("size", 0) > 1024 else []
        
        is_duplicate = len(name_dups) > 0 or len(content_dups) > 0
        item["is_duplicate"] = is_duplicate
        item["has_forbidden_chars"] = any(c in filename for c in self.FORBIDDEN_CHARS)
        item["is_readable_doc"] = extension in self.ALLOWED_SP_EXTS
        
        all_scores = {"Security (Risico's)": 100, "Locatie Beleid": 100, "Data Duplicatie": 100}
        all_reasons = []

        if extension in self.RISKY_EXTS:
            all_scores["Security (Risico's)"] = 0
            all_reasons.append("🚨 KRITIEK: Schadelijk bestand.")

        if mode == "sp" and extension not in self.ALLOWED_SP_EXTS:
            all_scores["Locatie Beleid"] = 0
            all_reasons.append(f"Locatie: Extensie {extension} mag niet op SP.")
        elif mode == "local":
            is_large_file = item["size"] >= (2 * 1024 * 1024 * 1024)
            if extension in self.ALLOWED_SP_EXTS and not is_large_file:
                all_scores["Locatie Beleid"] = 0
                all_reasons.append("Locatie: Bestand kan op SP en hoort niet lokaal.")

        if is_duplicate:
            all_scores["Data Duplicatie"] = 0
            merged_dups = list(set(name_dups + content_dups))
            
            clean_locs = []
            for loc in merged_dups[:2]:
                parts = loc.replace('\\', '/').split('/')
                if len(parts) >= 2: clean_locs.append(f".../{parts[-2]}/{parts[-1]}")
                else: clean_locs.append(loc)
                
            loc_str = ", ".join(clean_locs)
            if len(merged_dups) > 2:
                loc_str += f" (en {len(merged_dups)-2} meer)"
                
            if len(content_dups) > 0:
                all_reasons.append(f"Duplicatie: Identieke inhoud (grootte {item.get('size')}B) in {loc_str}")
            else:
                all_reasons.append(f"Duplicatie: Zeer vergelijkbare bestandsnaam in {loc_str}")

        for domein in self.base_domains:
            if item["mode"] == "sp": self.domain_scores_sp[domein].append(all_scores[domein])
            else: self.domain_scores_local[domein].append(all_scores[domein])

        for engine in self.active_engines:
            try:
                engine_data = engine.analyze(item, file_stream)
                for domein, score in engine_data["scores"].items():
                    all_scores[domein] = score
                    if item["mode"] == "sp": self.domain_scores_sp[domein].append(score)
                    else: self.domain_scores_local[domein].append(score)
                all_reasons.extend(engine_data["reasons"])
            except Exception as e:
                all_reasons.append(f"Engine Fout ({engine.__class__.__name__}): {str(e)}")

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

    def _get_file_stream(self, item):
        try:
            if item["mode"] == "local": return open(item["path"], "rb")
            elif item["mode"] == "sp": return io.BytesIO(item["ctx"].web.get_file_by_server_relative_url(item["sp_url"]).read())
        except Exception: return None

    def _get_base_filename(self, filename):
        name, _ = os.path.splitext(filename)
        name = re.sub(r'^\d{8}[_ \-]*', '', name)
        name = re.sub(r'[_ \-]+[a-zA-Z]$', '', name)
        name = re.sub(r'[_ \-]*(v|versie|version)\s*\d+(\.\d+)?$', '', name, flags=re.IGNORECASE)
        name = re.sub(r'[_ \-]*(kopie|copy|definitief|final)(\s*\(\d+\))?$', '', name, flags=re.IGNORECASE)
        name = re.sub(r'[^a-zA-Z0-9]', '', name)
        return name.lower()

    def _register_file(self, item):
        base_name = self._get_base_filename(item["name"])
        path = item["path"]
        
        if base_name not in self.file_registry_name: self.file_registry_name[base_name] = []
        if path not in self.file_registry_name[base_name]:
            self.file_registry_name[base_name].append(path)
            
        size = item.get("size", 0)
        ext = item.get("extension", "")
        if size > 1024:
            size_key = f"{size}_{ext}"
            if size_key not in self.file_registry_content: self.file_registry_content[size_key] = []
            if path not in self.file_registry_content[size_key]:
                self.file_registry_content[size_key].append(path)

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