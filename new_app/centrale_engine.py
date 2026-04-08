# ==============================================================================
# engine.py — Gebruikersmodel, database-laag en CentraleEngine (orchestrator)
#
#   Verantwoordelijkheden:
#     • User            — sessie-model (ingelogde gebruiker)
#     • Database        — init, gebruikersbeheer, scan-opslag (SQLite)
#     • Beveiliging     — bcrypt-hashing, account-lockout, audit-log
#     • CentraleEngine  — orkestreert lokale + SharePoint scans en delegeert
#                         analyses naar engines uit engines.py
#
#   Beveiligingsmodel:
#     - Wachtwoorden opgeslagen als bcrypt-hash (nooit plaintext)
#     - Rate limiting: 30s blokkade na 3 foute pogingen
#     - Account lockout: geblokkeerd na 5 foute pogingen (10 min)
#     - Audit log: elke inlogpoging opgeslagen met tijdstip en resultaat
#     - Generieke foutmeldingen: nooit specifiek wat er fout is
#
#   Zie engines.py voor ComplianceEngine en QualityEngine.
# ==============================================================================

# --- Stdlib imports ---
import io
import os
import random
import re
import smtplib
import sqlite3
import string
import sys
import time
from datetime import datetime, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- Third-party imports ---
import bcrypt
from office365.sharepoint.client_context import ClientContext
from requests_negotiate_sspi import HttpNegotiateAuth


# ==============================================================================
# HELPERS
# ==============================================================================

def get_app_dir():
    """Geeft de map terug waar de applicatie draait (ook bij frozen/exe)."""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_data_dir():
    """Geeft de gebruikersdata-map terug (AppData/Roaming/DocumentScanner)."""
    appdata = os.getenv('APPDATA')
    if appdata:
        p = os.path.join(appdata, 'DocumentScanner')
        os.makedirs(p, exist_ok=True)
        return p
    return get_app_dir()


DB_PATH = os.path.join(get_data_dir(), "app_data.db")


# ==============================================================================
# USER MODEL
# ==============================================================================

class User:
    """Eenvoudig user-model voor sessie-beheer."""
    def __init__(self, username, is_admin=False):
        self.username = username
        self.is_admin = is_admin



# ------------------------------------------------------------------------------
# Beveiligingsconstanten
# ------------------------------------------------------------------------------

MAX_ATTEMPTS_RATE_LIMIT = 3    # Aantal pogingen voor 30s rate limiting
MAX_ATTEMPTS_LOCKOUT    = 5    # Aantal pogingen voor account lockout
RATE_LIMIT_SECONDS      = 30   # Wachttijd bij rate limiting
LOCKOUT_SECONDS         = 600  # Lockout-duur: 10 minuten


# ------------------------------------------------------------------------------
# Wachtwoord-hashing (bcrypt)
# ------------------------------------------------------------------------------

def _hash_password(password: str) -> str:
    """
    Hash een wachtwoord met bcrypt (work factor 12).
    Geeft een UTF-8 string terug die veilig opgeslagen kan worden.
    """
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt(rounds=12)).decode('utf-8')


def _verify_password(password: str, stored_hash: str) -> bool:
    """
    Vergelijk een plaintext wachtwoord met een bcrypt-hash.
    Gebruikt constant-time vergelijking om timing-aanvallen te voorkomen.
    """
    try:
        return bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8'))
    except Exception:
        return False


# ------------------------------------------------------------------------------
# Database initialisatie
# ------------------------------------------------------------------------------

def init_db():
    """
    Initialiseer de SQLite-database.
    Maakt alle tabellen aan als ze nog niet bestaan en zorgt voor een
    standaard admin-account bij een lege database.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Gebruikers — wachtwoord altijd als bcrypt-hash opgeslagen
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role          TEXT NOT NULL DEFAULT 'User'
        )
    ''')

    # Rollen — definities van toegangsniveaus
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS roles (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT UNIQUE NOT NULL,
            beschrijving TEXT
        )
    ''')

    # Account lockout — bijhouden van geblokkeerde accounts
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS account_lockout (
            username      TEXT PRIMARY KEY,
            attempt_count INTEGER NOT NULL DEFAULT 0,
            locked_until  REAL    NOT NULL DEFAULT 0
        )
    ''')

    # Audit log — elke inlogpoging vastgelegd
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            tijdstip  TEXT NOT NULL,
            username  TEXT NOT NULL,
            actie     TEXT NOT NULL,
            resultaat TEXT NOT NULL
        )
    ''')

    # Scan-geschiedenis
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scans (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER,
            scan_datum TEXT NOT NULL,
            afdeling   TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scan_results (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_id      INTEGER,
            locatie      TEXT,
            bestandsnaam TEXT,
            score        TEXT,
            reden        TEXT,
            FOREIGN KEY (scan_id) REFERENCES scans (id)
        )
    ''')

    # Vul standaard rollen
    for naam, beschrijving in [
        ("Admin", "Volledige toegang, CRUD-rechten, export naar PowerBI"),
        ("User",  "Beperkte toegang, alleen eigen scan-resultaten"),
    ]:
        cursor.execute(
            "INSERT OR IGNORE INTO roles (name, beschrijving) VALUES (?, ?)",
            (naam, beschrijving)
        )

    # Maak standaard admin-account als de database leeg is
    cursor.execute("SELECT COUNT(*) FROM users")
    needs_admin = cursor.fetchone()[0] == 0

    conn.commit()
    conn.close()

    if needs_admin:
        create_user("admin", "admin", "Admin")


# ------------------------------------------------------------------------------
# Gebruikersbeheer
# ------------------------------------------------------------------------------

def create_user(username: str, password: str, role: str = "User") -> bool:
    """
    Maak een nieuwe gebruiker aan met een bcrypt-gehashte wachtwoord.
    Geeft True terug bij succes, False als de gebruikersnaam al bestaat.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
            (username, _hash_password(password), role)
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def user_exists(username: str) -> bool:
    """
    Controleer of een gebruikersnaam bestaat in de database.
    Gebruikt voor stap-1 van progressive disclosure.

    Let op: geef deze uitkomst NOOIT direct terug aan de gebruiker
    als 'gebruiker bestaat' — dit kan account-enumeration mogelijk maken.
    Gebruik alleen intern voor UX-flow.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM users WHERE username = ?", (username,))
    exists = cursor.fetchone() is not None
    conn.close()
    return exists


# ------------------------------------------------------------------------------
# Beveiliging: audit log
# ------------------------------------------------------------------------------

def log_login_attempt(username: str, success: bool) -> None:
    """
    Log een inlogpoging in de audit-tabel.
    Wordt altijd aangeroepen, ongeacht succes of falen.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO audit_log (tijdstip, username, actie, resultaat) VALUES (?, ?, ?, ?)",
        (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            username,
            "LOGIN",
            "SUCCES" if success else "MISLUKT"
        )
    )
    conn.commit()
    conn.close()


def reset_admin() -> None:
    """
    Noodherstel: Zet het 'admin' wachtwoord terug naar bcrypt 'admin'
    en verwijdert eventuele lockouts.
    Wordt alleen aangeroepen via de CLI flag --reset-admin.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Update wachtwoord
    cursor.execute(
        "UPDATE users SET password_hash = ? WHERE username = 'admin'",
        (_hash_password("admin"),)
    )
    
    # Verwijder lockout
    cursor.execute("DELETE FROM account_lockout WHERE username = 'admin'")
    
    conn.commit()
    conn.close()


def send_recovery_email(username: str, target_email: str) -> tuple[bool, str]:
    """
    Genereert een tijdelijk wachtwoord, slaat dit direct op in de database
    als de nieuwe bcrypt-hash, en verstuurt dit via Gmail SMTP naar target_email.
    
    Returns:
        (succes_bool, bericht_string)
    """
    # 1. Controleer of gebruiker bestaat
    if not user_exists(username):
        return False, "Gebruiker niet gevonden in het systeem."

    # 2. Genereer een willekeurig tijdelijk wachtwoord (8 tekens)
    tekens = string.ascii_letters + string.digits
    temp_password = ''.join(random.choice(tekens) for i in range(8))

    # 3. Update in database & annuleer lockouts
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE users SET password_hash = ? WHERE username = ?",
            (_hash_password(temp_password), username)
        )
        cursor.execute("DELETE FROM account_lockout WHERE username = ?", (username,))
        conn.commit()
        conn.close()
    except Exception as e:
        return False, f"Databasefout bij het resetten: {e}"

    # 4. Email configuratie (WAARSCHUWING: Vul hier je Google App-Wachtwoord in!)
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SENDER_EMAIL = "mickstruijs@gmail.com"
    SENDER_PASSWORD = "VUL_HIER_JE_APP_WACHTWOORD_IN"

    if SENDER_PASSWORD == "VUL_HIER_JE_APP_WACHTWOORD_IN":
        return False, f"Wachtwoord is veilig in de database gereset naar '{temp_password}'. Pas SENDER_PASSWORD aan in engine.py om de mail daadwerkelijk te sturen."

    # 5. Email opmaken
    msg = MIMEMultipart()
    msg['From'] = f"Document Scanner Beveiliging <{SENDER_EMAIL}>"
    msg['To'] = target_email
    msg['Subject'] = "Wachtwoord Herstel - Document Scanner"

    body = (
        f"Beste Beheerder,\n\n"
        f"Er is een wachtwoord-reset aangevraagd voor het systeem.\n\n"
        f"Je nieuwe inloggegevens zijn:\n"
        f"Gebruikersnaam: {username}\n"
        f"Wachtwoord: {temp_password}\n\n"
        f"Let op: Je account is weer ontgrendeld. Log zo snel mogelijk in.\n"
    )
    msg.attach(MIMEText(body, 'plain'))

    # 6. Email verzenden
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True, f"Er is een e-mail gestuurd naar {target_email} met het nieuwe wachtwoord."
    except Exception as e:
        return False, f"Wachtwoord is '{temp_password}', maar e-mail kon niet verstuurd worden: {e}"


# ------------------------------------------------------------------------------
# Beveiliging: rate limiting & account lockout
# ------------------------------------------------------------------------------

def get_lockout_status(username: str) -> dict:
    """
    Geeft de lockout-status van een account terug.

    Returns:
        {
            "locked":           bool   — True als account geblokkeerd is,
            "rate_limited":     bool   — True als 30s wachttijd actief is,
            "seconds_remaining": int   — seconden tot de blokkade vervalt,
            "attempt_count":    int    — totaal aantal foute pogingen
        }
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT attempt_count, locked_until FROM account_lockout WHERE username = ?",
        (username,)
    )
    row = cursor.fetchone()
    conn.close()

    now = time.time()

    if row is None:
        return {"locked": False, "rate_limited": False, "seconds_remaining": 0, "attempt_count": 0}

    attempt_count, locked_until = row
    seconds_remaining = max(0, int(locked_until - now))

    # Volledig geblokkeerd (na MAX_ATTEMPTS_LOCKOUT pogingen)
    if attempt_count >= MAX_ATTEMPTS_LOCKOUT and seconds_remaining > 0:
        return {"locked": True, "rate_limited": False, "seconds_remaining": seconds_remaining, "attempt_count": attempt_count}

    # Rate limiting (na MAX_ATTEMPTS_RATE_LIMIT pogingen, kortere wachttijd)
    if attempt_count >= MAX_ATTEMPTS_RATE_LIMIT and seconds_remaining > 0:
        return {"locked": False, "rate_limited": True, "seconds_remaining": seconds_remaining, "attempt_count": attempt_count}

    return {"locked": False, "rate_limited": False, "seconds_remaining": 0, "attempt_count": attempt_count}


def record_failed_attempt(username: str) -> dict:
    """
    Registreer een mislukte inlogpoging en pas rate limiting / lockout toe.
    Geeft de nieuwe lockout-status terug.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT attempt_count FROM account_lockout WHERE username = ?", (username,)
    )
    row = cursor.fetchone()
    new_count = (row[0] if row else 0) + 1

    # Bepaal blokkade-duur op basis van aantal pogingen
    if new_count >= MAX_ATTEMPTS_LOCKOUT:
        locked_until = time.time() + LOCKOUT_SECONDS       # 10 minuten
    elif new_count >= MAX_ATTEMPTS_RATE_LIMIT:
        locked_until = time.time() + RATE_LIMIT_SECONDS    # 30 seconden
    else:
        locked_until = 0

    cursor.execute(
        """INSERT INTO account_lockout (username, attempt_count, locked_until)
           VALUES (?, ?, ?)
           ON CONFLICT(username) DO UPDATE SET
               attempt_count = excluded.attempt_count,
               locked_until  = excluded.locked_until""",
        (username, new_count, locked_until)
    )
    conn.commit()
    conn.close()

    return get_lockout_status(username)


def reset_failed_attempts(username: str) -> None:
    """Reset de foutenteller na een succesvolle inlog."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM account_lockout WHERE username = ?", (username,))
    conn.commit()
    conn.close()


# ------------------------------------------------------------------------------
# Authenticatie
# ------------------------------------------------------------------------------

def verify_user(username: str, password: str) -> dict | None:
    """
    Verifieer inloggegevens volledig inclusief lockout-controle en audit-log.

    Returns:
        Gebruikersdict bij succes, None bij falen.
        Geef de caller NOOIT terug WAT er fout is (geen onderscheid
        tussen 'gebruiker bestaat niet' en 'wachtwoord verkeerd').
    """
    # 1. Controleer lockout VOOR de database-query
    status = get_lockout_status(username)
    if status["locked"] or status["rate_limited"]:
        return None  # Blokkade wordt afgehandeld door de UI

    # 2. Haal de gebruiker op uit de database
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, username, role, password_hash FROM users WHERE username = ?",
        (username,)
    )
    row = cursor.fetchone()
    conn.close()

    # 3. Verifieer wachtwoord (ook als gebruiker niet bestaat, om timing-aanvallen te voorkomen)
    if row and _verify_password(password, row[3]):
        # Succes: reset pogingen en log
        reset_failed_attempts(username)
        log_login_attempt(username, success=True)
        return {
            "id":       row[0],
            "username": row[1],
            "role":     row[2],
            "is_admin": row[2] == "Admin"
        }
    else:
        # Mislukking: registreer poging en log
        new_status = record_failed_attempt(username)
        log_login_attempt(username, success=False)
        return None


# ------------------------------------------------------------------------------
# Scan-opslag
# ------------------------------------------------------------------------------

def save_scan(user_id: int, afdeling: str, scan_datum: str, results: list) -> None:
    """Sla een voltooide scan op in de database."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO scans (user_id, scan_datum, afdeling) VALUES (?, ?, ?)",
        (user_id, scan_datum, afdeling)
    )
    scan_id = cursor.lastrowid

    for r in results:
        cursor.execute(
            """INSERT INTO scan_results
               (scan_id, locatie, bestandsnaam, score, reden)
               VALUES (?, ?, ?, ?, ?)""",
            (
                scan_id,
                r.get("Pad", ""),
                r.get("Naam", ""),
                str(r.get("Score_Totaal", "")),
                r.get("Reden", "")
            )
        )

    conn.commit()
    conn.close()


# ==============================================================================

class CentraleEngine:
    """Orkestreert lokale en SharePoint scans en delegeert naar actieve engines."""

    def __init__(self, local_paths, sharepoint_sites, active_engines, stop_event=None):
        self.local_paths = local_paths
        self.sharepoint_sites = sharepoint_sites
        self.active_engines = active_engines
        self.stop_event = stop_event

        self.results = []
        self.file_registry_name = {}
        self.file_registry_content = {}
        self.sp_bibliotheken_tracker = {}

        # Beveiligings- en locatie-baselines
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

        # 1. Lokale scan
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

                    # Data quality check: lege map
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
                        except OSError:
                            pass

        # 2. SharePoint scan
        for sp in self.sharepoint_sites:
            if self.stop_event and self.stop_event.is_set():
                q.put(("canceled", "Geannuleerd tijdens SharePoint connectie."))
                return

            site_url = sp["url"]
            self.sp_bibliotheken_tracker[site_url] = {
                "Open Bibliotheek": 0, "Gesloten Bibliotheek": 0, "Foutieve Bieb": 0
            }

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
                                "Type": "SP Structuur", "Naam": "Toegangsfout",
                                "Pad": f"{site_url}/{library.title}",
                                "Mode": "SP", "Score_Totaal": "0%",
                                "Reden": f"Kan SP-bibliotheek niet scannen: {str(lib_err)}"
                            })
            except Exception as e:
                q.put(("error", f"Kritieke SP Connectiefout op {site_url}: {str(e)}"))
                return

        if not all_items and not self.results:
            q.put(("error", "Geen bestanden of structuren gevonden om te scannen."))
            return

        # 3. Orchestratie & analyse
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
        if self.stop_event and self.stop_event.is_set():
            return

        try:
            ctx.load(folder, ["Folders", "Files"])
            ctx.execute_query()

            # Data quality check: lege SharePoint-map
            if len(folder.files) == 0 and len(folder.folders) == 0:
                map_naam = current_path.split('/')[-1] if '/' in current_path else current_path
                self.results.append({
                    "Type": "Structuur", "Naam": map_naam,
                    "Pad": f"SP: {current_path}",
                    "Mode": "SP", "Score_Totaal": "0%",
                    "Reden": "Data Vervuiling: Lege map gedetecteerd op SharePoint. Ruim deze op."
                })
                return

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
                except Exception:
                    continue

            for sub_folder in folder.folders:
                if sub_folder.name not in ["Forms", "_t", "_w", "Templates"]:
                    self._walk_sp_recursive(ctx, sub_folder, f"{current_path}/{sub_folder.name}", site_url, all_items)
        except Exception:
            pass

    def _analyze_item(self, item):
        file_stream = self._get_file_stream(item)

        filename = item["name"]
        base_name = self._get_base_filename(filename)
        extension = item["extension"]
        mode = item["mode"]
        size_key = f"{item.get('size', 0)}_{extension}"

        name_dups = [loc for loc in self.file_registry_name.get(base_name, []) if loc != item["path"]]
        content_dups = (
            [loc for loc in self.file_registry_content.get(size_key, []) if loc != item["path"]]
            if item.get("size", 0) > 1024 else []
        )

        is_duplicate = len(name_dups) > 0 or len(content_dups) > 0
        item["is_duplicate"] = is_duplicate
        item["has_forbidden_chars"] = any(c in filename for c in self.FORBIDDEN_CHARS)
        item["is_readable_doc"] = extension in self.ALLOWED_SP_EXTS

        all_scores = {"Security (Risico's)": 100, "Locatie Beleid": 100, "Data Duplicatie": 100}
        all_reasons = []

        # Security check: risicovolle extensie
        if extension in self.RISKY_EXTS:
            all_scores["Security (Risico's)"] = 0
            all_reasons.append("🚨 KRITIEK: Schadelijk bestand.")

        # Locatie-beleid check
        if mode == "sp" and extension not in self.ALLOWED_SP_EXTS:
            all_scores["Locatie Beleid"] = 0
            all_reasons.append(f"Locatie: Extensie {extension} mag niet op SP.")
        elif mode == "local":
            is_large_file = item["size"] >= (2 * 1024 * 1024 * 1024)
            if extension in self.ALLOWED_SP_EXTS and not is_large_file:
                all_scores["Locatie Beleid"] = 0
                all_reasons.append("Locatie: Bestand kan op SP en hoort niet lokaal.")

        # Duplicatie check
        if is_duplicate:
            all_scores["Data Duplicatie"] = 0
            merged_dups = list(set(name_dups + content_dups))

            clean_locs = []
            for loc in merged_dups[:2]:
                parts = loc.replace('\\', '/').split('/')
                if len(parts) >= 2:
                    clean_locs.append(f".../{parts[-2]}/{parts[-1]}")
                else:
                    clean_locs.append(loc)

            loc_str = ", ".join(clean_locs)
            if len(merged_dups) > 2:
                loc_str += f" (en {len(merged_dups)-2} meer)"

            if len(content_dups) > 0:
                all_reasons.append(f"Duplicatie: Identieke inhoud (grootte {item.get('size')}B) in {loc_str}")
            else:
                all_reasons.append(f"Duplicatie: Zeer vergelijkbare bestandsnaam in {loc_str}")

        # Scores registreren per domein
        for domein in self.base_domains:
            if item["mode"] == "sp":
                self.domain_scores_sp[domein].append(all_scores[domein])
            else:
                self.domain_scores_local[domein].append(all_scores[domein])

        # Actieve engines draaien
        for engine in self.active_engines:
            try:
                engine_data = engine.analyze(item, file_stream)
                for domein, score in engine_data["scores"].items():
                    all_scores[domein] = score
                    if item["mode"] == "sp":
                        self.domain_scores_sp[domein].append(score)
                    else:
                        self.domain_scores_local[domein].append(score)
                all_reasons.extend(engine_data["reasons"])
            except Exception as e:
                all_reasons.append(f"Engine Fout ({engine.__class__.__name__}): {str(e)}")

        # Resultaat samenstellen
        item_result = {
            "Type": "Bestand",
            "Naam": item["name"],
            "Pad": item["path"],
            "Mode": item["mode"].upper()
        }

        active_vals = [v for v in all_scores.values() if isinstance(v, int)]
        item_result["Score_Totaal"] = f"{int(sum(active_vals) / len(active_vals))}%" if active_vals else "0%"

        for dom in self.all_domains:
            val = all_scores.get(dom, "N/A")
            item_result[dom] = f"{val}%" if isinstance(val, int) else val

        item_result["Reden"] = " | ".join(all_reasons) if all_reasons else "Volledig Compliant"
        self.results.append(item_result)

        if file_stream:
            try:
                file_stream.close()
            except Exception:
                pass

    def _get_file_stream(self, item):
        try:
            if item["mode"] == "local":
                return open(item["path"], "rb")
            elif item["mode"] == "sp":
                return io.BytesIO(
                    item["ctx"].web.get_file_by_server_relative_url(item["sp_url"]).read()
                )
        except Exception:
            return None

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

        if base_name not in self.file_registry_name:
            self.file_registry_name[base_name] = []
        if path not in self.file_registry_name[base_name]:
            self.file_registry_name[base_name].append(path)

        size = item.get("size", 0)
        ext = item.get("extension", "")
        if size > 1024:
            size_key = f"{size}_{ext}"
            if size_key not in self.file_registry_content:
                self.file_registry_content[size_key] = []
            if path not in self.file_registry_content[size_key]:
                self.file_registry_content[size_key].append(path)

    def _validate_sp_library_name(self, site_url, lib_name):
        if lib_name == "Open Bibliotheek":
            self.sp_bibliotheken_tracker[site_url]["Open Bibliotheek"] += 1
        elif lib_name == "Gesloten Bibliotheek":
            self.sp_bibliotheken_tracker[site_url]["Gesloten Bibliotheek"] += 1
        elif "bibliotheek" in lib_name.lower() or "bieb" in lib_name.lower():
            self.sp_bibliotheken_tracker[site_url]["Foutieve Bieb"] += 1

    def _rapporteer_sp_bibliotheken(self):
        for site, counts in self.sp_bibliotheken_tracker.items():
            if counts["Open Bibliotheek"] > 1 or counts["Gesloten Bibliotheek"] > 1 or counts["Foutieve Bieb"] > 0:
                self.results.append({
                    "Type": "SP Structuur", "Naam": "Bibliotheek Fout",
                    "Pad": site, "Mode": "SP", "Score_Totaal": "0%",
                    "Reden": (
                        f"FOUT: Verkeerde bibliotheek formatie. "
                        f"Open: {counts['Open Bibliotheek']}, "
                        f"Gesloten: {counts['Gesloten Bibliotheek']}, "
                        f"Invalide: {counts['Foutieve Bieb']}."
                    )
                })