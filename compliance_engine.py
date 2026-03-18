import re
import os
import time
import PyPDF2
import docx
import openpyxl
import pptx
from datetime import datetime, timezone

class ComplianceEngine:
    def __init__(self, active_modules):
        self.active_modules = active_modules
        self.hard_rules = ["Security (Risico's)", "Data Duplicatie", "Locatie Beleid"]
        # Exporteer de domeinen zodat de CentraleEngine weet wat hij in het dashboard moet tekenen
        self.domains = self.hard_rules + active_modules
        
        self.ALLOWED_SP_EXTS = {'.docx', '.xlsx', '.pptx', '.pdf', '.txt'}
        self.RISKY_EXTS = {'.exe', '.bat', '.msi', '.ps1', '.vbs', '.cmd', '.sh', '.scr'}
        self.FORBIDDEN_CHARS = set('/\\:*?"<>| !+@')

    def analyze(self, item, stream, is_duplicate):
        scores = {mod: "N/A (Overgeslagen)" for mod in self.domains}
        reden = []
        
        filename = item["name"]
        filename_lower = filename.lower()
        extension = item["extension"]
        mode = item["mode"]
        in_werkomgeving = item.get("in_werkomgeving", False)

        # 1. BASELINE
        scores["Security (Risico's)"] = 100
        scores["Locatie Beleid"] = 100
        scores["Data Duplicatie"] = 100

        # 2. HARDE REGELS (Ongeacht map-locatie)
        if extension in self.RISKY_EXTS:
            scores["Security (Risico's)"] = 0
            reden.append("🚨 KRITIEK: Schadelijk bestand.")

        if mode == "sp" and extension not in self.ALLOWED_SP_EXTS:
            scores["Locatie Beleid"] = 0
            reden.append(f"Locatie: Extensie {extension} mag niet op SP.")
        elif mode == "local":
            is_large_file = item["size"] >= (2 * 1024 * 1024 * 1024)
            if extension in self.ALLOWED_SP_EXTS and not is_large_file:
                scores["Locatie Beleid"] = 0
                reden.append("Locatie: Bestand kan op SP en hoort niet lokaal.")

        if is_duplicate:
            scores["Data Duplicatie"] = 0
            reden.append("Duplicatie: Bestand bestaat lokaal én op SP.")
        elif scores["Locatie Beleid"] == 0:
            scores["Data Duplicatie"] = 0
            reden.append("Duplicatie: Faalt door onjuiste basislocatie.")

        # 3. OPTIONELE REGELS (Met Data Quality context)
        if "Naamgeving" in self.active_modules:
            if in_werkomgeving:
                scores["Naamgeving"] = "N/A (Werkomgeving)"
            elif any(c in filename for c in self.FORBIDDEN_CHARS):
                scores["Naamgeving"] = 0
                reden.append("Naamgeving: Bevat verboden tekens.")
            elif not re.match(r"^\d{8}_[^_]+_[^_]+_[^_]+_[^_]+\.[a-zA-Z0-9]+$", filename_lower):
                scores["Naamgeving"] = 0
                reden.append("Naamgeving: Fout format.")
            else:
                scores["Naamgeving"] = 100

        is_readable_doc = extension in self.ALLOWED_SP_EXTS
        file_is_locked = stream is None

        if "Metadata" in self.active_modules:
            if not is_readable_doc:
                scores["Metadata"] = "N/A"
            elif in_werkomgeving:
                scores["Metadata"] = "N/A (Werkomgeving)"
            elif file_is_locked:
                scores["Metadata"] = 0
                reden.append("Metadata: Bestand gelockt/onleesbaar.")
            elif self._check_metadata(stream, extension):
                scores["Metadata"] = 100
            else:
                scores["Metadata"] = 0
                reden.append("Metadata: Auteur/Status ontbreekt.")

        if "Rubricering" in self.active_modules:
            if not is_readable_doc:
                scores["Rubricering"] = "N/A"
            elif in_werkomgeving:
                scores["Rubricering"] = "N/A (Werkomgeving)"
            elif file_is_locked:
                scores["Rubricering"] = 0
                reden.append("Rubricering: Bestand gelockt.")
            else:
                pages_text = self._read_pages_sample(stream, extension)
                if pages_text:
                    labels = ["gerubriceerd", "ongerubriceerd", "gemerkt"]
                    is_compliant = True
                    for page in pages_text:
                        if sum(page.count(lbl) for lbl in labels) < 2:
                            is_compliant = False; break
                    scores["Rubricering"] = 100 if is_compliant else 0
                    if not is_compliant: reden.append("Rubricering: Onvoldoende gelabeld per pagina.")
                else:
                    scores["Rubricering"] = 0
                    reden.append("Rubricering: Document leeg of scanbaar als plaatje.")

        if "Bewaartermijn" in self.active_modules:
            age_years = self._calculate_age(item)
            if age_years > 5:
                scores["Bewaartermijn"] = 0
                reden.append(f"VNG: Te oud ({age_years:.1f} jaar).")
            else: scores["Bewaartermijn"] = 100

        # QA Check: Bereid stream voor op een eventuele VOLGENDE engine (zoals de QualityEngine)
        if stream:
            try: stream.seek(0)
            except Exception: pass

        return {
            "scores": scores,
            "reasons": reden
        }

    # -- HELPER METHODES --
    def _calculate_age(self, item):
        try:
            now = time.time()
            if item["mode"] == "local": return (now - item.get("time_modified", os.path.getmtime(item["path"]))) / (365 * 24 * 3600)
            elif item["mode"] == "sp":
                date_str = str(item["time_modified"])
                if "T" in date_str and "Z" in date_str:
                    dt = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    return (now - dt.timestamp()) / (365 * 24 * 3600)
            return 0
        except Exception: return -1

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