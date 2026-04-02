import re
import time
import os
import PyPDF2
import docx
import openpyxl
import pptx
import difflib  # Noodzakelijk voor de Levenshtein-afstand (60%-regel)
from datetime import datetime, timezone

class ComplianceEngine:
    """
    De core engine voor het valideren van bestanden, mappen en SharePoint-pagina's 
    tegen de geldende bedrijfs- en complianceregels.
    Ontworpen volgens het Open/Closed principe: uitbreidbaar zonder bestaande logica te breken.
    """
    
    def __init__(self, active_modules):
        # self.domains bepaalt welke controles (modules) er worden uitgevoerd.
        # Let op: Security en Locatie worden beheerd door de CentraleEngine, niet hier.
        self.domains = active_modules

    def analyze(self, item, stream):
        """
        Analyseert een specifiek item (bestand, map of pagina).
        
        Args:
            item (dict): Bevat alle metadata van het object (naam, extensie, type, etc.).
            stream (io.BytesIO / file): De binaire stroom van het bestand voor inhoudelijke checks.
            
        Returns:
            dict: Een dictionary met de berekende 'scores' per domein en een lijst met 'reasons' (foutmeldingen).
        """
        # Initialiseer alle actieve domeinen standaard op "N/A" om false positives te voorkomen
        scores = {mod: "N/A (Overgeslagen)" for mod in self.domains}
        reden = []
        
        # Veiligheidscheck: gebruik .get() om crashes te voorkomen als keys ontbreken (Data Quality)
        filename_lower = item.get("name", "").lower()
        extension = item.get("extension", "")
        in_werkomgeving = item.get("in_werkomgeving", False)

        # =========================================================
        # NIEUW: 5. HOMEPAGE UNIFORMITEIT (SharePoint Pages)
        # Doel: Garanderen dat startpagina's herkenbaar zijn ('Home') 
        # en altijd een contactpersoon (e-mail) bevatten.
        # =========================================================
        if "Homepage Uniformiteit" in self.domains:
            # Voer deze check alléén uit als het item expliciet als homepage is gemarkeerd
            if item.get("is_homepage", False):
                page_name = item.get("page_name", filename_lower)
                page_content = item.get("page_content", "")
                
                score, reasons = self._check_homepage_uniformity(page_name, page_content)
                scores["Homepage Uniformiteit"] = score
                reden.extend(reasons)
            else:
                scores["Homepage Uniformiteit"] = "N/A (Geen Startpagina)"

        # =========================================================
        # NIEUW: 6. WORKSPACE DUPLICATIE (60%-Regel)
        # Doel: Voorkomen van versnippering door lokale mappen te 
        # vergelijken met bestaande M365 samenwerkingsgroepen.
        # =========================================================
        if "Workspace Duplicatie" in self.domains:
            # QA: Voer dit alléén uit op mappen, het heeft geen zin om losse bestanden te testen
            if item.get("is_folder", False):
                folder_name = item.get("name", "")
                site_title = item.get("site_title", "")
                m365_groups = item.get("m365_groups", [])
                
                score, reasons = self._check_duplicate_workspace(folder_name, site_title, m365_groups)
                scores["Workspace Duplicatie"] = score
                reden.extend(reasons)
            else:
                scores["Workspace Duplicatie"] = "N/A (Bestand, geen map)"

        # =========================================================
        # OUDE CONTROLES (Bestandsniveau)
        # QA Fix: Toegevoegd `not item.get("is_folder")` en `not is_homepage` 
        # zodat mappen/pagina's niet onbedoeld afgekeurd worden op bestandsregels.
        # =========================================================
        is_file = not item.get("is_folder") and not item.get("is_homepage")
        
        # 1. NAAMGEVING CHECK
        if "Naamgeving" in self.domains and is_file:
            if in_werkomgeving:
                scores["Naamgeving"] = "N/A (Werkomgeving)"
            elif item.get("has_forbidden_chars", False):
                scores["Naamgeving"] = 0
                reden.append("Naamgeving: Bevat verboden tekens.")
            # Strikte RegEx validatie: YYYYMMDD_Rubricering_Afdeling_Onderwerp_Versie.ext
            elif not re.match(r"^\d{8}_[^_]+_[^_]+_[^_]+_[^_]+\.[a-zA-Z0-9]+$", filename_lower):
                scores["Naamgeving"] = 0
                reden.append("Naamgeving: Fout format (YYYYMMDD_Rubricering_Afdeling_Onderwerp_Versie).")
            else:
                scores["Naamgeving"] = 100

        # Pre-checks voor inhoudelijke inspecties
        is_readable_doc = item.get("is_readable_doc", False)
        file_is_locked = stream is None

        # 2. METADATA CHECK
        if "Metadata" in self.domains and is_file:
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

        # 3. RUBRICERING CHECK
        if "Rubricering" in self.domains and is_file:
            if not is_readable_doc:
                scores["Rubricering"] = "N/A"
            elif in_werkomgeving:
                scores["Rubricering"] = "N/A (Werkomgeving)"
            elif file_is_locked:
                scores["Rubricering"] = 0
                reden.append("Rubricering: Bestand gelockt.")
            else:
                # Lees de eerste paar pagina's om performance impact te beperken
                pages_text = self._read_pages_sample(stream, extension)
                if pages_text:
                    labels = ["gerubriceerd", "ongerubriceerd", "gemerkt"]
                    is_compliant = True
                    # Check of er op elke uitgelezen pagina tenminste 2 geldige labels staan
                    for page in pages_text:
                        if sum(page.count(lbl) for lbl in labels) < 2:
                            is_compliant = False; break
                    
                    scores["Rubricering"] = 100 if is_compliant else 0
                    if not is_compliant: 
                        reden.append("Rubricering: Onvoldoende gelabeld per pagina.")
                else:
                    scores["Rubricering"] = 0
                    reden.append("Rubricering: Document leeg of scanbaar als plaatje.")

        # 4. BEWAARTERMIJN CHECK
        if "Bewaartermijn" in self.domains and is_file:
            age_years = self._calculate_age(item)
            if age_years > 5:
                scores["Bewaartermijn"] = 0
                reden.append(f"VNG: Te oud ({age_years:.1f} jaar).")
            else: 
                scores["Bewaartermijn"] = 100

        # Post-processing: Zet de file pointer terug naar 0
        # Cruciaal om data-corruptie te voorkomen voor opvolgende processen die de stream nodig hebben
        if stream:
            try: stream.seek(0)
            except Exception: pass

        return {
            "scores": scores,
            "reasons": reden
        }

    # =========================================================
    # HELPER METHODES (Interne logica, niet direct van buitenaf aanroepen)
    # =========================================================

    def _check_homepage_uniformity(self, page_name: str, page_content: str) -> tuple:
        """
        Valideert de uniformiteit van een SharePoint startpagina.
        Checkt op harde naameis ('Home') en aanwezigheid van e-mail.
        """
        reasons = []
        score = 100
        standaard_naam = "Home"
        
        # Data Quality: Veiligheidscheck, forceer string type om crashes op null-waarden te voorkomen.
        page_name = str(page_name) if page_name else ""
        page_content = str(page_content) if page_content else ""

        # Check 1: Exacte naamgeving (case-sensitive)
        if page_name.strip() != standaard_naam:
            score = 0
            reasons.append(f"Uniformiteit: Startpagina heet '{page_name}', dit móét exact '{standaard_naam}' zijn.")
            
        # Check 2: E-mail aanwezigheid via robuuste RegEx
        email_pattern = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
        if not email_pattern.search(page_content):
            score = 0
            reasons.append("Uniformiteit: Geen geldig e-mailadres aangetroffen op de startpagina.")
            
        return score, reasons

    def _check_duplicate_workspace(self, folder_name: str, site_title: str, m365_group_names: list) -> tuple:
        """
        Implementeert de 60%-regel om dubbele samenwerkingsomgevingen te detecteren.
        Combineert mapnaam en sitetitel en vergelijkt deze met actieve M365 groepen.
        """
        # Als we geen data hebben, kunnen we niet oordelen. Return 'compliant'.
        if not folder_name or not m365_group_names:
            return 100, [] 
            
        # Data Quality: Uitsluiten van veelgebruikte stopwoorden die false positives (onterechte 60% matches) genereren.
        stop_words = {"team", "groep", "project", "afdeling", "samenwerking", "site", "map", "documenten"}
        
        def sanitize(text):
            """Interne helper om leestekens te strippen en stopwoorden te filteren."""
            clean = re.sub(r'[^\w\s]', '', str(text).lower())
            return [w for w in clean.split() if w not in stop_words]
            
        # Genereer de schone input-tokens
        folder_tokens = sanitize(folder_name)
        site_tokens = sanitize(site_title) if site_title else []
        combined_name = " ".join(folder_tokens + site_tokens)
        
        # Als na opschoning de string leeg is (bijv. de map heette "Team Documenten"), skip de controle.
        if not combined_name:
            return 100, []

        highest_ratio = 0.0
        matched_group = None
        
        # Vergelijk met de lijst van M365 groepen in de tenant
        for group in m365_group_names:
            group_clean = " ".join(sanitize(group))
            if not group_clean: 
                continue
            
            # Bereken Levenshtein ratio (0.0 tot 1.0)
            ratio = difflib.SequenceMatcher(None, combined_name, group_clean).ratio()
            if ratio > highest_ratio:
                highest_ratio = ratio
                matched_group = group
                
        # Toets aan de bedrijfsregel (>= 60%)
        if highest_ratio >= 0.60:
            return 0, [f"Samenwerking: Map overlapt voor {highest_ratio*100:.1f}% met de M365 Groep '{matched_group}'."]
            
        return 100, []

    # OUDE HELPER METHODES 
    # (Jouw _calculate_age, _check_metadata en _read_pages_sample blijven hier ongewijzigd staan, 
    # maar idealiter documenteer je deze op termijn op exact dezelfde manier).
    # ...