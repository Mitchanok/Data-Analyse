import os
import re
from datetime import datetime, timezone


class KwaliteitEngine:
    def __init__(self, active_domains=None):
        self.domains = [
            "Accuracy",
            "Completeness",
            "Consistency",
            "Uniqueness",
            "Timeliness",
            "Validity",
            "Granularity"
        ]
        self.active_domains = set(active_domains or self.domains)
        self.MAX_PATH_LENGTH = 260
        self.MAX_FOLDER_DEPTH = 4

        # Bestaande datum-prefix patronen
        self.DATE_PREFIX_PATTERNS = [
            re.compile(r"^(\d{8})[_ -].+"),
            re.compile(r"^(\d{4}-\d{2}-\d{2})[_ -].+"),
        ]

        # NATO §5.2.1 naamgevingsconventie: YY.PP.CC-TYP-NNN_Short-Description.ext
        self.NATO_NAMING_PATTERN = re.compile(
            r"^\d{2}\.\d{2}\.[\w]{2,4}-[A-Z]{3}-\d{3}"
        )

        # Erkende documenttype codes (NATO)
        self.ALLOWED_TYPE_CODES = {
            "INV", "CON", "RPT", "MIN", "LET", "POL", "TEM", "PRS", "NOT", "ANN"
        }

        # NATO §8: Patronen voor handmatig versiebeheer (triggeren Uniqueness-aftrek)
        self.MANUAL_VERSION_PATTERNS = [
            re.compile(r"copy\s+of", re.IGNORECASE),
            re.compile(r"kopie\s+van", re.IGNORECASE),
            re.compile(r"\(\d+\)"),              # (1), (2)
            re.compile(r"_v\d+", re.IGNORECASE), # _v2, _v3
            re.compile(r"_final_final", re.IGNORECASE),
            re.compile(r"_definitief_definitief", re.IGNORECASE),
            re.compile(r"\s-\s*\d$"),            # " - 1", " - 2"
        ]

        # NATO §3: Placeholder waarden die niet tellen als gevuld
        self.PLACEHOLDERS = {
            "tbd", "n/a", "na", "xxx", "nvt", "todo", "temp",
            "test", "-", "...", "placeholder", "invullen"
        }

        # Vage termen in bestandsnamen
        self.BAD_NAME_WORDS = {
            "nieuw", "new", "kopie", "copy", "temp", "final_final", "concept"
        }

        # NATO §4 Composite Timeliness Model: documentcategorieën
        # Point-in-time: eenmalig afgerond → nooit als 'verouderd' markeren
        self.POINT_IN_TIME_INDICATORS = {
            "con", "min", "definitief", "signed", "final",
            "getekend", "afgesloten", "gesloten"
        }
        # Living: actief bijgehouden → flag als >180 dagen niet gewijzigd
        self.LIVING_DOC_INDICATORS = {
            "rpt", "status", "update", "plan", "tracker",
            "dashboard", "overzicht", "voortgang"
        }
        # Reference: richtlijnen/templates → flag als >365 dagen niet herzien
        self.REFERENCE_DOC_INDICATORS = {
            "tem", "pol", "template", "policy", "procedure",
            "richtlijn", "handleiding", "instructie", "manual"
        }

    # =========================================================================
    # HOOFDANALYSE
    # =========================================================================

    def analyze(self, item, file_stream=None):
        scores = {}
        reasons = []

        if "Validity" in self.active_domains:
            score, msgs = self._check_validity(item)
            scores["Validity"] = score
            reasons.extend(msgs)

        if "Completeness" in self.active_domains:
            score, msgs = self._check_completeness(item)
            scores["Completeness"] = score
            reasons.extend(msgs)

        if "Consistency" in self.active_domains:
            score, msgs = self._check_consistency(item)
            scores["Consistency"] = score
            reasons.extend(msgs)

        if "Uniqueness" in self.active_domains:
            score, msgs = self._check_uniqueness(item)
            scores["Uniqueness"] = score
            reasons.extend(msgs)

        if "Timeliness" in self.active_domains:
            score, msgs = self._check_timeliness(item)
            scores["Timeliness"] = score
            reasons.extend(msgs)

        if "Granularity" in self.active_domains:
            score, msgs = self._check_granularity(item)
            scores["Granularity"] = score
            reasons.extend(msgs)

        if "Accuracy" in self.active_domains:
            score, msgs = self._check_accuracy(item)
            scores["Accuracy"] = score
            reasons.extend(msgs)

        return {"scores": scores, "reasons": reasons}

    # =========================================================================
    # VALIDITY  (NATO §5)
    # =========================================================================

    def _check_path_length(self, item):
        path_value = item.get("path", "")
        path_length = len(path_value)

        if path_length > self.MAX_PATH_LENGTH:
            return 0, [f"Padlengte: pad overschrijdt MAX_PATH ({path_length}>{self.MAX_PATH_LENGTH})."]
        if path_length > 220:
            return 50, [f"Padlengte: pad is lang en nadert de limiet ({path_length} tekens)."]
        return 100, []

    def _check_naamgeving(self, item):
        """
        NATO §5.2.1 Naamgevingsconventie: YY.PP.CC-TYP-NNN_Short-Description.ext
        Als dit niet van toepassing is, wordt ingeval op datum-prefix gelet.
        """
        score = 100
        reasons = []

        filename = item.get("name", "")
        stem = os.path.splitext(filename)[0].lower()

        is_nato = bool(self.NATO_NAMING_PATTERN.match(filename))

        if is_nato:
            # Valideer de documenttype code
            parts = stem.split("-")
            if len(parts) >= 2:
                type_code = parts[1].upper()
                if type_code not in self.ALLOWED_TYPE_CODES:
                    score -= 20
                    reasons.append(
                        f"Naamgeving: documenttype code '{type_code}' is niet herkend "
                        f"(verwacht: {'/'.join(sorted(self.ALLOWED_TYPE_CODES))})."
                    )
        else:
            # Terugval: controleer op datum-prefix
            has_date_prefix = any(p.match(filename) for p in self.DATE_PREFIX_PATTERNS)
            if not has_date_prefix:
                score -= 40
                reasons.append(
                    "Naamgeving: bestand voldoet niet aan de naamgevingsconventie "
                    "(YY.PP.CC-TYP-NNN of datum-prefix ontbreekt)."
                )

        if item.get("has_forbidden_chars", False):
            score -= 30
            reasons.append("Naamgeving: bestandsnaam bevat verboden tekens.")

        if any(word in stem for word in self.BAD_NAME_WORDS):
            score -= 30
            reasons.append("Naamgeving: bestandsnaam bevat tijdelijke of vage termen.")

        return max(score, 0), reasons

    def _check_syntaxis(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        is_nato = bool(self.NATO_NAMING_PATTERN.match(filename))

        if not extension:
            score -= 50
            reasons.append("Syntaxis: bestand heeft geen extensie.")

        # NATO-formaat bevat punten by design — geen dubbele-punt aftrek daarvoor
        if not is_nato and filename.count(".") > 1:
            score -= 20
            reasons.append("Syntaxis: bestandsnaam bevat meerdere punten.")

        if len(filename) < 8:
            score -= 30
            reasons.append("Syntaxis: bestandsnaam is erg kort en mogelijk niet beschrijvend.")

        if "  " in filename:
            score -= 20
            reasons.append("Syntaxis: bestandsnaam bevat dubbele spaties.")

        return max(score, 0), reasons

    def _check_mapdiepte(self, item):
        depth = self._calculate_depth(item)

        if depth > self.MAX_FOLDER_DEPTH:
            return 0, [f"Mapdiepte: bestand zit te diep in de structuur ({depth} niveaus)."]
        if depth == self.MAX_FOLDER_DEPTH:
            return 50, [f"Mapdiepte: bestand zit op de maximale toegestane diepte ({depth})."]
        return 100, []

    def _check_validity(self, item):
        scores = []
        reasons = []

        for fn in [
            self._check_path_length,
            self._check_naamgeving,
            self._check_syntaxis,
            self._check_mapdiepte,
        ]:
            score, msgs = fn(item)
            scores.append(score)
            reasons.extend(msgs)

        return int(sum(scores) / len(scores)), reasons

    # =========================================================================
    # COMPLETENESS  (NATO §3)
    # =========================================================================

    def _check_completeness(self, item):
        """
        NATO §3:
        - Bestandsnaam aanwezig en niet leeg/whitespace-only
        - Extensie aanwezig
        - Bestand heeft inhoud (>0 bytes, bij voorkeur >1 KB)
        - Geen placeholder-waarden (TBD, N/A, xxx, ...) in naam (§3.1.2)
        """
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        size = item.get("size", 0)
        stem = os.path.splitext(filename)[0].strip()

        if not filename.strip():
            score -= 50
            reasons.append("Completeness: bestandsnaam ontbreekt.")

        if not extension:
            score -= 25
            reasons.append("Completeness: bestand heeft geen extensie.")

        if size <= 0:
            score -= 50
            reasons.append("Completeness: bestand heeft geen inhoud (0 bytes).")
        elif size < 1024:
            score -= 25
            reasons.append("Completeness: bestand is mogelijk onvolledig of bijna leeg (<1 KB).")

        # §3.1.2 Placeholder detectie
        stem_words = re.split(r"[\s_\-\.]+", stem.lower())
        for word in stem_words:
            if word in self.PLACEHOLDERS:
                score -= 25
                reasons.append(
                    f"Completeness: bestandsnaam bevat een placeholder-waarde ('{word}'). "
                    "Dit telt als onvolledig per NATO §3.1.2."
                )
                break  # Eén melding per bestand

        return max(score, 0), reasons

    # =========================================================================
    # CONSISTENCY  (NATO §6)
    # =========================================================================

    def _check_consistency(self, item):
        """
        NATO §6:
        - Naamgevingsconventie consistent (NATO-formaat of datum-prefix)
        - Extensie passend voor de locatie (SP vs. lokaal)
        - Geen dubbele spaties of meerdere punten (legacy check)
        - Geen handmatige versiepatronen ('Copy of', '_v2') — dupliceert Uniqueness
          maar benadrukt het consistency-aspect (naam ≠ werkelijkheid)
        """
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        mode = item.get("mode", "")
        stem = os.path.splitext(filename)[0]
        is_nato = bool(self.NATO_NAMING_PATTERN.match(filename))

        # Naamgevingsconventie consistent?
        has_date_prefix = any(p.match(filename) for p in self.DATE_PREFIX_PATTERNS)
        if not is_nato and not has_date_prefix:
            score -= 40
            reasons.append("Consistency: bestand volgt geen consistente naamgevingsconventie.")

        # SharePoint: ongebruikelijke extensies (§6 cross-dataset)
        if mode == "sp" and extension not in {".docx", ".xlsx", ".pptx", ".pdf", ".txt"}:
            score -= 30
            reasons.append(f"Consistency: extensie {extension} is ongebruikelijk voor SharePoint-opslag.")

        # Dubbele spaties
        if "  " in filename:
            score -= 15
            reasons.append("Consistency: bestandsnaam bevat dubbele spaties.")

        # Meerdere punten (niet voor NATO-formaat dat punten vereist)
        if not is_nato and filename.count(".") > 1:
            score -= 15
            reasons.append("Consistency: bestandsnaam bevat meerdere punten.")

        # §6.1: Handmatig versiebeheer-patronen
        for pat in self.MANUAL_VERSION_PATTERNS:
            if pat.search(stem):
                score -= 25
                reasons.append(
                    "Consistency: bestandsnaam bevat handmatig versiebeheer-patroon "
                    "('Copy of', '_v2', '(1)', etc.) — inkonsistentie met versiebeheerbeleid."
                )
                break

        return max(score, 0), reasons

    # =========================================================================
    # UNIQUENESS  (NATO §8)
    # =========================================================================

    def _check_uniqueness(self, item):
        """
        NATO §8:
        - Cross-locatie duplicaat (bestaand, via is_duplicate flag)
        - Manueel versiebeheer detectie: 'Copy of', '_v2', '_FINAL_FINAL', '(1)'
          → indicate manual copies outside SharePoint versioning
        """
        score = 100
        reasons = []

        filename = item.get("name", "")
        stem = os.path.splitext(filename)[0]

        # §8.1 / §8.2: Bestand bestaat op meerdere locaties
        if item.get("is_duplicate", False):
            score -= 50
            reasons.append("Uniqueness: bestandsnaam komt op meerdere locaties of modi voor.")

        # §8.2 Name-pattern detectie voor handmatig versiebeheer
        for pat in self.MANUAL_VERSION_PATTERNS:
            if pat.search(stem):
                score -= 50
                reasons.append(
                    "Uniqueness: bestandsnaam bevat manueel versiebeheer-patroon "
                    "('Copy of', '_v2', '_FINAL_FINAL', '(1)', etc.). "
                    "Gebruik SharePoint versiebeheer in plaats van handmatige kopieën."
                )
                break

        return max(score, 0), reasons

    # =========================================================================
    # TIMELINESS  (NATO §4 - Composite Model)
    # =========================================================================

    def _check_timeliness(self, item):
        """
        NATO §4 Composite Timeliness Model:

        Categorie          | Verwachte versheid    | Drempelwaarden
        ── Point-in-time   | Eenmalig afgerond     | Altijd score 100 na afronding
        ── Living          | Regelmatig bijgehouden | >180 dagen → 50, >365 dagen → 0
        ── Reference       | Periodiek herzien     | >365 dagen → 50, >730 dagen → 0
        ── Onbekend        | Standaard leeftijd    | >3 jaar → 50, >5 jaar → 0

        Kernprincipe: timeliness gaat over levenscyclusgeschiktheid, NIET pure leeftijd.
        Een 5 jaar oud contract is volledig 'timely' als het nog van kracht is.
        """
        filename = item.get("name", "").lower()
        stem = os.path.splitext(filename)[0]

        modified_dt = None
        try:
            if item["mode"] == "local":
                ts = os.path.getmtime(item["path"])
                modified_dt = datetime.fromtimestamp(ts, tz=timezone.utc).astimezone()
            elif item["mode"] == "sp" and item.get("time_modified"):
                modified_dt = item["time_modified"]
        except Exception:
            pass

        if not modified_dt:
            return 50, ["Timeliness: wijzigingsdatum kon niet worden bepaald."]

        now = datetime.now(tz=modified_dt.tzinfo)
        age_days = (now - modified_dt).days

        # Categoriseer document op basis van naam-indicatoren
        is_point_in_time = any(ind in stem for ind in self.POINT_IN_TIME_INDICATORS)
        is_living = any(ind in stem for ind in self.LIVING_DOC_INDICATORS)
        is_reference = any(ind in stem for ind in self.REFERENCE_DOC_INDICATORS)

        # §4.3.1 Point-in-time: eenmaal definitief = altijd actueel
        if is_point_in_time:
            return 100, []

        # §4.3.1 Living documents: actief bijgehouden
        if is_living:
            if age_days > 365:
                return 0, ["Timeliness: actief document is meer dan 1 jaar niet bijgewerkt — mogelijk verouderd."]
            elif age_days > 180:
                return 50, ["Timeliness: actief document is meer dan 6 maanden niet bijgewerkt."]
            return 100, []

        # §4.3.1 Reference documents: periodiek herzien
        if is_reference:
            if age_days > 2 * 365:
                return 0, ["Timeliness: referentiedocument is meer dan 2 jaar niet herzien."]
            elif age_days > 365:
                return 50, ["Timeliness: referentiedocument is meer dan 1 jaar niet herzien."]
            return 100, []

        # Standaard (onbekend documenttype): basisdrempels
        if age_days > 5 * 365:
            return 0, ["Timeliness: bestand is ouder dan 5 jaar."]
        elif age_days > 3 * 365:
            return 50, ["Timeliness: bestand is ouder dan 3 jaar."]
        return 100, []

    # =========================================================================
    # GRANULARITY  (NATO §9)
    # =========================================================================

    def _check_granularity(self, item):
        """
        NATO §9:
        - Mapdiepte (bestaand)
        - Generieke bestandsnamen (uitgebreide lijst)
        - Datumgranulariteit: alleen jaarvermelding (2026) is onvoldoende detail
        """
        score = 100
        reasons = []

        depth = self._calculate_depth(item)
        filename = item.get("name", "").lower()
        stem = os.path.splitext(filename)[0]

        # Mapdiepte
        if depth > self.MAX_FOLDER_DEPTH:
            score -= 50
            reasons.append(f"Granularity: bestand zit te diep in de structuur ({depth} niveaus).")

        # Generieke namen — uitgebreid o.b.v. NATO §9
        generic_names = {
            "document", "bestand", "file", "new", "nieuw",
            "rapport", "report", "bijlage", "attachment",
            "onbekend", "unknown", "overig", "diversen", "misc"
        }
        if stem in generic_names:
            score -= 50
            reasons.append("Granularity: bestandsnaam is te generiek om classificatie mogelijk te maken.")

        # §9.1 Datumgranulariteit: naam bevat alleen een jaaraanduiding
        year_only = re.search(r"(?<!\d)(20\d{2})(?!\d)", stem)
        has_full_date = (
            any(p.match(filename) for p in self.DATE_PREFIX_PATTERNS)
            or bool(self.NATO_NAMING_PATTERN.match(filename))
        )
        if year_only and not has_full_date:
            score -= 25
            reasons.append(
                "Granularity: bestandsnaam bevat alleen een jaar (bijv. '2026') "
                "zonder volledige datum — te weinig granulariteit per NATO §9."
            )

        return max(score, 0), reasons

    # =========================================================================
    # ACCURACY  (NATO §7)
    # =========================================================================

    def _check_accuracy(self, item):
        """
        NATO §7 — wat geautomatiseerd kan worden:
        - Date logic: datum in bestandsnaam vs. werkelijke wijzigingsdatum
        - Structuurcontrole: naam zonder extensie
        - §5.2.3 Content-type heuristiek: documenttype code vs. extensie kruiscontrole
        """
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        stem = os.path.splitext(filename)[0]
        modified_dt = None

        # 1. Naam zonder extensie — structureel onbetrouwbaar
        if filename and not extension:
            score -= 30
            reasons.append(
                "Accuracy: bestandsstructuur is onvolledig, "
                "waardoor interpretatie onbetrouwbaar wordt."
            )

        # 2. Haal wijzigingsdatum op
        try:
            if item.get("mode") == "local":
                ts = os.path.getmtime(item["path"])
                modified_dt = datetime.fromtimestamp(ts, tz=timezone.utc).astimezone()
            elif item.get("mode") == "sp" and item.get("time_modified"):
                modified_dt = item["time_modified"]
        except Exception:
            modified_dt = None

        # 3. Datum in bestandsnaam vs. wijzigingsdatum (NATO §7.1 date logic)
        if modified_dt:
            for pattern in self.DATE_PREFIX_PATTERNS:
                match = pattern.match(filename)
                if not match:
                    continue

                raw_date = match.group(1)
                try:
                    if "-" in raw_date:
                        name_dt = datetime.strptime(raw_date, "%Y-%m-%d")
                    else:
                        name_dt = datetime.strptime(raw_date, "%Y%m%d")

                    delta_days = abs((modified_dt.date() - name_dt.date()).days)

                    if delta_days > 365:
                        score -= 50
                        reasons.append(
                            "Accuracy: datum in bestandsnaam wijkt sterk af van de wijzigingsdatum "
                            f"({delta_days} dagen verschil)."
                        )
                    elif delta_days > 30:
                        score -= 20
                        reasons.append(
                            "Accuracy: datum in bestandsnaam wijkt af van de wijzigingsdatum "
                            f"({delta_days} dagen verschil)."
                        )
                except ValueError:
                    score -= 20
                    reasons.append("Accuracy: datum in bestandsnaam heeft een ongeldig formaat.")
                break

        # 4. §5.2.3 Content-type kruiscontrole: type code vs. extensie
        nato_match = self.NATO_NAMING_PATTERN.match(filename)
        if nato_match:
            parts = stem.split("-")
            if len(parts) >= 2:
                type_code = parts[1].upper()

                # Documenten (INV/CON/RPT/LET/MIN) mogen geen data-extensie hebben
                if type_code in {"INV", "CON", "RPT", "LET", "MIN"} and extension in {".xlsx", ".csv", ".json"}:
                    score -= 20
                    reasons.append(
                        f"Accuracy: documenttype '{type_code}' is een tekstdocument "
                        f"maar heeft een data-extensie ({extension})."
                    )

                # Templates en policies mogen geen afbeelding/uitvoerbaar bestand zijn
                if type_code in {"TEM", "POL"} and extension in {".jpg", ".png", ".gif", ".exe", ".bat"}:
                    score -= 30
                    reasons.append(
                        f"Accuracy: documenttype '{type_code}' heeft een onverwachte extensie ({extension})."
                    )

        return max(score, 0), reasons

    # =========================================================================
    # HULPFUNCTIES
    # =========================================================================

    def _calculate_depth(self, item):
        path_value = item.get("path", "")

        if item.get("mode") == "local":
            rel = os.path.relpath(path_value, item["root_source"])
            parts = rel.split(os.sep)
            return max(len(parts) - 1, 0)

        if item.get("mode") == "sp":
            sp_path = path_value.replace("SP: ", "")
            parts = sp_path.split("/")
            return max(len(parts) - 2, 0)

        return 0