import os
import re
from datetime import datetime, timezone


class KwaliteitEngine:
    """Kwaliteitsdimensies zonder overlap met Centrale- en Compliance-engine.

    Deze engine beoordeelt alleen datakwaliteit en geen basisbeleid zoals:
    - Locatie/extensiebeleid (CentraleEngine)
    - Security-risico's (CentraleEngine)
    - Duplicatie over bronnen (CentraleEngine)
    - Rubricering/retentie/metadata-compliance (ComplianceEngine)
    """

    def __init__(self, active_domains=None):
        self.domains = [
            "Accuracy",
            "Completeness",
            "Consistency",
            "Uniqueness",
            "Timeliness",
            "Validity",
            "Granularity",
        ]
        self.active_domains = set(active_domains or self.domains)

        self.MAX_PATH_LENGTH = 260
        self.MAX_FOLDER_DEPTH = 4

        self.DATE_PREFIX_PATTERNS = [
            re.compile(r"^(\d{8})[_ -].+"),
            re.compile(r"^(\d{4}-\d{2}-\d{2})[_ -].+"),
        ]

        self.MANUAL_VERSION_PATTERNS = [
            re.compile(r"copy\s+of", re.IGNORECASE),
            re.compile(r"kopie\s+van", re.IGNORECASE),
            re.compile(r"\(\d+\)"),
            re.compile(r"_v\d+", re.IGNORECASE),
            re.compile(r"_final_final", re.IGNORECASE),
            re.compile(r"_definitief_definitief", re.IGNORECASE),
            re.compile(r"\s-\s*\d$"),
        ]

        self.PLACEHOLDERS = {
            "tbd",
            "n/a",
            "na",
            "xxx",
            "nvt",
            "todo",
            "temp",
            "test",
            "-",
            "...",
            "placeholder",
            "invullen",
        }

        self.BAD_NAME_WORDS = {"nieuw", "new", "kopie", "copy", "temp", "final_final", "concept"}

        self.POINT_IN_TIME_INDICATORS = {
            "contract",
            "con",
            "min",
            "definitief",
            "signed",
            "final",
            "getekend",
            "afgesloten",
            "gesloten",
        }
        self.LIVING_DOC_INDICATORS = {
            "rpt",
            "status",
            "update",
            "plan",
            "tracker",
            "dashboard",
            "overzicht",
            "voortgang",
        }
        self.REFERENCE_DOC_INDICATORS = {
            "tem",
            "pol",
            "template",
            "policy",
            "procedure",
            "richtlijn",
            "handleiding",
            "instructie",
            "manual",
        }

    # ------------------------------------------------------------------
    # Hoofdanalyse
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Validity
    # ------------------------------------------------------------------

    def _check_validity(self, item):
        scores = []
        reasons = []

        path_value = item.get("path", "")
        path_length = len(path_value)
        if path_length > self.MAX_PATH_LENGTH:
            scores.append(0)
            reasons.append(f"Validity: pad overschrijdt MAX_PATH ({path_length}>{self.MAX_PATH_LENGTH}).")
        elif path_length > 220:
            scores.append(50)
            reasons.append(f"Validity: pad is lang en nadert limiet ({path_length} tekens).")
        else:
            scores.append(100)

        filename = item.get("name", "")
        extension = item.get("extension", "")
        if not filename.strip():
            scores.append(0)
            reasons.append("Validity: bestandsnaam ontbreekt.")
        elif not extension:
            scores.append(25)
            reasons.append("Validity: bestand heeft geen extensie.")
        else:
            scores.append(100)

        depth = self._calculate_depth(item)
        if depth > self.MAX_FOLDER_DEPTH:
            scores.append(0)
            reasons.append(f"Validity: bestand zit te diep in de structuur ({depth} niveaus).")
        elif depth == self.MAX_FOLDER_DEPTH:
            scores.append(50)
            reasons.append(f"Validity: bestand zit op maximale mapdiepte ({depth}).")
        else:
            scores.append(100)

        return int(sum(scores) / len(scores)), reasons

    # ------------------------------------------------------------------
    # Completeness
    # ------------------------------------------------------------------

    def _check_completeness(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        size = item.get("size", 0)
        stem = os.path.splitext(filename)[0].strip()

        if not filename.strip():
            score -= 60
            reasons.append("Completeness: bestandsnaam ontbreekt.")

        if not extension:
            score -= 20
            reasons.append("Completeness: bestand heeft geen extensie.")

        if size <= 0:
            score -= 60
            reasons.append("Completeness: bestand heeft geen inhoud (0 bytes).")
        elif size < 1024:
            score -= 20
            reasons.append("Completeness: bestand is mogelijk onvolledig (<1 KB).")

        stem_words = re.split(r"[\s_\-\.]+", stem.lower())
        if any(word in self.PLACEHOLDERS for word in stem_words if word):
            score -= 20
            reasons.append("Completeness: bestandsnaam bevat placeholder-achtige termen.")

        return max(score, 0), reasons

    # ------------------------------------------------------------------
    # Consistency
    # ------------------------------------------------------------------

    def _check_consistency(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        stem = os.path.splitext(filename)[0]
        normalized = filename.lower()

        has_date_prefix = any(pattern.match(filename) for pattern in self.DATE_PREFIX_PATTERNS)
        if not has_date_prefix:
            score -= 25
            reasons.append("Consistency: bestandsnaam mist een consistente datum-prefix (YYYYMMDD of YYYY-MM-DD).")

        if "  " in filename:
            score -= 15
            reasons.append("Consistency: bestandsnaam bevat dubbele spaties.")

        if filename.count(".") > 1:
            score -= 10
            reasons.append("Consistency: bestandsnaam bevat meerdere punten.")

        if any(word in normalized for word in self.BAD_NAME_WORDS):
            score -= 20
            reasons.append("Consistency: bestandsnaam bevat tijdelijke/vage termen.")

        for pattern in self.MANUAL_VERSION_PATTERNS:
            if pattern.search(stem):
                score -= 30
                reasons.append("Consistency: handmatig versiebeheerpatroon gedetecteerd in bestandsnaam.")
                break

        return max(score, 0), reasons

    # ------------------------------------------------------------------
    # Uniqueness
    # ------------------------------------------------------------------

    def _check_uniqueness(self, item):
        """Alleen naamgedrag; bron-duplicatie wordt in CentraleEngine afgehandeld."""
        score = 100
        reasons = []
        stem = os.path.splitext(item.get("name", ""))[0]

        for pattern in self.MANUAL_VERSION_PATTERNS:
            if pattern.search(stem):
                score -= 60
                reasons.append("Uniqueness: bestandsnaam duidt op handmatige kopie/versie in plaats van centraal versiebeheer.")
                break

        return max(score, 0), reasons

    # ------------------------------------------------------------------
    # Timeliness
    # ------------------------------------------------------------------

    def _check_timeliness(self, item):
        """Versheid van informatie, niet retentie of wettelijke bewaartermijn."""
        filename = item.get("name", "").lower()
        stem = os.path.splitext(filename)[0]

        modified_dt = self._get_modified_datetime(item)
        if not modified_dt:
            return 50, ["Timeliness: wijzigingsdatum kon niet worden bepaald."]

        now = datetime.now(tz=modified_dt.tzinfo)
        age_days = (now - modified_dt).days

        status = str(item.get("status", "")).lower()
        if status in {"final", "definitief", "archived", "afgesloten"}:
            return 100, []

        if any(indicator in stem for indicator in self.POINT_IN_TIME_INDICATORS):
            return 100, []

        if any(indicator in stem for indicator in self.LIVING_DOC_INDICATORS):
            if age_days > 365:
                return 0, ["Timeliness: levend document is >365 dagen niet bijgewerkt."]
            if age_days > 180:
                return 50, ["Timeliness: levend document is >180 dagen niet bijgewerkt."]
            return 100, []

        if any(indicator in stem for indicator in self.REFERENCE_DOC_INDICATORS):
            if age_days > 730:
                return 0, ["Timeliness: referentiedocument is >730 dagen niet herzien."]
            if age_days > 365:
                return 50, ["Timeliness: referentiedocument is >365 dagen niet herzien."]
            return 100, []

        if age_days > 365:
            return 50, ["Timeliness: document is langer dan 365 dagen niet bijgewerkt."]
        return 100, []

    # ------------------------------------------------------------------
    # Granularity
    # ------------------------------------------------------------------

    def _check_granularity(self, item):
        score = 100
        reasons = []

        depth = self._calculate_depth(item)
        filename = item.get("name", "").lower()
        stem = os.path.splitext(filename)[0]

        if depth > self.MAX_FOLDER_DEPTH:
            score -= 40
            reasons.append(f"Granularity: bestand zit te diep in de structuur ({depth} niveaus).")

        generic_names = {
            "document",
            "bestand",
            "file",
            "new",
            "nieuw",
            "rapport",
            "report",
            "bijlage",
            "attachment",
            "onbekend",
            "unknown",
            "overig",
            "diversen",
            "misc",
        }
        if stem in generic_names:
            score -= 50
            reasons.append("Granularity: bestandsnaam is te generiek.")

        year_only = re.search(r"(?<!\d)(20\d{2})(?!\d)", stem)
        has_full_date = any(pattern.match(filename) for pattern in self.DATE_PREFIX_PATTERNS)
        if year_only and not has_full_date:
            score -= 20
            reasons.append("Granularity: bestandsnaam bevat alleen een jaaraanduiding zonder volledige datum.")

        return max(score, 0), reasons

    # ------------------------------------------------------------------
    # Accuracy
    # ------------------------------------------------------------------

    def _check_accuracy(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        modified_dt = self._get_modified_datetime(item)

        if not modified_dt:
            return 50, ["Accuracy: wijzigingsdatum ontbreekt; datumcontrole beperkt."]

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
                    reasons.append(f"Accuracy: datum in bestandsnaam wijkt sterk af van wijzigingsdatum ({delta_days} dagen).")
                elif delta_days > 30:
                    score -= 20
                    reasons.append(f"Accuracy: datum in bestandsnaam wijkt af van wijzigingsdatum ({delta_days} dagen).")
            except ValueError:
                score -= 20
                reasons.append("Accuracy: datum-prefix in bestandsnaam heeft ongeldig formaat.")
            break

        return max(score, 0), reasons

    # ------------------------------------------------------------------
    # Hulpfuncties
    # ------------------------------------------------------------------

    def _get_modified_datetime(self, item):
        try:
            if item.get("mode") == "local":
                ts = os.path.getmtime(item["path"])
                return datetime.fromtimestamp(ts, tz=timezone.utc).astimezone()

            if item.get("mode") == "sp" and item.get("time_modified"):
                raw = item.get("time_modified")
                if isinstance(raw, datetime):
                    return raw if raw.tzinfo else raw.replace(tzinfo=timezone.utc)

                raw_str = str(raw)
                for fmt in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                    try:
                        parsed = datetime.strptime(raw_str, fmt)
                        return parsed.replace(tzinfo=timezone.utc)
                    except ValueError:
                        continue
        except Exception:
            return None
        return None

    def _calculate_depth(self, item):
        path_value = item.get("path", "")

        if item.get("mode") == "local":
            rel = os.path.relpath(path_value, item["root_source"])
            return max(len(rel.split(os.sep)) - 1, 0)

        if item.get("mode") == "sp":
            sp_path = path_value.replace("SP: ", "")
            return max(len(sp_path.split("/")) - 2, 0)

        return 0
