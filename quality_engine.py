import os
import re
from datetime import datetime, timezone


class KwaliteitEngine:
    def __init__(self):
        self.domains = [
            "Accuracy",
            "Completeness",
            "Consistency",
            "Uniqueness",
            "Timeliness",
            "Validity",
            "Granularity"
        ]

        self.MAX_PATH_LENGTH = 260
        self.MAX_FOLDER_DEPTH = 4

        self.DATE_PREFIX_PATTERNS = [
            re.compile(r"^(\d{8})[_ -].+"),
            re.compile(r"^(\d{4}-\d{2}-\d{2})[_ -].+"),
        ]

        self.BAD_NAME_WORDS = {
            "nieuw", "new", "kopie", "copy", "temp", "final_final", "concept"
        }

    def analyze(self, item, file_stream=None):
        scores = {}
        reasons = []

        score, msgs = self._check_validity(item)
        scores["Validity"] = score
        reasons.extend(msgs)

        score, msgs = self._check_completeness(item)
        scores["Completeness"] = score
        reasons.extend(msgs)

        score, msgs = self._check_consistency(item)
        scores["Consistency"] = score
        reasons.extend(msgs)

        score, msgs = self._check_uniqueness(item)
        scores["Uniqueness"] = score
        reasons.extend(msgs)

        score, msgs = self._check_timeliness(item)
        scores["Timeliness"] = score
        reasons.extend(msgs)

        score, msgs = self._check_granularity(item)
        scores["Granularity"] = score
        reasons.extend(msgs)

        score, msgs = self._check_accuracy(item)
        scores["Accuracy"] = score
        reasons.extend(msgs)

        return {
            "scores": scores,
            "reasons": reasons
        }

    def _check_path_length(self, item):
        path_value = item.get("path", "")
        path_length = len(path_value)

        if path_length > self.MAX_PATH_LENGTH:
            return 0, [f"Padlengte: pad overschrijdt MAX_PATH ({path_length}>{self.MAX_PATH_LENGTH})."]

        if path_length > 220:
            return 50, [f"Padlengte: pad is lang en nadert de limiet ({path_length} tekens)."]

        return 100, []

def _check_accuracy(self, item):
    score = 100
    reasons = []

    filename = item.get("name", "")
    extension = item.get("extension", "")
    modified_dt = None

    # 1. Basiscontrole: naam zonder extensie is verdacht
    if filename and not extension:
        score -= 30
        reasons.append(
            "Accuracy: bestandsstructuur is onvolledig, waardoor interpretatie onbetrouwbaar wordt."
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

    # 3. Vergelijk datum in bestandsnaam met wijzigingsdatum
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
                        "Accuracy: datum in bestandsnaam wijkt sterk af van de wijzigingsdatum."
                    )
                elif delta_days > 30:
                    score -= 20
                    reasons.append(
                        "Accuracy: datum in bestandsnaam wijkt af van de wijzigingsdatum."
                    )

            except ValueError:
                score -= 20
                reasons.append("Accuracy: datum in bestandsnaam is ongeldig.")

            break

    return max(score, 0), reasons

    def _check_naamgeving(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        stem = os.path.splitext(filename)[0].lower()

        has_date_prefix = any(p.match(filename) for p in self.DATE_PREFIX_PATTERNS)
        if not has_date_prefix:
            score -= 40
            reasons.append("Naamgeving: bestand begint niet met een datum-prefix.")

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

        if not extension:
            score -= 50
            reasons.append("Syntaxis: bestand heeft geen extensie.")

        if filename.count(".") > 1:
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

    def _check_uniqueness(self, item):
        if item.get("is_duplicate", False):
            return 0, ["Uniqueness: bestandsnaam komt op meerdere locaties of modi voor."]

        return 100, []

    def _check_completeness(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        size = item.get("size", 0)

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
            reasons.append("Completeness: bestand is mogelijk onvolledig of bijna leeg.")

        return max(score, 0), reasons

    def _check_consistency(self, item):
        score = 100
        reasons = []

        filename = item.get("name", "")
        extension = item.get("extension", "")
        mode = item.get("mode", "")

        matches = [p.match(filename) for p in self.DATE_PREFIX_PATTERNS if p.match(filename)]
        if not matches:
            score -= 40
            reasons.append("Consistency: bestand volgt geen consistente datumconventie in de naam.")

        if mode == "sp" and extension not in {".docx", ".xlsx", ".pptx", ".pdf", ".txt"}:
            score -= 30
            reasons.append(f"Consistency: extensie {extension} is ongebruikelijk voor SharePoint-opslag.")

        if "  " in filename:
            score -= 15
            reasons.append("Consistency: bestandsnaam bevat dubbele spaties.")

        if filename.count(".") > 1:
            score -= 15
            reasons.append("Consistency: bestandsnaam bevat meerdere punten.")

        return max(score, 0), reasons

    def _check_timeliness(self, item):
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

        if age_days > 5 * 365:
            return 0, ["Timeliness: bestand is ouder dan 5 jaar."]
        elif age_days > 3 * 365:
            return 50, ["Timeliness: bestand is ouder dan 3 jaar."]
        else:
            return 100, []

    def _check_granularity(self, item):
        score = 100
        reasons = []

        depth = self._calculate_depth(item)
        filename = item.get("name", "").lower()
        stem = os.path.splitext(filename)[0]

        if depth > self.MAX_FOLDER_DEPTH:
            score -= 50
            reasons.append(f"Granularity: bestand zit te diep in de structuur ({depth} niveaus).")

        generic_names = {"document", "bestand", "file", "new", "nieuw"}
        if stem in generic_names:
            score -= 50
            reasons.append("Granularity: bestandsnaam is te generiek.")

        return max(score, 0), reasons

    def _check_validity(self, item):
        scores = []
        reasons = []

        for fn in [self._check_path_length, self._check_naamgeving, self._check_syntaxis]:
            score, msgs = fn(item)
            scores.append(score)
            reasons.extend(msgs)

        return int(sum(scores) / len(scores)), reasons

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