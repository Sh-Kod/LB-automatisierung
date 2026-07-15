"""
modules/naming_resolver.py

Smart Hybrid Naming Engine – Stufe 2: Resolver- und Konfliktlogik.
Führt strukturierte NamingEvidence-Objekte aus verschiedenen Quellen zusammen
und bestimmt deterministisch die finalen Werte für category, title und date.

Ausschließlich Python-Standardbibliothek (re, calendar, unicodedata, dataclasses).
Keine Dateizugriffe, Netzwerkaufrufe oder Seiteneffekte.
"""

import re
import calendar
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from modules.naming_engine import NamingEvidence, FilenameParseResult


# Zulässige Gesamtstatus-Konstanten
STATUS_AUTO_MERGED = "AUTO_MERGED"
STATUS_NEEDS_REVIEW = "NEEDS_REVIEW"
STATUS_CONFLICT = "CONFLICT"

# Zulässige Qualitäts-Konstanten
QUALITY_HIGH = "HIGH"
QUALITY_MEDIUM = "MEDIUM"
QUALITY_LOW = "LOW"

ALLOWED_FIELDS = ("category", "title", "date")
ALLOWED_QUALITIES = (QUALITY_HIGH, QUALITY_MEDIUM, QUALITY_LOW)


@dataclass
class ResolvedField:
    """Repräsentiert das aufgelöste Ergebnis eines einzelnen Naming-Feldes."""
    value: Optional[str] = None
    quality: Optional[str] = None
    sources: List[str] = field(default_factory=list)
    alternatives: List[str] = field(default_factory=list)
    conflict: bool = False
    reason: str = ""


@dataclass
class ResolvedResult:
    """Repräsentiert das Gesamtergebnis der Resolver-Auflösung über alle Naming-Felder."""
    category: ResolvedField = field(default_factory=ResolvedField)
    title: ResolvedField = field(default_factory=ResolvedField)
    date: ResolvedField = field(default_factory=ResolvedField)

    status: str = STATUS_NEEDS_REVIEW
    missing_fields: List[str] = field(default_factory=list)
    review_reasons: List[str] = field(default_factory=list)

    filename_is_generic: bool = False
    filename_is_ambiguous: bool = False


@dataclass
class _Candidate:
    """Interner Kandidat für einen normalisierten Wert eines Feldes."""
    normalized_value: str
    evidences: List[NamingEvidence] = field(default_factory=list)
    sources: List[str] = field(default_factory=list)
    strong_sources: set = field(default_factory=set)
    quality: str = QUALITY_LOW
    strong_source_count: int = 0
    strong_total_score: int = 0
    representative_original_value: str = ""


def _validate_date_value(val: str) -> bool:
    """Prüft, ob ein Datumswert exakt dem Format TT_MM entspricht und kalendarisch gültig ist."""
    parts = val.split("_")
    if len(parts) != 2:
        return False
    tag_str, monat_str = parts[0], parts[1]
    if not (tag_str.isdigit() and len(tag_str) == 2 and monat_str.isdigit() and len(monat_str) == 2):
        return False
    try:
        tag = int(tag_str)
        monat = int(monat_str)
        if not (1 <= monat <= 12):
            return False
        # Schaltjahr 2028 als neutrale Basis, damit 29_02 zulässig ist
        _, max_tage = calendar.monthrange(2028, monat)
        return 1 <= tag <= max_tage
    except (ValueError, TypeError):
        return False


def _normalize_for_comparison(field_name: str, value: str) -> str:
    """Normalisiert einen Wert rein für den Vergleich von Kandidaten (Originalwert bleibt erhalten)."""
    val = value.strip()
    if field_name == "title":
        norm = unicodedata.normalize("NFKC", val)
        norm = norm.casefold()
        norm = re.sub(r'[\s\-\._]+', '_', norm)
        return norm.strip('_')
    elif field_name == "category":
        norm = val.casefold()
        norm = re.sub(r'[\s\-\._]+', '', norm)
        return norm
    elif field_name == "date":
        return val
    return val.casefold()


def _validate_evidence(e: object, idx: int | str) -> None:
    """Validiert ein NamingEvidence streng auf korrekte Typen, erlaubte Werte, Normalisierung und Datumsformat."""
    if not isinstance(e, NamingEvidence):
        raise ValueError(f"Index {idx}: Element ist kein NamingEvidence-Objekt ({type(e).__name__})")
    if e.field not in ALLOWED_FIELDS:
        raise ValueError(f"Index {idx}: Unerlaubtes Feld '{e.field}' (erlaubt: {ALLOWED_FIELDS})")
    if not isinstance(e.value, str) or not e.value.strip():
        raise ValueError(f"Index {idx}: Feld '{e.field}' hat leeren oder ungültigen value ({repr(e.value)})")
    if not isinstance(e.source, str) or not e.source.strip():
        raise ValueError(f"Index {idx}: Feld '{e.field}' hat leeren oder ungültigen source ({repr(e.source)})")
    if e.quality not in ALLOWED_QUALITIES:
        raise ValueError(f"Index {idx}: Feld '{e.field}' hat unerlaubte quality '{e.quality}' (erlaubt: {ALLOWED_QUALITIES})")
    if isinstance(e.score, bool) or not isinstance(e.score, int):
        raise ValueError(f"Index {idx}: Feld '{e.field}' hat ungültigen score {repr(e.score)} (muss int sein)")
    if e.field == "date" and (e.value != e.value.strip() or not _validate_date_value(e.value)):
        raise ValueError(f"Index {idx}: Feld 'date' hat ungültiges Datumsformat '{e.value}' (erwartet TT_MM ohne äußere Leerzeichen und gültiger Kalendertag)")

    norm_check = _normalize_for_comparison(e.field, e.value)
    if not norm_check:
        raise ValueError(f"Index {idx}: Feld '{e.field}' ergibt nach Normalisierung keinen verwertbaren Wert ({repr(e.value)})")


def _quality_rank(quality: str) -> int:
    """Gibt den numerischen Rang einer Qualitätsstufe zurück (HIGH=3, MEDIUM=2, LOW=1)."""
    if quality == QUALITY_HIGH:
        return 3
    elif quality == QUALITY_MEDIUM:
        return 2
    return 1


def _candidate_sort_key(c: _Candidate) -> Tuple[int, int, int, str]:
    """
    Deterministische Sortierung für Kandidaten:
    - Primär nach Qualitätsstufe (HIGH > MEDIUM > LOW)
    - Bei HIGH: nur Anzahl starker Quellen und Summe der Scores starker Evidences, danach normalisierter Wert.
      LOW-Noise und schwache Scores fließen in die Rangfolge nicht ein.
    - Bei MEDIUM: nur MEDIUM-Quellenanzahl und -Score-Summe, danach normalisierter Wert.
    - Bei LOW: nur LOW-Quellenanzahl und -Score-Summe, danach normalisierter Wert.
    """
    if c.quality == QUALITY_HIGH:
        return (-3, -c.strong_source_count, -c.strong_total_score, c.normalized_value)
    elif c.quality == QUALITY_MEDIUM:
        med_src_count = len({e.source.strip() for e in c.evidences if e.quality == QUALITY_MEDIUM})
        med_score = sum(e.score for e in c.evidences if e.quality == QUALITY_MEDIUM)
        return (-2, -med_src_count, -med_score, c.normalized_value)
    else:
        low_src_count = len({e.source.strip() for e in c.evidences if e.quality == QUALITY_LOW})
        low_score = sum(e.score for e in c.evidences if e.quality == QUALITY_LOW)
        return (-1, -low_src_count, -low_score, c.normalized_value)


def resolve_naming(
    evidences: List[NamingEvidence],
    parse_result: Optional[FilenameParseResult] = None,
    required_fields: Tuple[str, ...] = ("title", "date"),
) -> ResolvedResult:
    """
    Stufe 2: Führt alle Evidences aus Parser, OCR und Bildmetadaten zusammen
    und berechnet deterministisch das ResolvedResult für category, title und date.

    Die Eingabeliste `evidences` bleibt unverändert.
    """
    # 0a. Eingabe-Container validieren
    if not isinstance(evidences, list):
        raise ValueError(f"evidences muss eine Liste sein, nicht {type(evidences).__name__}")
    if parse_result is not None and not isinstance(parse_result, FilenameParseResult):
        raise ValueError(f"parse_result muss FilenameParseResult oder None sein, nicht {type(parse_result).__name__}")

    # 0b. required_fields validieren
    if required_fields is None:
        raise ValueError("required_fields darf nicht None sein")
    if isinstance(required_fields, str):
        raise ValueError("required_fields muss ein Tuple sein, kein String")
    if not isinstance(required_fields, tuple):
        raise ValueError(f"required_fields muss ein Tuple sein, nicht {type(required_fields).__name__}")
    seen_rf = set()
    for i, rf_val in enumerate(required_fields):
        if not isinstance(rf_val, str):
            raise ValueError(f"required_fields[{i}]: Element muss ein String sein, nicht {type(rf_val).__name__}")
        if not rf_val.strip():
            raise ValueError(f"required_fields[{i}]: Leeres Pflichtfeld")
        if rf_val not in ALLOWED_FIELDS:
            raise ValueError(f"Unbekanntes Pflichtfeld '{rf_val}' (erlaubt: {ALLOWED_FIELDS})")
        if rf_val in seen_rf:
            raise ValueError(f"Doppeltes Pflichtfeld '{rf_val}'")
        seen_rf.add(rf_val)

    # 0c. Einzelne Evidences validieren (ohne Mutation)
    for idx, ev in enumerate(evidences):
        _validate_evidence(ev, idx)

    if parse_result is not None:
        for idx, ev in enumerate(parse_result.evidences):
            _validate_evidence(ev, f"parse_result_{idx}")

    # 1. Alle zu verarbeitenden Evidences sammeln (Parser + zusätzliche)
    all_evidences: List[NamingEvidence] = []
    if parse_result is not None and parse_result.evidences:
        all_evidences.extend(parse_result.evidences)
    all_evidences.extend(evidences)

    # 2. Deduplizierung nach (field, source, normalized_value)
    # Behalte das jeweils stärkste Evidence (höchste Qualität, dann höchster Score)
    dedup_map: Dict[Tuple[str, str, str], NamingEvidence] = {}
    for ev in all_evidences:
        norm_val = _normalize_for_comparison(ev.field, ev.value)
        key = (ev.field, ev.source.strip(), norm_val)
        if key not in dedup_map:
            dedup_map[key] = ev
        else:
            existing = dedup_map[key]
            if (_quality_rank(ev.quality) > _quality_rank(existing.quality)) or (
                _quality_rank(ev.quality) == _quality_rank(existing.quality) and ev.score > existing.score
            ):
                dedup_map[key] = ev

    # 3. Gruppierung nach Feld
    evidences_by_field: Dict[str, List[NamingEvidence]] = {f: [] for f in ALLOWED_FIELDS}
    for (field_name, _, _), ev in dedup_map.items():
        evidences_by_field[field_name].append(ev)

    # 4. Erkennung von Widersprüchen innerhalb derselben Quelle pro Feld
    intra_source_contradictions: Dict[str, List[str]] = {f: [] for f in ALLOWED_FIELDS}
    for field_name in ALLOWED_FIELDS:
        sources_to_norm_vals: Dict[str, set] = {}
        for ev in evidences_by_field[field_name]:
            src = ev.source.strip()
            norm = _normalize_for_comparison(field_name, ev.value)
            if src not in sources_to_norm_vals:
                sources_to_norm_vals[src] = set()
            sources_to_norm_vals[src].add(norm)
        for src, norm_set in sources_to_norm_vals.items():
            if len(norm_set) > 1:
                intra_source_contradictions[field_name].append(src)
        intra_source_contradictions[field_name].sort()

    # 5. Kandidaten pro Feld aufbauen, aggregieren und auflösen
    resolved_fields: Dict[str, ResolvedField] = {}
    field_review_reasons: List[str] = []

    for field_name in ALLOWED_FIELDS:
        ev_list = evidences_by_field[field_name]
        if not ev_list:
            resolved_fields[field_name] = ResolvedField(
                value=None, quality=None, sources=[], alternatives=[],
                conflict=False, reason="Keine Evidences vorhanden"
            )
            continue

        # 5a. Gruppieren nach normalized_value
        candidates_map: Dict[str, _Candidate] = {}
        for ev in ev_list:
            norm = _normalize_for_comparison(field_name, ev.value)
            if norm not in candidates_map:
                candidates_map[norm] = _Candidate(normalized_value=norm)
            cand = candidates_map[norm]
            cand.evidences.append(ev)
            if ev.source.strip() not in cand.sources:
                cand.sources.append(ev.source.strip())

        # 5b. Kandidaten aggregieren
        candidates: List[_Candidate] = []
        for cand in candidates_map.values():
            cand.sources.sort()

            # Quellen nach Qualitätsstufe aufteilen
            high_evidence_sources = {e.source.strip() for e in cand.evidences if e.quality == QUALITY_HIGH}
            medium_evidence_sources = {e.source.strip() for e in cand.evidences if e.quality == QUALITY_MEDIUM}

            # Aggregierte Qualität bestimmen
            if high_evidence_sources:
                cand.quality = QUALITY_HIGH
            elif len(medium_evidence_sources) >= 2:
                cand.quality = QUALITY_HIGH
            elif medium_evidence_sources:
                cand.quality = QUALITY_MEDIUM
            else:
                cand.quality = QUALITY_LOW

            # Starke Quellenunterstützung berechnen (für Konfliktentscheidung und Rangfolge):
            # - Quellen mit eigenem HIGH-Evidence tragen stark
            # - Unabhängige MEDIUM-Quellen, die zusammen die Hochstufung bewirken, tragen stark
            # - LOW-Quellen tragen niemals stark
            cand.strong_sources = set(high_evidence_sources)
            if len(medium_evidence_sources) >= 2:
                cand.strong_sources |= medium_evidence_sources

            cand.strong_source_count = len(cand.strong_sources)
            cand.strong_total_score = sum(
                e.score for e in cand.evidences
                if e.quality in (QUALITY_HIGH, QUALITY_MEDIUM) and e.source.strip() in cand.strong_sources
            )

            # Repräsentativen Originalwert deterministisch auswählen
            # Für Titel und Kategorie äußere Leerzeichen entfernen (`.strip()`)
            sorted_evs = sorted(
                cand.evidences,
                key=lambda e: (-_quality_rank(e.quality), -e.score, e.value.strip())
            )
            cand.representative_original_value = sorted_evs[0].value.strip()
            candidates.append(cand)

        # 5c. Kandidaten nach starker Rangfolge deterministisch sortieren
        candidates.sort(key=_candidate_sort_key)

        best = candidates[0]
        alts = [c.representative_original_value for c in candidates[1:]]

        # 5d. Konflikt- und Statusanalyse für das Feld
        if len(candidates) == 1:
            rf = ResolvedField(
                value=best.representative_original_value,
                quality=best.quality,
                sources=sorted(best.sources),
                alternatives=[],
                conflict=False,
                reason=f"Eindeutiger {best.quality}-Kandidat aus {len(best.sources)} Quelle(n)"
            )
        else:
            high_cands = [c for c in candidates if c.quality == QUALITY_HIGH]
            med_cands = [c for c in candidates if c.quality == QUALITY_MEDIUM]

            if len(high_cands) >= 2:
                # Prüfe echten Konflikt anhand starker Quellenunterstützung.
                # Ein einzelnes LOW-Evidence ist keine starke externe Bestätigung
                # und darf einen Intra-Source-Widerspruch nicht zum Konflikt hochstufen.
                all_strong_union = set()
                for c in high_cands:
                    all_strong_union |= c.strong_sources

                if len(all_strong_union) <= 1:
                    # Alle starke Unterstützung kommt aus höchstens einer Quelle
                    # → Intra-Source-Widerspruch, kein harter Konflikt
                    rf = ResolvedField(
                        value=best.representative_original_value,
                        quality=best.quality,
                        sources=sorted(best.sources),
                        alternatives=alts,
                        conflict=False,
                        reason=f"Widersprüchliche HIGH-Werte derselben Quelle ({best.representative_original_value} vs. {', '.join(alts)})"
                    )
                else:
                    # Starke Unterstützung aus mehreren unabhängigen Quellen → echter Konflikt
                    rf = ResolvedField(
                        value=best.representative_original_value,
                        quality=best.quality,
                        sources=sorted(best.sources),
                        alternatives=alts,
                        conflict=True,
                        reason=f"Konflikt: Mehrere HIGH-Kandidaten mit unabhängiger Quellenunterstützung ({best.representative_original_value} vs. {', '.join(c.representative_original_value for c in high_cands[1:])})"
                    )
            elif best.quality == QUALITY_HIGH and all(c.quality == QUALITY_LOW for c in candidates[1:]):
                # Klarer Gewinner gegen LOW
                rf = ResolvedField(
                    value=best.representative_original_value,
                    quality=best.quality,
                    sources=sorted(best.sources),
                    alternatives=alts,
                    conflict=False,
                    reason=f"Klarer Gewinner ({best.representative_original_value} [HIGH] vor {', '.join(alts)} [LOW])"
                )
            elif best.quality == QUALITY_HIGH and med_cands:
                # HIGH mit abweichendem MEDIUM-Kandidaten
                rf = ResolvedField(
                    value=best.representative_original_value,
                    quality=best.quality,
                    sources=sorted(best.sources),
                    alternatives=alts,
                    conflict=False,
                    reason=f"HIGH-Kandidat ({best.representative_original_value}) mit abweichender MEDIUM-Alternative ({', '.join(c.representative_original_value for c in med_cands)})"
                )
            else:
                # Kein HIGH-Kandidat und mehrere Werte (z.B. MEDIUM vs MEDIUM)
                rf = ResolvedField(
                    value=best.representative_original_value,
                    quality=best.quality,
                    sources=sorted(best.sources),
                    alternatives=alts,
                    conflict=False,
                    reason=f"Mehrere {best.quality}-Kandidaten ohne HIGH-Konsens ({best.representative_original_value} vs. {', '.join(alts)})"
                )

        resolved_fields[field_name] = rf

        # Review-Gründe pro Feld registrieren
        if rf.conflict:
            field_review_reasons.append(f"Konflikt im Feld '{field_name}': {rf.reason}")
        elif rf.quality in (QUALITY_MEDIUM, QUALITY_LOW):
            field_review_reasons.append(f"Feld '{field_name}' erreicht nur Qualität {rf.quality}")
        elif rf.quality == QUALITY_HIGH and any(c.quality == QUALITY_MEDIUM for c in candidates[1:]):
            field_review_reasons.append(f"Feld '{field_name}' hat abweichende MEDIUM-Alternativen")

        # Intra-Source-Widersprüche registrieren
        for src in intra_source_contradictions[field_name]:
            field_review_reasons.append(f"Widersprüchliche Werte der Quelle '{src}' für Feld '{field_name}'")

    # 6. Gesamtergebnis bilden
    cat_field = resolved_fields["category"]
    title_field = resolved_fields["title"]
    date_field = resolved_fields["date"]

    missing_fields: List[str] = []
    for req in required_fields:
        if req in resolved_fields and resolved_fields[req].value is None:
            missing_fields.append(req)
            field_review_reasons.append(f"Pflichtfeld '{req}' fehlt")
    missing_fields.sort()

    is_generic = parse_result.is_generic if parse_result is not None else False
    is_ambiguous = parse_result.is_ambiguous if parse_result is not None else False

    if is_ambiguous:
        field_review_reasons.append("Dateiname ist mehrdeutig (is_ambiguous=True)")

    unique_review_reasons = sorted(set(field_review_reasons))

    # 7. Gesamtstatus bestimmen
    has_conflict = any(resolved_fields[f].conflict for f in ALLOWED_FIELDS)
    has_missing_fields = bool(missing_fields)
    has_review_reasons = bool(unique_review_reasons)
    has_weak_resolved_fields = any(
        resolved_fields[f].value is not None and resolved_fields[f].quality != QUALITY_HIGH
        for f in ALLOWED_FIELDS
    )
    has_ambiguous_filename = is_ambiguous

    if has_conflict:
        status = STATUS_CONFLICT
    elif (
        has_missing_fields
        or has_review_reasons
        or has_weak_resolved_fields
        or has_ambiguous_filename
    ):
        status = STATUS_NEEDS_REVIEW
    else:
        status = STATUS_AUTO_MERGED

    return ResolvedResult(
        category=cat_field,
        title=title_field,
        date=date_field,
        status=status,
        missing_fields=missing_fields,
        review_reasons=unique_review_reasons,
        filename_is_generic=is_generic,
        filename_is_ambiguous=is_ambiguous,
    )
