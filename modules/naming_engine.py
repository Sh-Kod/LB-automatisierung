"""
modules/naming_engine.py

Smart Hybrid Naming Engine – Stufe 1: Dateinamen-Parser und Datenstrukturen.
Ausschließlich Python-Standardbibliothek (re, os, calendar, dataclasses).
"""

import os
import re
import calendar
from dataclasses import dataclass, field
from typing import Optional, List, Tuple


@dataclass
class NamingEvidence:
    """Repräsentiert einen einzelnen Nachweis (Evidence) für ein Naming-Feld."""
    field: str              # "category", "title", "date"
    value: str              # z.B. "FK", "Shrek", "18_10"
    source: str             # z.B. "filename_parser", "filename_category"
    quality: str            # "HIGH", "MEDIUM", "LOW" (auf Parser-Ebene)
    score: int              # Punktewert für Priorisierung
    raw_text: Optional[str] = None


@dataclass
class FilenameParseResult:
    """Repräsentiert das Ergebnis des Dateinamen-Parsers inklusive Metadaten."""
    evidences: List[NamingEvidence] = field(default_factory=list)
    is_generic: bool = False
    is_ambiguous: bool = False
    raw_stem: str = ""


# Bekannte Kategorie-Präfixe, absteigend nach Länge sortiert für eindeutige Treffer
_PREFIX_LIST: List[Tuple[str, str]] = [
    ("mein_erster_kinobesuch", "MeK"),
    ("zurueck_im_kino", "ZiK"),
    ("zuruck_im_kino", "ZiK"),
    ("zurück_im_kino", "ZiK"),
    ("filmklassiker", "FK"),
    ("traumkino", "TK"),
    ("zik", "ZiK"),
    ("mek", "MeK"),
    ("fk", "FK"),
    ("tk", "TK"),
]

# Muster für generische Dateinamen (z.B. bild_001, whatsapp_image...)
GENERIC_PATTERNS: List[str] = [
    r"^bild[_-]?\d*$",
    r"^image[_-]?\d*$",
    r"^img[_-]?\d+$",
    r"^scan[_-]?\d*$",
    r"^foto[_-]?\d*$",
    r"^screenshot[_-]?\d*$",
    r"^whatsapp[_-]?image.*$",
    r"^untitled[_-]?\d*$",
    r"^dsc[_-]?\d+$",
    r"^photo[_-]?\d*$",
]


def _get_basename(path_or_filename: str) -> str:
    r"""Ermittelt den Dateinamen robust gegenüber sowohl Windows- (\) als auch Unix- (/) Pfadtrennzeichen,
    unabhängig vom Betriebssystem, auf dem der Code läuft."""
    s = path_or_filename.replace("\\", "/")
    return os.path.basename(s)


def _clean_token_string(name_ohne_ext: str) -> str:
    """Wandelt Bindestriche, Punkte und Leerzeichen in Unterstriche um, behält aber die Groß-/Kleinschreibung
    und Umlaute bei, um die Originalschreibweise konservativ zu schützen."""
    s = re.sub(r'[\s\-\.]+', '_', name_ohne_ext)
    s = re.sub(r'[^a-zA-Z0-9_äöüÄÖÜß]', '', s)
    s = re.sub(r'_+', '_', s).strip('_')
    return s


def _validiere_datum(tag_str: str, monat_str: str, jahr_str: Optional[str] = None) -> bool:
    """Prüft via calendar.monthrange, ob Tag und Monat (sowie optional Jahr) ein echtes Kalenderdatum bilden.
    Ohne Jahr wird das Schaltjahr 2028 als neutrale Basis genommen, damit 29_02 grundsätzlich zulässig ist.
    Mit explizitem Jahr wird das tatsächliche Jahr geprüft."""
    if not (tag_str.isdigit() and monat_str.isdigit()):
        return False
    try:
        tag = int(tag_str)
        monat = int(monat_str)
        if not (1 <= monat <= 12):
            return False
        if jahr_str is not None and jahr_str.isdigit():
            jahr = int(jahr_str)
            if not (1900 <= jahr <= 2100):
                return False
        else:
            jahr = 2028  # neutrales Schaltjahr als Basis
        _, max_tage = calendar.monthrange(jahr, monat)
        return 1 <= tag <= max_tage
    except (ValueError, TypeError):
        return False


def parse_filename(filename_or_path: str) -> FilenameParseResult:
    """
    Stufe 1: Extrahiert NamingEvidences aus einem Dateinamen (unabhängig von Bild-/OCR-Daten).
    Gibt ein FilenameParseResult zurück.

    Konservative Grundsätze:
    - Vierstellige Jahreszahlen (1917, 1984, 2049, 2026) sind kein Datum und bleiben Teil des Titels,
      es sei denn, sie sind Teil eines vollständigen Datums mit Tag und Monat (TT_MM_YYYY).
    - Generische Dateinamen nach Datumsextraktion erzeugen KEIN Titel-Evidence (`is_generic=True`).
    - Mehrdeutige rein numerische Dateinamen (2001_12_05) erzeugen KEIN Titel-/Datum-Evidence (`is_ambiguous=True`).
    - Kalendarisch ungültige Datumsstrukturen (film_29_02_2025, film_31_04) lösen `is_ambiguous=True` aus,
      um fälschliche HIGH-Titel mit enthaltenem ungültigem Datum zu verhindern.
    - Titelnormalisierung bleibt konservativ (kein erzwungenes Title Case, Erhalt von WALL_E, Se7en etc.).
    """
    evidences: List[NamingEvidence] = []

    # 1. Dateiendung und Pfad robust (Windows & Unix) entfernen
    basename = _get_basename(filename_or_path)
    name_ohne_ext, _ = os.path.splitext(basename)
    if not name_ohne_ext:
        return FilenameParseResult(evidences=[], raw_stem="")

    # 2. Technisch normalisieren (unter Beibehaltung der Schreibung)
    norm = _clean_token_string(name_ohne_ext)
    if not norm:
        return FilenameParseResult(evidences=[], raw_stem=name_ohne_ext)

    # 3. Bekannte Kategoriepräfixe erkennen (case-insensitive)
    rest = norm
    norm_lower = norm.lower()
    for prefix, cat_code in _PREFIX_LIST:
        if norm_lower == prefix:
            evidences.append(NamingEvidence(
                field="category",
                value=cat_code,
                source="filename_category",
                quality="HIGH",
                score=35,
                raw_text=norm[:len(prefix)]
            ))
            rest = ""
            break
        elif norm_lower.startswith(prefix + "_"):
            evidences.append(NamingEvidence(
                field="category",
                value=cat_code,
                source="filename_category",
                quality="HIGH",
                score=35,
                raw_text=norm[:len(prefix)]
            ))
            rest = rest[len(prefix) + 1:]
            break

    if not rest:
        return FilenameParseResult(evidences=evidences, raw_stem=norm)

    tokens = [t for t in rest.split("_") if t]
    if not tokens:
        return FilenameParseResult(evidences=evidences, raw_stem=norm)

    # 4. Rein numerische Muster prüfen (z.B. 2001_12_05 vs 1984_15_03)
    if all(t.isdigit() for t in tokens):
        if len(tokens) == 1:
            evidences.append(NamingEvidence(
                field="title",
                value=tokens[0],
                source="filename_parser",
                quality="HIGH",
                score=30,
                raw_text=tokens[0]
            ))
            return FilenameParseResult(evidences=evidences, raw_stem=norm)
        elif len(tokens) == 2 and _validiere_datum(tokens[0], tokens[1]):
            datum_val = f"{int(tokens[0]):02d}_{int(tokens[1]):02d}"
            evidences.append(NamingEvidence(
                field="date",
                value=datum_val,
                source="filename_parser",
                quality="HIGH",
                score=40,
                raw_text=f"{tokens[0]}_{tokens[1]}"
            ))
            return FilenameParseResult(evidences=evidences, raw_stem=norm)
        elif len(tokens) == 3 and len(tokens[2]) == 4 and _validiere_datum(tokens[0], tokens[1], tokens[2]):
            datum_val = f"{int(tokens[0]):02d}_{int(tokens[1]):02d}"
            evidences.append(NamingEvidence(
                field="date",
                value=datum_val,
                source="filename_parser",
                quality="HIGH",
                score=40,
                raw_text=f"{tokens[0]}_{tokens[1]}"
            ))
            return FilenameParseResult(evidences=evidences, raw_stem=norm)
        elif len(tokens) == 3 and len(tokens[0]) == 4:
            # z.B. 2001_12_05 oder 1984_15_03
            is_iso_date = _validiere_datum(tokens[2], tokens[1], tokens[0])  # YYYY_MM_DD
            is_tt_mm = _validiere_datum(tokens[1], tokens[2])               # TT_MM

            if is_iso_date and is_tt_mm:
                return FilenameParseResult(
                    evidences=evidences,
                    is_generic=False,
                    is_ambiguous=True,
                    raw_stem=norm
                )
            elif is_tt_mm and not is_iso_date:
                datum_val = f"{int(tokens[1]):02d}_{int(tokens[2]):02d}"
                evidences.append(NamingEvidence(
                    field="date",
                    value=datum_val,
                    source="filename_parser",
                    quality="HIGH",
                    score=40,
                    raw_text=f"{tokens[1]}_{tokens[2]}"
                ))
                evidences.append(NamingEvidence(
                    field="title",
                    value=tokens[0],
                    source="filename_parser",
                    quality="HIGH",
                    score=30,
                    raw_text=tokens[0]
                ))
                return FilenameParseResult(evidences=evidences, raw_stem=norm)
            else:
                return FilenameParseResult(
                    evidences=evidences,
                    is_generic=False,
                    is_ambiguous=True,
                    raw_stem=norm
                )
        else:
            return FilenameParseResult(
                evidences=evidences,
                is_generic=False,
                is_ambiguous=True,
                raw_stem=norm
            )

    # 5. Datumsprüfung am Suffix oder Präfix mit/ohne Jahr & Erkennung ungültiger Datumsformate
    date_found = False
    date_val = ""
    raw_date = ""

    # Prüfung hinten (Suffix mit Jahr: TT_MM_YYYY)
    if len(tokens) >= 3 and tokens[-1].isdigit() and len(tokens[-1]) == 4 and tokens[-2].isdigit() and len(tokens[-2]) == 2 and tokens[-3].isdigit() and len(tokens[-3]) == 2:
        if _validiere_datum(tokens[-3], tokens[-2], tokens[-1]):
            date_val = f"{int(tokens[-3]):02d}_{int(tokens[-2]):02d}"
            raw_date = f"{tokens[-3]}_{tokens[-2]}_{tokens[-1]}"
            tokens = tokens[:-3]
            date_found = True
        else:
            # Datumsähnliches Muster (z.B. 29_02_2025 oder 31_04_2025 am Ende), aber kalendarisch ungültig!
            return FilenameParseResult(
                evidences=evidences,
                is_generic=False,
                is_ambiguous=True,
                raw_stem=norm
            )
    # Prüfung hinten (Suffix ohne Jahr: TT_MM)
    elif len(tokens) >= 2 and tokens[-1].isdigit() and len(tokens[-1]) == 2 and tokens[-2].isdigit() and len(tokens[-2]) == 2:
        if _validiere_datum(tokens[-2], tokens[-1]):
            date_val = f"{int(tokens[-2]):02d}_{int(tokens[-1]):02d}"
            raw_date = f"{tokens[-2]}_{tokens[-1]}"
            tokens = tokens[:-2]
            date_found = True
        else:
            # Datumsähnliches Suffix ohne Jahr (z.B. 31_02 oder 31_04 am Ende), aber kalendarisch ungültig!
            return FilenameParseResult(
                evidences=evidences,
                is_generic=False,
                is_ambiguous=True,
                raw_stem=norm
            )
    # Prüfung vorne (Präfix mit Jahr: TT_MM_YYYY)
    elif len(tokens) >= 3 and tokens[2].isdigit() and len(tokens[2]) == 4 and tokens[1].isdigit() and len(tokens[1]) == 2 and tokens[0].isdigit() and len(tokens[0]) == 2:
        if _validiere_datum(tokens[0], tokens[1], tokens[2]):
            date_val = f"{int(tokens[0]):02d}_{int(tokens[1]):02d}"
            raw_date = f"{tokens[0]}_{tokens[1]}_{tokens[2]}"
            tokens = tokens[3:]
            date_found = True
        else:
            # Datumsähnliches Präfix mit Jahr (z.B. 29_02_2025 vorne), kalendarisch ungültig!
            return FilenameParseResult(
                evidences=evidences,
                is_generic=False,
                is_ambiguous=True,
                raw_stem=norm
            )
    # Prüfung vorne (Präfix ohne Jahr: TT_MM)
    elif len(tokens) >= 2 and tokens[1].isdigit() and len(tokens[1]) == 2 and tokens[0].isdigit() and len(tokens[0]) == 2:
        if _validiere_datum(tokens[0], tokens[1]):
            date_val = f"{int(tokens[0]):02d}_{int(tokens[1]):02d}"
            raw_date = f"{tokens[0]}_{tokens[1]}"
            tokens = tokens[2:]
            date_found = True
        else:
            # Datumsähnliches Präfix ohne Jahr (z.B. 31_02 oder 30_02 vorne), kalendarisch ungültig!
            return FilenameParseResult(
                evidences=evidences,
                is_generic=False,
                is_ambiguous=True,
                raw_stem=norm
            )

    if date_found and date_val:
        evidences.append(NamingEvidence(
            field="date",
            value=date_val,
            source="filename_parser",
            quality="HIGH",
            score=40,
            raw_text=raw_date
        ))

    # 6. Verbleibende Titel-Tokens verarbeiten & auf generische Muster nach Datumsextraktion prüfen
    if not tokens:
        return FilenameParseResult(evidences=evidences, raw_stem=norm)

    raw_title_remaining = "_".join(tokens)
    if any(re.match(pat, raw_title_remaining, re.IGNORECASE) for pat in GENERIC_PATTERNS):
        return FilenameParseResult(
            evidences=evidences,  # Datum und Kategorie (falls vorhanden) bleiben erhalten!
            is_generic=True,
            is_ambiguous=False,
            raw_stem=norm
        )

    evidences.append(NamingEvidence(
        field="title",
        value=raw_title_remaining,
        source="filename_parser",
        quality="HIGH",
        score=30,
        raw_text=raw_title_remaining
    ))

    return FilenameParseResult(evidences=evidences, raw_stem=norm)
