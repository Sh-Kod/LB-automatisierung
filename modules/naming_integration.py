"""
modules/naming_integration.py

Smart Hybrid Naming Engine – Stufe 4: Integrationsschicht.
Verbindet Dateinamen-Parser (Stufe 1), Resolver (Stufe 2) und OCR-Evidence (Stufe 3)
zu einer einheitlichen Pipeline für DCP-Namensvorschläge und Telegram-Dialoge.

Ausschließlich Python-Standardbibliothek plus modulinterne Abhängigkeiten.
"""

from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple
import os
import re
import unicodedata

from modules.naming_engine import parse_filename, FilenameParseResult, NamingEvidence
from modules.naming_resolver import (
    resolve_naming,
    ResolvedResult,
    STATUS_AUTO_MERGED,
    STATUS_NEEDS_REVIEW,
    STATUS_CONFLICT,
)
from modules.naming_ocr import (
    analyze_image_with_tesseract,
    NamingOCRResult,
)


# Standard-Kategorie-Präfixe für DCP-Namen
CATEGORY_PREFIX_MAP: Dict[str, str] = {
    "MeK": "LB_MeK",
    "ZiK": "LB_ZiK",
    "TK": "LB_TK",
    "FK": "LB_FK",
}

# Dialog-Aktions-Konstanten
ACTION_SET_NAME = "SET_NAME"
ACTION_ASK_CUSTOM = "ASK_CUSTOM"
ACTION_ASK_DATE = "ASK_DATE"
ACTION_ASK_TITLE = "ASK_TITLE"
ACTION_SKIP = "SKIP"


# ═══════════════════════════════════════════════════════════════════════════
# Öffentliche Dataclasses
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class NamingProcessResult:
    """Gesamtergebnis des Naming-Integrationsprozesses für ein Bild."""
    image_path: str
    parse_result: FilenameParseResult
    ocr_result: NamingOCRResult
    resolved: ResolvedResult
    dcp_name_proposal: Optional[str]
    status: str
    dialog_message: str
    options: Dict[str, Tuple[str, Optional[str]]] = field(default_factory=dict)


# ═══════════════════════════════════════════════════════════════════════════
# Hilfsfunktionen zur DCP-Namensbereinigung und -zusammensetzung
# ═══════════════════════════════════════════════════════════════════════════

def clean_dcp_name(name: str) -> str:
    """Wandelt Umlaute um und entfernt unerlaubte Zeichen aus DCP-Namen.
    Entspricht der zentralen _bereinige_dcp_name-Bereinigung."""
    if not name:
        return ""
    for orig, repl in [
        ("ä", "ae"), ("ö", "oe"), ("ü", "ue"),
        ("Ä", "Ae"), ("Ö", "Oe"), ("Ü", "Ue"),
        ("ß", "ss"),
    ]:
        name = name.replace(orig, repl)
    # Weitere diakritische Zeichen (z.B. é -> e, è -> e) normalisieren
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    # Nicht-ASCII bzw. Sonderzeichen entfernen
    name = re.sub(r"[^a-zA-Z0-9_\-]", "_", name)
    return re.sub(r"_+", "_", name).strip("_")


def build_dcp_name(
    category: Optional[str] = None,
    title: Optional[str] = None,
    date: Optional[str] = None,
) -> Optional[str]:
    """Erzeugt einen wohlgeformten DCP-Namen aus den Feldern.
    Format: LB_<Kategorie>_<Titel>_<Datum> (bzw. LB_<Titel>_<Datum> wenn Kategorie fehlt).
    Gibt None zurück, wenn kein Titel vorhanden ist."""
    if not title or not title.strip():
        return None

    clean_title = clean_dcp_name(title.strip())
    if not clean_title:
        return None

    # Präfix bestimmen
    if category and category.strip():
        cat = category.strip()
        prefix = CATEGORY_PREFIX_MAP.get(cat, f"LB_{cat}" if not cat.startswith("LB_") else cat)
    else:
        prefix = "LB"

    # Teile zusammenfügen
    parts = [prefix, clean_title]
    if date and date.strip():
        clean_date = clean_dcp_name(date.strip())
        if clean_date:
            parts.append(clean_date)

    raw_result = "_".join(parts)
    return clean_dcp_name(raw_result)


def build_alternative_dcp_names(
    resolved: ResolvedResult,
) -> List[str]:
    """Erzeugt alternative DCP-Namensvorschläge aus den Alternativen des Resolvers."""
    proposals: List[str] = []

    # 1. Primärer Vorschlag
    p1 = build_dcp_name(resolved.category.value, resolved.title.value, resolved.date.value)
    if p1:
        proposals.append(p1)

    # 2. Alternative Titel
    for alt_title in resolved.title.alternatives:
        p = build_dcp_name(resolved.category.value, alt_title, resolved.date.value)
        if p and p not in proposals:
            proposals.append(p)

    # 3. Alternative Kategorien
    for alt_cat in resolved.category.alternatives:
        p = build_dcp_name(alt_cat, resolved.title.value, resolved.date.value)
        if p and p not in proposals:
            proposals.append(p)

    # 4. Alternative Datumswerte
    for alt_date in resolved.date.alternatives:
        p = build_dcp_name(resolved.category.value, resolved.title.value, alt_date)
        if p and p not in proposals:
            proposals.append(p)

    return proposals


# ═══════════════════════════════════════════════════════════════════════════
# Telegram-Dialog-Formatierung
# ═══════════════════════════════════════════════════════════════════════════

def format_telegram_dialog(
    resolved: ResolvedResult,
    proposal: Optional[str],
    alternatives: Optional[List[str]] = None,
) -> Tuple[str, Dict[str, Tuple[str, Optional[str]]]]:
    """Erstellt den formatierten Dialogtext und die Aktionszuordnung für Telegram.

    Returns:
        (dialog_message, options_dict)
    """
    t = "─" * 30
    options: Dict[str, Tuple[str, Optional[str]]] = {}
    alts = alternatives or []

    # ── 1. AUTO_MERGED ───────────────────────────────────────────────────
    if resolved.status == STATUS_AUTO_MERGED and proposal:
        msg_lines = [
            t,
            "Vorschlag (automatisch erkannt):",
            f"{proposal}\n",
            "[1]  Übernehmen",
            "[2]  Eigenen Namen eingeben",
            "[3]  Überspringen",
            t,
            "(Timeout: 60 Min)",
        ]
        options = {
            "1": (ACTION_SET_NAME, proposal),
            "2": (ACTION_ASK_CUSTOM, None),
            "3": (ACTION_SKIP, None),
        }
        return "\n".join(msg_lines), options

    # ── 2. CONFLICT ──────────────────────────────────────────────────────
    if resolved.status == STATUS_CONFLICT:
        msg_lines = [
            t,
            "⚠️ Konflikt erkannt:",
        ]
        for reason in resolved.review_reasons:
            msg_lines.append(f"• {reason}")
        msg_lines.append("")

        if len(alts) >= 2:
            msg_lines.append("Mögliche Varianten:")
            opt_idx = 1
            for alt in alts[:3]:  # max 3 Varianten anzeigen
                msg_lines.append(f"[{opt_idx}]  {alt}")
                options[str(opt_idx)] = (ACTION_SET_NAME, alt)
                opt_idx += 1

            msg_lines.append(f"[{opt_idx}]  Eigenen Namen eingeben")
            options[str(opt_idx)] = (ACTION_ASK_CUSTOM, None)
            opt_idx += 1

            msg_lines.append(f"[{opt_idx}]  Überspringen")
            options[str(opt_idx)] = (ACTION_SKIP, None)
        else:
            if proposal:
                msg_lines.append(f"Erkannter Vorschlag:\n{proposal}\n")
                msg_lines.append("[1]  Vorschlag übernehmen")
                msg_lines.append("[2]  Eigenen Namen eingeben")
                msg_lines.append("[3]  Überspringen")
                options = {
                    "1": (ACTION_SET_NAME, proposal),
                    "2": (ACTION_ASK_CUSTOM, None),
                    "3": (ACTION_SKIP, None),
                }
            else:
                msg_lines.append("Kein vollständiger Vorschlag möglich.\n")
                msg_lines.append("[1]  Namen eingeben")
                msg_lines.append("[2]  Überspringen")
                options = {
                    "1": (ACTION_ASK_CUSTOM, None),
                    "2": (ACTION_SKIP, None),
                }

        msg_lines.append(t)
        msg_lines.append("(Timeout: 60 Min)")
        return "\n".join(msg_lines), options

    # ── 3. NEEDS_REVIEW ──────────────────────────────────────────────────
    missing = set(resolved.missing_fields) if resolved.missing_fields else set()
    has_missing_required = bool(missing & {"title", "date"})

    msg_lines = [
        t,
        "⚠️ Prüfung erforderlich:",
    ]
    for reason in resolved.review_reasons:
        msg_lines.append(f"• {reason}")
    msg_lines.append("")

    if has_missing_required:
        # Fehlende Pflichtfelder: kein unvollständiger Name zur direkten Übernahme
        if resolved.title.value:
            msg_lines.append(f"Erkannter Titel: {resolved.title.value}")
        if resolved.category.value:
            msg_lines.append(f"Erkannte Kategorie: {resolved.category.value}")
        if resolved.date.value:
            msg_lines.append(f"Erkanntes Datum: {resolved.date.value}")
        msg_lines.append("")

        opt_idx = 1
        if "date" in missing and "title" not in missing:
            # Titel vorhanden, Datum fehlt
            msg_lines.append(f"[{opt_idx}]  Datum eingeben (TT_MM)")
            options[str(opt_idx)] = (ACTION_ASK_DATE, None)
            opt_idx += 1
        elif "title" in missing and "date" not in missing:
            # Datum vorhanden, Titel fehlt
            msg_lines.append(f"[{opt_idx}]  Filmtitel eingeben")
            options[str(opt_idx)] = (ACTION_ASK_TITLE, None)
            opt_idx += 1
        else:
            # Beide fehlen
            msg_lines.append(f"[{opt_idx}]  Filmtitel eingeben")
            options[str(opt_idx)] = (ACTION_ASK_TITLE, None)
            opt_idx += 1

        msg_lines.append(f"[{opt_idx}]  Vollständigen Namen eingeben")
        options[str(opt_idx)] = (ACTION_ASK_CUSTOM, None)
        opt_idx += 1

        msg_lines.append(f"[{opt_idx}]  Überspringen")
        options[str(opt_idx)] = (ACTION_SKIP, None)

    elif proposal:
        msg_lines.append("Erkannter Vorschlag:")
        msg_lines.append(f"{proposal}\n")
        msg_lines.append("[1]  Vorschlag übernehmen")
        msg_lines.append("[2]  Eigenen Namen eingeben")
        msg_lines.append("[3]  Überspringen")
        options = {
            "1": (ACTION_SET_NAME, proposal),
            "2": (ACTION_ASK_CUSTOM, None),
            "3": (ACTION_SKIP, None),
        }
    else:
        msg_lines.append("Kein vollständiger Vorschlag möglich.\n")
        msg_lines.append("[1]  Namen eingeben")
        msg_lines.append("[2]  Überspringen")
        options = {
            "1": (ACTION_ASK_CUSTOM, None),
            "2": (ACTION_SKIP, None),
        }

    msg_lines.append(t)
    msg_lines.append("(Timeout: 60 Min)")
    return "\n".join(msg_lines), options


def evaluate_dialog_response(
    antwort: Optional[str],
    options: Dict[str, Tuple[str, Optional[str]]],
) -> Tuple[str, Optional[str]]:
    """Wertet die Telegram-Benutzerantwort anhand der verfügbaren Optionen aus.

    Returns:
        (ACTION_SET_NAME, "final_name")
        (ACTION_ASK_CUSTOM, None)
        (ACTION_SKIP, None)
    """
    if antwort is None:
        return (ACTION_SKIP, None)

    a = antwort.strip()
    if a.lower() == "/skip":
        return (ACTION_SKIP, None)

    # 1. Ziffernauswahl aus den Optionen
    if a in options:
        return options[a]

    # 2. Direkte Namenseingabe (z.B. Benutzer tippt 'LB_FK_MeinFilm_01_02')
    return (ACTION_SET_NAME, a)


# ═══════════════════════════════════════════════════════════════════════════
# Tesseract-Adapter & Hauptprozess
# ═══════════════════════════════════════════════════════════════════════════

def create_tesseract_runner(tesseract_cmd: Optional[str] = None) -> Callable:
    """Erzeugt eine aufrufbare OCR-Runner-Funktion für analyze_image_with_tesseract().
    Setzt tesseract_cmd nur lokal für die Ausführung."""
    def _runner(image, language: str, pass_name: str) -> str:
        import pytesseract
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
        return pytesseract.image_to_string(image, lang=language)
    return _runner


def process_image_naming(
    image_path: str,
    ocr_runner: Optional[Callable] = None,
    tesseract_cmd: Optional[str] = None,
    language: str = "deu+eng",
    filename_override: Optional[str] = None,
) -> NamingProcessResult:
    """Führt die gesamte Smart Hybrid Naming Pipeline für ein Bild aus.

    Ablauf:
    1. Dateinamen-Parser (Stufe 1)
    2. Tesseract-OCR-Evidences (Stufe 3)
    3. Resolver (Stufe 2)
    4. DCP-Namenszusammensetzung
    5. Dialog-Formatierung

    Args:
        image_path: Pfad zur Bilddatei.
        ocr_runner: Optionaler injizierbarer OCR-Runner.
        tesseract_cmd: Pfad zur tesseract.exe (aus config.yaml).
        language: Tesseract-Sprachcode.
        filename_override: Optionaler Dateiname (falls abweichend von image_path).

    Returns:
        NamingProcessResult mit vollständigen Daten, Vorschlag und Dialogtext.
    """
    # 1. Dateinamen parsen
    parse_source = filename_override if filename_override else image_path
    parse_res = parse_filename(parse_source)

    # 2. OCR ausführen
    runner = ocr_runner
    if runner is None and tesseract_cmd:
        runner = create_tesseract_runner(tesseract_cmd)

    try:
        ocr_res = analyze_image_with_tesseract(
            image_path,
            ocr_runner=runner,
            language=language,
        )
    except Exception as e:
        ocr_res = NamingOCRResult(
            evidences=[],
            passes=[],
            warnings=[],
            errors=[f"Unerwarteter OCR-Fehler: {e}"],
        )

    # 3. Evidences zusammenführen
    all_evidences: List[NamingEvidence] = list(parse_res.evidences) + list(ocr_res.evidences)

    # 4. Resolver ausführen
    resolved = resolve_naming(all_evidences, parse_result=parse_res)

    # 5. DCP-Vorschlag erzeugen
    proposal = build_dcp_name(
        category=resolved.category.value,
        title=resolved.title.value,
        date=resolved.date.value,
    )

    # 6. Alternativen bei Konflikten berechnen
    alternatives = build_alternative_dcp_names(resolved)

    # 7. Dialog formatieren
    dialog_msg, options = format_telegram_dialog(
        resolved=resolved,
        proposal=proposal,
        alternatives=alternatives,
    )

    return NamingProcessResult(
        image_path=image_path,
        parse_result=parse_res,
        ocr_result=ocr_res,
        resolved=resolved,
        dcp_name_proposal=proposal,
        status=resolved.status,
        dialog_message=dialog_msg,
        options=options,
    )
