"""
modules/naming_ocr.py

Smart Hybrid Naming Engine – Stufe 3: OCR-Evidence-Schicht.
Erzeugt strukturierte NamingEvidence-Objekte aus OCR-Texten und Bildern.

Zwei klar getrennte Verantwortlichkeiten:
  A. analyze_ocr_texts()  – reine, seiteneffektfreie Textanalyse
  B. analyze_image_with_tesseract() – Bild laden, Varianten erzeugen, OCR ausführen

Alle Tesseract-Pässe gehören logisch zu einer einzigen Quelle.
Quellennamen sind feldspezifisch: tesseract_category, tesseract_title, tesseract_date.

Ausschließlich Python-Standardbibliothek plus Pillow und pytesseract (bereits in requirements.txt).
Keine globalen Seiteneffekte, keine print()-Ausgaben, keine Konfigurationsdateien.
"""

import re
import calendar
import unicodedata
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Set, Tuple

from modules.naming_engine import NamingEvidence

# Optionale Abhängigkeiten – auf Modul-Ebene importiert für Testbarkeit (patch)
try:
    from PIL import Image, ImageOps
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False


# ═══════════════════════════════════════════════════════════════════════════
# Öffentliche Dataclasses
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class OCRPassResult:
    """Ergebnis eines einzelnen OCR-Durchlaufs (Pass)."""
    pass_name: str
    text: str
    error: Optional[str] = None


@dataclass
class NamingOCRResult:
    """Gesamtergebnis der OCR-Analyse mit Evidences, Durchlauf-Details, Warnungen und Fehlern."""
    evidences: List[NamingEvidence] = field(default_factory=list)
    passes: List[OCRPassResult] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════════════════
# Interne Konstanten
# ═══════════════════════════════════════════════════════════════════════════

# Feste Reihenfolge der Tesseract-Pässe
_PASS_NAMES: Tuple[str, ...] = ("original", "grayscale", "autocontrast")

# ── Kategorie-Erkennung ──────────────────────────────────────────────────

# Vollständige eindeutige Phrasen → HIGH, score=50
_CATEGORY_PHRASES_HIGH: List[Tuple[str, str]] = [
    ("mein erster kinobesuch", "MeK"),
    ("zurück im kino", "ZiK"),
    ("traumkino", "TK"),
    ("filmklassiker", "FK"),
]

# Bekannte OCR-Fehlschreibweisen → MEDIUM, score=40
_CATEGORY_PHRASES_MEDIUM: List[Tuple[str, str]] = [
    ("zuruck im kino", "ZiK"),
    ("zurck im kino", "ZiK"),
    ("kinobesuch", "MeK"),
]

# Isolierte eindeutige Abkürzungen → MEDIUM, score=35
_CATEGORY_ABBREVIATIONS: Dict[str, str] = {
    "mek": "MeK",
    "zik": "ZiK",
    "tk": "TK",
    "fk": "FK",
}

# Alle Kategorie-Phrasen zusammen (für Titel-Bereinigung)
_ALL_CATEGORY_PHRASES: List[str] = [
    p for p, _ in _CATEGORY_PHRASES_HIGH
] + [
    p for p, _ in _CATEGORY_PHRASES_MEDIUM
]

# ── Datum-Erkennung ──────────────────────────────────────────────────────

# Erkennt Datumsformate: TT.MM, TT/MM, TT-MM, TT_MM, TT.MM. (mit optionalem Trailing-Punkt)
# Negative Lookbehind/Lookahead verhindert Treffer innerhalb von Jahreszahlen oder längeren Zahlen.
_DATE_PATTERN = re.compile(
    r'(?<!\d)(\d{1,2})[\.\/\-_](\d{1,2})\.?(?!\d)'
)

# ── Titel-Filterung ─────────────────────────────────────────────────────

# Mindestanzahl alphanumerischer Zeichen für einen gültigen Titel
_MIN_ALNUM_CHARS = 1

# Mindestlänge des Gesamttitels (nach Strip)
_MIN_TITLE_LENGTH = 1

# Muster für reines OCR-Rauschen (nur Satzzeichen, Symbole, Whitespace)
_NOISE_PATTERN = re.compile(r'^[\s\W]+$')

# ── Kinoplakat-Metazeilen-Filter ────────────────────────────────────────
#
# Die folgenden Muster erkennen typische Nicht-Titel-Zeilen auf Kinoplakaten.
# Sie sind bewusst konservativ formuliert, um echte Filmtitel nicht zu zerstören.
#
# Grundprinzip: Eine Zeile, die VOLLSTÄNDIG einem bekannten Meta-Muster entspricht,
# wird verworfen. Zeilen, die nur teilweise passen, bleiben erhalten.

# Uhrzeiten: "20 UHR", "19:30 UHR", "20:00 UHR", "20.00 UHR"
_TIME_PATTERN = re.compile(
    r'^\d{1,2}(?:[\.:]\d{2})?\s*UHR$',
    re.IGNORECASE
)

# Altersfreigaben: "FSK 6", "FSK 0", "AB 6 JAHREN", "AB 12 JAHREN", "AB 0 JAHREN"
_FSK_PATTERN = re.compile(
    r'^(?:FSK\s*\d+|AB\s+\d+\s+JAHREN?)$',
    re.IGNORECASE
)

# URLs: "WWW.KINO.DE", "KINO.DE", "WWW.LICHTBLICK-KINO.DE"
_URL_PATTERN = re.compile(
    r'^(?:(?:HTTPS?://)?WWW\.|[\w\-]+\.(?:DE|COM|NET|ORG|AT|CH|EU|INFO))',
    re.IGNORECASE
)

# Crew-/Cast-/Verleih-Präfixe: "EIN FILM VON ...", "REGIE ...", "DARSTELLER ...", "VERLEIH ..."
# Hinweis: Nur wenn auf das Präfix weiterer Text folgt (sonst könnte "MIT" ein Titel sein)
_CREW_PATTERN = re.compile(
    r'^(?:EIN\s+FILM\s+VON|REGIE|DARSTELLER|DREHBUCH|PRODUKTION|PRODUZIERT\s+VON|VERLEIH|NACH\s+(?:EINEM|DEM)\s+(?:ROMAN|BUCH))\s+.+$',
    re.IGNORECASE
)

# "MIT ..." als Cast-Zeile: mindestens ein weiterer Name nach MIT.
# Erkennt sowohl gemischte Schreibweise (Mit Michelle Williams) als auch
# reine Großschreibung aus OCR (MIT MICHELLE WILLIAMS, MIT WILLIAMS).
# "MIT" allein oder als Satzanfang ("Mit Herz und Hand") wird NICHT gefiltert.
_MIT_CREW_PATTERN = re.compile(
    r'^MIT\s+(?:[A-ZÄÖÜ][a-zäöüß]+\s+[A-ZÄÖÜ]|[A-ZÄÖÜ]{2,}(?:\s+[A-ZÄÖÜ]{2,})*$)',
    re.UNICODE | re.IGNORECASE
)

# Bekannte Promo-/Meta-Einzelphrasen (exakte Übereinstimmung, case-insensitive)
_META_PHRASES: Set[str] = {
    # Kinowerbung
    "jetzt im kino", "nur im kino", "ab jetzt im kino",
    "demnächst", "demnachst", "demnaechst",
    "premiere", "vorpremiere", "sonderpremiere",
    "special screening", "preview",
    "vorverkauf", "tickets online", "karten online",
    # Zeitangaben (inkl. Restfragmente nach Datumsextraktion, z.B. "20.00 UHR" -> "UHR")
    "heute", "morgen", "uhr", "einlass",
    "ab donnerstag", "ab freitag",
    "ab samstag", "ab sonntag", "ab montag", "ab dienstag", "ab mittwoch",
    "nur heute", "nur morgen",
    # Format / Sprache
    "original version", "originalversion",
    "originalfassung", "original fassung",
    "omu", "ov", "omeu", "omeU",
    "2d", "3d", "imax", "dolby atmos",
    # Verleih / Präsentation
    "präsentiert", "prasentiert",
    "verleih", "im verleih von",
    "eine produktion von",
}

# Zusätzliche Regex-basierte Meta-Muster (ganzzeilig)
_META_PATTERNS: List[re.Pattern] = [
    _TIME_PATTERN,
    _FSK_PATTERN,
    _URL_PATTERN,
    _CREW_PATTERN,
]


# ═══════════════════════════════════════════════════════════════════════════
# Interne Hilfsfunktionen
# ═══════════════════════════════════════════════════════════════════════════

def _normalize_for_comparison(field_name: str, value: str) -> str:
    """Normalisiert einen Wert rein für den Vergleich (identisch zum Resolver)."""
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


def _quality_rank(quality: str) -> int:
    """Gibt den numerischen Rang einer Qualitätsstufe zurück."""
    if quality == "HIGH":
        return 3
    elif quality == "MEDIUM":
        return 2
    return 1


def _validate_date(day: int, month: int) -> bool:
    """Prüft via calendar.monthrange, ob Tag und Monat ein echtes Kalenderdatum bilden.
    Schaltjahr 2028 als neutrale Basis (damit 29_02 zulässig ist)."""
    if not (1 <= month <= 12):
        return False
    try:
        _, max_days = calendar.monthrange(2028, month)
        return 1 <= day <= max_days
    except (ValueError, TypeError):
        return False


def _is_likely_year(text: str) -> bool:
    """Prüft, ob ein Text eine isolierte Jahreszahl ist (1900-2099)."""
    stripped = text.strip()
    if stripped.isdigit() and len(stripped) == 4:
        year = int(stripped)
        return 1900 <= year <= 2099
    return False


def _line_is_pure_date(line: str) -> bool:
    """Prüft, ob eine Zeile ausschließlich ein Datum enthält (nach Bereinigung nichts übrig bleibt)."""
    cleaned = _DATE_PATTERN.sub('', line)
    cleaned = re.sub(r'[\s\.\-/\_]+', '', cleaned)
    return len(cleaned) == 0


def _remove_category_from_line(line: str) -> str:
    """Entfernt erkannte Kategoriephrasen und -abkürzungen aus einer Zeile für die Titelextraktion."""
    result = line
    lower = result.lower()

    # Lange Phrasen zuerst entfernen (absteigend nach Länge)
    for phrase in sorted(_ALL_CATEGORY_PHRASES, key=len, reverse=True):
        idx = lower.find(phrase)
        if idx != -1:
            result = result[:idx] + " " + result[idx + len(phrase):]
            lower = result.lower()

    # Abkürzungen nur als eigenständige Tokens entfernen
    for abbrev in _CATEGORY_ABBREVIATIONS:
        result = re.sub(
            r'\b' + re.escape(abbrev) + r'\b',
            ' ',
            result,
            flags=re.IGNORECASE
        )

    return re.sub(r'\s+', ' ', result).strip()


def _extract_dates_from_line(line: str) -> Tuple[List[NamingEvidence], List[str], str]:
    """Extrahiert Datums-Evidences aus einer Zeile.

    Returns:
        Tuple von (evidences, warnings, remaining_line_without_dates)
    """
    evidences: List[NamingEvidence] = []
    warnings: List[str] = []

    matches = list(_DATE_PATTERN.finditer(line))
    for match in matches:
        day_str, month_str = match.group(1), match.group(2)
        try:
            day = int(day_str)
            month = int(month_str)
        except (ValueError, TypeError):
            continue

        # Jahreszahl-Schutz: Wenn das Match Teil einer vierstelligen Zahl ist, überspringen
        # Prüfe ob vor dem Match weitere Ziffern stehen, die zusammen eine Jahreszahl bilden könnten
        start = match.start()
        full_match_text = match.group(0)

        if _validate_date(day, month):
            canonical = f"{day:02d}_{month:02d}"
            evidences.append(NamingEvidence(
                field="date",
                value=canonical,
                source="tesseract_date",
                quality="HIGH",
                score=45,
                raw_text=full_match_text.rstrip('.')
            ))
        else:
            warnings.append(
                f"Ungültiges Datum in OCR gefunden: '{full_match_text.rstrip('.')}' "
                f"(Tag={day}, Monat={month})"
            )

    # Datum-Matches aus der Zeile entfernen
    remaining = line
    for match in reversed(matches):  # reversed um Indizes nicht zu verschieben
        remaining = remaining[:match.start()] + " " + remaining[match.end():]
    remaining = re.sub(r'\s+', ' ', remaining).strip()

    return evidences, warnings, remaining


def _extract_categories_from_line(line: str) -> List[NamingEvidence]:
    """Extrahiert Kategorie-Evidences aus einer Zeile."""
    evidences: List[NamingEvidence] = []
    lower = line.lower()

    # Vollständige Phrasen (HIGH, score=50)
    for phrase, cat_code in _CATEGORY_PHRASES_HIGH:
        if phrase in lower:
            evidences.append(NamingEvidence(
                field="category",
                value=cat_code,
                source="tesseract_category",
                quality="HIGH",
                score=50,
                raw_text=phrase,
            ))

    # OCR-Varianten (MEDIUM, score=40)
    for phrase, cat_code in _CATEGORY_PHRASES_MEDIUM:
        if phrase in lower:
            evidences.append(NamingEvidence(
                field="category",
                value=cat_code,
                source="tesseract_category",
                quality="MEDIUM",
                score=40,
                raw_text=phrase,
            ))

    # Isolierte Abkürzungen (MEDIUM, score=35) – nur als eigenständige Tokens
    tokens = re.findall(r'\b[a-zA-ZäöüÄÖÜß]+\b', lower)
    for token in tokens:
        if token in _CATEGORY_ABBREVIATIONS:
            evidences.append(NamingEvidence(
                field="category",
                value=_CATEGORY_ABBREVIATIONS[token],
                source="tesseract_category",
                quality="MEDIUM",
                score=35,
                raw_text=token,
            ))

    return evidences


def _is_meta_line(candidate: str) -> bool:
    """Erkennt typische Kinoplakat-Metazeilen, die keine Filmtitel sind.

    Prüft gegen:
    - Uhrzeiten (20 UHR, 19:30 UHR)
    - Altersfreigaben (FSK 6, AB 12 JAHREN)
    - URLs (WWW.KINO.DE, KINO.DE)
    - Crew-/Cast-Zeilen (EIN FILM VON ..., REGIE ..., MIT Vorname Nachname)
    - Bekannte Promo-Phrasen (JETZT IM KINO, PREMIERE, OV, 2D, 3D, ...)

    Diese Funktion ist bewusst konservativ: nur exakte Muster-Treffer
    werden als Meta-Zeilen klassifiziert, um echte Filmtitel zu schützen.
    """
    stripped = candidate.strip()
    if not stripped:
        return False

    # 1. Exakte Phrasen-Übereinstimmung (case-insensitive)
    if stripped.lower() in _META_PHRASES:
        return True

    # 2. Regex-basierte strukturelle Muster
    for pattern in _META_PATTERNS:
        if pattern.match(stripped):
            return True

    # 3. "MIT ..."-Crew-Zeilen (nur wenn Vorname + Nachname folgt)
    if _MIT_CREW_PATTERN.match(stripped):
        return True

    return False


def _is_valid_title_candidate(candidate: str) -> bool:
    """Prüft, ob ein Titel-Kandidat die Mindestanforderungen erfüllt
    und nicht als Kinoplakat-Metazeile erkannt wird.

    Filterregeln (in Prüfreihenfolge):
    1. Nicht leer, Mindestlänge erfüllt
    2. Mindestens ein alphanumerisches Zeichen
    3. Kein reines OCR-Rauschen (nur Satzzeichen/Symbole)
    4. Keine erkannte Kinoplakat-Metazeile
    """
    stripped = candidate.strip()
    if not stripped:
        return False
    if len(stripped) < _MIN_TITLE_LENGTH:
        return False

    # Mindestens ein Buchstabe oder Ziffer erforderlich
    alnum_count = sum(1 for c in stripped if c.isalnum())
    if alnum_count < _MIN_ALNUM_CHARS:
        return False

    # Reines Rauschen (nur Satzzeichen/Symbole) ablehnen
    if _NOISE_PATTERN.match(stripped):
        return False

    # Kinoplakat-Metazeilen ablehnen
    if _is_meta_line(stripped):
        return False

    return True


def _deduplicate_evidences(evidences: List[NamingEvidence]) -> List[NamingEvidence]:
    """Dedupliziert Evidences nach (field, source, normalized_value).
    Behält die stärkste Evidence (höchste Qualität, dann höchster Score).
    Sortiert deterministisch: category → title → date, dann Qualität, Score, normalisierter Wert."""
    dedup_map: Dict[Tuple[str, str, str], NamingEvidence] = {}
    for ev in evidences:
        norm_val = _normalize_for_comparison(ev.field, ev.value)
        key = (ev.field, ev.source, norm_val)
        if key not in dedup_map:
            dedup_map[key] = ev
        else:
            existing = dedup_map[key]
            if (_quality_rank(ev.quality) > _quality_rank(existing.quality)) or (
                _quality_rank(ev.quality) == _quality_rank(existing.quality)
                and ev.score > existing.score
            ):
                dedup_map[key] = ev

    result = list(dedup_map.values())

    field_order = {"category": 0, "title": 1, "date": 2}
    result.sort(key=lambda e: (
        field_order.get(e.field, 99),
        -_quality_rank(e.quality),
        -e.score,
        _normalize_for_comparison(e.field, e.value),
    ))

    return result


# ═══════════════════════════════════════════════════════════════════════════
# Öffentliche API: Reine Textanalyse
# ═══════════════════════════════════════════════════════════════════════════

def analyze_ocr_texts(
    texts: list,
    source: str = "tesseract",
) -> NamingOCRResult:
    """
    Reine, seiteneffektfreie Funktion: Analysiert OCR-Texte und erzeugt NamingEvidence-Objekte.

    Args:
        texts: Liste von OCR-Textstrings (ein String pro Tesseract-Pass).
        source: Logischer Quellenname (wird zu tesseract_category, tesseract_title, tesseract_date).

    Returns:
        NamingOCRResult mit deduplizierten Evidences, Warnungen und leerer Pass-Liste.

    Raises:
        ValueError: Bei ungültigen Eingabetypen oder -werten.

    Diese Funktion darf:
    - keine Dateien öffnen
    - kein Tesseract aufrufen
    - keine Konfiguration laden
    - keine globale Variable verändern
    - keine Netzwerkzugriffe durchführen
    - nichts ausgeben oder drucken
    """
    # ── Eingabevalidierung ───────────────────────────────────────────────
    if not isinstance(texts, list):
        raise ValueError(f"texts muss eine Liste sein, nicht {type(texts).__name__}")
    for idx, t in enumerate(texts):
        if not isinstance(t, str):
            raise ValueError(f"texts[{idx}] muss ein String sein, nicht {type(t).__name__}")
    if not isinstance(source, str) or not source.strip():
        raise ValueError(f"source muss ein nicht-leerer String sein, nicht {repr(source)}")

    # ── Quellennamen ─────────────────────────────────────────────────────
    src = source.strip()
    src_category = f"{src}_category"
    src_title = f"{src}_title"
    src_date = f"{src}_date"

    all_evidences: List[NamingEvidence] = []
    all_warnings: Set[str] = set()

    for text in texts:
        lines = text.split('\n')
        for raw_line in lines:
            line = raw_line.strip()
            if not line:
                continue

            # 1. Datum extrahieren
            date_evs, date_warnings, line_without_dates = _extract_dates_from_line(line)

            # Quellennamen auf konfigurierten Source anpassen
            for ev in date_evs:
                all_evidences.append(NamingEvidence(
                    field=ev.field,
                    value=ev.value,
                    source=src_date,
                    quality=ev.quality,
                    score=ev.score,
                    raw_text=ev.raw_text,
                ))
            all_warnings.update(date_warnings)

            # 2. Kategorien extrahieren
            cat_evs = _extract_categories_from_line(line)
            for ev in cat_evs:
                all_evidences.append(NamingEvidence(
                    field=ev.field,
                    value=ev.value,
                    source=src_category,
                    quality=ev.quality,
                    score=ev.score,
                    raw_text=ev.raw_text,
                ))

            # 3. Titel extrahieren (aus der datums- und kategorie-bereinigten Zeile)
            # Zuerst prüfen: Ist die Original-Zeile eine reine Datumszeile?
            if _line_is_pure_date(line):
                continue

            # Kategorien aus der Zeile entfernen
            title_candidate = _remove_category_from_line(line_without_dates)
            title_candidate = title_candidate.strip()

            if not title_candidate:
                continue

            # Jahreszahlen als potenzielle Titel behalten (z.B. "1917")
            # aber reine Datumswerte nicht als Titel verwenden
            if _line_is_pure_date(title_candidate):
                continue

            if _is_valid_title_candidate(title_candidate):
                all_evidences.append(NamingEvidence(
                    field="title",
                    value=title_candidate,
                    source=src_title,
                    quality="MEDIUM",
                    score=35,
                    raw_text=title_candidate,
                ))

    # ── Deduplizierung und Sortierung ────────────────────────────────────
    dedup_evidences = _deduplicate_evidences(all_evidences)

    return NamingOCRResult(
        evidences=dedup_evidences,
        passes=[],
        warnings=sorted(all_warnings),
        errors=[],
    )


# ═══════════════════════════════════════════════════════════════════════════
# Öffentliche API: Bild- und Tesseract-Ausführung
# ═══════════════════════════════════════════════════════════════════════════

def analyze_image_with_tesseract(
    image_path: str,
    ocr_runner: Optional[Callable] = None,
    language: str = "deu+eng",
) -> NamingOCRResult:
    """
    Lädt ein Bild, erzeugt kontrollierte Bildvarianten, führt einen injizierbaren
    OCR-Runner pro Variante aus und übergibt die Texte an analyze_ocr_texts().

    Args:
        image_path: Pfad zur Bilddatei (str oder Path-artiges Objekt).
        ocr_runner: Aufrufbare Funktion mit Signatur (image, language, pass_name) -> str.
                    Wenn None, wird pytesseract.image_to_string als Standardimplementierung verwendet.
        language: Tesseract-Sprachcode (Standard: "deu+eng").

    Returns:
        NamingOCRResult mit Evidences, Durchlauf-Details, Warnungen und Fehlern.

    Raises:
        ValueError: Bei ungültigen Eingabetypen oder -werten.
    """
    # ── Eingabevalidierung ───────────────────────────────────────────────
    if not isinstance(image_path, str):
        raise ValueError(f"image_path muss ein String sein, nicht {type(image_path).__name__}")
    if not image_path.strip():
        raise ValueError("image_path darf nicht leer sein")
    if not isinstance(language, str) or not language.strip():
        raise ValueError(f"language muss ein nicht-leerer String sein, nicht {repr(language)}")
    if ocr_runner is not None and not callable(ocr_runner):
        raise ValueError(f"ocr_runner muss aufrufbar (callable) oder None sein, nicht {type(ocr_runner).__name__}")

    # ── Standard-OCR-Runner ──────────────────────────────────────────────
    if ocr_runner is None:
        try:
            import pytesseract

            def _default_runner(image, lang, pass_name):
                return pytesseract.image_to_string(image, lang=lang)

            ocr_runner = _default_runner
        except ImportError:
            return NamingOCRResult(
                evidences=[], passes=[], warnings=[],
                errors=["pytesseract ist nicht installiert"],
            )

    # ── Bild laden ───────────────────────────────────────────────────────
    passes: List[OCRPassResult] = []
    errors: List[str] = []

    if not _PIL_AVAILABLE:
        return NamingOCRResult(
            evidences=[], passes=[], warnings=[],
            errors=["Pillow (PIL) ist nicht installiert"],
        )

    try:
        img = Image.open(image_path)
        img.load()  # Sofort laden, um I/O-Fehler hier zu fangen
    except Exception as e:
        return NamingOCRResult(
            evidences=[], passes=[], warnings=[],
            errors=[f"Fehler beim Laden des Bildes: {e}"],
        )

    # ── Kontrollierte Tesseract-Pässe (feste Reihenfolge) ────────────────
    # Pass 1: original
    try:
        text_original = ocr_runner(img, language, "original")
        passes.append(OCRPassResult(pass_name="original", text=text_original))
    except Exception as e:
        passes.append(OCRPassResult(pass_name="original", text="", error=str(e)))
        errors.append(f"Pass 'original' fehlgeschlagen: {e}")

    # Pass 2: grayscale
    gray_img = None
    try:
        gray_img = ImageOps.grayscale(img)
        text_gray = ocr_runner(gray_img, language, "grayscale")
        passes.append(OCRPassResult(pass_name="grayscale", text=text_gray))
    except Exception as e:
        passes.append(OCRPassResult(pass_name="grayscale", text="", error=str(e)))
        errors.append(f"Pass 'grayscale' fehlgeschlagen: {e}")

    # Pass 3: autocontrast (auf Basis des Grayscale-Bildes)
    try:
        if gray_img is not None:
            auto_img = ImageOps.autocontrast(gray_img)
        else:
            # Fallback: Autocontrast direkt auf Original (Grayscale-Konvertierung inline)
            auto_img = ImageOps.autocontrast(ImageOps.grayscale(img))
        text_auto = ocr_runner(auto_img, language, "autocontrast")
        passes.append(OCRPassResult(pass_name="autocontrast", text=text_auto))
    except Exception as e:
        passes.append(OCRPassResult(pass_name="autocontrast", text="", error=str(e)))
        errors.append(f"Pass 'autocontrast' fehlgeschlagen: {e}")

    # ── Texte aus erfolgreichen Pässen sammeln und analysieren ────────────
    successful_texts = [p.text for p in passes if p.error is None and p.text]

    if not successful_texts:
        return NamingOCRResult(
            evidences=[], passes=passes, warnings=[],
            errors=errors if errors else ["Alle OCR-Pässe haben leere Ergebnisse geliefert"],
        )

    ocr_result = analyze_ocr_texts(successful_texts, source="tesseract")

    # Passes und Fehler zusammenführen
    return NamingOCRResult(
        evidences=ocr_result.evidences,
        passes=passes,
        warnings=ocr_result.warnings,
        errors=errors + ocr_result.errors,
    )
