"""
tests/test_naming_resolver.py

Unit-Tests für den Naming-Resolver in modules/naming_resolver.py (Branch 2).
Umfasst alle Pflicht-Tests, Konfliktlogik mit starker Quellenunterstützung,
Normalisierung, Reihenfolgeunabhängigkeit, Eingabevalidierung und Determinismus.
"""

import copy
import itertools
import unittest

from modules.naming_engine import NamingEvidence, FilenameParseResult
from modules.naming_resolver import (
    resolve_naming,
    ResolvedField,
    ResolvedResult,
    STATUS_AUTO_MERGED,
    STATUS_NEEDS_REVIEW,
    STATUS_CONFLICT,
    QUALITY_HIGH,
    QUALITY_MEDIUM,
    QUALITY_LOW,
)


class TestNamingResolver(unittest.TestCase):

    # ── 1-11: Grundlegende Auflösungs- und Statusfälle ──────────────────

    def test_eindeutiges_high_evidence(self):
        """1. Ein eindeutiges HIGH-Evidence für title und date -> AUTO_MERGED."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.status, STATUS_AUTO_MERGED)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.title.quality, QUALITY_HIGH)
        self.assertEqual(res.date.value, "18_10")
        self.assertEqual(res.date.quality, QUALITY_HIGH)
        self.assertFalse(res.title.conflict)
        self.assertFalse(res.date.conflict)
        self.assertEqual(res.missing_fields, [])

    def test_gleiche_werte_aus_mehreren_quellen(self):
        """2. Gleiche Werte aus mehreren Quellen (HIGH + MEDIUM) -> HIGH, kein Konflikt."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "SHREK", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.title.quality, QUALITY_HIGH)
        self.assertEqual(set(res.title.sources), {"filename_parser", "tesseract_header"})
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    def test_zwei_unabhaengige_medium_quellen(self):
        """3. Zwei unabhängige MEDIUM-Quellen -> zu HIGH verstärkt."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("title", "shrek", "easyocr_header", QUALITY_MEDIUM, 45),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.quality, QUALITY_HIGH)
        self.assertEqual(set(res.title.sources), {"tesseract_header", "easyocr_header"})
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    def test_doppelte_evidences_derselben_quelle(self):
        """4. Doppelte Evidences derselben Quelle -> keine Verstärkung zu HIGH."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("title", "SHREK", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.quality, QUALITY_MEDIUM)
        self.assertEqual(res.title.sources, ["tesseract_header"])
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_high_gegen_low(self):
        """5. HIGH gegen LOW -> HIGH-Wert gewählt, LOW in alternatives, kein Konflikt."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "Shrek2", "ocr_noise", QUALITY_LOW, 10),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.title.quality, QUALITY_HIGH)
        self.assertEqual(res.title.alternatives, ["Shrek2"])
        self.assertFalse(res.title.conflict)
        self.assertIn("Shrek2", res.title.reason)
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    def test_high_gegen_abweichendes_medium(self):
        """6. HIGH gegen abweichendes MEDIUM -> HIGH-Wert gewählt, NEEDS_REVIEW."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.title.alternatives, ["Titanic"])
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_high_gegen_high(self):
        """7. HIGH gegen HIGH aus verschiedenen Quellen -> CONFLICT."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertTrue(res.title.conflict)
        self.assertEqual(res.status, STATUS_CONFLICT)

    def test_zwei_unterschiedliche_medium_werte(self):
        """8. Zwei unterschiedliche MEDIUM-Werte -> Shrek gewinnt deterministisch (Score 50 > 45)."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("title", "Titanic", "easyocr_header", QUALITY_MEDIUM, 45),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.title.quality, QUALITY_MEDIUM)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.title.alternatives, ["Titanic"])
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_titel_fehlt(self):
        """9. Titel fehlt -> missing_fields=['title'], NEEDS_REVIEW."""
        evidences = [
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertIsNone(res.title.value)
        self.assertEqual(res.date.value, "18_10")
        self.assertIn("title", res.missing_fields)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_datum_fehlt(self):
        """10. Datum fehlt -> missing_fields=['date'], NEEDS_REVIEW."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.value, "Shrek")
        self.assertIsNone(res.date.value)
        self.assertIn("date", res.missing_fields)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_kategorie_fehlt(self):
        """11. Kategorie fehlt (optionales Feld) -> AUTO_MERGED wenn title & date HIGH."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertIsNone(res.category.value)
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    # ── 12-14: FilenameParseResult-Integration ───────────────────────────

    def test_generischer_dateiname_mit_ocr_titel(self):
        """12. Generischer Dateiname + OCR-Titel (HIGH) & Datum (HIGH) -> AUTO_MERGED."""
        parse_res = FilenameParseResult(
            evidences=[NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40)],
            is_generic=True,
            is_ambiguous=False,
            raw_stem="bild_001_18_10",
        )
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
        ]
        res = resolve_naming(evidences, parse_result=parse_res)
        self.assertTrue(res.filename_is_generic)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.date.value, "18_10")
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    def test_mehrdeutiger_dateiname_mit_eindeutiger_ocr(self):
        """13. Mehrdeutiger Dateiname + eindeutige OCR -> Werte vorhanden, aber NEEDS_REVIEW."""
        parse_res = FilenameParseResult(
            evidences=[],
            is_generic=False,
            is_ambiguous=True,
            raw_stem="2001_12_05",
        )
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("date", "18_10", "tesseract_header", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences, parse_result=parse_res)
        self.assertTrue(res.filename_is_ambiguous)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.date.value, "18_10")
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_mehrdeutiger_dateiname_ohne_weitere_evidences(self):
        """14. Mehrdeutiger Dateiname ohne Evidences -> NEEDS_REVIEW, fehlende Pflichtfelder."""
        parse_res = FilenameParseResult(evidences=[], is_ambiguous=True, raw_stem="film_31_04")
        res = resolve_naming([], parse_result=parse_res)
        self.assertTrue(res.filename_is_ambiguous)
        self.assertIn("title", res.missing_fields)
        self.assertIn("date", res.missing_fields)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    # ── 15-16: Normalisierung und Unicode-Erhalt ─────────────────────────

    def test_vergleichsnormalisierung_titel(self):
        """15. Blade Runner, blade_runner, BLADE-RUNNER -> gleicher Kandidat, zu HIGH verstärkt."""
        evidences = [
            NamingEvidence("title", "Blade Runner", "filename_parser", QUALITY_MEDIUM, 30),
            NamingEvidence("title", "blade_runner", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("title", "BLADE-RUNNER", "easyocr_header", QUALITY_MEDIUM, 45),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.quality, QUALITY_HIGH)
        self.assertEqual(set(res.title.sources), {"filename_parser", "tesseract_header", "easyocr_header"})
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    def test_unicode_erhalt_titel(self):
        """16. Amélie vs AMÉLIE -> gleicher Kandidat, Originalwert mit höherem Score bleibt."""
        evidences = [
            NamingEvidence("title", "Amélie", "filename_parser", QUALITY_MEDIUM, 30),
            NamingEvidence("title", "AMÉLIE", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertEqual(res.title.quality, QUALITY_HIGH)
        self.assertEqual(res.title.value, "AMÉLIE")
        self.assertEqual(res.status, STATUS_AUTO_MERGED)

    # ── 17: Reihenfolgeunabhängigkeit (vollständiger Dataclass-Vergleich) ─

    def test_reihenfolgeunabhaengigkeit(self):
        """17. Auflösung muss in jeder Permutation exakt dasselbe ResolvedResult ergeben."""
        evidences = [
            NamingEvidence("title", "Shrek", "filename_parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("title", "Shrek2", "ocr_noise", QUALITY_LOW, 10),
            NamingEvidence("category", "FK", "filename_category", QUALITY_HIGH, 35),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        baseline = resolve_naming(evidences)

        for perm in itertools.permutations(evidences):
            result = resolve_naming(list(perm))
            self.assertEqual(result, baseline)

    # ── 18-19: Eingabevalidierung und Unveränderbarkeit ──────────────────

    def test_ungueltige_evidences_werfen_valueerror(self):
        """18. Ungültige Evidences werden mit ValueError abgelehnt."""
        invalid_cases = [
            NamingEvidence("unknown_field", "Shrek", "parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "Shrek", "parser", "SUPER_HIGH", 30),
            NamingEvidence("title", "", "parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "   ", "parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "Shrek", "", QUALITY_HIGH, 30),
            NamingEvidence("title", "Shrek", "parser", QUALITY_HIGH, True),
            NamingEvidence("title", "Shrek", "parser", QUALITY_HIGH, 10.5),
            NamingEvidence("date", "31_02", "parser", QUALITY_HIGH, 40),
            NamingEvidence("date", "18-10", "parser", QUALITY_HIGH, 40),
            "Kein NamingEvidence Objekt",
        ]
        for idx, inv_ev in enumerate(invalid_cases):
            with self.subTest(case=idx):
                with self.assertRaises(ValueError):
                    resolve_naming([inv_ev] if isinstance(inv_ev, NamingEvidence) else [inv_ev])  # type: ignore

    def test_eingabeliste_bleibt_unveraendert(self):
        """19. Eingabeliste und enthaltene Objekte dürfen durch resolve_naming nicht mutiert werden."""
        ev1 = NamingEvidence("title", "Shrek", "parser", QUALITY_HIGH, 30)
        ev2 = NamingEvidence("date", "18_10", "parser", QUALITY_HIGH, 40)
        evidences = [ev1, ev2]
        evidences_copy = copy.deepcopy(evidences)

        _ = resolve_naming(evidences)
        self.assertEqual(evidences, evidences_copy)
        self.assertEqual(evidences[0].value, "Shrek")
        self.assertEqual(evidences[1].value, "18_10")

    # ── 20: Intra-Source-Widersprüche ────────────────────────────────────

    def test_widersprueche_innerhalb_derselben_quelle(self):
        """20. Widersprüche innerhalb derselben Quelle -> NEEDS_REVIEW, kein CONFLICT."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)
        self.assertTrue(any("tesseract_header" in r and "Widersprüchliche" in r for r in res.review_reasons))

    # ── 21-24: Konfliktlogik mit starker Quellenunterstützung ────────────

    def test_schwacher_externer_hinweis_kein_konflikt(self):
        """21. Gleiche Quelle gibt zwei HIGH-Werte, einer erhält zusätzlich LOW extern
        -> NEEDS_REVIEW, kein CONFLICT. LOW darf keinen Intra-Source-Widerspruch hochstufen."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "ocr_noise", QUALITY_LOW, 10),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    def test_starke_externe_bestaetigung_konflikt(self):
        """22. Gleiche Quelle gibt zwei HIGH-Werte, einer erhält zusätzlich unabhängiges HIGH
        -> echter CONFLICT, weil starke unabhängige Quellenunterstützung vorliegt."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "easyocr_header", QUALITY_HIGH, 45),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertTrue(res.title.conflict)
        self.assertEqual(res.status, STATUS_CONFLICT)

    def test_medium_promotion_externe_bestaetigung(self):
        """23. Gleiche Quelle gibt zwei HIGH-Werte, einer wird durch zwei unabhängige
        MEDIUM-Quellen bestätigt. Die MEDIUM-Promotion bewirkt starke externe
        Unterstützung -> CONFLICT, weil Titanic von mehreren unabhängigen Quellen
        stark getragen wird."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "med_src_1", QUALITY_MEDIUM, 40),
            NamingEvidence("title", "Titanic", "med_src_2", QUALITY_MEDIUM, 40),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertTrue(res.title.conflict)
        self.assertEqual(res.status, STATUS_CONFLICT)

    def test_low_allein_keine_starke_bestaetigung(self):
        """24. Ein externer LOW-Hinweis darf niemals allein einen Kandidaten zu einem
        starken extern bestätigten Kandidaten machen. Beide Kandidaten haben nur
        tesseract_header als starke Quelle."""
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Shrek", "noise_a", QUALITY_LOW, 5),
            NamingEvidence("title", "Titanic", "tesseract_header", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "noise_b", QUALITY_LOW, 5),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        # Beide Kandidaten haben starke Unterstützung nur von tesseract_header
        # LOW-Quellen noise_a und noise_b sind nicht stark
        self.assertFalse(res.title.conflict)
        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)

    # ── 25-27: Eingabevalidierung für Container und Argumente ────────────

    def test_ungueltige_required_fields_valueerror(self):
        """25. Ungültige required_fields werden mit ValueError abgelehnt."""
        valid_evidences = [
            NamingEvidence("title", "Shrek", "parser", QUALITY_HIGH, 30),
            NamingEvidence("date", "18_10", "parser", QUALITY_HIGH, 40),
        ]
        invalid_cases = [
            (("titel", "date"), "Unbekanntes Pflichtfeld"),
            (("", "date"), "Leeres Pflichtfeld"),
            (("title", "title"), "Doppeltes Pflichtfeld"),
            (("title", 123), "Kein String-Element"),
            ("title", "String statt Tuple"),
            (None, "None"),
        ]
        for rf_val, label in invalid_cases:
            with self.subTest(label=label):
                with self.assertRaises(ValueError):
                    resolve_naming(valid_evidences, required_fields=rf_val)  # type: ignore

    def test_ungueltiges_parse_result_valueerror(self):
        """26. Ungültiges parse_result (kein FilenameParseResult) wird mit ValueError abgelehnt."""
        with self.assertRaises(ValueError):
            resolve_naming([], parse_result={"is_generic": True})  # type: ignore

    def test_ungueltiger_evidences_container_valueerror(self):
        """27. Ungültiger Evidences-Container (kein Listentyp) wird mit ValueError abgelehnt."""
        invalid_containers = [
            (None, "None"),
            ("kein Evidence", "String"),
            ({"title": "Shrek"}, "Dict"),
        ]
        for container, label in invalid_containers:
            with self.subTest(label=label):
                with self.assertRaises(ValueError):
                    resolve_naming(container)  # type: ignore

    # ── 28: Determinismus mit parse_result ────────────────────────────────

    def test_determinismus_mit_parse_result(self):
        """28. Reihenfolgeunabhängigkeit mit parse_result: Vollständiger Dataclass-Vergleich."""
        parse_res = FilenameParseResult(
            evidences=[
                NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
                NamingEvidence("category", "FK", "filename_category", QUALITY_HIGH, 35),
            ],
            is_generic=False,
            is_ambiguous=False,
            raw_stem="fk_shrek_18_10",
        )
        evidences = [
            NamingEvidence("title", "Shrek", "tesseract_header", QUALITY_MEDIUM, 50),
            NamingEvidence("title", "shrek", "easyocr_header", QUALITY_MEDIUM, 45),
            NamingEvidence("title", "Shrek2", "ocr_noise", QUALITY_LOW, 10),
        ]
        baseline = resolve_naming(evidences, parse_result=parse_res)

        for perm in itertools.permutations(evidences):
            result = resolve_naming(list(perm), parse_result=parse_res)
            self.assertEqual(result, baseline)

    # ── 29-33: Erweiterte Randfälle & starke Kandidatensortierung ─────────

    def test_leere_normalisierte_werte_werfen_valueerror(self):
        """29. Werte wie '---', '___' oder '...' werden nach Normalisierung abgelehnt (ValueError)."""
        invalid_evs = [
            NamingEvidence("title", "---", "parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "___", "parser", QUALITY_HIGH, 30),
            NamingEvidence("title", "...", "parser", QUALITY_HIGH, 30),
            NamingEvidence("category", "---", "parser", QUALITY_HIGH, 30),
            NamingEvidence("category", "_._", "parser", QUALITY_HIGH, 30),
        ]
        for idx, inv_ev in enumerate(invalid_evs):
            with self.subTest(case=idx, value=inv_ev.value):
                with self.assertRaises(ValueError) as ctx:
                    resolve_naming([inv_ev])
                self.assertIn("ergibt nach Normalisierung keinen verwertbaren Wert", str(ctx.exception))

    def test_aeussere_leerzeichen_titel_kategorie_und_datum(self):
        """30. Äußere Leerzeichen: bei Titel/Kategorie im Ergebnis entfernt, bei Datum mit ValueError abgelehnt."""
        # Titel und Kategorie: Ergebnis ohne äußere Leerzeichen, Eingabeobjekte unberührt
        ev_title = NamingEvidence("title", " Shrek ", "tesseract_header", QUALITY_HIGH, 50)
        ev_cat = NamingEvidence("category", " FK ", "filename_category", QUALITY_HIGH, 35)
        ev_date = NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40)

        evidences = [ev_title, ev_cat, ev_date]
        ev_copy = copy.deepcopy(evidences)

        res = resolve_naming(evidences)
        self.assertEqual(res.title.value, "Shrek")
        self.assertEqual(res.category.value, "FK")
        self.assertEqual(evidences, ev_copy)  # Keine Mutation der Eingabe

        # Datum: äußere Leerzeichen sind verboten und lösen ValueError aus
        invalid_date_evs = [
            NamingEvidence("date", " 18_10 ", "parser", QUALITY_HIGH, 40),
            NamingEvidence("date", "18_10 ", "parser", QUALITY_HIGH, 40),
            NamingEvidence("date", " 18_10", "parser", QUALITY_HIGH, 40),
        ]
        for idx, inv_date in enumerate(invalid_date_evs):
            with self.subTest(case=idx, date_val=inv_date.value):
                with self.assertRaises(ValueError) as ctx:
                    resolve_naming([inv_date])
                self.assertIn("ungültiges Datumsformat", str(ctx.exception))

    def test_low_noise_beeinflusst_konfliktgewinner_nicht(self):
        """31. LOW-Evidences verändern nicht die starke Rangfolge oder den Gewinner im Konflikt."""
        # Baseline
        base_evs = [
            NamingEvidence("title", "Shrek", "source_a", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "source_b", QUALITY_HIGH, 40),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        base_res = resolve_naming(base_evs)
        self.assertTrue(base_res.title.conflict)
        self.assertEqual(base_res.title.value, "Shrek")

        # Mit massivem LOW-Noise für Titanic
        noise_evs = base_evs + [
            NamingEvidence("title", "Titanic", "noise_1", QUALITY_LOW, 1000),
            NamingEvidence("title", "Titanic", "noise_2", QUALITY_LOW, 1000),
            NamingEvidence("title", "Titanic", "noise_3", QUALITY_LOW, 1000),
        ]
        for perm in itertools.permutations(noise_evs):
            res = resolve_naming(list(perm))
            self.assertTrue(res.title.conflict)
            self.assertEqual(res.title.value, "Shrek")  # Shrek bleibt Gewinner trotz 3000 LOW-Score
            self.assertIn("Titanic", res.title.alternatives)

    def test_starke_externe_unterstuetzung_bevorzugt(self):
        """32. Kandidat mit stärkerer echter Unterstützung (2 HIGH-Quellen vs 1 HIGH-Quelle) gewinnt."""
        evidences = [
            NamingEvidence("title", "Shrek", "source_a", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "source_b", QUALITY_HIGH, 40),
            NamingEvidence("title", "Titanic", "source_c", QUALITY_HIGH, 40),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        res = resolve_naming(evidences)
        self.assertTrue(res.title.conflict)
        self.assertEqual(res.title.value, "Titanic")  # 2 starke Quellen schlagen 1 starke Quelle

    def test_identische_starke_kandidaten_tiebreaker(self):
        """33. Bei identischer starker Unterstützung entscheidet deterministisch der normalisierte String."""
        evidences = [
            NamingEvidence("title", "Shrek", "source_a", QUALITY_HIGH, 50),
            NamingEvidence("title", "Titanic", "source_b", QUALITY_HIGH, 50),
            NamingEvidence("date", "18_10", "filename_parser", QUALITY_HIGH, 40),
        ]
        baseline = resolve_naming(evidences)
        self.assertEqual(baseline.title.value, "Shrek")  # 'shrek' < 'titanic' lexikographisch

        for perm in itertools.permutations(evidences):
            result = resolve_naming(list(perm))
            self.assertEqual(result, baseline)


if __name__ == "__main__":
    unittest.main()
