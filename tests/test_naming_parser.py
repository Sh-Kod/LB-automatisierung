"""
tests/test_naming_parser.py

Unit-Tests für den Dateinamen-Parser in modules/naming_engine.py (Branch 1).
Führt alle vom Nutzer geforderten Mindesttests, Metadaten-Vertragsprüfungen,
generischen Tests nach Datumsextraktion, explizite Jahresvalidierungen sowie sämtliche Edge Cases aus.
"""

import unittest
from modules.naming_engine import parse_filename, NamingEvidence, FilenameParseResult


class TestNamingParser(unittest.TestCase):

    def _get_evidence_by_field(self, evidences: list[NamingEvidence], field: str) -> NamingEvidence | None:
        for e in evidences:
            if e.field == field:
                return e
        return None

    def _assert_evidence_metadata_valid(self, evidences: list[NamingEvidence]):
        """Prüft, ob alle Evidences saubere Metadaten, keine leeren Werte und auf Parser-Ebene
        nur die Qualitäten HIGH, MEDIUM oder LOW haben."""
        fields_seen = set()
        for e in evidences:
            self.assertIsNotNone(e.field)
            self.assertIsNotNone(e.value)
            self.assertNotEqual(e.value, "", f"Leerer Stringwert für Feld {e.field}")
            self.assertIsNotNone(e.source)
            self.assertIn(e.quality, {"HIGH", "MEDIUM", "LOW"}, f"Unerlaubte Qualität auf Parser-Ebene: {e.quality}")
            self.assertIsInstance(e.score, int)
            self.assertNotIn(e.field, fields_seen, f"Doppeltes Evidence-Feld: {e.field}")
            fields_seen.add(e.field)

    def test_vertrag_metadaten_konkret(self):
        """Testet gezielt den exakten öffentlichen Metadaten-Vertrag für zentrale Fälle."""
        # 1. shrek_18_10.jpg
        res = parse_filename("shrek_18_10.jpg")
        self._assert_evidence_metadata_valid(res.evidences)

        title_ev = self._get_evidence_by_field(res.evidences, "title")
        self.assertIsNotNone(title_ev)
        self.assertEqual(title_ev.field, "title")
        self.assertEqual(title_ev.value, "shrek")
        self.assertEqual(title_ev.source, "filename_parser")
        self.assertEqual(title_ev.quality, "HIGH")
        self.assertEqual(title_ev.score, 30)

        date_ev = self._get_evidence_by_field(res.evidences, "date")
        self.assertIsNotNone(date_ev)
        self.assertEqual(date_ev.field, "date")
        self.assertEqual(date_ev.value, "18_10")
        self.assertEqual(date_ev.source, "filename_parser")
        self.assertEqual(date_ev.quality, "HIGH")
        self.assertEqual(date_ev.score, 40)

        # 2. fk_shrek_18_10.jpg
        res_fk = parse_filename("fk_shrek_18_10.jpg")
        self._assert_evidence_metadata_valid(res_fk.evidences)

        category_ev = self._get_evidence_by_field(res_fk.evidences, "category")
        self.assertIsNotNone(category_ev)
        self.assertEqual(category_ev.field, "category")
        self.assertEqual(category_ev.value, "FK")
        self.assertEqual(category_ev.source, "filename_category")
        self.assertEqual(category_ev.quality, "HIGH")
        self.assertEqual(category_ev.score, 35)

    def test_mindesttests(self):
        """Prüft die ursprünglichen 13 Pflicht-Dateinamen."""
        cases = [
            ("zombie-_dawn_of_the_dead_11_10.jpg", None, "zombie_dawn_of_the_dead", "11_10"),
            ("mein_name_ist_nobody_23_08.png", None, "mein_name_ist_nobody", "23_08"),
            ("shrek_18_10.jpg", None, "shrek", "18_10"),
            ("apollo_13.jpg", None, "apollo_13", None),
            ("catch_22_15_03.jpg", None, "catch_22", "15_03"),
            ("2001_a_space_odyssey.jpg", None, "2001_a_space_odyssey", None),
            ("traumkino_inception_05.12.2026.jpg", "TK", "inception", "05_12"),
            ("zurueck-im-kino_matrix_22-11.jpg", "ZiK", "matrix", "22_11"),
            ("mek_frozen_08_06.jpg", "MeK", "frozen", "08_06"),
            ("03_01_shrek.jpg", None, "shrek", "03_01"),
            ("werbung_sommer_2026.jpg", None, "werbung_sommer_2026", None),
            ("fk_shrek_18_10.jpg", "FK", "shrek", "18_10"),
        ]
        for filename, exp_cat, exp_title, exp_date in cases:
            with self.subTest(filename=filename):
                res = parse_filename(filename)
                self.assertIsInstance(res, FilenameParseResult)
                self._assert_evidence_metadata_valid(res.evidences)

                cat_ev = self._get_evidence_by_field(res.evidences, "category")
                if exp_cat is None: self.assertIsNone(cat_ev)
                else: self.assertEqual(cat_ev.value, exp_cat)

                title_ev = self._get_evidence_by_field(res.evidences, "title")
                if exp_title is None: self.assertIsNone(title_ev)
                else: self.assertEqual(title_ev.value, exp_title)

                date_ev = self._get_evidence_by_field(res.evidences, "date")
                if exp_date is None: self.assertIsNone(date_ev)
                else: self.assertEqual(date_ev.value, exp_date)

    def test_generische_titel_mit_und_ohne_datum(self):
        """Prüft generische Dateinamen, auch nach Datumsextraktion und mit Kategorien."""
        cases = [
            # (Dateiname, Erw. is_generic, Erw. Kategorie, Erw. Datum)
            ("bild_001.jpg", True, None, None),
            ("image_123.png", True, None, None),
            ("img_4567.jpg", True, None, None),
            ("scan.jpg", True, None, None),
            ("foto.jpg", True, None, None),
            ("whatsapp_image_2026.jpeg", True, None, None),
            ("screenshot.png", True, None, None),
            ("untitled.jpg", True, None, None),
            ("bild_001_18_10.jpg", True, None, "18_10"),
            ("image_123_11_10.jpg", True, None, "11_10"),
            ("fk_bild_001_18_10.jpg", True, "FK", "18_10"),
            ("whatsapp_image_2026_15_07.jpg", True, None, "15_07"),
        ]
        for filename, exp_gen, exp_cat, exp_date in cases:
            with self.subTest(filename=filename):
                res = parse_filename(filename)
                self.assertEqual(res.is_generic, exp_gen, f"{filename} sollte is_generic={exp_gen} sein")
                self.assertFalse(res.is_ambiguous)
                # Bei generischen Namen darf niemals ein Titel-Evidence entstehen
                self.assertIsNone(self._get_evidence_by_field(res.evidences, "title"), f"{filename} hat unerwartetes Titel-Evidence")

                cat_ev = self._get_evidence_by_field(res.evidences, "category")
                if exp_cat is None: self.assertIsNone(cat_ev)
                else: self.assertEqual(cat_ev.value, exp_cat)

                date_ev = self._get_evidence_by_field(res.evidences, "date")
                if exp_date is None: self.assertIsNone(date_ev)
                else: self.assertEqual(date_ev.value, exp_date)

    def test_schaltjahr_und_explizite_jahrespruefung(self):
        """Prüft die Datumsvalidierung mit explizitem Jahr (Schaltjahr vs. Nicht-Schaltjahr)."""
        cases = [
            # (Dateiname, Erw. is_ambiguous, Erw. Titel, Erw. Datum)
            ("film_29_02_2025.jpg", True, None, None),   # 2025 kein Schaltjahr
            ("film_29_02_2028.jpg", False, "film", "29_02"), # 2028 ist Schaltjahr
            ("29_02_2025_film.jpg", True, None, None),   # 2025 kein Schaltjahr (vorne)
            ("29_02_2028_film.jpg", False, "film", "29_02"), # 2028 ist Schaltjahr (vorne)
            ("film_31_12_2025.jpg", False, "film", "31_12"),
            ("film_31_04_2025.jpg", True, None, None),   # April hat max. 30 Tage
            ("film_28_02_2025.jpg", False, "film", "28_02"),
        ]
        for filename, exp_ambig, exp_title, exp_date in cases:
            with self.subTest(filename=filename):
                res = parse_filename(filename)
                self.assertEqual(res.is_ambiguous, exp_ambig, f"Falsches Ambiguitätsflag für {filename}")
                self._assert_evidence_metadata_valid(res.evidences)

                title_ev = self._get_evidence_by_field(res.evidences, "title")
                if exp_title is None: self.assertIsNone(title_ev)
                else: self.assertEqual(title_ev.value, exp_title)

                date_ev = self._get_evidence_by_field(res.evidences, "date")
                if exp_date is None: self.assertIsNone(date_ev)
                else: self.assertEqual(date_ev.value, exp_date)

    def test_mehrdeutige_rein_numerische_muster(self):
        """Prüft, ob rein numerische Dateinamen wie 2001_12_05 saubere Ambiguitätsflags erzeugen."""
        res = parse_filename("2001_12_05.jpg")
        self.assertTrue(res.is_ambiguous)
        self.assertFalse(res.is_generic)
        self.assertEqual(len(res.evidences), 0)
        self.assertEqual(res.raw_stem, "2001_12_05")

    def test_zusaetzliche_edge_cases_und_pfade(self):
        r"""Prüft alle geforderten zusätzlichen Edge Cases und plattformunabhängige Pfade (\ und /)."""
        cases = [
            ("blade_runner_2049.jpg", None, "blade_runner_2049", None),
            ("1917.jpg", None, "1917", None),
            ("1984.jpg", None, "1984", None),
            ("1984_15_03.jpg", None, "1984", "15_03"),
            ("film_2026.jpg", None, "film_2026", None),
            ("film_11_10_2026.jpg", None, "film", "11_10"),
            ("29_02_film.jpg", None, "film", "29_02"),
            ("zurück_im_kino_film_01_12.jpg", "ZiK", "film", "01_12"),
            ("FK_SHREK_18_10.JPG", "FK", "SHREK", "18_10"),
            (r"C:\Import\FK_SHREK_18_10.JPG", "FK", "SHREK", "18_10"),
            ("C:/Import/FK_SHREK_18_10.JPG", "FK", "SHREK", "18_10"),
        ]
        for path, exp_cat, exp_title, exp_date in cases:
            with self.subTest(path=path):
                res = parse_filename(path)
                self._assert_evidence_metadata_valid(res.evidences)

                cat_ev = self._get_evidence_by_field(res.evidences, "category")
                if exp_cat is None: self.assertIsNone(cat_ev)
                else: self.assertEqual(cat_ev.value, exp_cat)

                title_ev = self._get_evidence_by_field(res.evidences, "title")
                if exp_title is None: self.assertIsNone(title_ev)
                else: self.assertEqual(title_ev.value, exp_title)

                date_ev = self._get_evidence_by_field(res.evidences, "date")
                if exp_date is None: self.assertIsNone(date_ev)
                else: self.assertEqual(date_ev.value, exp_date)

    def test_konservative_normalisierung_und_schreibweisen(self):
        """Prüft, dass offizielle/spezielle Schreibweisen und Umlaute nicht beschädigt werden."""
        cases = [
            ("WALL_E.jpg", "WALL_E"), ("Se7en.jpg", "Se7en"), ("xXx.jpg", "xXx"),
            ("Spider-Man.jpg", "Spider_Man"), ("DAS_BOOT.jpg", "DAS_BOOT"),
            ("M_Eine_Stadt_sucht_einen_Mörder.jpg", "M_Eine_Stadt_sucht_einen_Mörder"),
        ]
        for filename, exp_title in cases:
            with self.subTest(filename=filename):
                res = parse_filename(filename)
                title_ev = self._get_evidence_by_field(res.evidences, "title")
                self.assertIsNotNone(title_ev)
                self.assertEqual(title_ev.value, exp_title)

    def test_ungueltige_datumswerte_ohne_jahr(self):
        """Prüft, ob ungültige Kalendertage/Monate ohne explizites Jahr korrekt abgelehnt und als Ambiguität markiert werden."""
        invalid_cases = [
            "film_31_02.jpg", "film_31_04.jpg", "film_30_02.jpg",
            "film_00_10.jpg", "film_10_00.jpg", "film_32_12.jpg",
        ]
        for filename in invalid_cases:
            with self.subTest(filename=filename):
                res = parse_filename(filename)
                self.assertTrue(res.is_ambiguous, f"{filename} sollte als is_ambiguous markiert sein")
                self.assertIsNone(self._get_evidence_by_field(res.evidences, "date"))
                self.assertIsNone(self._get_evidence_by_field(res.evidences, "title"))


if __name__ == "__main__":
    unittest.main()
