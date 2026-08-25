"""
tests/test_naming_ocr.py

Unit-Tests für die OCR-Evidence-Schicht in modules/naming_ocr.py (Branch 3).
Umfasst Kategorie-, Datum-, Titel-Erkennung, Tesseract-Pässe,
Determinismus, Eingabevalidierung und Deduplizierung.

Alle Tests verwenden Dependency Injection (fake OCR-Runner).
Kein lokal installiertes Tesseract erforderlich.
"""

import copy
import itertools
import unittest
from unittest.mock import patch, MagicMock

from modules.naming_engine import NamingEvidence
from modules.naming_ocr import (
    analyze_ocr_texts,
    analyze_image_with_tesseract,
    NamingOCRResult,
    OCRPassResult,
)


class TestAnalyzeOcrTextsCategories(unittest.TestCase):
    """Kategorie-Erkennung aus OCR-Text."""

    def test_mein_erster_kinobesuch(self):
        """Vollständige Phrase 'MEIN ERSTER KINOBESUCH' -> MeK, HIGH, score=50."""
        res = analyze_ocr_texts(["MEIN ERSTER KINOBESUCH"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "MeK" and e.quality == "HIGH" and e.score == 50 for e in cats))

    def test_mek_als_eigenstaendiges_token(self):
        """Abkürzung 'MEK' als eigenständiges Token -> MeK, MEDIUM, score=35."""
        res = analyze_ocr_texts(["MEK"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "MeK" and e.quality == "MEDIUM" and e.score == 35 for e in cats))

    def test_zurueck_im_kino(self):
        """'ZURÜCK IM KINO' -> ZiK, HIGH."""
        res = analyze_ocr_texts(["ZURÜCK IM KINO"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "ZiK" and e.quality == "HIGH" for e in cats))

    def test_zuruck_im_kino_ocr_variante(self):
        """OCR-Variante 'ZURUCK IM KINO' -> ZiK, MEDIUM, score=40."""
        res = analyze_ocr_texts(["ZURUCK IM KINO"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "ZiK" and e.quality == "MEDIUM" and e.score == 40 for e in cats))

    def test_zurck_im_kino_ocr_variante(self):
        """OCR-Variante 'ZURCK IM KINO' -> ZiK, MEDIUM, score=40."""
        res = analyze_ocr_texts(["ZURCK IM KINO"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "ZiK" and e.quality == "MEDIUM" and e.score == 40 for e in cats))

    def test_zik_abkuerzung(self):
        """'ZIK' als eigenständiges Token -> ZiK, MEDIUM, score=35."""
        res = analyze_ocr_texts(["ZIK"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "ZiK" and e.quality == "MEDIUM" and e.score == 35 for e in cats))

    def test_traumkino(self):
        """'TRAUMKINO' -> TK, HIGH."""
        res = analyze_ocr_texts(["TRAUMKINO"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "TK" and e.quality == "HIGH" for e in cats))

    def test_tk_abkuerzung(self):
        """'TK' als eigenständiges Token -> TK, MEDIUM, score=35."""
        res = analyze_ocr_texts(["TK"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "TK" and e.quality == "MEDIUM" and e.score == 35 for e in cats))

    def test_filmklassiker(self):
        """'FILMKLASSIKER' -> FK, HIGH."""
        res = analyze_ocr_texts(["FILMKLASSIKER"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "FK" and e.quality == "HIGH" for e in cats))

    def test_fk_abkuerzung(self):
        """'FK' als eigenständiges Token -> FK, MEDIUM, score=35."""
        res = analyze_ocr_texts(["FK"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "FK" and e.quality == "MEDIUM" and e.score == 35 for e in cats))

    def test_keine_teilwort_fehlklassifizierung(self):
        """FK innerhalb eines Wortes ('AFKA') darf nicht als Kategorie erkannt werden."""
        res = analyze_ocr_texts(["AFKA"])
        cats = [e for e in res.evidences if e.field == "category"]
        cat_vals = [e.value for e in cats]
        self.assertNotIn("FK", cat_vals)

    def test_zwei_kategorien_derselben_quelle(self):
        """Zwei unterschiedliche Kategorien aus derselben Quelle -> beide als Evidences."""
        res = analyze_ocr_texts(["FILMKLASSIKER\nTRAUMKINO"])
        cats = [e for e in res.evidences if e.field == "category"]
        cat_vals = {e.value for e in cats}
        self.assertIn("FK", cat_vals)
        self.assertIn("TK", cat_vals)
        # Alle müssen dieselbe logische Quelle haben
        for e in cats:
            self.assertEqual(e.source, "tesseract_category")

    def test_kinobesuch_als_medium(self):
        """'KINOBESUCH' (ohne 'MEIN ERSTER') -> MeK, MEDIUM, score=40."""
        res = analyze_ocr_texts(["KINOBESUCH"])
        cats = [e for e in res.evidences if e.field == "category"]
        self.assertTrue(any(e.value == "MeK" and e.quality == "MEDIUM" and e.score == 40 for e in cats))


class TestAnalyzeOcrTextsDates(unittest.TestCase):
    """Datumserkennung aus OCR-Text."""

    def test_datum_punkt(self):
        """'18.10' -> 18_10, HIGH, score=45."""
        res = analyze_ocr_texts(["18.10"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertTrue(any(e.value == "18_10" and e.quality == "HIGH" and e.score == 45 for e in dates))

    def test_datum_slash(self):
        """'18/10' -> 18_10."""
        res = analyze_ocr_texts(["18/10"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_datum_bindestrich(self):
        """'18-10' -> 18_10."""
        res = analyze_ocr_texts(["18-10"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_datum_unterstrich(self):
        """'18_10' -> 18_10."""
        res = analyze_ocr_texts(["18_10"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_datum_schaltjahr_29_02(self):
        """'29.02' -> 29_02 (gültig mit Schaltjahr-Basis 2028)."""
        res = analyze_ocr_texts(["29.02"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertTrue(any(e.value == "29_02" for e in dates))

    def test_ungueltig_31_04(self):
        """'31.04' -> ungültig, Warnung, keine Evidence."""
        res = analyze_ocr_texts(["31.04"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 0)
        self.assertTrue(any("31" in w and "04" in w for w in res.warnings))

    def test_ungueltig_31_02(self):
        """'31.02' -> ungültig."""
        res = analyze_ocr_texts(["31.02"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 0)
        self.assertTrue(len(res.warnings) > 0)

    def test_ungueltig_monat_13(self):
        """Monat 13 -> ungültig."""
        res = analyze_ocr_texts(["18.13"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 0)

    def test_ungueltig_tag_00(self):
        """Tag 0 -> ungültig."""
        res = analyze_ocr_texts(["00.10"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 0)

    def test_jahreszahl_nicht_als_datum(self):
        """Jahreszahlen wie '1917' oder '2001' dürfen nicht als Datum interpretiert werden."""
        res = analyze_ocr_texts(["1917"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 0)

    def test_mehrere_gueltige_datumswerte(self):
        """Zwei verschiedene gültige Datumsangaben erzeugen zwei Evidences."""
        res = analyze_ocr_texts(["18.10 und 25.12"])
        dates = [e for e in res.evidences if e.field == "date"]
        date_vals = {e.value for e in dates}
        self.assertIn("18_10", date_vals)
        self.assertIn("25_12", date_vals)

    def test_datum_mit_trailing_punkt(self):
        """'18.10.' (mit optionalem Trailing-Punkt) -> 18_10."""
        res = analyze_ocr_texts(["18.10."])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_datum_quelle_korrekt(self):
        """Alle Date-Evidences verwenden 'tesseract_date' als Quelle."""
        res = analyze_ocr_texts(["18.10"])
        dates = [e for e in res.evidences if e.field == "date"]
        for e in dates:
            self.assertEqual(e.source, "tesseract_date")


class TestAnalyzeOcrTextsTitles(unittest.TestCase):
    """Titel-Erkennung aus OCR-Text."""

    def test_shrek(self):
        """'Shrek' -> Titel-Evidence, MEDIUM, score=35."""
        res = analyze_ocr_texts(["Shrek"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any(e.value == "Shrek" and e.quality == "MEDIUM" and e.score == 35 for e in titles))

    def test_blade_runner(self):
        """'Blade Runner' -> Titel-Evidence."""
        res = analyze_ocr_texts(["Blade Runner"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any("Blade Runner" in e.value for e in titles))

    def test_wall_e(self):
        """'WALL-E' bleibt als Titel erhalten."""
        res = analyze_ocr_texts(["WALL-E"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any("WALL" in e.value for e in titles))

    def test_se7en(self):
        """'Se7en' bleibt als Titel erhalten."""
        res = analyze_ocr_texts(["Se7en"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any("Se7en" in e.value for e in titles))

    def test_1917(self):
        """'1917' als eigenständiger Titel (kein Datum)."""
        res = analyze_ocr_texts(["1917"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any("1917" in e.value for e in titles))
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 0)

    def test_das_boot(self):
        """'Das Boot' -> Titel-Evidence."""
        res = analyze_ocr_texts(["Das Boot"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any("Das Boot" in e.value for e in titles))

    def test_titel_mit_umlauten(self):
        """Titel mit Umlauten bleiben erhalten."""
        res = analyze_ocr_texts(["Für Elise"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any("Für Elise" in e.value for e in titles))

    def test_kategoriezeile_nicht_als_titel(self):
        """Eine reine Kategoriezeile darf nicht als Titel erscheinen."""
        res = analyze_ocr_texts(["FILMKLASSIKER"])
        titles = [e for e in res.evidences if e.field == "title"]
        title_vals = [e.value.lower() for e in titles]
        self.assertFalse(any("filmklassiker" in v for v in title_vals))

    def test_datumszeile_nicht_als_titel(self):
        """Eine reine Datumszeile darf nicht als Titel erscheinen."""
        res = analyze_ocr_texts(["18.10"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertEqual(len(titles), 0)

    def test_reine_satzzeichen(self):
        """Reines Satzzeichen-Rauschen erzeugt keinen Titel."""
        res = analyze_ocr_texts(["---", "...", "***", "!!!"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertEqual(len(titles), 0)

    def test_leertext(self):
        """Leerer Text erzeugt keine Evidences."""
        res = analyze_ocr_texts(["", "   ", "\n\n"])
        self.assertEqual(len(res.evidences), 0)

    def test_ocr_rauschen(self):
        """Reines OCR-Rauschen (Sonderzeichen) erzeugt keinen Titel."""
        res = analyze_ocr_texts(["@#$%"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertEqual(len(titles), 0)

    def test_mehrere_plausible_titel(self):
        """Mehrere Titel auf verschiedenen Zeilen erzeugen mehrere Evidences."""
        res = analyze_ocr_texts(["Shrek\nBlade Runner"])
        titles = [e for e in res.evidences if e.field == "title"]
        title_vals = {e.value for e in titles}
        self.assertIn("Shrek", title_vals)
        self.assertIn("Blade Runner", title_vals)

    def test_gleiche_titel_verschiedene_schreibung(self):
        """Gleiche Titel in unterschiedlicher Schreibung werden dedupliziert."""
        res = analyze_ocr_texts(["SHREK", "Shrek", "shrek"])
        titles = [e for e in res.evidences if e.field == "title"]
        # Alle normalisieren zu "shrek", also nur ein deduplizierter Eintrag
        self.assertEqual(len(titles), 1)

    def test_aeussere_leerzeichen_entfernen(self):
        """Äußere Leerzeichen werden im Titel entfernt."""
        res = analyze_ocr_texts(["  Shrek  "])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any(e.value == "Shrek" for e in titles))

    def test_eingabetext_unveraendert(self):
        """Die Eingabetexte dürfen nicht verändert werden."""
        texts = ["FILMKLASSIKER\nShrek\n18.10"]
        texts_copy = copy.deepcopy(texts)
        _ = analyze_ocr_texts(texts)
        self.assertEqual(texts, texts_copy)

    def test_titel_quelle_korrekt(self):
        """Alle Title-Evidences verwenden 'tesseract_title' als Quelle."""
        res = analyze_ocr_texts(["Shrek"])
        titles = [e for e in res.evidences if e.field == "title"]
        for e in titles:
            self.assertEqual(e.source, "tesseract_title")

    def test_titel_qualitaet_hoechstens_medium(self):
        """Titel aus OCR dürfen höchstens MEDIUM sein."""
        res = analyze_ocr_texts(["Shrek\nShrek\nShrek"])
        titles = [e for e in res.evidences if e.field == "title"]
        for e in titles:
            self.assertIn(e.quality, ("MEDIUM", "LOW"))


class TestAnalyzeOcrTextsIntegration(unittest.TestCase):
    """Integrationstests: gemischte OCR-Texte mit Kategorie, Datum und Titel."""

    def test_vollstaendiges_kinoplakat(self):
        """Typisches Kinoplakat mit Kategorie, Titel und Datum."""
        res = analyze_ocr_texts(["FILMKLASSIKER\nShrek\n18.10."])
        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]
        dates = [e for e in res.evidences if e.field == "date"]

        self.assertTrue(any(e.value == "FK" for e in cats))
        self.assertTrue(any(e.value == "Shrek" for e in titles))
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_determinismus_textpermutationen(self):
        """Ergebnis ist unabhängig von der Reihenfolge der Texte."""
        texts = ["FILMKLASSIKER", "Shrek", "18.10"]
        baseline = analyze_ocr_texts(texts)

        for perm in itertools.permutations(texts):
            result = analyze_ocr_texts(list(perm))
            self.assertEqual(len(result.evidences), len(baseline.evidences))
            for e1, e2 in zip(result.evidences, baseline.evidences):
                self.assertEqual(e1.field, e2.field)
                self.assertEqual(e1.value, e2.value)
                self.assertEqual(e1.quality, e2.quality)
                self.assertEqual(e1.score, e2.score)
                self.assertEqual(e1.source, e2.source)

    def test_deduplizierung_nach_field_source_normvalue(self):
        """Doppelte Evidences (gleicher normalisierter Wert) werden dedupliziert."""
        res = analyze_ocr_texts(["18.10", "18/10", "18_10"])
        dates = [e for e in res.evidences if e.field == "date"]
        self.assertEqual(len(dates), 1)
        self.assertEqual(dates[0].value, "18_10")

    def test_evidence_sortierung(self):
        """Evidences werden in der Reihenfolge category → title → date sortiert."""
        res = analyze_ocr_texts(["FK\nShrek\n18.10"])
        fields = [e.field for e in res.evidences]
        # Jedes category-Evidence muss vor title kommen, title vor date
        cat_indices = [i for i, f in enumerate(fields) if f == "category"]
        title_indices = [i for i, f in enumerate(fields) if f == "title"]
        date_indices = [i for i, f in enumerate(fields) if f == "date"]
        if cat_indices and title_indices:
            self.assertLess(max(cat_indices), min(title_indices))
        if title_indices and date_indices:
            self.assertLess(max(title_indices), min(date_indices))

    def test_custom_source_name(self):
        """Benutzerdefinierter Quellenname wird korrekt weitergegeben."""
        res = analyze_ocr_texts(["Shrek"], source="custom_ocr")
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(all(e.source == "custom_ocr_title" for e in titles))


class TestAnalyzeOcrTextsValidation(unittest.TestCase):
    """Eingabevalidierung für analyze_ocr_texts."""

    def test_texts_kein_liste(self):
        """texts muss eine Liste sein."""
        with self.assertRaises(ValueError):
            analyze_ocr_texts("kein Liste")

    def test_texts_element_kein_string(self):
        """Jedes Element in texts muss ein String sein."""
        with self.assertRaises(ValueError):
            analyze_ocr_texts([123])

    def test_texts_none(self):
        """texts darf nicht None sein."""
        with self.assertRaises(ValueError):
            analyze_ocr_texts(None)  # type: ignore

    def test_source_leer(self):
        """source darf nicht leer sein."""
        with self.assertRaises(ValueError):
            analyze_ocr_texts([], source="")

    def test_source_kein_string(self):
        """source muss ein String sein."""
        with self.assertRaises(ValueError):
            analyze_ocr_texts([], source=123)  # type: ignore

    def test_leere_liste_gueltig(self):
        """Leere Liste ist gültig und liefert keine Evidences."""
        res = analyze_ocr_texts([])
        self.assertEqual(len(res.evidences), 0)
        self.assertEqual(len(res.warnings), 0)
        self.assertEqual(len(res.errors), 0)


class TestAnalyzeImageWithTesseract(unittest.TestCase):
    """Tests für die Bild- und Tesseract-Ausführungsschicht."""

    @patch('modules.naming_ocr.Image')
    @patch('modules.naming_ocr.ImageOps')
    def _run_with_mocks(self, mock_imageops_cls, mock_image_cls, ocr_runner=None):
        """Hilfsmethode: Erstellt gemockte PIL-Importe und führt analyze_image_with_tesseract aus."""
        # Das wird durch den Import-Patch abgefangen
        pass

    def test_alle_paesse_erfolgreich(self):
        """Alle drei Pässe liefern gültigen Text."""
        def fake_runner(img, lang, pass_name):
            texts = {
                "original": "FILMKLASSIKER\nShrek\n18.10",
                "grayscale": "FILMKLASSIKER\nShrek\n18.10",
                "autocontrast": "Shrek\n18.10",
            }
            return texts.get(pass_name, "")

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_gray = MagicMock()
            mock_imageops.grayscale.return_value = mock_gray
            mock_auto = MagicMock()
            mock_imageops.autocontrast.return_value = mock_auto

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        self.assertEqual(len(res.passes), 3)
        self.assertEqual(len(res.errors), 0)
        self.assertTrue(all(p.error is None for p in res.passes))

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]
        dates = [e for e in res.evidences if e.field == "date"]

        self.assertTrue(any(e.value == "FK" for e in cats))
        self.assertTrue(any(e.value == "Shrek" for e in titles))
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_ein_pass_schlaegt_fehl(self):
        """Ein Pass schlägt fehl, die anderen liefern Ergebnisse."""
        def fake_runner(img, lang, pass_name):
            if pass_name == "grayscale":
                raise RuntimeError("Tesseract segfault")
            return "Shrek\n18.10"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_gray = MagicMock()
            mock_imageops.grayscale.return_value = mock_gray
            mock_auto = MagicMock()
            mock_imageops.autocontrast.return_value = mock_auto

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        self.assertEqual(len(res.passes), 3)
        # Genau ein Fehler (grayscale)
        failed = [p for p in res.passes if p.error is not None]
        self.assertEqual(len(failed), 1)
        self.assertEqual(failed[0].pass_name, "grayscale")
        self.assertTrue(len(res.errors) >= 1)

        # Ergebnisse aus original und autocontrast weiterhin vorhanden
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(any(e.value == "Shrek" for e in titles))

    def test_alle_paesse_schlagen_fehl(self):
        """Alle Pässe schlagen fehl -> leere Evidences, Fehlerliste."""
        def failing_runner(img, lang, pass_name):
            raise RuntimeError(f"Pass {pass_name} failed")

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_gray = MagicMock()
            mock_imageops.grayscale.return_value = mock_gray
            mock_auto = MagicMock()
            mock_imageops.autocontrast.return_value = mock_auto

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=failing_runner)

        self.assertEqual(len(res.evidences), 0)
        self.assertEqual(len(res.passes), 3)
        self.assertTrue(all(p.error is not None for p in res.passes))
        self.assertTrue(len(res.errors) >= 3)

    def test_bild_kann_nicht_geoeffnet_werden(self):
        """Fehler beim Laden des Bildes -> strukturierter Fehler."""
        with patch('modules.naming_ocr.Image') as mock_image_mod:
            mock_image_mod.open.side_effect = FileNotFoundError("No such file")

            res = analyze_image_with_tesseract("nicht_vorhanden.jpg",
                                                ocr_runner=lambda i, l, p: "")

        self.assertEqual(len(res.evidences), 0)
        self.assertEqual(len(res.passes), 0)
        self.assertTrue(any("Fehler beim Laden des Bildes" in e for e in res.errors))

    def test_beschaedigte_bilddatei(self):
        """Beschädigtes Bild (load() schlägt fehl) -> strukturierter Fehler."""
        with patch('modules.naming_ocr.Image') as mock_image_mod:
            mock_img = MagicMock()
            mock_img.load.side_effect = OSError("Corrupted image data")
            mock_image_mod.open.return_value = mock_img

            res = analyze_image_with_tesseract("beschaedigt.jpg",
                                                ocr_runner=lambda i, l, p: "")

        self.assertEqual(len(res.evidences), 0)
        self.assertTrue(any("Fehler beim Laden" in e for e in res.errors))

    def test_paesse_liefern_denselben_titel(self):
        """Alle Pässe liefern denselben Titel -> dedupliziert zu einem Evidence."""
        def fake_runner(img, lang, pass_name):
            return "Shrek"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_imageops.grayscale.return_value = MagicMock()
            mock_imageops.autocontrast.return_value = MagicMock()

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        titles = [e for e in res.evidences if e.field == "title"]
        self.assertEqual(len(titles), 1)
        self.assertEqual(titles[0].value, "Shrek")

    def test_paesse_liefern_verschiedene_titel(self):
        """Pässe liefern unterschiedliche Titel -> mehrere Evidences mit derselben Quelle."""
        def fake_runner(img, lang, pass_name):
            if pass_name == "original":
                return "Shrek"
            elif pass_name == "grayscale":
                return "Titanic"
            return "Shrek"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_imageops.grayscale.return_value = MagicMock()
            mock_imageops.autocontrast.return_value = MagicMock()

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        titles = [e for e in res.evidences if e.field == "title"]
        title_vals = {e.value for e in titles}
        self.assertIn("Shrek", title_vals)
        self.assertIn("Titanic", title_vals)
        # Alle mit derselben logischen Quelle
        for e in titles:
            self.assertEqual(e.source, "tesseract_title")

    def test_paesse_nicht_als_unabhaengige_quellen(self):
        """Verschiedene Pässe dürfen NICHT als unabhängige Resolver-Quellen erscheinen."""
        def fake_runner(img, lang, pass_name):
            return "FILMKLASSIKER\nShrek\n18.10"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_imageops.grayscale.return_value = MagicMock()
            mock_imageops.autocontrast.return_value = MagicMock()

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        all_sources = {e.source for e in res.evidences}
        # Quellen dürfen nur tesseract_category, tesseract_title, tesseract_date sein
        for src in all_sources:
            self.assertTrue(src.startswith("tesseract_"), f"Unerwartete Quelle: {src}")
        # Keine pass-spezifischen Quellen
        self.assertNotIn("tesseract_original", all_sources)
        self.assertNotIn("tesseract_grayscale", all_sources)
        self.assertNotIn("tesseract_autocontrast", all_sources)


class TestAnalyzeImageValidation(unittest.TestCase):
    """Eingabevalidierung für analyze_image_with_tesseract."""

    def test_image_path_kein_string(self):
        """image_path muss ein String sein."""
        with self.assertRaises(ValueError):
            analyze_image_with_tesseract(123, ocr_runner=lambda i, l, p: "")  # type: ignore

    def test_image_path_leer(self):
        """image_path darf nicht leer sein."""
        with self.assertRaises(ValueError):
            analyze_image_with_tesseract("", ocr_runner=lambda i, l, p: "")

    def test_ocr_runner_nicht_aufrufbar(self):
        """ocr_runner muss callable oder None sein."""
        with self.assertRaises(ValueError):
            analyze_image_with_tesseract("test.jpg", ocr_runner="nicht_callable")  # type: ignore

    def test_language_leer(self):
        """language darf nicht leer sein."""
        with self.assertRaises(ValueError):
            analyze_image_with_tesseract("test.jpg", ocr_runner=lambda i, l, p: "",
                                          language="")


class TestDeterminismAndImmutability(unittest.TestCase):
    """Determinismus und Eingabe-Unveränderbarkeit."""

    def test_gleiche_ergebnisse_unabhaengig_von_pass_reihenfolge(self):
        """Auch wenn Texte aus verschiedenen Pässen in anderer Reihenfolge kommen,
        ist das Ergebnis identisch."""
        texts_a = ["FILMKLASSIKER\n18.10", "Shrek"]
        texts_b = ["Shrek", "FILMKLASSIKER\n18.10"]

        res_a = analyze_ocr_texts(texts_a)
        res_b = analyze_ocr_texts(texts_b)

        self.assertEqual(len(res_a.evidences), len(res_b.evidences))
        for e1, e2 in zip(res_a.evidences, res_b.evidences):
            self.assertEqual(e1.field, e2.field)
            self.assertEqual(e1.value, e2.value)
            self.assertEqual(e1.quality, e2.quality)
            self.assertEqual(e1.score, e2.score)
            self.assertEqual(e1.source, e2.source)

    def test_gleiche_evidence_reihenfolge(self):
        """Evidence-Reihenfolge ist bei identischem Input immer gleich."""
        texts = ["FK\nShrek\n18.10"]
        res1 = analyze_ocr_texts(texts)
        res2 = analyze_ocr_texts(texts)
        for e1, e2 in zip(res1.evidences, res2.evidences):
            self.assertEqual(e1, e2)

    def test_gleiche_warnungsreihenfolge(self):
        """Warnungen sind sortiert und deterministisch."""
        texts = ["31.04\n31.02"]
        res1 = analyze_ocr_texts(texts)
        res2 = analyze_ocr_texts(texts)
        self.assertEqual(res1.warnings, res2.warnings)

    def test_keine_mutation_der_eingabelisten(self):
        """Die Eingabeliste und ihre Strings dürfen nicht mutiert werden."""
        texts = ["FILMKLASSIKER\nShrek\n18.10"]
        texts_copy = copy.deepcopy(texts)
        _ = analyze_ocr_texts(texts)
        self.assertEqual(texts, texts_copy)

    def test_keine_mutation_bildobjekte(self):
        """Das vom Aufrufer übergebene PIL-Image darf nicht verändert werden."""
        original_img = MagicMock()
        call_count = [0]

        def counting_runner(img, lang, pass_name):
            call_count[0] += 1
            return "Shrek"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_image_mod.open.return_value = original_img
            mock_imageops.grayscale.return_value = MagicMock()
            mock_imageops.autocontrast.return_value = MagicMock()

            _ = analyze_image_with_tesseract("test.jpg", ocr_runner=counting_runner)

        # Image.open wurde aufgerufen, aber das Original-Mock-Objekt sollte nicht
        # durch Bildoperationen direkt verändert worden sein (grayscale/autocontrast
        # erzeugen neue Objekte via ImageOps)
        mock_imageops.grayscale.assert_called_once_with(original_img)
        # autocontrast wird auf das Grayscale-Bild angewendet, nicht auf das Original
        gray_result = mock_imageops.grayscale.return_value
        mock_imageops.autocontrast.assert_called_once_with(gray_result)


class TestAnalyzeImageWithTesseractPasses(unittest.TestCase):
    """Spezifische Tests für die Pass-Struktur."""

    def test_pass_reihenfolge_fest(self):
        """Die Pässe erscheinen immer in der Reihenfolge original, grayscale, autocontrast."""
        def fake_runner(img, lang, pass_name):
            return f"Text from {pass_name}"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_imageops.grayscale.return_value = MagicMock()
            mock_imageops.autocontrast.return_value = MagicMock()

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        self.assertEqual(len(res.passes), 3)
        self.assertEqual(res.passes[0].pass_name, "original")
        self.assertEqual(res.passes[1].pass_name, "grayscale")
        self.assertEqual(res.passes[2].pass_name, "autocontrast")

    def test_pass_texte_vollstaendig(self):
        """OCRPassResult enthält den vollständigen erkannten Text."""
        def fake_runner(img, lang, pass_name):
            return f"OCR text for {pass_name}"

        with patch('modules.naming_ocr.Image') as mock_image_mod, \
             patch('modules.naming_ocr.ImageOps') as mock_imageops:
            mock_img = MagicMock()
            mock_image_mod.open.return_value = mock_img
            mock_imageops.grayscale.return_value = MagicMock()
            mock_imageops.autocontrast.return_value = MagicMock()

            res = analyze_image_with_tesseract("test.jpg", ocr_runner=fake_runner)

        for p in res.passes:
            self.assertIn("OCR text for", p.text)
            self.assertIsNone(p.error)


class TestMetaLineFiltering(unittest.TestCase):
    """Kinoplakat-Metazeilen dürfen nicht als Titel-Evidences erscheinen."""

    def _assert_no_title(self, line: str):
        """Hilfsmethode: Prüft, dass eine Zeile keinen Titel erzeugt."""
        res = analyze_ocr_texts([line])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertEqual(len(titles), 0,
                         f'"{line}" sollte keinen Titel erzeugen,'
                         f' erzeugt aber: {[e.value for e in titles]}')

    # ── Uhrzeiten ────────────────────────────────────────────────────────
    def test_uhrzeit_ganzzahl(self):
        """'20 UHR' ist kein Titel."""
        self._assert_no_title("20 UHR")

    def test_uhrzeit_mit_minuten_doppelpunkt(self):
        """'19:30 UHR' ist kein Titel."""
        self._assert_no_title("19:30 UHR")

    def test_uhrzeit_mit_minuten_punkt(self):
        """'20.00 UHR' ist kein Titel."""
        self._assert_no_title("20.00 UHR")

    def test_uhrzeit_20_00(self):
        """'20:00 UHR' ist kein Titel."""
        self._assert_no_title("20:00 UHR")

    # ── Altersfreigaben ──────────────────────────────────────────────────
    def test_fsk_6(self):
        """'FSK 6' ist kein Titel."""
        self._assert_no_title("FSK 6")

    def test_fsk_0(self):
        """'FSK 0' ist kein Titel."""
        self._assert_no_title("FSK 0")

    def test_ab_6_jahren(self):
        """'AB 6 JAHREN' ist kein Titel."""
        self._assert_no_title("AB 6 JAHREN")

    def test_ab_12_jahren(self):
        """'AB 12 JAHREN' ist kein Titel."""
        self._assert_no_title("AB 12 JAHREN")

    def test_ab_16_jahren(self):
        """'AB 16 JAHREN' ist kein Titel."""
        self._assert_no_title("AB 16 JAHREN")

    def test_ab_0_jahren(self):
        """'AB 0 JAHREN' ist kein Titel."""
        self._assert_no_title("AB 0 JAHREN")

    # ── Promo-Phrasen ────────────────────────────────────────────────────
    def test_jetzt_im_kino(self):
        """'JETZT IM KINO' ist kein Titel."""
        self._assert_no_title("JETZT IM KINO")

    def test_nur_im_kino(self):
        """'NUR IM KINO' ist kein Titel."""
        self._assert_no_title("NUR IM KINO")

    def test_premiere(self):
        """'PREMIERE' ist kein Titel."""
        self._assert_no_title("PREMIERE")

    def test_demnachst(self):
        """'DEMNÄCHST' ist kein Titel."""
        self._assert_no_title("DEMNÄCHST")

    def test_heute(self):
        """'HEUTE' ist kein Titel."""
        self._assert_no_title("HEUTE")

    def test_morgen(self):
        """'MORGEN' ist kein Titel."""
        self._assert_no_title("MORGEN")

    def test_ab_donnerstag(self):
        """'AB DONNERSTAG' ist kein Titel."""
        self._assert_no_title("AB DONNERSTAG")

    def test_vorverkauf(self):
        """'VORVERKAUF' ist kein Titel."""
        self._assert_no_title("VORVERKAUF")

    def test_tickets_online(self):
        """'TICKETS ONLINE' ist kein Titel."""
        self._assert_no_title("TICKETS ONLINE")

    def test_karten_online(self):
        """'KARTEN ONLINE' ist kein Titel."""
        self._assert_no_title("KARTEN ONLINE")

    def test_special_screening(self):
        """'SPECIAL SCREENING' ist kein Titel."""
        self._assert_no_title("SPECIAL SCREENING")

    def test_original_version(self):
        """'ORIGINAL VERSION' ist kein Titel."""
        self._assert_no_title("ORIGINAL VERSION")

    def test_nur_heute(self):
        """'NUR HEUTE' ist kein Titel."""
        self._assert_no_title("NUR HEUTE")

    def test_nur_morgen(self):
        """'NUR MORGEN' ist kein Titel."""
        self._assert_no_title("NUR MORGEN")

    # ── Technik-Labels ───────────────────────────────────────────────────
    def test_omu(self):
        """'OMU' ist kein Titel."""
        self._assert_no_title("OMU")

    def test_ov(self):
        """'OV' ist kein Titel."""
        self._assert_no_title("OV")

    def test_2d(self):
        """'2D' ist kein Titel."""
        self._assert_no_title("2D")

    def test_3d(self):
        """'3D' ist kein Titel."""
        self._assert_no_title("3D")

    def test_praesentiert(self):
        """'PRÄSENTIERT' ist kein Titel."""
        self._assert_no_title("PRÄSENTIERT")

    # ── URLs ─────────────────────────────────────────────────────────────
    def test_url_www(self):
        """'WWW.KINO.DE' ist kein Titel."""
        self._assert_no_title("WWW.KINO.DE")

    def test_url_domain(self):
        """'KINO.DE' ist kein Titel."""
        self._assert_no_title("KINO.DE")

    def test_url_subdomain(self):
        """'WWW.LICHTBLICK-KINO.DE' ist kein Titel."""
        self._assert_no_title("WWW.LICHTBLICK-KINO.DE")

    # ── Crew/Cast-Zeilen ─────────────────────────────────────────────────
    def test_ein_film_von(self):
        """'EIN FILM VON STEVEN SPIELBERG' ist kein Titel."""
        self._assert_no_title("EIN FILM VON STEVEN SPIELBERG")

    def test_regie(self):
        """'REGIE SPIELBERG' ist kein Titel."""
        self._assert_no_title("REGIE SPIELBERG")

    def test_darsteller(self):
        """'DARSTELLER XY' ist kein Titel."""
        self._assert_no_title("DARSTELLER XY")

    def test_mit_vorname_nachname(self):
        """'Mit Michelle Williams' ist kein Titel."""
        self._assert_no_title("Mit Michelle Williams")

    def test_mit_nachname_allcaps(self):
        """'MIT WILLIAMS' (OCR-Großbuchstaben) ist kein Titel."""
        self._assert_no_title("MIT WILLIAMS")

    def test_mit_michelle_williams_allcaps(self):
        """'MIT MICHELLE WILLIAMS' (volle OCR-Caps) ist kein Titel."""
        self._assert_no_title("MIT MICHELLE WILLIAMS")

    def test_verleih_mit_name(self):
        """'VERLEIH X' ist kein Titel."""
        self._assert_no_title("VERLEIH X")


class TestTitlePreservation(unittest.TestCase):
    """Echte Filmtitel dürfen durch die Metazeilen-Filter NICHT zerstört werden."""

    def _assert_has_title(self, line: str, expected_substring: str = None):
        """Hilfsmethode: Prüft, dass eine Zeile einen Titel erzeugt."""
        res = analyze_ocr_texts([line])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(len(titles) > 0, f'"{line}" sollte einen Titel erzeugen')
        if expected_substring:
            title_vals = [e.value for e in titles]
            self.assertTrue(
                any(expected_substring in v for v in title_vals),
                f'Titel "{expected_substring}" nicht in {title_vals}'
            )

    def test_1917(self):
        """Numerischer Titel '1917' bleibt erhalten."""
        self._assert_has_title("1917", "1917")

    def test_se7en(self):
        """Alphanumerischer Titel 'Se7en' bleibt erhalten."""
        self._assert_has_title("Se7en", "Se7en")

    def test_wall_e_bindestrich(self):
        """'WALL-E' bleibt erhalten."""
        self._assert_has_title("WALL-E")

    def test_wall_e_unterstrich(self):
        """'WALL_E' bleibt erhalten."""
        self._assert_has_title("WALL_E")

    def test_das_boot(self):
        """'Das Boot' bleibt erhalten."""
        self._assert_has_title("Das Boot", "Das Boot")

    def test_amelie(self):
        """'Amélie' mit Akzent bleibt erhalten."""
        self._assert_has_title("Amélie")

    def test_fuer_elise(self):
        """'Für Elise' mit Umlaut bleibt erhalten."""
        self._assert_has_title("Für Elise")

    def test_m_einzelbuchstabe(self):
        """Einbuchstabiger Titel 'M' bleibt erhalten."""
        self._assert_has_title("M", "M")

    def test_it(self):
        """'IT' bleibt erhalten (nicht als Technik-Label verwechselt)."""
        self._assert_has_title("IT", "IT")

    def test_her(self):
        """'Her' bleibt erhalten."""
        self._assert_has_title("Her", "Her")

    def test_us(self):
        """'Us' bleibt erhalten."""
        self._assert_has_title("Us", "Us")

    def test_up(self):
        """'Up' bleibt erhalten."""
        self._assert_has_title("Up", "Up")

    def test_no(self):
        """'No' bleibt erhalten."""
        self._assert_has_title("No", "No")

    def test_3_als_titel(self):
        """Einstelliger numerischer Titel '3' bleibt erhalten."""
        self._assert_has_title("3", "3")

    def test_mit_allein_kein_filter(self):
        """'MIT' allein könnte ein Titel sein und wird NICHT gefiltert."""
        res = analyze_ocr_texts(["MIT"])
        titles = [e for e in res.evidences if e.field == "title"]
        self.assertTrue(len(titles) > 0, "'MIT' allein sollte nicht gefiltert werden")


class TestRealisticOCRBlocks(unittest.TestCase):
    """Realistische OCR-Textblöcke von Kinoplakaten: Integrationstests."""

    def test_kritischer_testfall(self):
        """FILMKLASSIKER + SHREK + Meta-Zeilen + Datum -> nur FK, Shrek, 18_10."""
        text = "FILMKLASSIKER\nSHREK\nAB 6 JAHREN\n20 UHR\nJETZT IM KINO\n18.10"
        res = analyze_ocr_texts([text])

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]
        dates = [e for e in res.evidences if e.field == "date"]

        self.assertTrue(any(e.value == "FK" for e in cats))
        self.assertEqual(len(titles), 1)
        self.assertEqual(titles[0].value, "SHREK")
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_beispiel_a_traumkino_das_boot(self):
        """TRAUMKINO + DAS BOOT + 20:00 UHR + NUR IM KINO + 21.11."""
        text = "TRAUMKINO\nDAS BOOT\n20:00 UHR\nNUR IM KINO\n21.11"
        res = analyze_ocr_texts([text])

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]
        dates = [e for e in res.evidences if e.field == "date"]

        self.assertTrue(any(e.value == "TK" for e in cats))
        self.assertEqual(len(titles), 1)
        self.assertIn("DAS BOOT", titles[0].value)
        self.assertTrue(any(e.value == "21_11" for e in dates))

    def test_beispiel_b_wall_e_fsk_url(self):
        """ZURÜCK IM KINO + WALL-E + FSK 0 + 18.10. + URL."""
        text = "ZURÜCK IM KINO\nWALL-E\nFSK 0\n18.10.\nWWW.LICHTBLICK-KINO.DE"
        res = analyze_ocr_texts([text])

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]
        dates = [e for e in res.evidences if e.field == "date"]

        self.assertTrue(any(e.value == "ZiK" for e in cats))
        self.assertEqual(len(titles), 1)
        self.assertIn("WALL-E", titles[0].value)
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_beispiel_c_se7en_filter_schutz(self):
        """MEIN ERSTER KINOBESUCH + SE7EN + AB 16 JAHREN + 19:30 UHR."""
        text = "MEIN ERSTER KINOBESUCH\nSE7EN\nAB 16 JAHREN\n19:30 UHR"
        res = analyze_ocr_texts([text])

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]

        self.assertTrue(any(e.value == "MeK" for e in cats))
        self.assertEqual(len(titles), 1)
        self.assertIn("SE7EN", titles[0].value)

    def test_beispiel_d_1917_als_titel(self):
        """FILMKLASSIKER + 1917 + 20 UHR -> 1917 bleibt Titel."""
        text = "FILMKLASSIKER\n1917\n20 UHR"
        res = analyze_ocr_texts([text])

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]

        self.assertTrue(any(e.value == "FK" for e in cats))
        self.assertEqual(len(titles), 1)
        self.assertIn("1917", titles[0].value)

    def test_beispiel_e_crew_zeilen_gefiltert(self):
        """FILMKLASSIKER + DIE FABELMANS + Crew-Zeilen + 18.10."""
        text = "FILMKLASSIKER\nDIE FABELMANS\nEIN FILM VON STEVEN SPIELBERG\nMIT MICHELLE WILLIAMS\n18.10"
        res = analyze_ocr_texts([text])

        cats = [e for e in res.evidences if e.field == "category"]
        titles = [e for e in res.evidences if e.field == "title"]
        dates = [e for e in res.evidences if e.field == "date"]

        self.assertTrue(any(e.value == "FK" for e in cats))
        self.assertEqual(len(titles), 1)
        self.assertEqual(titles[0].value, "DIE FABELMANS")
        self.assertTrue(any(e.value == "18_10" for e in dates))

    def test_realistischer_block_determinismus(self):
        """Permutationen desselben Textblocks liefern identische Evidences."""
        lines = [
            "FILMKLASSIKER",
            "SHREK",
            "AB 6 JAHREN",
            "20 UHR",
            "JETZT IM KINO",
            "18.10",
        ]
        baseline = analyze_ocr_texts(["\n".join(lines)])
        split_res = analyze_ocr_texts([
            "\n".join(lines[:2]),
            "\n".join(lines[2:4]),
            "\n".join(lines[4:]),
        ])

        self.assertEqual(len(baseline.evidences), len(split_res.evidences))
        for e1, e2 in zip(baseline.evidences, split_res.evidences):
            self.assertEqual(e1.field, e2.field)
            self.assertEqual(e1.value, e2.value)
            self.assertEqual(e1.quality, e2.quality)
            self.assertEqual(e1.score, e2.score)

    def test_eingabetext_unveraendert_realistisch(self):
        """Realistische Eingabetexte bleiben nach der Analyse unverändert."""
        texts = ["FILMKLASSIKER\nSHREK\nAB 6 JAHREN\n20 UHR\n18.10"]
        texts_copy = copy.deepcopy(texts)
        _ = analyze_ocr_texts(texts)
        self.assertEqual(texts, texts_copy)


if __name__ == "__main__":
    unittest.main()
