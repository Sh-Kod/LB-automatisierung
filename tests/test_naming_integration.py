"""
tests/test_naming_integration.py

Umfassende Testsuite für die Integrationsschicht (modules/naming_integration.py).
Prüft DCP-Namensbildung, Resolver-Integration, Dialog-Formatierung,
Tesseract-Adapter, Fehlertoleranz und alle Naming-Szenarien.
"""

import sys
import unittest
from unittest.mock import MagicMock, patch

from modules.naming_engine import NamingEvidence, FilenameParseResult
from modules.naming_resolver import (
    STATUS_AUTO_MERGED,
    STATUS_NEEDS_REVIEW,
    STATUS_CONFLICT,
    ResolvedResult,
    ResolvedField,
)
from modules.naming_integration import (
    clean_dcp_name,
    build_dcp_name,
    build_alternative_dcp_names,
    format_telegram_dialog,
    evaluate_dialog_response,
    create_tesseract_runner,
    process_image_naming,
    ACTION_SET_NAME,
    ACTION_ASK_CUSTOM,
    ACTION_ASK_DATE,
    ACTION_ASK_TITLE,
    ACTION_SKIP,
    CATEGORY_PREFIX_MAP,
)


class TestCleanDcpName(unittest.TestCase):
    """Tests für clean_dcp_name."""

    def test_umlaute_umwandlung(self):
        """Umlaute und ß werden korrekt ersetzt."""
        self.assertEqual(clean_dcp_name("äöüÄÖÜß"), "aeoeueAeOeUess")
        self.assertEqual(clean_dcp_name("Für_Elise"), "Fuer_Elise")
        self.assertEqual(clean_dcp_name("Amélie"), "Amelie")

    def test_sonderzeichen_entfernung(self):
        """Sonderzeichen und Leerzeichen werden zu Unterstrichen."""
        self.assertEqual(clean_dcp_name("Blade Runner 2049!"), "Blade_Runner_2049")
        self.assertEqual(clean_dcp_name("Film @ Cinema #1"), "Film_Cinema_1")

    def test_mehrfache_unterstriche_bereinigung(self):
        """Mehrfache Unterstriche und äußere Unterstriche werden bereinigt."""
        self.assertEqual(clean_dcp_name("___LB___FK___Shrek___"), "LB_FK_Shrek")

    def test_leere_eingabe(self):
        """Leere Eingaben liefern leeren String."""
        self.assertEqual(clean_dcp_name(""), "")
        self.assertEqual(clean_dcp_name(None), "")


class TestBuildDcpName(unittest.TestCase):
    """Tests für build_dcp_name."""

    def test_standard_kategorien(self):
        """Standard-Kategorien MeK, ZiK, TK, FK erzeugen korrekte LB_-Präfixe."""
        self.assertEqual(build_dcp_name("FK", "Shrek", "18_10"), "LB_FK_Shrek_18_10")
        self.assertEqual(build_dcp_name("TK", "Das Boot", "21_11"), "LB_TK_Das_Boot_21_11")
        self.assertEqual(build_dcp_name("ZiK", "WALL_E", "18_10"), "LB_ZiK_WALL_E_18_10")
        self.assertEqual(build_dcp_name("MeK", "Ponyo", "05_05"), "LB_MeK_Ponyo_05_05")

    def test_benutzerdefinierte_kategorie(self):
        """Unbekannte/eigene Kategorien erhalten LB_-Präfix."""
        self.assertEqual(build_dcp_name("Sonder", "Film", "12_12"), "LB_Sonder_Film_12_12")
        self.assertEqual(build_dcp_name("LB_Custom", "Film", "12_12"), "LB_Custom_Film_12_12")

    def test_ohne_kategorie(self):
        """Ohne Kategorie wird Standard-LB-Präfix verwendet."""
        self.assertEqual(build_dcp_name(None, "Shrek", "18_10"), "LB_Shrek_18_10")
        self.assertEqual(build_dcp_name("", "Shrek", "18_10"), "LB_Shrek_18_10")

    def test_ohne_datum(self):
        """Ohne Datum wird der Name ohne Datumsanhang gebildet."""
        self.assertEqual(build_dcp_name("FK", "Shrek", None), "LB_FK_Shrek")
        self.assertEqual(build_dcp_name("FK", "Shrek", ""), "LB_FK_Shrek")

    def test_ohne_titel_liefert_none(self):
        """Ohne Titel kann kein gültiger DCP-Name gebildet werden."""
        self.assertIsNone(build_dcp_name("FK", None, "18_10"))
        self.assertIsNone(build_dcp_name("FK", "", "18_10"))
        self.assertIsNone(build_dcp_name("FK", "   ", "18_10"))

    def test_umlaute_im_titel(self):
        """Umlaute im Titel werden für DCP bereinigt."""
        self.assertEqual(build_dcp_name("FK", "Für Elise", "18_10"), "LB_FK_Fuer_Elise_18_10")
        self.assertEqual(build_dcp_name("FK", "Amélie", "18_10"), "LB_FK_Amelie_18_10")


class TestEvaluateDialogResponse(unittest.TestCase):
    """Tests für evaluate_dialog_response."""

    def setUp(self):
        self.options = {
            "1": (ACTION_SET_NAME, "LB_FK_Shrek_18_10"),
            "2": (ACTION_ASK_CUSTOM, None),
            "3": (ACTION_SKIP, None),
        }

    def test_option_1_uebernehmen(self):
        """Option '1' übernimmt den vorgeschlagenen Namen."""
        aktion, wert = evaluate_dialog_response("1", self.options)
        self.assertEqual(aktion, ACTION_SET_NAME)
        self.assertEqual(wert, "LB_FK_Shrek_18_10")

    def test_option_2_custom(self):
        """Option '2' fordert manuelle Eingabe an."""
        aktion, wert = evaluate_dialog_response("2", self.options)
        self.assertEqual(aktion, ACTION_ASK_CUSTOM)
        self.assertIsNone(wert)

    def test_option_3_skip(self):
        """Option '3' überspringt das Bild."""
        aktion, wert = evaluate_dialog_response("3", self.options)
        self.assertEqual(aktion, ACTION_SKIP)
        self.assertIsNone(wert)

    def test_slash_skip_befehl(self):
        """'/skip' überspringt immer unabhängig von Optionen."""
        aktion, wert = evaluate_dialog_response("/skip", self.options)
        self.assertEqual(aktion, ACTION_SKIP)
        self.assertIsNone(wert)

    def test_timeout_none(self):
        """None (Timeout) führt zu Skip."""
        aktion, wert = evaluate_dialog_response(None, self.options)
        self.assertEqual(aktion, ACTION_SKIP)
        self.assertIsNone(wert)

    def test_direkte_namenseingabe(self):
        """Direkt eingegebener Text wird als Name übernommen."""
        aktion, wert = evaluate_dialog_response("LB_FK_Mein_Eigener_Film_12_05", self.options)
        self.assertEqual(aktion, ACTION_SET_NAME)
        self.assertEqual(wert, "LB_FK_Mein_Eigener_Film_12_05")


class TestFormatTelegramDialog(unittest.TestCase):
    """Tests für format_telegram_dialog."""

    def test_auto_merged_dialog(self):
        """AUTO_MERGED erzeugt klaren Bestätigungsdialog."""
        res = ResolvedResult(
            category=ResolvedField(value="FK", quality="HIGH"),
            title=ResolvedField(value="Shrek", quality="HIGH"),
            date=ResolvedField(value="18_10", quality="HIGH"),
            status=STATUS_AUTO_MERGED,
        )
        msg, opts = format_telegram_dialog(res, "LB_FK_Shrek_18_10")
        self.assertIn("LB_FK_Shrek_18_10", msg)
        self.assertIn("[1]  Übernehmen", msg)
        self.assertIn("[2]  Eigenen Namen eingeben", msg)
        self.assertIn("[3]  Überspringen", msg)
        self.assertEqual(opts["1"], (ACTION_SET_NAME, "LB_FK_Shrek_18_10"))
        self.assertEqual(opts["2"], (ACTION_ASK_CUSTOM, None))
        self.assertEqual(opts["3"], (ACTION_SKIP, None))

    def test_needs_review_datum_fehlt_bietet_datum_eingabe(self):
        """NEEDS_REVIEW mit fehlendem Datum bietet Datum-Eingabe statt Übernahme an."""
        res = ResolvedResult(
            category=ResolvedField(value="FK", quality="HIGH"),
            title=ResolvedField(value="Shrek", quality="HIGH"),
            date=ResolvedField(value=None, quality=None),
            status=STATUS_NEEDS_REVIEW,
            missing_fields=["date"],
            review_reasons=["Pflichtfeld 'date' fehlt"],
        )
        msg, opts = format_telegram_dialog(res, "LB_FK_Shrek")
        self.assertIn("Prüfung erforderlich", msg)
        self.assertIn("Pflichtfeld 'date' fehlt", msg)
        self.assertIn("Erkannter Titel: Shrek", msg)
        # Kein [1] Vorschlag übernehmen!
        self.assertNotIn("Vorschlag übernehmen", msg)
        # Stattdessen Datum-Eingabe
        self.assertIn("Datum eingeben", msg)
        self.assertEqual(opts["1"], (ACTION_ASK_DATE, None))
        # Vollständiger Name + Skip weiterhin verfügbar
        has_custom = any(a == ACTION_ASK_CUSTOM for a, _ in opts.values())
        has_skip = any(a == ACTION_SKIP for a, _ in opts.values())
        self.assertTrue(has_custom)
        self.assertTrue(has_skip)

    def test_needs_review_titel_fehlt_bietet_titel_eingabe(self):
        """NEEDS_REVIEW mit fehlendem Titel bietet Titel-Eingabe an."""
        res = ResolvedResult(
            category=ResolvedField(value="FK", quality="HIGH"),
            title=ResolvedField(value=None, quality=None),
            date=ResolvedField(value="18_10", quality="HIGH"),
            status=STATUS_NEEDS_REVIEW,
            missing_fields=["title"],
            review_reasons=["Pflichtfeld 'title' fehlt"],
        )
        msg, opts = format_telegram_dialog(res, None)
        self.assertNotIn("Vorschlag übernehmen", msg)
        self.assertIn("Filmtitel eingeben", msg)
        self.assertIn("Erkanntes Datum: 18_10", msg)
        self.assertEqual(opts["1"], (ACTION_ASK_TITLE, None))

    def test_needs_review_beide_fehlen(self):
        """NEEDS_REVIEW mit beiden fehlenden Pflichtfeldern bietet Titel-Eingabe an."""
        res = ResolvedResult(
            category=ResolvedField(value=None),
            title=ResolvedField(value=None),
            date=ResolvedField(value=None),
            status=STATUS_NEEDS_REVIEW,
            missing_fields=["date", "title"],
            review_reasons=["Pflichtfeld 'date' fehlt", "Pflichtfeld 'title' fehlt"],
        )
        msg, opts = format_telegram_dialog(res, None)
        self.assertNotIn("Vorschlag übernehmen", msg)
        self.assertIn("Filmtitel eingeben", msg)
        self.assertEqual(opts["1"], (ACTION_ASK_TITLE, None))
        has_custom = any(a == ACTION_ASK_CUSTOM for a, _ in opts.values())
        has_skip = any(a == ACTION_SKIP for a, _ in opts.values())
        self.assertTrue(has_custom)
        self.assertTrue(has_skip)

    def test_needs_review_ohne_fehlende_pflichtfelder_bietet_uebernahme(self):
        """NEEDS_REVIEW ohne fehlende Pflichtfelder (z.B. nur MEDIUM) bietet Übernahme an."""
        res = ResolvedResult(
            category=ResolvedField(value="FK", quality="MEDIUM"),
            title=ResolvedField(value="Shrek", quality="MEDIUM"),
            date=ResolvedField(value="18_10", quality="MEDIUM"),
            status=STATUS_NEEDS_REVIEW,
            missing_fields=[],
            review_reasons=["Feld 'title' erreicht nur Qualität MEDIUM"],
        )
        msg, opts = format_telegram_dialog(res, "LB_FK_Shrek_18_10")
        self.assertIn("Vorschlag übernehmen", msg)
        self.assertEqual(opts["1"], (ACTION_SET_NAME, "LB_FK_Shrek_18_10"))

    def test_conflict_dialog_mit_varianten(self):
        """CONFLICT zeigt Konfliktgründe und Varianten zur Auswahl an."""
        res = ResolvedResult(
            category=ResolvedField(value="TK", quality="HIGH"),
            title=ResolvedField(
                value="Das_Boot",
                quality="HIGH",
                conflict=True,
                alternatives=["Shrek"],
                reason="HIGH-Kandidat (Das_Boot) vs. HIGH-Kandidat (Shrek)",
            ),
            date=ResolvedField(value="21_11", quality="HIGH"),
            status=STATUS_CONFLICT,
            review_reasons=["Konflikt im Feld 'title': Das_Boot vs. Shrek"],
        )
        alts = ["LB_TK_Das_Boot_21_11", "LB_TK_Shrek_21_11"]
        msg, opts = format_telegram_dialog(res, "LB_TK_Das_Boot_21_11", alts)
        self.assertIn("Konflikt erkannt", msg)
        self.assertIn("Mögliche Varianten:", msg)
        self.assertIn("[1]  LB_TK_Das_Boot_21_11", msg)
        self.assertIn("[2]  LB_TK_Shrek_21_11", msg)
        self.assertEqual(opts["1"], (ACTION_SET_NAME, "LB_TK_Das_Boot_21_11"))
        self.assertEqual(opts["2"], (ACTION_SET_NAME, "LB_TK_Shrek_21_11"))
        self.assertEqual(opts["3"], (ACTION_ASK_CUSTOM, None))
        self.assertEqual(opts["4"], (ACTION_SKIP, None))


class TestProcessImageNamingIntegration(unittest.TestCase):
    """End-to-End Tests für process_image_naming."""

    def _make_fake_runner(self, text_to_return: str):
        """Erzeugt einen Fake-OCR-Runner."""
        def _runner(image, language, pass_name):
            return text_to_return
        return _runner

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_1_parser_und_ocr_vollstaendig_uebereinstimmend(self, mock_open):
        """Szenario 1: Dateiname und OCR stimmen perfekt überein -> AUTO_MERGED."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("FILMKLASSIKER\nShrek\n18.10")
        res = process_image_naming(
            "C:/bilder/FK_Shrek_18_10.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_AUTO_MERGED)
        self.assertEqual(res.dcp_name_proposal, "LB_FK_Shrek_18_10")
        self.assertEqual(res.resolved.category.value, "FK")
        self.assertEqual(res.resolved.title.value, "Shrek")
        self.assertEqual(res.resolved.date.value, "18_10")
        self.assertIn("[1]  Übernehmen", res.dialog_message)

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_2_filename_high_ocr_bestaetigt(self, mock_open):
        """Szenario 2: Dateiname liefert HIGH-Titel und HIGH-Datum, OCR bestätigt Titel."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("Shrek")
        res = process_image_naming(
            "C:/bilder/Shrek_18_10.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_AUTO_MERGED)
        self.assertEqual(res.dcp_name_proposal, "LB_Shrek_18_10")
        self.assertEqual(res.resolved.title.value, "Shrek")
        self.assertEqual(res.resolved.date.value, "18_10")

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_3_ocr_mit_zusatzrauschen(self, mock_open):
        """Szenario 3: Reales Kinoplakat mit Rauschen (FSK, Uhrzeit, URLs) wird korrekt gefiltert."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        ocr_text = "FILMKLASSIKER\nSHREK\nAB 6 JAHREN\n20 UHR\nJETZT IM KINO\n18.10"
        fake_runner = self._make_fake_runner(ocr_text)

        res = process_image_naming(
            "C:/bilder/image.jpg",  # generischer Dateiname
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)
        self.assertEqual(res.dcp_name_proposal, "LB_FK_SHREK_18_10")
        self.assertEqual(res.resolved.category.value, "FK")
        self.assertEqual(res.resolved.title.value, "SHREK")
        self.assertEqual(res.resolved.date.value, "18_10")

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_4_generischer_dateiname_mit_guter_ocr(self, mock_open):
        """Szenario 4: Generischer Dateiname IMG_001.jpg mit starker OCR."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("TRAUMKINO\nDas Boot\n21.11")
        res = process_image_naming(
            "C:/bilder/IMG_001.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.dcp_name_proposal, "LB_TK_Das_Boot_21_11")
        self.assertEqual(res.resolved.category.value, "TK")
        self.assertEqual(res.resolved.title.value, "Das Boot")
        self.assertEqual(res.resolved.date.value, "21_11")

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_5_unterschiedliche_titel_konflikt(self, mock_open):
        """Szenario 5: Dateiname und OCR liefern unterschiedliche Titel."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("TRAUMKINO\nShrek\n21.11")
        res = process_image_naming(
            "C:/bilder/TK_Das_Boot_21_11.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)
        self.assertEqual(res.resolved.title.value, "Das_Boot")

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_6_unterschiedliche_kategorien_konflikt(self, mock_open):
        """Szenario 6: Zwei HIGH-Kategorien (MeK im Dateinamen vs. ZiK im OCR) -> CONFLICT."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("ZURÜCK IM KINO\nShrek\n18.10")
        res = process_image_naming(
            "C:/bilder/MeK_Shrek_18_10.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_CONFLICT)
        self.assertTrue(res.resolved.category.conflict)
        self.assertIn("Konflikt", res.dialog_message)

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_7_unterschiedliche_daten_konflikt(self, mock_open):
        """Szenario 7: Zwei HIGH-Daten (18_10 im Dateinamen vs. 21_11 im OCR) -> CONFLICT."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("FILMKLASSIKER\nShrek\n21.11")
        res = process_image_naming(
            "C:/bilder/FK_Shrek_18_10.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_CONFLICT)
        self.assertTrue(res.resolved.date.conflict)

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_8_fehlender_titel(self, mock_open):
        """Szenario 8: Nur Datum vorhanden, kein Titel -> NEEDS_REVIEW, Titel-Eingabe angeboten."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("18.10")
        res = process_image_naming(
            "C:/bilder/18_10.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)
        self.assertIn("title", res.resolved.missing_fields)
        self.assertIsNone(res.dcp_name_proposal)
        # Nicht Übernehmen, sondern Titel-Eingabe
        self.assertNotIn("Vorschlag übernehmen", res.dialog_message)
        self.assertIn("Filmtitel eingeben", res.dialog_message)

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_9_fehlendes_datum(self, mock_open):
        """Szenario 9: Nur Titel vorhanden, kein Datum -> NEEDS_REVIEW, Datum-Eingabe angeboten."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        fake_runner = self._make_fake_runner("FILMKLASSIKER\nShrek")
        res = process_image_naming(
            "C:/bilder/FK_Shrek.jpg",
            ocr_runner=fake_runner,
        )

        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)
        self.assertIn("date", res.resolved.missing_fields)
        # Unvollständiger Name LB_FK_Shrek nicht direkt übernehmbar
        self.assertNotIn("Vorschlag übernehmen", res.dialog_message)
        self.assertIn("Datum eingeben", res.dialog_message)
        self.assertEqual(res.options["1"], (ACTION_ASK_DATE, None))

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_10_ocr_ausfall_robustheit(self, mock_open):
        """Szenario 10: OCR-Runner wirft Ausnahme -> kein Absturz, Dateiname wird genutzt."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        def broken_runner(img, lang, pass_name):
            raise RuntimeError("Tesseract crash")

        res = process_image_naming(
            "C:/bilder/FK_Shrek_18_10.jpg",
            ocr_runner=broken_runner,
        )

        self.assertEqual(res.status, STATUS_AUTO_MERGED)
        self.assertEqual(res.dcp_name_proposal, "LB_FK_Shrek_18_10")
        self.assertTrue(len(res.ocr_result.errors) > 0)

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_11_spezielle_titel_geschuetzt(self, mock_open):
        """Szenario 11: Spezielle Titel wie 1917, Se7en, WALL-E bleiben erhalten."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        # 1917
        res1 = process_image_naming(
            "C:/bilder/1917_18_10.jpg",
            ocr_runner=self._make_fake_runner("FILMKLASSIKER\n1917\n18.10"),
        )
        self.assertEqual(res1.dcp_name_proposal, "LB_FK_1917_18_10")

        # Se7en
        res2 = process_image_naming(
            "C:/bilder/Se7en_18_10.jpg",
            ocr_runner=self._make_fake_runner("FILMKLASSIKER\nSe7en\n18.10"),
        )
        self.assertEqual(res2.dcp_name_proposal, "LB_FK_Se7en_18_10")

        # WALL-E -> WALL_E im Parser
        res3 = process_image_naming(
            "C:/bilder/WALL-E_18_10.jpg",
            ocr_runner=self._make_fake_runner("ZURÜCK IM KINO\nWALL_E\n18.10"),
        )
        self.assertEqual(res3.dcp_name_proposal, "LB_ZiK_WALL_E_18_10")

    @patch("modules.naming_ocr.Image.open")
    def test_szenario_12_filename_override_parameter(self, mock_open):
        """Szenario 12: filename_override überschreibt den Dateinamen für den Parser."""
        mock_img = MagicMock()
        mock_open.return_value = mock_img

        res = process_image_naming(
            "C:/temp/download_12345.tmp",
            ocr_runner=self._make_fake_runner(""),
            filename_override="FK_Shrek_18_10.jpg",
        )

        self.assertEqual(res.status, STATUS_AUTO_MERGED)
        self.assertEqual(res.dcp_name_proposal, "LB_FK_Shrek_18_10")


class TestMissingFieldsDialog(unittest.TestCase):
    """Tests für den gehärteten Dialog bei fehlenden Pflichtfeldern."""

    def _make_fake_runner(self, text):
        def _runner(image, language, pass_name):
            return text
        return _runner

    @patch("modules.naming_ocr.Image.open")
    def test_datum_fehlt_kein_uebernahme_button(self, mock_open):
        """Bei fehlendem Datum darf LB_FK_Shrek nicht via [1] Übernehmen bestätigt werden."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/FK_Shrek.jpg",
            ocr_runner=self._make_fake_runner("FILMKLASSIKER\nShrek"),
        )
        # Kein SET_NAME in den Optionen
        for opt_key, (action, val) in res.options.items():
            self.assertNotEqual(action, ACTION_SET_NAME,
                f"Option {opt_key} bietet SET_NAME an obwohl Datum fehlt")

    @patch("modules.naming_ocr.Image.open")
    def test_titel_fehlt_kein_uebernahme_button(self, mock_open):
        """Bei fehlendem Titel darf kein Name zur direkten Übernahme angeboten werden."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/18_10.jpg",
            ocr_runner=self._make_fake_runner("18.10"),
        )
        for opt_key, (action, val) in res.options.items():
            self.assertNotEqual(action, ACTION_SET_NAME,
                f"Option {opt_key} bietet SET_NAME an obwohl Titel fehlt")

    @patch("modules.naming_ocr.Image.open")
    def test_beide_fehlen_kein_uebernahme_button(self, mock_open):
        """Bei fehlendem Titel und Datum darf kein Name zur direkten Übernahme angeboten werden."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/image.jpg",
            ocr_runner=self._make_fake_runner(""),
        )
        for opt_key, (action, val) in res.options.items():
            self.assertNotEqual(action, ACTION_SET_NAME,
                f"Option {opt_key} bietet SET_NAME an obwohl beide Felder fehlen")

    def test_build_dcp_name_ergaenzt_datum(self):
        """Benutzer ergänzt fehlendes Datum -> vollständiger DCP-Name."""
        name = build_dcp_name(category="FK", title="Shrek", date="18_10")
        self.assertEqual(name, "LB_FK_Shrek_18_10")

    def test_build_dcp_name_ergaenzt_titel(self):
        """Benutzer ergänzt fehlenden Titel -> vollständiger DCP-Name."""
        name = build_dcp_name(category="FK", title="Shrek", date="18_10")
        self.assertEqual(name, "LB_FK_Shrek_18_10")

    @patch("modules.naming_ocr.Image.open")
    def test_dialog_zeigt_erkannte_felder_bei_fehlendem_datum(self, mock_open):
        """Bei fehlendem Datum zeigt der Dialog die bereits erkannten Felder an."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/FK_Shrek.jpg",
            ocr_runner=self._make_fake_runner("FILMKLASSIKER\nShrek"),
        )
        self.assertIn("Erkannter Titel: Shrek", res.dialog_message)
        self.assertIn("Erkannte Kategorie: FK", res.dialog_message)

    @patch("modules.naming_ocr.Image.open")
    def test_dialog_zeigt_erkanntes_datum_bei_fehlendem_titel(self, mock_open):
        """Bei fehlendem Titel zeigt der Dialog das bereits erkannte Datum an."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/18_10.jpg",
            ocr_runner=self._make_fake_runner("18.10"),
        )
        self.assertIn("Erkanntes Datum: 18_10", res.dialog_message)

    @patch("modules.naming_ocr.Image.open")
    def test_skip_und_custom_weiterhin_verfuegbar_bei_fehlendem_datum(self, mock_open):
        """Bei fehlendem Datum bleiben /skip und vollständige Namenseingabe verfügbar."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/FK_Shrek.jpg",
            ocr_runner=self._make_fake_runner("FILMKLASSIKER\nShrek"),
        )
        actions = [a for a, _ in res.options.values()]
        self.assertIn(ACTION_SKIP, actions)
        self.assertIn(ACTION_ASK_CUSTOM, actions)


class TestOCRAusfallRobustheit(unittest.TestCase):
    """Tests für OCR-/Tesseract-Ausfall-Szenarien."""

    @patch("modules.naming_ocr.Image.open")
    def test_leerer_tesseract_pfad_kein_absturz(self, mock_open):
        """Leerer Tesseract-Pfad erzeugt keinen Absturz."""
        mock_open.return_value = MagicMock()
        res = process_image_naming(
            "C:/bilder/FK_Shrek_18_10.jpg",
            tesseract_cmd="",
        )
        # Kein Runner erzeugt, kein OCR -> nur Dateiname
        self.assertIsNotNone(res)
        self.assertEqual(res.resolved.title.value, "Shrek")
        self.assertEqual(res.resolved.date.value, "18_10")

    @patch("modules.naming_ocr.Image.open")
    def test_ocr_runner_exception_kein_absturz(self, mock_open):
        """OCR-Runner wirft verschiedene Exceptions -> kein Absturz."""
        mock_open.return_value = MagicMock()

        for exc in [RuntimeError("crash"), OSError("not found"), ValueError("bad")]:
            def broken(img, lang, pass_name, e=exc):
                raise e
            res = process_image_naming(
                "C:/bilder/FK_Shrek_18_10.jpg",
                ocr_runner=broken,
            )
            self.assertEqual(res.status, STATUS_AUTO_MERGED)
            self.assertEqual(res.dcp_name_proposal, "LB_FK_Shrek_18_10")
            self.assertTrue(len(res.ocr_result.errors) > 0)

    @patch("modules.naming_ocr.Image.open")
    def test_ocr_ausfall_mit_generischem_dateinamen(self, mock_open):
        """OCR-Ausfall + generischer Dateiname -> NEEDS_REVIEW, kein Absturz."""
        mock_open.return_value = MagicMock()

        def broken_runner(img, lang, pass_name):
            raise RuntimeError("Tesseract crash")

        res = process_image_naming(
            "C:/bilder/IMG_001.jpg",
            ocr_runner=broken_runner,
        )

        self.assertEqual(res.status, STATUS_NEEDS_REVIEW)
        self.assertIsNone(res.dcp_name_proposal)
        self.assertTrue(len(res.ocr_result.errors) > 0)
        # Benutzer kann weiterhin manuell eingeben oder überspringen
        actions = [a for a, _ in res.options.values()]
        self.assertTrue(ACTION_ASK_CUSTOM in actions or ACTION_ASK_TITLE in actions)
        self.assertIn(ACTION_SKIP, actions)

    @patch("modules.naming_ocr.Image.open")
    def test_ocr_ausfall_mit_gutem_dateinamen(self, mock_open):
        """OCR-Ausfall + guter Dateiname -> AUTO_MERGED, Dateiname genügt."""
        mock_open.return_value = MagicMock()

        def broken_runner(img, lang, pass_name):
            raise RuntimeError("Tesseract crash")

        res = process_image_naming(
            "C:/bilder/FK_Shrek_18_10.jpg",
            ocr_runner=broken_runner,
        )

        self.assertEqual(res.status, STATUS_AUTO_MERGED)
        self.assertEqual(res.dcp_name_proposal, "LB_FK_Shrek_18_10")
        self.assertIn("[1]  Übernehmen", res.dialog_message)

    @patch("modules.naming_ocr.Image.open")
    def test_ocr_ausfall_dialog_funktioniert(self, mock_open):
        """Bei OCR-Ausfall funktioniert der Dialog weiterhin mit evaluate_dialog_response."""
        mock_open.return_value = MagicMock()

        def broken_runner(img, lang, pass_name):
            raise RuntimeError("Tesseract crash")

        res = process_image_naming(
            "C:/bilder/FK_Shrek_18_10.jpg",
            ocr_runner=broken_runner,
        )

        # Benutzer wählt [1] Übernehmen
        aktion, wert = evaluate_dialog_response("1", res.options)
        self.assertEqual(aktion, ACTION_SET_NAME)
        self.assertEqual(wert, "LB_FK_Shrek_18_10")

        # Benutzer wählt /skip
        aktion2, wert2 = evaluate_dialog_response("/skip", res.options)
        self.assertEqual(aktion2, ACTION_SKIP)


class TestTesseractRunnerFactory(unittest.TestCase):
    """Tests für create_tesseract_runner."""

    def test_create_tesseract_runner_execution(self):
        """Der erzeugte Runner ruft pytesseract korrekt mit der Sprache auf."""
        mock_pytesseract = MagicMock()
        mock_pytesseract.image_to_string.return_value = "Test OCR Result"

        with patch.dict("sys.modules", {"pytesseract": mock_pytesseract}):
            runner = create_tesseract_runner(tesseract_cmd="C:\\Tesseract-OCR\\tesseract.exe")
            img = MagicMock()
            result = runner(img, "deu+eng", "original")

            self.assertEqual(result, "Test OCR Result")
            self.assertEqual(mock_pytesseract.pytesseract.tesseract_cmd, "C:\\Tesseract-OCR\\tesseract.exe")
            mock_pytesseract.image_to_string.assert_called_once_with(img, lang="deu+eng")


if __name__ == "__main__":
    unittest.main()
