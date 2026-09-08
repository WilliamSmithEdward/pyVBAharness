"""Command-line argument routing, which runs before any Office process.

A positional path is routed on its extension. Reading a workbook as VBA
source used to fail deep in a decoder with "'utf-8' codec can't decode byte
0xf6", after check had already paid for an Excel start.
"""
import pytest

from pyvbaharness import apps
from pyvbaharness.__main__ import (
    _CliError,
    _read_source,
    _resolve_inputs,
    main,
)

# A real .xlsm is a zip; only the first bytes matter for the routing test.
WORKBOOK_BYTES = b"PK\x03\x04\x14\x00\x06\x00\x08\x00\xf6\xa1\x00\x00"
SOURCE = "Public Sub Main()\nEnd Sub\n"


@pytest.fixture
def files(tmp_path):
    def make(name, data):
        path = tmp_path / name
        if isinstance(data, bytes):
            path.write_bytes(data)
        else:
            path.write_text(data, encoding="utf-8")
        return path
    return make


class TestDocumentExtensions:
    @pytest.mark.parametrize("suffix, app", [
        (".xlsm", "excel"), (".xlsb", "excel"), (".xlsx", "excel"),
        (".XLSM", "excel"),
        (".docm", "word"), (".docx", "word"),
        (".pptm", "powerpoint"), (".pptx", "powerpoint"),
        (".accdb", "access"), (".mdb", "access"),
    ])
    def test_office_extensions_are_owned(self, suffix, app):
        assert apps.app_for_document(suffix) == app

    @pytest.mark.parametrize("suffix", [".bas", ".cls", ".frm", ".vba",
                                        ".txt", ""])
    def test_source_extensions_are_not_documents(self, suffix):
        """VBA source has no reserved extension, so anything unrecognised is
        source rather than a guess."""
        assert apps.app_for_document(suffix) is None

    def test_every_save_format_is_also_a_document_extension(self):
        for detail in apps.APPS.values():
            missing = set(detail.save_formats) - set(detail.document_extensions)
            assert not missing, f"{detail.key}: {missing}"


class TestResolveInputs:
    def test_source_files_stay_sources(self, files):
        one = files("Mod1.bas", SOURCE)
        two = files("Mod2.cls", SOURCE)
        sources, workbook = _resolve_inputs([str(one), str(two)])
        assert sources == [one, two]
        assert workbook is None

    def test_workbook_routes_away_from_source_reading(self, files):
        book = files("model.xlsm", WORKBOOK_BYTES)
        sources, workbook = _resolve_inputs([str(book)])
        assert sources == []
        assert workbook == book

    def test_missing_file_is_named(self, files, tmp_path):
        with pytest.raises(_CliError, match="no such file"):
            _resolve_inputs([str(tmp_path / "absent.bas")])

    def test_non_excel_document_says_which_host(self, files):
        book = files("report.docm", WORKBOOK_BYTES)
        with pytest.raises(_CliError, match="Word document"):
            _resolve_inputs([str(book)])

    def test_two_workbooks_are_refused(self, files):
        first = files("a.xlsm", WORKBOOK_BYTES)
        second = files("b.xlsm", WORKBOOK_BYTES)
        with pytest.raises(_CliError, match="one workbook at a time"):
            _resolve_inputs([str(first), str(second)])

    def test_workbook_mixed_with_sources_is_refused(self, files):
        book = files("a.xlsm", WORKBOOK_BYTES)
        module = files("Mod1.bas", SOURCE)
        with pytest.raises(_CliError, match="not both"):
            _resolve_inputs([str(book), str(module)])


class TestReadSource:
    def test_plain_utf8(self, files):
        assert _read_source(files("Mod1.bas", SOURCE)) == SOURCE

    def test_bom_is_stripped(self, files):
        path = files("Mod1.bas", b"\xef\xbb\xbf" + SOURCE.encode("utf-8"))
        assert _read_source(path) == SOURCE

    def test_crlf_is_normalised(self, files):
        """Text-mode reading collapses CRLF, which is what the VBE receives
        either way. Pinned because .bas files on disk are CRLF."""
        path = files("Mod1.bas", SOURCE.replace("\n", "\r\n").encode("utf-8"))
        assert _read_source(path) == SOURCE

    def test_ansi_export_is_read_not_rejected(self, files):
        """VBIDE Export writes the ANSI code page, so a module with an
        accented character is not valid UTF-8."""
        path = files("Mod1.bas",
                     'Public Sub Main()\r\n    Debug.Print "café"\r\n'
                     'End Sub\r\n'.encode("cp1252"))
        assert "café" in _read_source(path)


class TestExitCodes:
    """A bad invocation reports one line and exit 2, never a traceback."""

    def test_check_with_no_arguments(self, capsys):
        assert main(["check"]) == 2
        assert "source files or a workbook" in capsys.readouterr().out

    def test_check_missing_file(self, capsys, tmp_path):
        assert main(["check", str(tmp_path / "absent.bas")]) == 2
        assert "no such file" in capsys.readouterr().out

    def test_check_positional_and_flag_workbook(self, capsys, files):
        book = files("a.xlsm", WORKBOOK_BYTES)
        assert main(["check", str(book), "--workbook", str(book)]) == 2
        assert "one workbook at a time" in capsys.readouterr().out

    def test_run_rejects_a_word_document(self, capsys, files):
        book = files("report.docm", WORKBOOK_BYTES)
        assert main(["run", str(book)]) == 2
        assert "WordSession" in capsys.readouterr().out

    def test_bad_module_name_is_caught_before_excel(self, capsys, files):
        """The stem check runs during argument validation, so it costs no
        Excel start and cannot leave a half-populated workbook."""
        path = files("9bad.bas", SOURCE)
        assert main(["check", str(path)]) == 2
        assert "valid VBA module name" in capsys.readouterr().out
