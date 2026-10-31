"""Live regression for private Word conversion, command guards and long numbering.

Creates only a new test document and uses a private conversion instance. The
controlled failure check simulates a *returned* failed macro; it never clears a
pre-existing busy flag or interrupts an active native call.
"""

from __future__ import annotations

import argparse
import json
import time
import uuid
import winreg
from contextlib import contextmanager
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

from axmath_mcp.backend import TO_AXMATH
from axmath_mcp.font_metrics import read_equation_metrics
from axmath_mcp.office import OfficeBackend
from axmath_mcp.word_workspace import configure_private_printer

ROOT = Path(__file__).resolve().parents[1]
KEY = r"Software\AxMath\WordCmds"
MARKOV = r"$$P(X_{n+1}=j\mid X_n=i,X_{n-1}=i_{n-1},...,X_0=i_0)=P(X_{n+1}=j\mid X_n=i)$$"
STATIONARY = r"$$\pi _j=\sum_{i\in S}{\pi _ip_{ij},\qquad \sum_{j\in S}{\pi _j}}=1$$"


def read_reg(key, name):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as handle:
        return winreg.QueryValueEx(handle, name)[0]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exercise-office", action="store_true")
    args = parser.parse_args()
    if not args.exercise_office:
        parser.error("Use --exercise-office to explicitly run the live tests.")
    assert read_reg(KEY, "ConvertOccupy") == "0", "Finish active AxMath work first."
    printer_key = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
    printer_before = read_reg(printer_key, "Device")
    output = ROOT / "artifacts" / ("resilience-" + uuid.uuid4().hex[:8])
    output.mkdir(parents=True)
    pythoncom.CoInitialize()
    target_app = document = office = None
    results = {"ok": False, "cases": [], "observed_peak_conversion_documents": 0}
    try:
        office = OfficeBackend()
        target_app = win32com.client.DispatchEx("Word.Application")
        target_app.Visible = False
        configure_private_printer(target_app)  # only this newly-created test app
        document = target_app.Documents.Add(Visible=False)
        document.PageSetup.PageWidth = 595.3
        document.PageSetup.PageHeight = 841.9
        document.PageSetup.LeftMargin = document.PageSetup.RightMargin = 73.7
        document.PageSetup.TopMargin = document.PageSetup.BottomMargin = 60
        document.Styles(-1).Font.Name = "Arial"
        document.Styles(-1).Font.NameFarEast = "宋体"
        document.Styles(-1).Font.Size = 14
        document.Styles(-1).ParagraphFormat.LineSpacingRule = 3
        document.Styles(-1).ParagraphFormat.LineSpacing = 20
        did = office.word._remember_doc(document)
        workspace = office.word._workspace
        original_scratch = workspace.scratch

        @contextmanager
        def observed_scratch():
            with original_scratch() as scratch:
                results["observed_peak_conversion_documents"] = max(
                    results["observed_peak_conversion_documents"],
                    scratch.Application.Documents.Count,
                )
                yield scratch

        workspace.scratch = observed_scratch

        def check_cleanup():
            state = office.word.conversion_status()
            assert state["temporary_document_count"] == state["pending_count"] == 0, state
            assert document.Application.Documents.Count == 1
            assert workspace.get_app().Documents.Count == 0
            assert not workspace.get_app().Visible
            assert read_reg(KEY, "ConvertOccupy") == "0"
            assert read_reg(printer_key, "Device") == printer_before

        def heading(title):
            if document.Paragraphs.Last.Range.Text != "\r":
                document.Range(document.Content.End - 1, document.Content.End - 1).InsertAfter("\r")
            document.Paragraphs.Last.Range.Style = document.Styles(-1)
            if title:
                document.Range(document.Content.End - 1, document.Content.End - 1).InsertAfter(
                    title + "\r"
                )
                document.Paragraphs.Last.Range.Style = document.Styles(-1)

        before = workspace.status()["word_started"]
        try:
            office.insert_equation(did, r"d(i)=\gcd\{n\ge1:p_{ii}^{(n)}>0\}", font_size_pt=12)
        except ValueError as error:
            assert "\\gcd" in str(error)
            results["direct_gcd_rejected_before_word"] = (
                not before and not workspace.status()["word_started"]
            )
        else:
            raise AssertionError("Direct gcd must fail before native conversion.")
        assert read_reg(KEY, "ConvertOccupy") == "0"

        cases = [
            ("Original Markov property", MARKOV, "display"),
            ("Original stationary distribution", STATIONARY, "numbered"),
            ("forall", r"\forall x\in S", "numbered"),
            ("mathbf", r"\mathbf{P}=\begin{pmatrix}0.8&0.3\\0.2&0.7\end{pmatrix}", "numbered"),
            ("ge", r"n\ge1", "numbered"),
            (
                "Period with explicit operator",
                r"d(i)=\operatorname{gcd}\{n\ge1:p_{ii}^{(n)}>0\}",
                "numbered",
            ),
        ]
        for name, source, layout in cases:
            heading(name)
            started = time.monotonic()
            result = office.insert_equation(did, source, layout=layout, font_size_pt=12)
            if name.startswith("Original Markov"):
                result = office.format_equation(did, result["equation_id"], layout="numbered")
                assert result["font_size_pt"] == 12
                assert result["width_pt"] > 300
                assert (
                    read_equation_metrics(office.word._eq(did, result["equation_id"]).Range)[
                        "baseline_offset_pt"
                    ]
                    < 0
                )
            assert result["layout"] == "numbered"
            check_cleanup()
            results["cases"].append(
                {"name": name, "seconds": round(time.monotonic() - started, 3), **result}
            )
            print(json.dumps({"passed": name, "width_pt": result["width_pt"]}), flush=True)
            heading("")

        first = office.word._eq(did, results["cases"][0]["equation_id"])
        first_start = first.Range.Paragraphs(1).Range.Start
        document.Range(first_start, first_start).InsertBefore("\r")
        document.Bookmarks.Add("NativeNumberProbe", document.Range(first_start, first_start))
        extra = office.insert_equation(
            did,
            "a=b",
            position="bookmark",
            bookmark="NativeNumberProbe",
            layout="numbered",
            font_size_pt=12,
        )
        assert office.refresh_equation_numbers(did)["numbers"] == [str(i) for i in range(1, 8)]
        updated = office.update_equation(did, extra["equation_id"], "a=b+c")
        assert updated["number_text"] == "(1)"
        display = office.format_equation(did, extra["equation_id"], layout="display")
        assert display["number_text"] is None
        numbered = office.format_equation(did, extra["equation_id"], layout="numbered")
        assert numbered["number_text"] == "(1)"
        office.word._eq(did, extra["equation_id"]).Range.Paragraphs(1).Range.Delete()
        assert office.refresh_equation_numbers(did)["numbers"] == [str(i) for i in range(1, 7)]
        document.Activate()
        document.Range(0, 0).Select()
        target_app.Run("AMCUpdateFields.Main")
        assert office.refresh_equation_numbers(did)["numbers"] == [str(i) for i in range(1, 7)]
        assert (
            read_equation_metrics(office.word._eq(did, results["cases"][0]["equation_id"]).Range)[
                "baseline_offset_pt"
            ]
            < 0
        )
        results["native_insert_update_reformat_delete_and_vendor_refresh"] = True

        # Check that transfer into an ordinary mixed paragraph does not introduce
        # a paragraph break or overwrite surrounding font/style properties.
        heading("Mixed 中文 / English")
        start = document.Content.End - 1
        document.Range(start, start).InsertAfter("中文 before TOKEN after English.\r")
        token = document.Range(start + 10, start + 15)
        assert token.Text == "TOKEN"
        token.Select()
        inline = office.insert_equation(
            did,
            r"\frac{a+b}{c+d}",
            position="selection",
            expected_selection="TOKEN",
            font_size_pt=12,
        )
        paragraph = office.word._eq(did, inline["equation_id"]).Range.Paragraphs(1).Range
        assert paragraph.Text == "中文 before \x01 after English.\r"
        assert document.Range(start, start + 1).Font.Size == 14
        assert paragraph.Characters(paragraph.Characters.Count - 1).Font.Size == 14
        check_cleanup()
        results["mixed_text_and_fonts_preserved"] = True

        # Exercise recovery only after a simulated synchronous macro has exited.
        with office.word._conversion() as converter:

            class FailedMacro:
                _oleobj_ = converter._oleobj_
                Documents = converter.Documents

                def Run(self, *_):
                    with winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER, KEY, 0, winreg.KEY_SET_VALUE
                    ) as handle:
                        winreg.SetValueEx(handle, "ConvertOccupy", 0, winreg.REG_SZ, "1")
                    raise pywintypes.com_error(
                        -2147352567, "simulated returned macro error", None, None
                    )

            proxy = FailedMacro()
            office.word._conversion_app = proxy
            try:
                office.word._run_macro(proxy, TO_AXMATH)
            except RuntimeError:
                pass
            else:
                raise AssertionError("Expected simulated macro failure")
            finally:
                office.word._conversion_app = converter
        assert read_reg(KEY, "ConvertOccupy") == "0"
        results["simulated_returned_failure_recovered"] = True
        heading("Next insertion after failure")
        office.insert_equation(did, "x^2+y^2=z^2", layout="numbered", font_size_pt=12)
        check_cleanup()
        results["next_native_insert_passed"] = True
        results["workspace"] = office.word.conversion_status()
        results["system_default_printer_unchanged"] = (
            read_reg(printer_key, "Device") == printer_before
        )
        file = output / "regression.docx"
        office.save_document(did, str(file))
        from zipfile import ZipFile

        with ZipFile(file) as archive:
            xml = archive.read("word/document.xml").decode("utf-8")
        assert "AMMPlaceRM" in xml and "SEQ AMEqn" in xml
        assert "AxMathEquation" not in xml
        results["only_native_number_fields"] = True
        document.ExportAsFixedFormat(str(file.with_suffix(".pdf")), 17)
        results.update(ok=True, docx=str(file), pdf=str(file.with_suffix(".pdf")))
    finally:
        if document is not None:
            document.Close(0)
        if office is not None:
            office.shutdown()
        if target_app is not None and target_app.Documents.Count == 0:
            target_app.Quit(0)
        pythoncom.CoUninitialize()
        (output / "verification.json").write_text(
            json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(json.dumps({"ok": results["ok"], "output": str(output)}), flush=True)


if __name__ == "__main__":
    main()
