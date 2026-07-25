from pyvbaharness.dialog_policy import (
    CLASS_COMPILE_ERROR,
    CLASS_EXCEL_MODAL,
    CLASS_RUNTIME_ERROR,
    CLASS_VBA_MODAL,
    IDCANCEL,
    IDHELP,
    IDNO,
    IDOK,
    IDYES,
    ButtonCapture,
    DialogCapture,
    classify,
    decide,
    fingerprint,
)


def vbe_runtime_dialog() -> DialogCapture:
    return DialogCapture(
        title="Microsoft Visual Basic",
        texts=["Run-time error '9':", "Subscript out of range"],
        buttons=[
            ButtonCapture("Continue", 0),
            ButtonCapture("End", 0),
            ButtonCapture("&Debug", 0),
            ButtonCapture("&Help", IDHELP),
        ],
    )


class TestClassification:
    def test_compile_error(self):
        dialog = DialogCapture(
            title="Microsoft Visual Basic for Applications",
            texts=["Compile error:", "Syntax error"],
            buttons=[ButtonCapture("OK", IDOK), ButtonCapture("Help", IDHELP)],
        )
        assert classify(dialog) == CLASS_COMPILE_ERROR

    def test_runtime_error(self):
        assert classify(vbe_runtime_dialog()) == CLASS_RUNTIME_ERROR

    def test_vba_modal(self):
        dialog = DialogCapture(title="Microsoft Visual Basic",
                               texts=["Something"], buttons=[])
        assert classify(dialog) == CLASS_VBA_MODAL

    def test_excel_modal(self):
        dialog = DialogCapture(title="Microsoft Excel", texts=["Save?"],
                               buttons=[])
        assert classify(dialog) == CLASS_EXCEL_MODAL


class TestDecisions:
    def test_compile_error_is_never_dismissed(self):
        dialog = DialogCapture(
            title="Microsoft Visual Basic for Applications",
            texts=["Compile error:", "Expected: expression"],
            buttons=[ButtonCapture("OK", IDOK)],
        )
        decision = decide(dialog)
        assert not decision.safe_to_dismiss
        assert decision.reason == "compile-error-blocks-runner"

    def test_runtime_error_dismissed_with_end(self):
        decision = decide(vbe_runtime_dialog())
        assert decision.safe_to_dismiss
        assert decision.button is not None
        assert decision.button.text == "End"

    def test_vba_modal_without_decision_buttons_uses_ok(self):
        dialog = DialogCapture(
            title="Microsoft Visual Basic",
            texts=["Out of memory"],
            buttons=[ButtonCapture("OK", IDOK), ButtonCapture("Help", IDHELP)],
        )
        decision = decide(dialog)
        assert decision.safe_to_dismiss
        assert decision.button is not None
        assert decision.button.control_id == IDOK

    def test_informational_ok_dismissed(self):
        dialog = DialogCapture(
            title="Microsoft Excel",
            texts=["hello from MsgBox"],
            buttons=[ButtonCapture("OK", IDOK)],
        )
        decision = decide(dialog)
        assert decision.safe_to_dismiss
        assert decision.reason == "dismiss-informational-ok"

    def test_yes_no_is_blocked(self):
        dialog = DialogCapture(
            title="Microsoft Excel",
            texts=["Overwrite the file?"],
            buttons=[ButtonCapture("&Yes", IDYES), ButtonCapture("&No", IDNO)],
        )
        decision = decide(dialog)
        assert not decision.safe_to_dismiss
        assert decision.reason == "decision-or-unknown-dialog"

    def test_ok_cancel_is_blocked(self):
        dialog = DialogCapture(
            title="Microsoft Excel",
            texts=["Proceed?"],
            buttons=[ButtonCapture("OK", IDOK),
                     ButtonCapture("Cancel", IDCANCEL)],
        )
        assert not decide(dialog).safe_to_dismiss

    def test_buttonless_dialog_is_blocked(self):
        dialog = DialogCapture(title="Microsoft Excel", texts=["Working..."],
                               buttons=[])
        assert not decide(dialog).safe_to_dismiss


class TestFingerprint:
    def test_same_content_same_key(self):
        a = vbe_runtime_dialog()
        b = vbe_runtime_dialog()
        assert fingerprint(a, 100) == fingerprint(b, 100)

    def test_handle_and_text_distinguish(self):
        a = vbe_runtime_dialog()
        assert fingerprint(a, 100) != fingerprint(a, 101)
