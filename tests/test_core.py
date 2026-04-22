from pathlib import Path

from app import (
    AuditTrail,
    app_data_dir_for_platform,
    layout_mode_for_width,
    safe_cell_text,
    safe_csv_value,
)


def test_layout_mode_boundaries() -> None:
    assert layout_mode_for_width(445) == "compact"
    assert layout_mode_for_width(700) == "compact"
    assert layout_mode_for_width(701) == "wide"


def test_app_data_dir_platform_paths() -> None:
    home = Path("/home/tester")
    assert app_data_dir_for_platform("linux", home) == home / ".local" / "share" / "MembershipVerifier"
    assert app_data_dir_for_platform("darwin", home) == home / "Library" / "Application Support" / "MembershipVerifier"
    assert app_data_dir_for_platform("win32", home) == home / "AppData" / "Local" / "MembershipVerifier"


def test_safe_cell_text_blocks_formula_prefix() -> None:
    for value in ["=SUM(A1:A2)", "+1", "-1", "@cmd"]:
        try:
            safe_cell_text(value)
            assert False, "Expected ValueError"
        except ValueError:
            pass


def test_safe_csv_value_prefixes_formula_values() -> None:
    assert safe_csv_value("=HELLO") == "'=HELLO"
    assert safe_csv_value("normal") == "normal"


def test_audit_chain_detects_tampering(tmp_path: Path) -> None:
    audit = AuditTrail(tmp_path)
    audit.log("event_one", {"value": "1"})
    audit.log("event_two", {"value": "2"})
    assert audit.verify_chain() is True

    data = audit.audit_file.read_text(encoding="utf-8").splitlines()
    data[1] = data[1].replace("event_two", "event_hacked")
    audit.audit_file.write_text("\n".join(data) + "\n", encoding="utf-8")
    assert audit.verify_chain() is False


def test_audit_files_are_hidden_like(tmp_path: Path) -> None:
    audit = AuditTrail(tmp_path)
    assert audit.audit_file.name.startswith(".")
    assert audit.seed_file.name.startswith(".")
