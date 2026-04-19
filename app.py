from __future__ import annotations

import csv
import os
import re
import shutil
import subprocess
import sys
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import openpyxl


def normalize(value: str) -> str:
    value = value or ""
    return re.sub(r"[^a-z0-9]", "", value.strip().lower())


def canonical_header(value: str) -> str:
    return normalize(value)


def split_first_name(first_name: str) -> list[str]:
    if not first_name:
        return []
    parts = [p.strip() for p in re.split(r"[/&]| and ", first_name, flags=re.IGNORECASE)]
    parts = [p for p in parts if p]
    return parts or [first_name.strip()]


def parse_yes_no(value: str) -> str:
    cleaned = (value or "").strip().lower()
    if cleaned in {"yes", "y", "true", "1"}:
        return "Yes"
    if cleaned in {"no", "n", "false", "0", ""}:
        return "No"
    return "No"


def safe_cell_text(value: str) -> str:
    text = (value or "").strip()
    if text.startswith(("=", "+", "-", "@")):
        raise ValueError("Values cannot start with =, +, -, or @.")
    return text


def safe_csv_value(value: str) -> str:
    text = str(value or "")
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


@dataclass
class MemberRecord:
    first_name: str
    last_name: str
    email: str
    membership_type: str
    membership_number: str
    includes_cart: str
    includes_range: str
    membership_amount_used: int
    sheet_name: str
    row_number: int

    @property
    def signature(self) -> tuple[str, int]:
        return (self.sheet_name, self.row_number)

    @property
    def display_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()


@dataclass
class SheetConfig:
    sheet_name: str
    header_row: int
    index_map: dict[str, int]


class MembershipDatabase:
    COL_ALIASES = {
        "first_name": {"firstname"},
        "last_name": {"lastname"},
        "email": {"email"},
        "membership_type": {"membershiptype"},
        "membership_number": {"membershipnumber", "membernumber"},
        "membership_amount_used": {"membershipamountused", "amountused", "membershipused"},
        "includes_cart": {"includescart", "incldescart", "includescart", "includecart"},
        "includes_range": {"includesrange", "includerange"},
    }

    def __init__(self) -> None:
        self.workbook = None
        self.workbook_path: Path | None = None
        self.loaded_mtime_ns: int | None = None
        self.sheet_configs: dict[str, SheetConfig] = {}
        self.records: list[MemberRecord] = []
        self.by_signature: dict[tuple[str, int], MemberRecord] = {}
        self.by_membership_number: dict[str, list[MemberRecord]] = {}
        self.by_email: dict[str, list[MemberRecord]] = {}
        self.by_name: dict[str, list[MemberRecord]] = {}

    def clear(self) -> None:
        self.sheet_configs.clear()
        self.records.clear()
        self.by_signature.clear()
        self.by_membership_number.clear()
        self.by_email.clear()
        self.by_name.clear()

    def load_excel(self, excel_path: Path) -> None:
        self.clear()
        self.workbook_path = excel_path
        self.loaded_mtime_ns = excel_path.stat().st_mtime_ns
        self.workbook = openpyxl.load_workbook(excel_path)

        for sheet_name in self.workbook.sheetnames:
            if "total" in sheet_name.lower():
                continue
            sheet = self.workbook[sheet_name]
            config = self._find_sheet_config(sheet_name, sheet)
            if not config:
                continue
            self.sheet_configs[sheet_name] = config
            self._load_sheet_records(sheet_name, sheet, config)

        self._rebuild_indexes()

    def _find_sheet_config(self, sheet_name: str, sheet) -> SheetConfig | None:
        header_row = 0
        raw_headers: list[str] = []
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=30, values_only=True), start=1):
            cells = [str(c).strip() if c is not None else "" for c in row]
            canon = {canonical_header(c) for c in cells if c}
            if "firstname" in canon and "lastname" in canon:
                header_row = row_idx
                raw_headers = cells
                break
        if not header_row:
            return None

        source = {canonical_header(v): idx for idx, v in enumerate(raw_headers, start=1) if v}
        index_map: dict[str, int] = {}
        for logical, aliases in self.COL_ALIASES.items():
            for alias in aliases:
                if alias in source:
                    index_map[logical] = source[alias]
                    break

        required = {"first_name", "last_name", "membership_type"}
        if not required.issubset(index_map.keys()):
            return None

        return SheetConfig(sheet_name=sheet_name, header_row=header_row, index_map=index_map)

    def _cell_text(self, row, idx: int | None) -> str:
        if not idx:
            return ""
        value = row[idx - 1]
        if value is None:
            return ""
        return str(value).strip()

    def _cell_int(self, row, idx: int | None) -> int:
        if not idx:
            return 0
        value = row[idx - 1]
        if value is None or str(value).strip() == "":
            return 0
        try:
            return int(float(str(value).strip()))
        except ValueError:
            return 0

    def _load_sheet_records(self, sheet_name: str, sheet, config: SheetConfig) -> None:
        for row_idx, row in enumerate(
            sheet.iter_rows(min_row=config.header_row + 1, max_row=sheet.max_row, values_only=True),
            start=config.header_row + 1,
        ):
            first_name = self._cell_text(row, config.index_map.get("first_name"))
            last_name = self._cell_text(row, config.index_map.get("last_name"))
            email = self._cell_text(row, config.index_map.get("email"))
            membership_number = self._cell_text(row, config.index_map.get("membership_number"))
            if not (first_name or last_name or email or membership_number):
                continue

            record = MemberRecord(
                first_name=first_name,
                last_name=last_name,
                email=email,
                membership_type=self._cell_text(row, config.index_map.get("membership_type")),
                membership_number=membership_number,
                includes_cart=self._cell_text(row, config.index_map.get("includes_cart")) or "No",
                includes_range=self._cell_text(row, config.index_map.get("includes_range")) or "No",
                membership_amount_used=self._cell_int(row, config.index_map.get("membership_amount_used")),
                sheet_name=sheet_name,
                row_number=row_idx,
            )
            self.records.append(record)
            self.by_signature[record.signature] = record

    def _add_to_index(self, index: dict[str, list[MemberRecord]], key: str, record: MemberRecord) -> None:
        if not key:
            return
        index.setdefault(key, []).append(record)

    def _rebuild_indexes(self) -> None:
        self.by_membership_number.clear()
        self.by_email.clear()
        self.by_name.clear()

        for record in self.records:
            self._add_to_index(self.by_membership_number, normalize(record.membership_number), record)
            self._add_to_index(self.by_email, normalize(record.email), record)
            first_names = split_first_name(record.first_name)
            last_name = record.last_name.strip()
            for first in first_names:
                self._add_to_index(self.by_name, normalize(f"{first} {last_name}"), record)
                self._add_to_index(self.by_name, normalize(f"{last_name} {first}"), record)
            self._add_to_index(self.by_name, normalize(record.display_name), record)

    def lookup(self, scan_text: str) -> list[MemberRecord]:
        raw = (scan_text or "").strip()
        if not raw:
            return []

        key = normalize(raw)
        matches: list[MemberRecord] = []
        if key in self.by_membership_number:
            matches.extend(self.by_membership_number[key])
        if not matches and key in self.by_email:
            matches.extend(self.by_email[key])
        if not matches:
            for candidate in self._name_candidates(raw):
                ckey = normalize(candidate)
                if ckey in self.by_name:
                    matches.extend(self.by_name[ckey])
        if not matches:
            matches = self._fallback_contains(raw)

        unique: list[MemberRecord] = []
        seen = set()
        for item in matches:
            if item.signature in seen:
                continue
            seen.add(item.signature)
            unique.append(item)
        return unique

    def _name_candidates(self, raw: str) -> list[str]:
        values = [raw]
        if "," in raw:
            parts = [p.strip() for p in raw.split(",") if p.strip()]
            if len(parts) >= 2:
                values.extend([f"{parts[1]} {parts[0]}", f"{parts[0]} {parts[1]}"])
        else:
            parts = [p for p in raw.split() if p]
            if len(parts) >= 2:
                values.extend([f"{parts[0]} {parts[-1]}", f"{parts[-1]} {parts[0]}"])
        return values

    def _fallback_contains(self, raw: str) -> list[MemberRecord]:
        needle = normalize(raw)
        if not needle:
            return []
        results = []
        for record in self.records:
            haystack = normalize(
                f"{record.first_name} {record.last_name} {record.email} {record.membership_number}"
            )
            if needle in haystack:
                results.append(record)
        return results

    def get_record(self, signature: tuple[str, int]) -> MemberRecord | None:
        return self.by_signature.get(signature)

    def _save_workbook_atomic(self) -> None:
        if not self.workbook_path or not self.workbook:
            raise RuntimeError("No workbook is loaded.")

        current_mtime = self.workbook_path.stat().st_mtime_ns
        if self.loaded_mtime_ns is not None and current_mtime != self.loaded_mtime_ns:
            raise RuntimeError(
                "Workbook changed on disk since load. Reload before saving to prevent data loss."
            )

        backup_dir = self.workbook_path.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = backup_dir / f"{self.workbook_path.stem}_{timestamp}{self.workbook_path.suffix}"
        shutil.copy2(self.workbook_path, backup_path)

        temp_path = self.workbook_path.with_name(f"{self.workbook_path.name}.tmp")
        try:
            self.workbook.save(temp_path)
            os.replace(temp_path, self.workbook_path)
        finally:
            if temp_path.exists():
                temp_path.unlink(missing_ok=True)

        self.loaded_mtime_ns = self.workbook_path.stat().st_mtime_ns

    def update_record(self, signature: tuple[str, int], updates: dict[str, str | int]) -> MemberRecord:
        record = self.get_record(signature)
        if not record:
            raise ValueError("Selected member record was not found.")

        config = self.sheet_configs[record.sheet_name]
        sheet = self.workbook[record.sheet_name]
        row = record.row_number

        def write_if_present(field: str, value):
            col = config.index_map.get(field)
            if col:
                sheet.cell(row=row, column=col).value = value

        if "first_name" in updates:
            record.first_name = safe_cell_text(str(updates["first_name"]))
            write_if_present("first_name", record.first_name)
        if "last_name" in updates:
            record.last_name = safe_cell_text(str(updates["last_name"]))
            write_if_present("last_name", record.last_name)
        if "email" in updates:
            record.email = safe_cell_text(str(updates["email"]))
            write_if_present("email", record.email)
        if "membership_type" in updates:
            record.membership_type = safe_cell_text(str(updates["membership_type"]))
            write_if_present("membership_type", record.membership_type)
        if "membership_number" in updates:
            record.membership_number = safe_cell_text(str(updates["membership_number"]))
            write_if_present("membership_number", record.membership_number)
        if "includes_cart" in updates:
            record.includes_cart = parse_yes_no(str(updates["includes_cart"]))
            write_if_present("includes_cart", record.includes_cart)
        if "includes_range" in updates:
            record.includes_range = parse_yes_no(str(updates["includes_range"]))
            write_if_present("includes_range", record.includes_range)
        if "membership_amount_used" in updates:
            try:
                amount = int(str(updates["membership_amount_used"]).strip())
            except ValueError as exc:
                raise ValueError("Membership Amount Used must be a whole number.") from exc
            if amount < 0:
                raise ValueError("Membership Amount Used cannot be negative.")
            record.membership_amount_used = amount
            write_if_present("membership_amount_used", amount)

        self._save_workbook_atomic()
        self._rebuild_indexes()
        return record


@dataclass
class ScanEvent:
    signature: tuple[str, int]
    previous_amount: int
    new_amount: int
    scan_value: str
    scanned_at: str


class MembershipApp:
    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("Membership Card Verifier")
        self.root.geometry("1280x820")
        self.root.minsize(860, 620)

        self.db = MembershipDatabase()
        self.excel_path_var = tk.StringVar()
        self.scan_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Load an Excel file to begin.")
        self.last_scan_text = ""
        self.scan_in_progress = False
        self.current_matches: list[MemberRecord] = []
        self.current_selection: tuple[str, int] | None = None
        self.last_scan_events: list[ScanEvent] = []
        self.log_path = self._default_log_path()
        self.selected_match_id = tk.StringVar(value="")
        self.match_cards: list[ctk.CTkFrame] = []

        self.detail_vars = {
            "first_name": tk.StringVar(),
            "last_name": tk.StringVar(),
            "email": tk.StringVar(),
            "membership_type": tk.StringVar(),
            "membership_number": tk.StringVar(),
            "includes_cart": tk.StringVar(value="No"),
            "includes_range": tk.StringVar(value="No"),
            "membership_amount_used": tk.StringVar(value="0"),
        }

        self._build_styles()
        self._build_ui()
        self._ensure_log_file()
        self._auto_load_default_file()

    def _build_styles(self) -> None:
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

    def _build_ui(self) -> None:
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        container = ctk.CTkFrame(self.root, fg_color="#F2F7FF", corner_radius=0)
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(2, weight=1)
        container.grid_columnconfigure(0, weight=1)

        head = ctk.CTkFrame(container, fg_color="#246BFD", corner_radius=0, height=110)
        head.grid(row=0, column=0, sticky="ew")
        head.grid_columnconfigure(0, weight=1)
        head.grid_propagate(False)
        ctk.CTkLabel(
            head,
            text="Membership Scanner",
            text_color="#FFFFFF",
            font=ctk.CTkFont(size=34, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=26, pady=(16, 0))
        ctk.CTkLabel(
            head,
            text="Scan, verify, update, and audit membership usage in real time",
            text_color="#D9E7FF",
            font=ctk.CTkFont(size=15),
        ).grid(row=1, column=0, sticky="w", padx=26, pady=(2, 0))

        controls_wrap = ctk.CTkFrame(container, fg_color="transparent")
        controls_wrap.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 8))
        controls_wrap.grid_columnconfigure(0, weight=2)
        controls_wrap.grid_columnconfigure(1, weight=1)

        chip_row = ctk.CTkFrame(controls_wrap, fg_color="transparent")
        chip_row.grid(row=0, column=0, sticky="w", pady=(0, 10))
        ctk.CTkLabel(
            chip_row,
            text="Live Verification",
            fg_color="#DFF6E8",
            text_color="#0A6A3E",
            corner_radius=20,
            font=ctk.CTkFont(size=12, weight="bold"),
            padx=14,
            pady=6,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkLabel(
            chip_row,
            text="Auto Usage Tracking",
            fg_color="#E1ECFF",
            text_color="#244F9C",
            corner_radius=20,
            font=ctk.CTkFont(size=12, weight="bold"),
            padx=14,
            pady=6,
        ).pack(side=tk.LEFT)

        file_card, file_content = self._make_card(controls_wrap, "Database")
        file_card.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        file_content.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(file_content, textvariable=self.excel_path_var, height=42, corner_radius=14).grid(
            row=0, column=0, sticky="ew"
        )
        ctk.CTkButton(file_content, text="Browse", command=self.choose_file, height=42, corner_radius=14).grid(
            row=0, column=1, padx=(8, 0)
        )
        ctk.CTkButton(file_content, text="Reload", command=self.reload_database, height=42, corner_radius=14).grid(
            row=0, column=2, padx=(8, 0)
        )
        ctk.CTkButton(file_content, text="Open Logs", command=self.open_log_folder, height=42, corner_radius=14).grid(
            row=0, column=3, padx=(8, 0)
        )

        scan_card, scan_content = self._make_card(controls_wrap, "Scanner")
        scan_card.grid(row=2, column=0, sticky="nsew", padx=(0, 8))
        scan_content.grid_columnconfigure(0, weight=1)
        self.scan_entry = ctk.CTkEntry(
            scan_content,
            textvariable=self.scan_var,
            height=50,
            corner_radius=14,
            font=ctk.CTkFont(size=19, weight="bold"),
            placeholder_text="Scan membership card barcode...",
        )
        self.scan_entry.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.scan_entry.bind("<Return>", self.on_scan_enter)
        self.scan_entry.bind("<KeyPress>", self.on_scan_keypress)

        action_row = ctk.CTkFrame(scan_content, fg_color="transparent")
        action_row.grid(row=1, column=0, sticky="ew")
        for col in range(4):
            action_row.grid_columnconfigure(col, weight=1)
        ctk.CTkButton(
            action_row,
            text="Verify",
            command=self.verify_scan,
            height=42,
            corner_radius=14,
            fg_color="#FF9F45",
            hover_color="#EF8C2D",
            text_color="#163046",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ctk.CTkButton(action_row, text="Confirm Selected", command=self.confirm_selected_scan, height=42, corner_radius=14).grid(
            row=0, column=1, sticky="ew", padx=6
        )
        ctk.CTkButton(action_row, text="Undo Last Scan", command=self.undo_last_scan, height=42, corner_radius=14).grid(
            row=0, column=2, sticky="ew", padx=6
        )
        ctk.CTkButton(action_row, text="Clear", command=self.clear_scan, height=42, corner_radius=14).grid(
            row=0, column=3, sticky="ew", padx=(6, 0)
        )

        status_card, status_content = self._make_card(controls_wrap, "Status")
        status_card.grid(row=2, column=1, sticky="nsew", padx=(8, 0))
        self.status_label = ctk.CTkLabel(
            status_content,
            textvariable=self.status_var,
            anchor="w",
            text_color="#123B58",
            justify="left",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.status_label.pack(fill=tk.BOTH, expand=True, pady=(2, 4))

        split = ctk.CTkFrame(container, fg_color="transparent")
        split.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        split.grid_columnconfigure(0, weight=3)
        split.grid_columnconfigure(1, weight=2)
        split.grid_rowconfigure(0, weight=1)

        left_card, left_content = self._make_card(split, "Match Results")
        right_card, right_content = self._make_card(split, "Edit Member")
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._build_result_table(left_content)
        self._build_detail_editor(right_content)

    def _make_card(self, parent, title: str) -> tuple[tk.Frame, tk.Frame]:
        shell = ctk.CTkFrame(parent, fg_color="#FFFFFF", corner_radius=24, border_width=1, border_color="#DCEBFF")
        card = shell
        ctk.CTkLabel(
            card,
            text=title,
            text_color="#1A4A8C",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(anchor="w", padx=16, pady=(14, 8))
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))
        return shell, content

    def _build_result_table(self, parent) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        self.results_scroll = ctk.CTkScrollableFrame(parent, fg_color="#F8FBFF", corner_radius=16)
        self.results_scroll.grid(row=0, column=0, sticky="nsew")
        self.results_scroll.grid_columnconfigure(0, weight=1)

        self.empty_results_label = ctk.CTkLabel(
            self.results_scroll,
            text="No scans yet. Scan a card to see matches.",
            text_color="#5C7596",
            font=ctk.CTkFont(size=15),
        )
        self.empty_results_label.grid(row=0, column=0, padx=10, pady=24)

    def _build_detail_editor(self, parent) -> None:
        field_defs = [
            ("first_name", "First Name", "entry"),
            ("last_name", "Last Name", "entry"),
            ("email", "Email", "entry"),
            ("membership_type", "Membership Type", "entry"),
            ("membership_number", "Membership Number", "entry"),
            ("includes_cart", "Includes Cart", "yesno"),
            ("includes_range", "Includes Range", "yesno"),
            ("membership_amount_used", "Membership Amount Used", "entry"),
        ]

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        editor_scroll = ctk.CTkScrollableFrame(parent, fg_color="#F8FBFF", corner_radius=16)
        editor_scroll.grid(row=0, column=0, sticky="nsew")
        editor_scroll.grid_columnconfigure(1, weight=1)

        for row_idx, (key, label, field_type) in enumerate(field_defs):
            ctk.CTkLabel(
                editor_scroll,
                text=label,
                text_color="#5E7393",
                font=ctk.CTkFont(size=13, weight="bold"),
            ).grid(row=row_idx, column=0, sticky="w", pady=(6, 2), padx=(0, 8))
            if field_type == "yesno":
                widget = ctk.CTkOptionMenu(
                    editor_scroll,
                    variable=self.detail_vars[key],
                    values=["Yes", "No"],
                    height=40,
                    corner_radius=12,
                )
            else:
                widget = ctk.CTkEntry(editor_scroll, textvariable=self.detail_vars[key], height=40, corner_radius=12)
            widget.grid(row=row_idx, column=1, sticky="ew", pady=(6, 2))

        btn_row = ctk.CTkFrame(editor_scroll, fg_color="transparent")
        btn_row.grid(row=len(field_defs), column=0, columnspan=2, sticky="ew", pady=(16, 0))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            btn_row,
            text="Save Changes",
            command=self.save_member_changes,
            height=42,
            corner_radius=14,
            fg_color="#2E89FF",
            hover_color="#1F6FE3",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ctk.CTkButton(btn_row, text="Revert", command=self.refresh_editor_from_selection, height=42, corner_radius=14).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

    def _default_startup_dir(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path.cwd()

    def _app_data_dir(self) -> Path:
        if sys.platform.startswith("win"):
            base = Path.home() / "AppData" / "Local"
        elif sys.platform == "darwin":
            base = Path.home() / "Library" / "Application Support"
        else:
            base = Path.home() / ".local" / "share"
        return base / "MembershipVerifier"

    def _default_log_path(self) -> Path:
        return self._app_data_dir() / "scan_history.csv"

    def _ensure_log_file(self) -> None:
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        if self.log_path.exists():
            return
        with self.log_path.open("w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(
                [
                    "timestamp",
                    "event",
                    "scan_value",
                    "result",
                    "match_count",
                    "member_name",
                    "membership_number",
                    "membership_amount_used",
                    "sheet",
                    "row",
                    "excel_file",
                ]
            )

    def _append_scan_log(
        self,
        event: str,
        scan_value: str,
        result: str,
        matches: list[MemberRecord],
        target: MemberRecord | None = None,
    ) -> None:
        row_target = target if target else (matches[0] if len(matches) == 1 else None)
        row = [
            datetime.now().isoformat(timespec="seconds"),
            safe_csv_value(event),
            safe_csv_value(scan_value),
            safe_csv_value(result),
            len(matches),
            safe_csv_value(row_target.display_name if row_target else ""),
            safe_csv_value(row_target.membership_number if row_target else ""),
            row_target.membership_amount_used if row_target else "",
            safe_csv_value(row_target.sheet_name if row_target else ""),
            row_target.row_number if row_target else "",
            safe_csv_value(self.excel_path_var.get().strip()),
        ]
        with self.log_path.open("a", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(row)

    def _set_status(self, message: str, level: str) -> None:
        self.status_var.set(message)
        color = {"ok": "#0E7A46", "warn": "#9A6200", "error": "#A32B2B"}.get(level, "#123B58")
        self.status_label.configure(text_color=color)

    def _auto_load_default_file(self) -> None:
        search_dir = self._default_startup_dir()
        files = sorted(search_dir.glob("*.xlsx"))
        if files:
            self.excel_path_var.set(str(files[0]))
            self.reload_database()

    def open_log_folder(self) -> None:
        folder = self.log_path.parent
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception as exc:
            messagebox.showerror("Open Folder", f"Could not open log folder:\n{exc}")

    def choose_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Membership Excel File",
            filetypes=[("Excel Workbook", "*.xlsx *.xlsm"), ("All Files", "*.*")],
        )
        if not selected:
            return
        self.excel_path_var.set(selected)
        self.reload_database()

    def reload_database(self) -> None:
        path_text = self.excel_path_var.get().strip()
        if not path_text:
            self._set_status("No file selected.", "warn")
            return

        excel_path = Path(path_text)
        if not excel_path.exists():
            self._set_status("Excel file not found.", "error")
            return

        try:
            self.db.load_excel(excel_path)
        except Exception as exc:
            self._set_status(f"Failed to load database: {exc}", "error")
            messagebox.showerror("Load Error", f"Could not load workbook:\n{exc}")
            return

        self._clear_tree()
        self.current_matches = []
        self.current_selection = None
        self.scan_in_progress = False
        self.selected_match_id.set("")
        self._clear_detail_editor()
        self._set_status(f"Loaded {len(self.db.records)} membership records.", "ok")
        self.scan_entry.focus_set()

    def on_scan_enter(self, _event: tk.Event) -> None:
        self.verify_scan()

    def on_scan_keypress(self, event: tk.Event) -> None:
        if event.keysym == "Return":
            return
        if self.scan_in_progress:
            self.scan_var.set("")
            self._clear_tree()
            self.current_matches = []
            self.current_selection = None
            self.selected_match_id.set("")
            self._clear_detail_editor()
            self.scan_in_progress = False

    def verify_scan(self) -> None:
        text = self.scan_var.get().strip()
        if not text:
            self._set_status("No scan value provided.", "warn")
            return
        if not self.db.records:
            self._set_status("Load an Excel membership file first.", "warn")
            return

        self.last_scan_text = text
        matches = self.db.lookup(text)
        self.current_matches = matches
        self.scan_in_progress = True
        self._show_matches(matches)

        if not matches:
            self._append_scan_log("scan", text, "not_found", [])
            self._set_status(f"No active membership found for: {text}", "error")
            return

        if len(matches) == 1:
            self._apply_scan_for_record(matches[0], text)
            return

        self._append_scan_log("scan", text, "multiple_matches", matches)
        self._set_status(
            f"Multiple matches found ({len(matches)}). Select one and click Confirm Selected.",
            "warn",
        )

    def _show_matches(self, matches: list[MemberRecord]) -> None:
        self._clear_tree()

        if not matches:
            self.empty_results_label = ctk.CTkLabel(
                self.results_scroll,
                text="No matching membership found.",
                text_color="#7C8EA9",
                font=ctk.CTkFont(size=15),
            )
            self.empty_results_label.grid(row=0, column=0, padx=10, pady=24)
            return

        if hasattr(self, "empty_results_label"):
            self.empty_results_label.destroy()

        for idx, member in enumerate(matches):
            match_id = f"{member.sheet_name}|{member.row_number}"
            card = ctk.CTkFrame(self.results_scroll, fg_color="#FFFFFF", corner_radius=16, border_width=1, border_color="#D9E6FB")
            card.grid(row=idx, column=0, sticky="ew", padx=8, pady=6)
            card.grid_columnconfigure(0, weight=1)

            rb = ctk.CTkRadioButton(
                card,
                text=member.display_name,
                variable=self.selected_match_id,
                value=match_id,
                command=self.on_tree_select,
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color="#143C70",
            )
            rb.grid(row=0, column=0, sticky="w", padx=14, pady=(10, 4))

            details = (
                f"{member.membership_type}   |   #{member.membership_number or 'N/A'}   |   "
                f"Cart: {parse_yes_no(member.includes_cart)}   |   Range: {parse_yes_no(member.includes_range)}   |   "
                f"Used: {member.membership_amount_used}"
            )
            ctk.CTkLabel(
                card,
                text=details,
                text_color="#4F6787",
                font=ctk.CTkFont(size=13),
            ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 4))

            ctk.CTkLabel(
                card,
                text=f"{member.email or 'No email'}  |  Sheet: {member.sheet_name}  |  Row: {member.row_number}",
                text_color="#7A8DAA",
                font=ctk.CTkFont(size=12),
            ).grid(row=2, column=0, sticky="w", padx=16, pady=(0, 10))

            self.match_cards.append(card)

        if len(matches) == 1:
            self.current_selection = matches[0].signature
            self.selected_match_id.set(f"{matches[0].sheet_name}|{matches[0].row_number}")
            self.refresh_editor_from_selection()

    def confirm_selected_scan(self) -> None:
        record = self._selected_record()
        if not record:
            self._set_status("Select one member to confirm this scan.", "warn")
            return
        if not self.last_scan_text:
            self._set_status("No recent scan to confirm.", "warn")
            return
        self._apply_scan_for_record(record, self.last_scan_text)

    def _apply_scan_for_record(self, record: MemberRecord, scan_value: str) -> None:
        previous = record.membership_amount_used
        new_value = previous + 1
        try:
            updated = self.db.update_record(record.signature, {"membership_amount_used": new_value})
        except Exception as exc:
            self._set_status(f"Failed to update usage count: {exc}", "error")
            messagebox.showerror("Save Error", f"Could not save usage update:\n{exc}")
            return

        self.last_scan_events.append(
            ScanEvent(
                signature=updated.signature,
                previous_amount=previous,
                new_amount=new_value,
                scan_value=scan_value,
                scanned_at=datetime.now().isoformat(timespec="seconds"),
            )
        )

        self.current_matches = [self.db.get_record(r.signature) or r for r in self.current_matches]
        self._show_matches(self.current_matches)
        self.current_selection = updated.signature
        self.refresh_editor_from_selection()
        self._append_scan_log("scan", scan_value, "verified", [updated], target=updated)
        self._set_status(
            f"Verified: {updated.display_name} | Usage count is now {updated.membership_amount_used}.",
            "ok",
        )
        self.scan_var.set("")
        self.scan_entry.focus_set()

    def undo_last_scan(self) -> None:
        if not self.last_scan_events:
            self._set_status("No scan to undo.", "warn")
            return

        event = self.last_scan_events.pop()
        record = self.db.get_record(event.signature)
        if not record:
            self._set_status("Undo failed: record no longer exists.", "error")
            return

        try:
            updated = self.db.update_record(record.signature, {"membership_amount_used": event.previous_amount})
        except Exception as exc:
            self._set_status(f"Undo failed: {exc}", "error")
            messagebox.showerror("Undo Error", f"Could not undo scan:\n{exc}")
            return

        self.current_matches = [self.db.get_record(r.signature) or r for r in self.current_matches]
        self._show_matches(self.current_matches)
        self.current_selection = updated.signature
        self.refresh_editor_from_selection()
        self._append_scan_log("undo", event.scan_value, "scan_reverted", [updated], target=updated)
        self._set_status(
            f"Undo complete: {updated.display_name} usage reverted to {updated.membership_amount_used}.",
            "ok",
        )

    def on_tree_select(self, _event=None) -> None:
        self.refresh_editor_from_selection()

    def _selected_record(self) -> MemberRecord | None:
        selected = self.selected_match_id.get().strip()
        if not selected:
            return None
        if "|" not in selected:
            return None
        sheet_name, row_text = selected.split("|", 1)
        try:
            row = int(row_text)
        except ValueError:
            return None
        signature = (sheet_name, row)
        self.current_selection = signature
        return self.db.get_record(signature)

    def refresh_editor_from_selection(self) -> None:
        record = self._selected_record()
        if not record:
            if self.current_selection:
                record = self.db.get_record(self.current_selection)
            if not record:
                self._clear_detail_editor()
                return
        self.detail_vars["first_name"].set(record.first_name)
        self.detail_vars["last_name"].set(record.last_name)
        self.detail_vars["email"].set(record.email)
        self.detail_vars["membership_type"].set(record.membership_type)
        self.detail_vars["membership_number"].set(record.membership_number)
        self.detail_vars["includes_cart"].set(parse_yes_no(record.includes_cart))
        self.detail_vars["includes_range"].set(parse_yes_no(record.includes_range))
        self.detail_vars["membership_amount_used"].set(str(record.membership_amount_used))

    def _clear_detail_editor(self) -> None:
        self.detail_vars["first_name"].set("")
        self.detail_vars["last_name"].set("")
        self.detail_vars["email"].set("")
        self.detail_vars["membership_type"].set("")
        self.detail_vars["membership_number"].set("")
        self.detail_vars["includes_cart"].set("No")
        self.detail_vars["includes_range"].set("No")
        self.detail_vars["membership_amount_used"].set("0")

    def save_member_changes(self) -> None:
        if not self.current_selection:
            self._set_status("Select a member before saving changes.", "warn")
            return

        updates = {
            "first_name": self.detail_vars["first_name"].get(),
            "last_name": self.detail_vars["last_name"].get(),
            "email": self.detail_vars["email"].get(),
            "membership_type": self.detail_vars["membership_type"].get(),
            "membership_number": self.detail_vars["membership_number"].get(),
            "includes_cart": self.detail_vars["includes_cart"].get(),
            "includes_range": self.detail_vars["includes_range"].get(),
            "membership_amount_used": self.detail_vars["membership_amount_used"].get(),
        }

        try:
            updated = self.db.update_record(self.current_selection, updates)
        except Exception as exc:
            self._set_status(f"Save failed: {exc}", "error")
            messagebox.showerror("Save Error", f"Could not save member details:\n{exc}")
            return

        self.current_matches = [self.db.get_record(r.signature) or r for r in self.current_matches]
        if not self.current_matches:
            self.current_matches = [updated]
        self._show_matches(self.current_matches)
        self.current_selection = updated.signature
        self.refresh_editor_from_selection()
        self._append_scan_log("edit", "", "member_updated", [updated], target=updated)
        self._set_status(f"Saved changes for {updated.display_name}.", "ok")

    def clear_scan(self) -> None:
        self.scan_var.set("")
        self._clear_tree()
        self._clear_detail_editor()
        self.current_matches = []
        self.current_selection = None
        self.selected_match_id.set("")
        self.scan_in_progress = False
        self._set_status("Ready for next scan.", "ok")
        self.scan_entry.focus_set()

    def _clear_tree(self) -> None:
        for card in self.match_cards:
            card.destroy()
        self.match_cards.clear()
        if hasattr(self, "empty_results_label"):
            self.empty_results_label.destroy()


def main() -> None:
    root = ctk.CTk()
    app = MembershipApp(root)
    app.scan_entry.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
