"""FIRST Global 2025 Team Video Slotter GUI.

This module launches a simple Tkinter application with two tabs for
future expansion.
"""

from __future__ import annotations

import importlib.util

if importlib.util.find_spec("tkinter") is None:  # pragma: no cover - import-time guard
    raise ModuleNotFoundError(
        "tkinter is not available in this Python installation. "
        "On macOS with Homebrew Python, install it with 'brew install python-tk@3.13'."
    )

import itertools
import json
import re
import threading
import tkinter as tk
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

try:  # pragma: no cover - optional dependency
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ModuleNotFoundError:  # pragma: no cover - handled at runtime
    Credentials = None  # type: ignore[assignment]
    InstalledAppFlow = None  # type: ignore[assignment]

try:  # pragma: no cover - optional dependency
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ModuleNotFoundError:  # pragma: no cover - handled at runtime
    build = None  # type: ignore[assignment]
    HttpError = None  # type: ignore[assignment]

try:  # pragma: no cover - optional dependency
    from google.auth.transport.requests import Request
except ModuleNotFoundError:  # pragma: no cover - handled at runtime
    Request = None  # type: ignore[assignment]


FLAG_EMOJI_PATTERN = re.compile(r"^\s*([\U0001F1E6-\U0001F1FF]{2})")
PLACEHOLDER_CODE_PATTERN = re.compile(
    r"^TEAM VIDEO PLACEHOLDER ([A-Z]{3})\b",
    re.IGNORECASE,
)


def _get_widget_background(widget: tk.Misc, fallback: tk.Misc) -> str:
    """Return a usable background color for a widget.

    ttk widgets such as ``ttk.Frame`` don't expose a ``-background`` option,
    which results in a ``TclError`` when ``cget("background")`` is called.
    This helper gracefully handles that case by falling back to the top-level
    window's background or, ultimately, the fallback widget's current
    background color.
    """

    for candidate in (widget, widget.winfo_toplevel(), fallback):
        try:
            background = str(candidate.cget("background"))
        except tk.TclError:
            continue
        if background:
            return background
    return str(fallback.cget("background"))


class ApplicationConsole:
    """Display log messages emitted by tools across the application."""

    def __init__(self, parent: tk.Misc) -> None:
        self.parent = parent
        self._text_widget: Optional[tk.Text] = None

    def render(self, row: int) -> None:
        frame = ttk.Frame(self.parent, padding=(16, 0, 16, 16))
        frame.grid(row=row, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Console", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        text_widget = tk.Text(
            frame,
            height=8,
            wrap="word",
            borderwidth=1,
            relief="solid",
            cursor="xterm",
        )
        background = _get_widget_background(frame, text_widget)
        text_widget.configure(background=background, state="disabled")
        text_widget.grid(row=1, column=0, sticky="nsew")

        def _on_key(event: tk.Event) -> str | None:  # type: ignore[name-defined]
            control_pressed = bool(event.state & 0x4)
            command_pressed = bool(event.state & 0x200000)
            if control_pressed or command_pressed:
                if event.keysym.lower() in {"a", "c"}:
                    return None
                return "break"
            if event.keysym in {
                "Left",
                "Right",
                "Up",
                "Down",
                "Home",
                "End",
                "Next",
                "Prior",
                "Shift_L",
                "Shift_R",
                "Control_L",
                "Control_R",
                "Meta_L",
                "Meta_R",
            }:
                return None
            if event.keysym in {"BackSpace", "Delete", "Return", "KP_Enter"}:
                return "break"
            if event.char and event.char.isprintable():
                return "break"
            return None

        text_widget.bind("<Key>", _on_key)
        frame.rowconfigure(1, weight=1)
        self._text_widget = text_widget

    def log(self, message: str) -> None:
        if not message:
            return
        widget = self._text_widget
        if widget is None:
            return

        def _append() -> None:
            widget.configure(state="normal")
            if widget.index("end-1c") != "1.0":
                widget.insert(tk.END, "\n")
            widget.insert(tk.END, message)
            widget.see(tk.END)
            widget.configure(state="disabled")

        widget.after(0, _append)


def create_main_window() -> tk.Tk:
    """Create and configure the main application window."""
    root = tk.Tk()
    root.title("FIRST Global 2025 Team Video Slotter")
    root.geometry("800x600")

    # Configure a grid layout so the notebook expands with the window.
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    title_label = ttk.Label(
        root,
        text="FIRST Global 2025 Team Video Slotter",
        font=("Helvetica", 18, "bold"),
        anchor="center",
    )
    title_label.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="ew")

    notebook = ttk.Notebook(root)
    notebook.grid(row=1, column=0, padx=16, pady=(16, 8), sticky="nsew")

    team_videos_frame = ttk.Frame(notebook)
    config_frame = ttk.Frame(notebook)
    tools_frame = ttk.Frame(notebook)

    for frame in (team_videos_frame, config_frame, tools_frame):
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

    notebook.add(team_videos_frame, text="Team Videos")
    notebook.add(config_frame, text="Config")
    notebook.add(tools_frame, text="Tools")

    placeholder_videos_label = ttk.Label(
        team_videos_frame,
        text="Team Videos content will go here.",
        anchor="center",
    )
    placeholder_videos_label.grid(padx=12, pady=12)

    console = ApplicationConsole(root)
    root.rowconfigure(2, weight=0)
    console.render(row=2)
    (
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
    ) = build_config_tab(config_frame, console)
    build_tools_tab(
        tools_frame,
        console,
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
    )

    return root


def build_config_tab(
    parent: ttk.Frame, console: ApplicationConsole
) -> Tuple[
    "GoogleDriveCredentialsManager",
    "ROSDocumentLoaderUI",
    "MatchScheduleImporterUI",
]:
    """Populate the Config tab with shared application settings."""

    parent.columnconfigure(0, weight=1)

    load_drive_frame = ttk.Frame(parent, padding=(12, 12, 12, 6))
    load_drive_frame.grid(row=0, column=0, sticky="nsew")
    load_drive_frame.columnconfigure(1, weight=1)

    ttk.Label(
        load_drive_frame,
        text="Load Google Drive Credentials",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    credentials_manager = GoogleDriveCredentialsManager(load_drive_frame, console)
    credentials_manager.render(row=1)

    separator_one = ttk.Separator(parent, orient="horizontal")
    separator_one.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

    ros_document_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    ros_document_frame.grid(row=2, column=0, sticky="nsew")
    ros_document_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_document_frame,
        text="Load ROS Document",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ros_document_loader = ROSDocumentLoaderUI(
        ros_document_frame, console, credentials_manager=credentials_manager
    )
    ros_document_loader.render(row=1)

    separator_two = ttk.Separator(parent, orient="horizontal")
    separator_two.grid(row=3, column=0, sticky="ew", padx=12, pady=6)

    match_schedule_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    match_schedule_frame.grid(row=4, column=0, sticky="nsew")
    match_schedule_frame.columnconfigure(1, weight=1)

    ttk.Label(
        match_schedule_frame,
        text="Import Match Schedule",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    match_schedule_importer = MatchScheduleImporterUI(match_schedule_frame, console)
    match_schedule_importer.render(row=1)

    return credentials_manager, ros_document_loader, match_schedule_importer


def build_tools_tab(
    parent: ttk.Frame,
    console: ApplicationConsole,
    credentials_manager: "GoogleDriveCredentialsManager",
    ros_document_loader: "ROSDocumentLoaderUI",
    match_schedule_importer: "MatchScheduleImporterUI",
) -> None:
    """Populate the Tools tab with utilities that operate on ROS documents."""

    parent.columnconfigure(0, weight=1)

    ros_frame = ttk.Frame(parent, padding=(12, 12, 12, 6))
    ros_frame.grid(row=0, column=0, sticky="nsew")
    ros_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_frame,
        text="ROS Placeholder Generator",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ROSPlaceholderGeneratorUI(ros_frame, credentials_manager, ros_document_loader, console).render(
        row=1
    )

    separator = ttk.Separator(parent, orient="horizontal")
    separator.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

    match_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    match_frame.grid(row=2, column=0, sticky="nsew")
    match_frame.columnconfigure(1, weight=1)

    ttk.Label(
        match_frame,
        text="Match Number Generator",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    MatchNumberGeneratorUI(
        match_frame,
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
        console,
    ).render(row=1)


class MatchScheduleImporterUI:
    """Import match schedules from JSON files within the Config tab."""

    FIELD_OPTIONS = tuple(str(number) for number in range(1, 6))

    def __init__(self, parent: ttk.Frame, console: ApplicationConsole) -> None:
        self.parent = parent
        self.console = console
        self.file_path_var = tk.StringVar()
        self.field_number_var = tk.StringVar(value=self.FIELD_OPTIONS[0])
        self._status_var = tk.StringVar()
        self._matches: List[Dict[str, Any]] = []
        self._imported_field: Optional[int] = None
        self._matches_by_date: Dict[str, List[Dict[str, Any]]] = {}

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Selected file:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        file_display = ttk.Entry(
            self.parent,
            textvariable=self.file_path_var,
            state="readonly",
        )
        file_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(self.parent, text="Field:").grid(
            row=row + 1, column=0, sticky="w", pady=(8, 0)
        )

        field_selector = ttk.Combobox(
            self.parent,
            textvariable=self.field_number_var,
            values=self.FIELD_OPTIONS,
            state="readonly",
            width=5,
        )
        field_selector.grid(row=row + 1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        import_button = ttk.Button(
            self.parent,
            text="Import Match Schedule",
            command=self.import_schedule,
        )
        import_button.grid(row=row + 2, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 3, column=0, columnspan=2, sticky="w")

        self._set_status("No match schedule imported yet.", log=False)

    def import_schedule(self) -> None:
        field_number = self._get_selected_field()
        if field_number is None:
            self._set_status("Select a valid field number between 1 and 5.")
            return

        filename = filedialog.askopenfilename(
            parent=self.parent,
            title="Select match schedule JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*")],
        )
        if not filename:
            return

        try:
            with open(filename, "r", encoding="utf-8") as stream:
                data = json.load(stream)
        except FileNotFoundError:
            self._set_status("Selected file could not be found.")
            return
        except json.JSONDecodeError as exc:
            self._set_status(f"Failed to parse JSON: {exc}")
            return
        except OSError as exc:
            self._set_status(f"Unable to read file: {exc}")
            return

        matches = data.get("matches") if isinstance(data, dict) else None
        if not isinstance(matches, list):
            self._set_status("JSON file does not contain a 'matches' list.")
            return

        self._matches = matches
        self._imported_field = field_number
        self._matches_by_date = self._group_matches_by_date(matches, field_number)
        self.file_path_var.set(filename)

        field_counts = self._count_field_matches(matches, field_number)
        total_matches = len(matches)
        field_total = sum(field_counts.values())

        self._log_field_counts(field_counts, field_number)

        if field_counts:
            summary = (
                f"Imported {total_matches} matches. {field_total} occur on Field {field_number}."
            )
        else:
            summary = (
                f"Imported {total_matches} matches. No matches found on Field {field_number}."
            )
        self._set_status(summary)

    def _get_selected_field(self) -> Optional[int]:
        try:
            field_number = int(self.field_number_var.get())
        except (TypeError, ValueError):
            return None
        if field_number < 1 or field_number > 5:
            return None
        return field_number

    def _count_field_matches(
        self, matches: Sequence[Dict[str, Any]], field_number: int
    ) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for match in matches:
            field = match.get("field")
            if field != field_number:
                continue
            scheduled_time = match.get("scheduledTime")
            if not isinstance(scheduled_time, str):
                continue
            date_str = self._extract_date(scheduled_time)
            if not date_str:
                continue
            counts[date_str] = counts.get(date_str, 0) + 1
        return dict(sorted(counts.items()))

    def _group_matches_by_date(
        self, matches: Sequence[Dict[str, Any]], field_number: int
    ) -> Dict[str, List[Dict[str, Any]]]:
        grouped: Dict[str, List[Dict[str, Any]]] = {}
        for match in matches:
            if match.get("field") != field_number:
                continue
            scheduled_time = match.get("scheduledTime")
            if not isinstance(scheduled_time, str):
                continue
            date_str = self._extract_date(scheduled_time)
            if not date_str:
                continue
            grouped.setdefault(date_str, []).append(match)

        for date, items in grouped.items():
            items.sort(key=self._match_sort_key)

        return {date: list(grouped[date]) for date in sorted(grouped)}

    def _extract_date(self, timestamp: str) -> Optional[str]:
        timestamp = timestamp.strip()
        if not timestamp:
            return None
        normalized = timestamp.replace("Z", "+00:00")
        try:
            return datetime.fromisoformat(normalized).date().isoformat()
        except ValueError:
            if "T" in timestamp:
                candidate = timestamp.split("T", 1)[0]
                if re.fullmatch(r"\d{4}-\d{2}-\d{2}", candidate):
                    return candidate
            return None

    def _parse_datetime(self, timestamp: Any) -> Optional[datetime]:
        if not isinstance(timestamp, str):
            return None
        normalized = timestamp.strip()
        if not normalized:
            return None
        normalized = normalized.replace("Z", "+00:00")
        try:
            return datetime.fromisoformat(normalized)
        except ValueError:
            return None

    def _match_sort_key(self, match: Dict[str, Any]) -> Tuple[Any, Any, Any]:
        scheduled_dt = self._parse_datetime(match.get("scheduledTime"))
        if scheduled_dt is not None and scheduled_dt.tzinfo is not None:
            scheduled_dt = scheduled_dt.astimezone(timezone.utc).replace(tzinfo=None)
        if scheduled_dt is not None:
            scheduled_key: Tuple[Any, ...] = (
                0,
                scheduled_dt.toordinal(),
                scheduled_dt.hour,
                scheduled_dt.minute,
                scheduled_dt.second,
                scheduled_dt.microsecond,
            )
        else:
            scheduled_key = (1, 0, 0, 0, 0, 0, 0)

        match_number = self.extract_match_number(match)
        fallback = (
            match.get("matchKey")
            or match.get("description")
            or match.get("name")
            or ""
        )
        return (
            scheduled_key,
            match_number if match_number is not None else float("inf"),
            fallback,
        )

    def _log_field_counts(self, counts: Dict[str, int], field_number: int) -> None:
        if not counts:
            self.console.log(
                f"[Match Schedule Importer] No matches on Field {field_number} were found."
            )
            return
        for date, total in counts.items():
            message = f"{date} has {total} matches on Field {field_number}"
            self.console.log(f"[Match Schedule Importer] {message}")

    def _set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[Match Schedule Importer] {message}")

    def has_loaded_schedule(self) -> bool:
        return bool(self._matches)

    def get_imported_field_number(self) -> Optional[int]:
        return self._imported_field

    def get_matches_by_date_for_selected_field(self) -> List[Tuple[str, List[Dict[str, Any]]]]:
        return [
            (date, list(matches))
            for date, matches in sorted(self._matches_by_date.items())
        ]

    def extract_match_number(self, match: Dict[str, Any]) -> Optional[int]:
        for key in (
            "matchNumber",
            "match_number",
            "matchNumberDisplay",
            "matchnumber",
            "matchNo",
            "id",
        ):
            number = self._coerce_match_number(match.get(key))
            if number is not None:
                return number

        for fallback_key in ("matchKey", "match"):
            number = self._coerce_match_number(match.get(fallback_key))
            if number is not None:
                return number

        return None

    def describe_match(self, match: Mapping[str, Any]) -> str:
        """Return a human-readable summary of a match entry for diagnostics."""

        def _format_value(value: Any) -> Optional[str]:
            if value is None or isinstance(value, bool):
                return None
            if isinstance(value, (int, float)):
                if isinstance(value, float) and not value.is_integer():
                    return str(value)
                return str(int(value)) if isinstance(value, float) else str(value)
            if isinstance(value, str):
                text = value.strip()
                if text:
                    return text
                return None
            return None

        details: List[str] = []
        for label in ("matchKey", "description", "name", "id"):
            formatted = _format_value(match.get(label))
            if formatted:
                details.append(f"{label}={formatted}")
                break

        scheduled = _format_value(match.get("scheduledTime"))
        if scheduled:
            details.append(f"scheduledTime={scheduled}")

        field_value = match.get("field")
        field_formatted = _format_value(field_value)
        if field_formatted:
            details.append(f"field={field_formatted}")

        if details:
            return ", ".join(details)

        keys = ", ".join(sorted(map(str, match.keys())))
        if keys:
            return f"available keys: {keys}"
        return "no additional details available"

    def _coerce_match_number(self, value: Any) -> Optional[int]:
        if isinstance(value, bool):
            return None
        if isinstance(value, int):
            return value
        if isinstance(value, float) and value.is_integer():
            return int(value)
        if isinstance(value, str):
            stripped = value.strip()
            if not stripped:
                return None
            if stripped.isdigit():
                return int(stripped)
            digits = re.findall(r"\d+", stripped)
            if digits:
                return int(digits[-1])
        return None

class GoogleDriveCredentialsManager:
    """Handle selection and loading of Google Drive OAuth credentials."""

    SCOPES = [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    STORAGE_PATH = Path.home() / ".fgc_team_filler" / "google_credentials.json"
    USER_POLL_INTERVAL_MS = 60_000

    def __init__(self, parent: ttk.Frame, console: ApplicationConsole) -> None:
        self.parent = parent
        self.console = console
        self.credentials_path_var = tk.StringVar()
        self.logged_in_user_var = tk.StringVar(value="No credentials loaded")
        self._credentials: Optional[Credentials] = None
        self._user_poll_after_id: Optional[str] = None
        self._last_user_message: Optional[str] = None
        self._credential_listeners: List[Callable[[Optional[Credentials]], None]] = []

    @property
    def credentials(self) -> Optional[Credentials]:
        return self._credentials

    def get_valid_credentials(
        self, *, log_status: bool = True
    ) -> Tuple[Optional[Credentials], Optional[str]]:
        """Return credentials that are refreshed and valid, or an error message."""

        credentials = self._credentials
        if credentials is None:
            return None, None

        if credentials.expired and getattr(credentials, "refresh_token", None):
            if Request is None:
                message = (
                    "Stored credentials have expired and google-auth-transport-requests is missing. "
                    "Install it with 'pip install google-auth'."
                )
                if log_status:
                    self.set_status(message)
                return None, message
            try:
                credentials.refresh(Request())
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to refresh credentials: {exc}"[:500]
                if log_status:
                    self.set_status(message)
                return None, message
            else:
                self._credentials = credentials
                self._persist_credentials(credentials)
                self._notify_credentials_listeners()

        if not credentials.valid:
            message = "Stored credentials are invalid. Please re-authorize."
            if log_status:
                self.set_status(message)
            return None, message

        return credentials, None

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Credentials file:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        path_display = ttk.Entry(
            self.parent,
            textvariable=self.credentials_path_var,
            state="readonly",
        )
        path_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        select_button = ttk.Button(
            self.parent,
            text="Select credentials.json",
            command=self.select_credentials_file,
        )
        select_button.grid(row=row + 1, column=0, sticky="w", pady=8)

        authorize_button = ttk.Button(
            self.parent,
            text="Authorize",
            command=self.authorize,
        )
        authorize_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        ttk.Label(self.parent, text="Currently Logged-in User:").grid(
            row=row + 2, column=0, sticky="w"
        )

        user_display = ttk.Entry(
            self.parent,
            textvariable=self.logged_in_user_var,
            state="readonly",
        )
        user_display.grid(row=row + 2, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))

        self.set_status("No credentials loaded.", log=False)
        self._load_persisted_credentials()
        self._schedule_user_poll(0)

    def add_credentials_listener(
        self, callback: Callable[[Optional[Credentials]], None]
    ) -> None:
        self._credential_listeners.append(callback)
        try:
            callback(self._credentials)
        except Exception:
            pass

    def _notify_credentials_listeners(self) -> None:
        for listener in list(self._credential_listeners):
            try:
                listener(self._credentials)
            except Exception:
                continue

    def select_credentials_file(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.parent,
            title="Select Google Drive credentials.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*")],
        )
        if filename:
            self.credentials_path_var.set(filename)
            self.set_status("Credentials file selected. Click Authorize to continue.")

    def authorize(self) -> None:
        credentials_path = self.credentials_path_var.get().strip()
        if not credentials_path:
            self.set_status("Please select a credentials.json file first.")
            return

        if InstalledAppFlow is None:
            self.set_status(
                "google-auth-oauthlib is not installed. Install it with 'pip install google-auth-oauthlib'."
            )
            return

        self.set_status("Starting OAuth flow...")

        def _run_flow() -> None:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_path, self.SCOPES
                )
                credentials = flow.run_local_server(port=0, open_browser=True)
            except Exception as exc:  # pragma: no cover - user driven
                self.parent.after(
                    0,
                    lambda: self.set_status(f"Authorization failed: {exc}"[:500]),
                )
                return

            def _on_success() -> None:
                self._credentials = credentials
                message = self._persist_credentials(credentials)
                if message is None:
                    message = "Authorization successful. Credentials are loaded in memory."
                self.set_status(message)
                self._schedule_user_poll(0)
                self._notify_credentials_listeners()

            self.parent.after(0, _on_success)

        threading.Thread(target=_run_flow, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        if log:
            self.console.log(f"[Drive Credentials] {message}")

    def _load_persisted_credentials(self) -> None:
        if Credentials is None:
            return

        storage_path = self.STORAGE_PATH
        if not storage_path.exists():
            return

        try:
            credentials = Credentials.from_authorized_user_file(
                str(storage_path), self.SCOPES
            )
        except Exception as exc:  # pragma: no cover - file parsing
            self.set_status(f"Failed to load saved credentials: {exc}"[:500])
            return

        if not credentials.valid and getattr(credentials, "refresh_token", None):
            if Request is None:
                self.set_status(
                    "Saved credentials are expired but google-auth is missing the Requests transport."
                )
                return
            try:
                credentials.refresh(Request())
            except Exception as exc:  # pragma: no cover - network interaction
                self.set_status(f"Failed to refresh saved credentials: {exc}"[:500])
                return

        if credentials.valid:
            self._credentials = credentials
            self.credentials_path_var.set(str(storage_path))
            self.set_status("Loaded saved credentials. You're ready to go.")
            self._persist_credentials(credentials)
            self._schedule_user_poll(0)
            self._notify_credentials_listeners()
        else:
            self.set_status("Saved credentials are invalid. Please re-authorize.")

    def _persist_credentials(self, credentials: Credentials) -> Optional[str]:
        if Credentials is None:
            return None

        storage_path = self.STORAGE_PATH
        try:
            storage_path.parent.mkdir(parents=True, exist_ok=True)
            storage_path.write_text(credentials.to_json())
        except Exception as exc:  # pragma: no cover - filesystem issues
            return (
                "Authorization succeeded, but saving credentials failed: "
                f"{exc}"[:500]
            )
        self.credentials_path_var.set(str(storage_path))
        return f"Authorization successful. Credentials saved to {storage_path}."

    def _schedule_user_poll(self, delay_ms: int) -> None:
        if self._user_poll_after_id is not None:
            try:
                self.parent.after_cancel(self._user_poll_after_id)
            except Exception:
                pass
        self._user_poll_after_id = self.parent.after(delay_ms, self._poll_logged_in_user)

    def _poll_logged_in_user(self) -> None:
        credentials, error = self.get_valid_credentials(log_status=False)
        if credentials is None:
            if error:
                self._set_logged_in_user("Unavailable")
                self._log_user_lookup_message(error)
            else:
                self._set_logged_in_user("No credentials loaded")
                self._last_user_message = None
            self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)
            return

        if build is None:
            message = (
                "google-api-python-client is not installed. Install it to fetch the logged-in user."
            )
            self._set_logged_in_user("Library missing")
            self._log_user_lookup_message(message)
            self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)
            return

        threading.Thread(
            target=self._fetch_logged_in_user, args=(credentials,), daemon=True
        ).start()

    def _fetch_logged_in_user(self, credentials: Credentials) -> None:
        try:
            service = build("drive", "v3", credentials=credentials, cache_discovery=False)
            about = service.about().get(fields="user(displayName,emailAddress)").execute()
        except Exception as exc:  # pragma: no cover - network interaction
            message = f"Failed to fetch logged-in user: {exc}"[:500]
            self.parent.after(0, lambda: self._handle_user_lookup_failure(message))
            return

        user_info = about.get("user", {}) if isinstance(about, dict) else {}
        display_name = user_info.get("displayName")
        email = user_info.get("emailAddress")
        self.parent.after(
            0,
            lambda: self._handle_user_lookup_success(
                str(display_name) if display_name else "",
                str(email) if email else "",
            ),
        )

    def _handle_user_lookup_success(self, display_name: str, email: str) -> None:
        parts = [part for part in (display_name.strip(), email.strip()) if part]
        display_text = " • ".join(parts) if parts else "Unknown user"
        self._set_logged_in_user(display_text)
        self._log_user_lookup_message(f"Logged in as {display_text}.")
        self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)

    def _handle_user_lookup_failure(self, message: str) -> None:
        self._set_logged_in_user("Lookup failed")
        self._log_user_lookup_message(message)
        self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)

    def _set_logged_in_user(self, value: str) -> None:
        self.logged_in_user_var.set(value)

    def _log_user_lookup_message(self, message: str) -> None:
        if not message:
            return
        if message == self._last_user_message:
            return
        self.console.log(f"[Drive Credentials] {message}")
        self._last_user_message = message


class ROSDocumentLoaderUI:
    """Allow the user to select and persist the ROS spreadsheet URL."""

    STORAGE_PATH = Path.home() / ".fgc_team_filler" / "ros_document_url.txt"

    def __init__(
        self,
        parent: ttk.Frame,
        console: ApplicationConsole,
        *,
        credentials_manager: Optional["GoogleDriveCredentialsManager"] = None,
    ) -> None:
        self.parent = parent
        self.console = console
        self._credentials_manager = credentials_manager
        self.sheet_url_var = tk.StringVar(value=self._load_saved_url())
        self.document_name_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._listeners: List[Callable[[str], None]] = []
        self._name_listeners: List[Callable[[str], None]] = []
        self._resolved_names: Dict[str, str] = {}
        self._lookup_in_progress: Set[str] = set()

        self.sheet_url_var.trace_add("write", self._on_url_var_changed)
        self._update_document_name(self.sheet_url_var.get())

        if self._credentials_manager is not None:
            self._credentials_manager.add_credentials_listener(
                self._on_credentials_changed
            )

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Google Sheets URL:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        entry = ttk.Entry(self.parent, textvariable=self.sheet_url_var)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(self.parent, text="File name:").grid(
            row=row + 1, column=0, sticky="w", pady=(4, 0)
        )

        name_display = ttk.Entry(
            self.parent,
            textvariable=self.document_name_var,
            state="readonly",
        )
        name_display.grid(row=row + 1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))

        save_button = ttk.Button(
            self.parent,
            text="Save Document URL",
            command=self.save_document_url,
        )
        save_button.grid(row=row + 2, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 3, column=0, columnspan=2, sticky="w")

        initial_url = self.sheet_url_var.get().strip()
        if initial_url:
            self._set_status("Loaded saved ROS document URL.", log=False)
        else:
            self._set_status("Enter a Google Sheets link to use with ROS tools.", log=False)

    def add_listener(self, callback: Callable[[str], None]) -> None:
        self._listeners.append(callback)

    def add_name_listener(self, callback: Callable[[str], None]) -> None:
        self._name_listeners.append(callback)

    def get_document_url(self) -> str:
        return self.sheet_url_var.get().strip()

    def get_document_name(self) -> str:
        return self.document_name_var.get().strip()

    def save_document_url(self) -> None:
        url = self.sheet_url_var.get().strip()
        if not url:
            self._set_status("Please enter a Google Sheets link before saving.")
            return

        spreadsheet_id = extract_spreadsheet_id(url)
        if not spreadsheet_id:
            self._set_status("Unable to determine spreadsheet ID from the provided link.")
            return

        storage_path = self.STORAGE_PATH
        try:
            storage_path.parent.mkdir(parents=True, exist_ok=True)
            storage_path.write_text(url)
        except Exception as exc:  # pragma: no cover - filesystem issues
            self._set_status(f"Failed to save document URL: {exc}"[:500])
            return

        self._set_status("ROS document URL saved.")
        self._update_document_name(url)
        for listener in list(self._listeners):
            try:
                listener(url)
            except Exception:
                continue

    def _load_saved_url(self) -> str:
        storage_path = self.STORAGE_PATH
        if not storage_path.exists():
            return ""
        try:
            return storage_path.read_text().strip()
        except Exception:
            return ""

    def _set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[ROS Document] {message}")

    def _on_url_var_changed(self, *_: Any) -> None:
        self._update_document_name(self.sheet_url_var.get())

    def set_document_name(self, name: str) -> None:
        text = (name or "").strip()
        if text:
            spreadsheet_id = extract_spreadsheet_id(self.sheet_url_var.get().strip())
            if spreadsheet_id:
                self._resolved_names[spreadsheet_id] = text
        display = text if text else "Unknown"
        self._set_document_name_display(display)

    def _update_document_name(self, url: str) -> None:
        url = (url or "").strip()
        if not url:
            self._set_document_name_display("Not set")
            return

        spreadsheet_id = extract_spreadsheet_id(url)
        if spreadsheet_id:
            cached_name = self._resolved_names.get(spreadsheet_id)
            if cached_name:
                self._set_document_name_display(cached_name)
                return

        name = derive_document_name(url)
        if not name:
            if spreadsheet_id:
                name = f"Spreadsheet {spreadsheet_id}"
        display = name or "Unknown"

        self._set_document_name_display(display)
        if spreadsheet_id:
            self._maybe_lookup_remote_title(spreadsheet_id)

    def _set_document_name_display(self, display: str) -> None:
        self.document_name_var.set(display)
        for listener in list(self._name_listeners):
            try:
                listener(display)
            except Exception:
                continue

    def _on_credentials_changed(self, credentials: Optional[Credentials]) -> None:
        if credentials is None:
            return
        url = self.sheet_url_var.get().strip()
        spreadsheet_id = extract_spreadsheet_id(url)
        if spreadsheet_id:
            self._maybe_lookup_remote_title(spreadsheet_id, credentials=credentials)

    def _maybe_lookup_remote_title(
        self,
        spreadsheet_id: str,
        *,
        credentials: Optional[Credentials] = None,
    ) -> None:
        if not spreadsheet_id:
            return
        if spreadsheet_id in self._resolved_names:
            return
        if spreadsheet_id in self._lookup_in_progress:
            return
        if build is None:
            return

        manager = self._credentials_manager
        if credentials is None:
            if manager is None:
                return
            credentials, error = manager.get_valid_credentials(log_status=False)
            if credentials is None:
                if error:
                    self.console.log(f"[ROS Document] {error}")
                return
        self._lookup_in_progress.add(spreadsheet_id)

        def _worker() -> None:
            try:
                spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to fetch spreadsheet details: {exc}"[:500]
                self.parent.after(
                    0,
                    lambda: self._finalize_remote_lookup(
                        spreadsheet_id, None, error_message=message
                    ),
                )
                return

            title = str(spreadsheet.get("properties", {}).get("title", "")).strip()
            self.parent.after(
                0,
                lambda: self._finalize_remote_lookup(spreadsheet_id, title or None),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _finalize_remote_lookup(
        self,
        spreadsheet_id: str,
        title: Optional[str],
        *,
        error_message: Optional[str] = None,
    ) -> None:
        self._lookup_in_progress.discard(spreadsheet_id)
        if error_message:
            self.console.log(f"[ROS Document] {error_message}")
        if not title:
            return

        self._resolved_names[spreadsheet_id] = title
        current_id = extract_spreadsheet_id(self.sheet_url_var.get().strip())
        if current_id == spreadsheet_id:
            self._set_document_name_display(title)


class ROSPlaceholderGeneratorUI:
    """User interface wrapper for the ROS placeholder generator tool."""

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: GoogleDriveCredentialsManager,
        document_loader: ROSDocumentLoaderUI,
        console: ApplicationConsole,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.document_loader = document_loader
        self.console = console
        self.current_document_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._default_status = "Load a ROS document and Google credentials to begin."

        self.document_loader.add_listener(self._on_document_url_changed)
        self.document_loader.add_name_listener(self._on_document_name_changed)
        self._on_document_url_changed(self.document_loader.get_document_url())
        self._on_document_name_changed(self.document_loader.get_document_name())

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Current ROS Document:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        document_display = ttk.Entry(
            self.parent,
            textvariable=self.current_document_var,
            state="readonly",
        )
        document_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        generate_button = ttk.Button(
            self.parent,
            text="Generate Placeholders",
            command=self.generate_placeholders,
        )
        generate_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 2, column=0, columnspan=2, sticky="w")

        self.set_status(self._default_status, log=False)

    def _on_document_url_changed(self, url: str) -> None:
        if not url:
            self.set_status("Save a ROS document URL before generating placeholders.", log=False)

    def _on_document_name_changed(self, name: str) -> None:
        display_value = name if name else "Not set"
        self.current_document_var.set(display_value)

    def generate_placeholders(self) -> None:
        document_url = self.document_loader.get_document_url()
        if not document_url:
            self.set_status("Save a ROS document URL before generating placeholders.")
            return

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status("Load Google Drive credentials before running this tool.")
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return

        spreadsheet_id = extract_spreadsheet_id(document_url)
        if not spreadsheet_id:
            self.set_status("Unable to determine spreadsheet ID from the saved document URL.")
            return

        self.set_status("Contacting Google Sheets API...")

        def _worker() -> None:
            try:
                report, diagnostics, spreadsheet_title = generate_placeholders_for_sheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            self.parent.after(
                0,
                lambda: self._handle_placeholder_success(
                    report, diagnostics, spreadsheet_title
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[ROS Placeholder Generator] {message}")

    def _handle_placeholder_success(
        self,
        report: Dict[str, List[str]],
        diagnostics: Sequence[str],
        spreadsheet_title: str,
    ) -> None:
        if spreadsheet_title:
            self.document_loader.set_document_name(spreadsheet_title)

        console_message = format_report(report, diagnostics)

        if report:
            total_updates = sum(len(entries) for entries in report.values())
            sheet_count = len(report)
            cell_word = "cell" if total_updates == 1 else "cells"
            sheet_word = "sheet" if sheet_count == 1 else "sheets"
            status_message = (
                "Placeholder updates applied to "
                f"{total_updates} {cell_word} across {sheet_count} {sheet_word}."
            )
            breakdowns = []
            for sheet, entries in report.items():
                entry_count = len(entries)
                entry_cell_word = "cell" if entry_count == 1 else "cells"
                breakdowns.append(
                    f"{entry_count} {entry_cell_word} updated on sheet ({sheet})"
                )
            if breakdowns:
                status_message += " " + ", ".join(breakdowns)
        else:
            status_message = "No matching placeholders were found."

        self.console.log(f"[ROS Placeholder Generator] {console_message}")
        self.set_status(status_message, log=False)


class MatchNumberGeneratorUI:
    """User interface for numbering RANKING MATCH entries in ROS documents."""

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: GoogleDriveCredentialsManager,
        document_loader: ROSDocumentLoaderUI,
        match_schedule_importer: MatchScheduleImporterUI,
        console: ApplicationConsole,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.document_loader = document_loader
        self.match_schedule_importer = match_schedule_importer
        self.console = console
        self.current_document_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._default_status = "Load a ROS document and Google credentials to begin."

        self.document_loader.add_listener(self._on_document_url_changed)
        self.document_loader.add_name_listener(self._on_document_name_changed)
        self._on_document_url_changed(self.document_loader.get_document_url())
        self._on_document_name_changed(self.document_loader.get_document_name())

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Current ROS Document:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        document_display = ttk.Entry(
            self.parent,
            textvariable=self.current_document_var,
            state="readonly",
        )
        document_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        generate_button = ttk.Button(
            self.parent,
            text="Generate Match Numbers",
            command=self.generate_match_numbers,
        )
        generate_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 2, column=0, columnspan=2, sticky="w")

        self.set_status(self._default_status, log=False)

    def _on_document_url_changed(self, url: str) -> None:
        if not url:
            self.set_status("Save a ROS document URL before generating match numbers.", log=False)

    def _on_document_name_changed(self, name: str) -> None:
        display_value = name if name else "Not set"
        self.current_document_var.set(display_value)

    def generate_match_numbers(self) -> None:
        document_url = self.document_loader.get_document_url()
        if not document_url:
            self.set_status("Save a ROS document URL before generating match numbers.")
            return

        if not self.match_schedule_importer.has_loaded_schedule():
            messagebox.showerror(
                "Match Schedule Required",
                "Import a match schedule JSON before generating match numbers.",
                parent=self.parent.winfo_toplevel(),
            )
            self.set_status("Import a match schedule JSON before generating match numbers.")
            return

        schedule_by_date = self.match_schedule_importer.get_matches_by_date_for_selected_field()
        if not schedule_by_date:
            field_number = self.match_schedule_importer.get_imported_field_number()
            if field_number is None:
                schedule_message = "The imported schedule does not include matches for the selected field."
            else:
                schedule_message = (
                    f"The imported schedule does not include matches for Field {field_number}."
                )
            messagebox.showerror(
                "Match Schedule Missing Matches",
                schedule_message,
                parent=self.parent.winfo_toplevel(),
            )
            self.set_status(schedule_message)
            return

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status("Load Google Drive credentials before running this tool.")
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return

        spreadsheet_id = extract_spreadsheet_id(document_url)
        if not spreadsheet_id:
            self.set_status("Unable to determine spreadsheet ID from the saved document URL.")
            return

        self.set_status("Contacting Google Sheets API...")

        def _worker() -> None:
            try:
                spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            sheet_entries, existing_numbers_found = inspect_ranking_match_numbers(
                spreadsheet
            )
            spreadsheet_title = str(
                spreadsheet.get("properties", {}).get("title", "")
            )
            self.parent.after(
                0,
                lambda: self._handle_match_analysis(
                    credentials,
                    spreadsheet_id,
                    sheet_entries,
                    existing_numbers_found,
                    spreadsheet_title,
                    schedule_by_date,
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[Match Number Generator] {message}")

    def _handle_match_analysis(
        self,
        credentials: Credentials,
        spreadsheet_id: str,
        sheet_entries: Sequence[Dict[str, Any]],
        existing_numbers_found: bool,
        spreadsheet_title: str,
        schedule_by_date: Sequence[Tuple[str, Sequence[Dict[str, Any]]]],
    ) -> None:
        if spreadsheet_title:
            self.document_loader.set_document_name(spreadsheet_title)

        total_matches = sum(len(entry.get("matches", [])) for entry in sheet_entries)
        diagnostics = self._gather_analysis_diagnostics(sheet_entries)

        if total_matches == 0:
            if diagnostics:
                self._log_diagnostics(diagnostics)
            self.set_status("No RANKING MATCH cells were updated.")
            return

        match_numbers_by_sheet, mismatch_error, alignment_notes = (
            self._derive_match_number_assignments(sheet_entries, schedule_by_date)
        )
        if mismatch_error:
            if diagnostics:
                self._log_diagnostics(diagnostics)
            messagebox.showerror(
                "Match Schedule Mismatch",
                mismatch_error,
                parent=self.parent.winfo_toplevel(),
            )
            self.console.log(f"[Match Number Generator] {mismatch_error}")
            self.set_status("Match schedule does not match the ROS document.")
            return

        for note in alignment_notes:
            self.console.log(f"[Match Number Generator] {note}")

        if existing_numbers_found:
            response = messagebox.askyesno(
                "Existing Match Numbers Found",
                "Existing Match Numbers found, do you want to renumber?",
                parent=self.parent.winfo_toplevel(),
            )
            if not response:
                if diagnostics:
                    self._log_diagnostics(diagnostics)
                self.console.log(
                    "[Match Number Generator] Existing match numbers left unchanged by user request."
                )
                self.set_status("Existing match numbers left unchanged.")
                return
            renumber_all = True
        else:
            renumber_all = False

        self._apply_match_numbers(
            credentials,
            spreadsheet_id,
            sheet_entries,
            renumber_all=renumber_all,
            initial_diagnostics=diagnostics,
            match_numbers_by_sheet=match_numbers_by_sheet,
        )

    def _derive_match_number_assignments(
        self,
        sheet_entries: Sequence[Dict[str, Any]],
        schedule_by_date: Sequence[Tuple[str, Sequence[Dict[str, Any]]]],
    ) -> Tuple[Dict[str, List[int]], Optional[str], List[str]]:
        sheet_match_entries = [
            entry for entry in sheet_entries if entry.get("matches")
        ]
        assignments: Dict[str, List[int]] = {}
        alignment_notes: List[str] = []

        schedule_count = len(schedule_by_date)
        sheet_count = len(sheet_match_entries)

        if schedule_count > sheet_count:
            extra_date, extra_matches = schedule_by_date[sheet_count]
            message = (
                f"The number of matches ({len(extra_matches)}) in the schedule for {extra_date} "
                "doesn't match the number of RANKING MATCH slots in "
                "N/A (no sheet available)."
            )
            return {}, message, []

        if sheet_count > schedule_count:
            extra_sheet = sheet_match_entries[schedule_count]
            sheet_name = str(extra_sheet.get("title", "Untitled"))
            message = (
                "The number of matches (0) in the schedule for N/A doesn't match the number "
                f"of RANKING MATCH slots in {sheet_name}."
            )
            return {}, message, []

        for (date, matches_for_date), sheet_entry in zip(
            schedule_by_date, sheet_match_entries
        ):
            sheet_name = str(sheet_entry.get("title", "Untitled"))
            sheet_matches: Sequence[Dict[str, Any]] = sheet_entry.get("matches", [])
            schedule_total = len(matches_for_date)
            slot_total = len(sheet_matches)
            alignment_notes.append(
                f"Verified {schedule_total} matches for {date} align with sheet {sheet_name}."
            )

            if schedule_total != slot_total:
                message = (
                    f"The number of matches ({schedule_total}) in the schedule for {date} "
                    f"doesn't match the number of RANKING MATCH slots in {sheet_name}."
                )
                return {}, message, []

            numbers: List[int] = []
            for match in matches_for_date:
                match_number = self.match_schedule_importer.extract_match_number(match)
                if match_number is None:
                    match_details = self.match_schedule_importer.describe_match(match)
                    message = (
                        "The imported schedule entry is missing a match number. "
                        f"Date: {date}. Match details: {match_details}. "
                        "Ensure this entry includes an 'id' or 'matchNumber' value."
                    )
                    return {}, message, []
                numbers.append(match_number)

            assignments[sheet_name] = numbers

        return assignments, None, alignment_notes

    def _apply_match_numbers(
        self,
        credentials: Credentials,
        spreadsheet_id: str,
        sheet_entries: Sequence[Dict[str, Any]],
        *,
        renumber_all: bool,
        initial_diagnostics: Sequence[str],
        match_numbers_by_sheet: Dict[str, Sequence[int]],
    ) -> None:
        self.set_status("Applying ranking match numbers...")

        def _worker() -> None:
            try:
                report, diagnostics = apply_ranking_match_number_updates(
                    credentials,
                    spreadsheet_id,
                    sheet_entries,
                    renumber_all=renumber_all,
                    initial_diagnostics=initial_diagnostics,
                    match_numbers_by_sheet=match_numbers_by_sheet,
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            self.parent.after(
                0,
                lambda: self._handle_match_success(report, diagnostics),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_match_success(
        self, report: Dict[str, List[str]], diagnostics: Sequence[str]
    ) -> None:
        if diagnostics:
            self._log_diagnostics(diagnostics)

        console_message = format_report(
            report,
            (),
            success_header="Ranking match numbers applied:",
            empty_message="No RANKING MATCH cells were updated.",
        )
        if report:
            total_updates = sum(len(entries) for entries in report.values())
            sheet_count = len(report)
            sheet_word = "sheet" if sheet_count == 1 else "sheets"
            status_message = (
                "Ranking match numbers applied to "
                f"{total_updates} cell{'s' if total_updates != 1 else ''} across "
                f"{sheet_count} {sheet_word}."
            )
            breakdowns = []
            for sheet, entries in report.items():
                entry_count = len(entries)
                cell_word = "cell" if entry_count == 1 else "cells"
                breakdowns.append(
                    f"{entry_count} {cell_word} applied on sheet ({sheet})"
                )
            if breakdowns:
                status_message += " " + ", ".join(breakdowns)
        else:
            status_message = "No RANKING MATCH cells were updated."

        self.console.log(f"[Match Number Generator] {console_message}")
        self.set_status(status_message, log=False)

    def _gather_analysis_diagnostics(
        self, sheet_entries: Sequence[Dict[str, Any]]
    ) -> List[str]:
        diagnostics: List[str] = []
        for entry in sheet_entries:
            diagnostics.extend(entry.get("diagnostics", []))
        return diagnostics

    def _log_diagnostics(self, diagnostics: Sequence[str]) -> None:
        for message in diagnostics:
            self.console.log(f"[Match Number Generator] {message}")

def extract_spreadsheet_id(url: str) -> str:
    """Extract the spreadsheet ID from a Google Sheets URL."""

    parts = url.split("/")
    if "spreadsheets" in parts:
        try:
            index = parts.index("d")
        except ValueError:
            pass
        else:
            if index + 1 < len(parts):
                return parts[index + 1]

    if "spreadsheets" in url and "#gid" in url:
        # Fall back to query parsing
        import urllib.parse

        parsed = urllib.parse.urlparse(url)
        path_parts = parsed.path.split("/")
        if "d" in path_parts:
            d_index = path_parts.index("d")
            if d_index + 1 < len(path_parts):
                return path_parts[d_index + 1]

    return ""


def derive_document_name(url: str) -> str:
    """Best-effort extraction of a human-friendly name from a Sheets URL."""

    if not url:
        return ""

    import urllib.parse

    parsed = urllib.parse.urlparse(url)
    path_parts = [part for part in parsed.path.split("/") if part]

    for part in reversed(path_parts):
        lowered = part.lower()
        if lowered in {"edit", "view", "copy"}:
            continue
        if lowered == "d":
            continue
        if lowered in {"spreadsheets", "file"}:
            continue
        return urllib.parse.unquote(part)

    query = urllib.parse.parse_qs(parsed.query)
    for key in ("name", "title", "resourcekey"):
        values = query.get(key)
        if values:
            return urllib.parse.unquote(values[0])

    fragment_query = urllib.parse.parse_qs(parsed.fragment)
    for key in ("name", "title"):
        values = fragment_query.get(key)
        if values:
            return urllib.parse.unquote(values[0])

    return ""


def _fetch_spreadsheet(
    credentials: Credentials, spreadsheet_id: str
) -> Tuple[Dict[str, Any], bool, Any]:
    """Retrieve spreadsheet metadata and return it with theme support info."""

    service = build("sheets", "v4", credentials=credentials)
    spreadsheets_resource = service.spreadsheets()
    base_fields = (
        "sheets("
        "properties(title,sheetId,index),"
        "data(rowData(values("
        "userEnteredValue,"
        "formattedValue,"
        "effectiveFormat(backgroundColor,backgroundColorStyle),"
        "userEnteredFormat(backgroundColor,backgroundColorStyle)"
        "))))"
    )
    fields_with_theme = base_fields + ",spreadsheetTheme"
    get_kwargs: Dict[str, Any] = {
        "spreadsheetId": spreadsheet_id,
        "includeGridData": True,
        "fields": fields_with_theme,
    }
    theme_supported = True
    if _method_supports_parameter(spreadsheets_resource.get, "supportsAllDrives"):
        get_kwargs["supportsAllDrives"] = True

    try:
        spreadsheet = spreadsheets_resource.get(**get_kwargs).execute()
    except Exception as exc:
        if HttpError is not None and isinstance(exc, HttpError):
            status = getattr(exc, "status_code", None) or getattr(exc, "resp", None)
            if hasattr(status, "status"):
                status = getattr(status, "status")
            message = str(exc)
            if status == 400 and "not supported for this document" in message.lower():
                raise ValueError(
                    "The selected file is stored in Office Compatibility mode. "
                    "Open it in Google Sheets and use File → Save as Google Sheets, "
                    "then try again."
                ) from exc
            if status == 400 and "spreadsheettheme" in message.lower():
                fallback_kwargs = dict(get_kwargs, fields=base_fields)
                spreadsheet = spreadsheets_resource.get(**fallback_kwargs).execute()
                theme_supported = False
            else:
                raise
        else:
            raise

    return spreadsheet, theme_supported, service


def generate_placeholders_for_sheet(
    credentials: Credentials, spreadsheet_id: str
) -> Tuple[Dict[str, List[str]], List[str], str]:
    """Scan the spreadsheet and replace placeholder cells.

    Returns the applied updates together with diagnostics describing how the
    spreadsheet was inspected (for example, which column contained the ``TASK``
    header and sample values that began with a flag emoji), and the spreadsheet
    title.
    """

    spreadsheet, _theme_supported, service = _fetch_spreadsheet(
        credentials, spreadsheet_id
    )
    spreadsheet_title = str(
        spreadsheet.get("properties", {}).get("title", "")
    )

    updates: Dict[str, List[Tuple[str, str]]] = {}
    diagnostics: List[str] = [
        "Placeholder detection looks for rows whose TASK value begins with an emoji flag.",
    ]
    data_updates: List[Dict[str, Sequence[Sequence[str]]]] = []

    for sheet in spreadsheet.get("sheets", []):
        properties = sheet.get("properties", {})
        title = properties.get("title", "Untitled")
        sheet_index = properties.get("index", 0)
        letter_prefix = chr(ord("A") + sheet_index)

        sheet_data = sheet.get("data", [])
        if not sheet_data:
            continue

        task_column = find_task_column(sheet_data)
        if task_column is None:
            diagnostics.append(f"{title}: No TASK column found within the first 10 rows.")
            continue

        column_letter = column_index_to_letter(task_column)
        diagnostics.append(
            f"{title}: TASK column located at index {task_column} (column {column_letter})."
        )

        matching_cells, existing_codes = find_placeholder_cells(
            sheet_data, task_column, diagnostics, title
        )
        if not matching_cells:
            diagnostics.append(
                f"{title}: No cells beginning with a flag emoji were found in column {column_letter}."
            )
            continue

        updates[title] = []
        code_iter = placeholder_code_iter(letter_prefix)

        highest_existing_index = _highest_placeholder_index(existing_codes, letter_prefix)
        if highest_existing_index is not None:
            code_iter = itertools.islice(
                code_iter,
                highest_existing_index + 1,
                None,
            )
        for row_index in matching_cells:
            code = next(code_iter)
            cell_a1 = column_index_to_letter(task_column) + str(row_index + 1)
            text = f"TEAM VIDEO PLACEHOLDER {code}"
            updates[title].append((cell_a1, text))
            data_updates.append(
                {
                    "range": single_cell_range(title, cell_a1),
                    "values": [[text]],
                }
            )

    if data_updates:
        values_resource = service.spreadsheets().values()
        batch_update_kwargs = {
            "spreadsheetId": spreadsheet_id,
            "body": {
                "valueInputOption": "USER_ENTERED",
                "data": data_updates,
            },
        }
        if _method_supports_parameter(values_resource.batchUpdate, "supportsAllDrives"):
            batch_update_kwargs["supportsAllDrives"] = True

        values_resource.batchUpdate(**batch_update_kwargs).execute()

    return (
        {
            sheet: [f"{cell}: {text}" for cell, text in entries]
            for sheet, entries in updates.items()
        },
        diagnostics,
        spreadsheet_title,
    )


def inspect_ranking_match_numbers(
    spreadsheet: Dict[str, Any]
) -> Tuple[List[Dict[str, Any]], bool]:
    """Return per-sheet match metadata and whether numbering already exists."""

    sheets: List[Dict[str, Any]] = []
    existing_numbers_found = False

    for sheet in spreadsheet.get("sheets", []):
        if not isinstance(sheet, dict):
            continue

        properties = sheet.get("properties", {})
        title = properties.get("title", "Untitled")
        index = int(properties.get("index", 0))
        sheet_data = sheet.get("data", [])

        entry: Dict[str, Any] = {
            "title": title,
            "index": index,
            "matches": [],
            "diagnostics": [],
        }

        if not sheet_data:
            entry["diagnostics"].append(f"{title}: No data available to inspect.")
            sheets.append(entry)
            continue

        matches: List[Dict[str, Any]] = []
        for grid_data in sheet_data:
            if not isinstance(grid_data, dict):
                continue
            start_row = grid_data.get("startRow", 0)
            start_column = grid_data.get("startColumn", 0)
            for row_offset, row in enumerate(grid_data.get("rowData", [])):
                if not isinstance(row, dict):
                    continue
                row_index = start_row + row_offset
                values = row.get("values", [])
                for column_offset, cell in enumerate(values):
                    if not isinstance(cell, dict):
                        continue
                    column_index = start_column + column_offset
                    cell_text = _extract_cell_text(cell)
                    if not cell_text:
                        continue
                    normalized = cell_text.strip()
                    upper_normalized = normalized.upper()
                    if upper_normalized == "RANKING MATCH" or upper_normalized.startswith(
                        "RANKING MATCH #"
                    ):
                        matches.append(
                            {
                                "row_index": row_index,
                                "column_index": column_index,
                                "original_text": normalized,
                            }
                        )
                        if upper_normalized.startswith("RANKING MATCH #"):
                            existing_numbers_found = True

        if matches:
            entry["matches"] = matches
        else:
            entry["diagnostics"].append(f"{title}: No RANKING MATCH cells found.")

        sheets.append(entry)

    return sheets, existing_numbers_found


def apply_ranking_match_number_updates(
    credentials: Credentials,
    spreadsheet_id: str,
    sheet_entries: Sequence[Dict[str, Any]],
    *,
    renumber_all: bool,
    initial_diagnostics: Optional[Sequence[str]] = None,
    match_numbers_by_sheet: Optional[Mapping[str, Sequence[int]]] = None,
) -> Tuple[Dict[str, List[str]], List[str]]:
    """Apply numbering updates to the provided match metadata."""

    updates: Dict[str, List[Tuple[str, str]]] = {}
    diagnostics: List[str] = list(initial_diagnostics or [])
    data_updates: List[Dict[str, Any]] = []

    counter = 1
    provided_numbers: Dict[str, List[int]] = (
        {
            sheet: [int(number) for number in numbers]
            for sheet, numbers in (match_numbers_by_sheet or {}).items()
        }
    )
    for entry in sorted(
        sheet_entries, key=lambda item: (int(item.get("index", 0)), item.get("title", ""))
    ):
        for message in entry.get("diagnostics", []):
            diagnostics.append(message)

        matches: Sequence[Dict[str, Any]] = entry.get("matches", [])
        if not matches:
            continue

        title = str(entry.get("title", "Untitled"))
        sheet_updates: List[Tuple[str, str]] = []
        assigned_numbers = provided_numbers.get(title)
        use_schedule_numbers = bool(assigned_numbers)

        if use_schedule_numbers and len(assigned_numbers) != len(matches):
            diagnostics.append(
                f"{title}: Provided match numbers ({len(assigned_numbers)}) do not match the number of slots ({len(matches)})."
            )
            continue

        sorted_matches = sorted(
            matches, key=lambda item: (item["row_index"], item["column_index"])
        )
        for idx, match in enumerate(sorted_matches):
            row_index = int(match["row_index"])
            column_index = int(match["column_index"])
            cell_a1 = column_index_to_letter(column_index) + str(row_index + 1)
            if use_schedule_numbers:
                match_number = assigned_numbers[idx]
            else:
                match_number = counter
                counter += 1
            new_text = f"RANKING MATCH #{match_number}"

            if renumber_all or match.get("original_text", "") != new_text:
                sheet_updates.append((cell_a1, new_text))
                data_updates.append(
                    {
                        "range": single_cell_range(title, cell_a1),
                        "values": [[new_text]],
                    }
                )

        diagnostics.append(
            f"{title}: Numbered {len(matches)} RANKING MATCH cell(s); "
            f"updated {len(sheet_updates)} of them."
        )

        if sheet_updates:
            updates[title] = sheet_updates

    if data_updates:
        if build is None:
            raise ModuleNotFoundError(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )

        service = build("sheets", "v4", credentials=credentials)
        values_resource = service.spreadsheets().values()
        batch_update_kwargs = {
            "spreadsheetId": spreadsheet_id,
            "body": {
                "valueInputOption": "USER_ENTERED",
                "data": data_updates,
            },
        }
        if _method_supports_parameter(values_resource.batchUpdate, "supportsAllDrives"):
            batch_update_kwargs["supportsAllDrives"] = True

        values_resource.batchUpdate(**batch_update_kwargs).execute()

    return (
        {
            sheet: [f"{cell}: {text}" for cell, text in entries]
            for sheet, entries in updates.items()
        },
        diagnostics,
    )


def generate_ranking_match_numbers(
    credentials: Credentials, spreadsheet_id: str, *, renumber_all: bool = False
) -> Tuple[Dict[str, List[str]], List[str]]:
    """Number every RANKING MATCH cell found in the spreadsheet."""

    spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
        credentials, spreadsheet_id
    )
    sheet_entries, _existing_numbers = inspect_ranking_match_numbers(spreadsheet)
    initial_diagnostics: List[str] = []
    for entry in sheet_entries:
        initial_diagnostics.extend(entry.get("diagnostics", []))
    return apply_ranking_match_number_updates(
        credentials,
        spreadsheet_id,
        sheet_entries,
        renumber_all=renumber_all,
        initial_diagnostics=initial_diagnostics,
    )


def _extract_cell_text(cell: Dict[str, Any]) -> str:
    """Return the string contents of a cell, if available."""

    formatted = cell.get("formattedValue")
    if isinstance(formatted, str) and formatted.strip():
        return formatted.strip()

    user_entered = cell.get("userEnteredValue")
    if isinstance(user_entered, dict):
        string_value = user_entered.get("stringValue")
        if isinstance(string_value, str) and string_value.strip():
            return string_value.strip()

    effective = cell.get("effectiveValue")
    if isinstance(effective, dict):
        string_value = effective.get("stringValue")
        if isinstance(string_value, str) and string_value.strip():
            return string_value.strip()

    return ""

def find_task_column(sheet_data: Sequence[dict]) -> Optional[int]:
    """Locate the zero-based column index containing the TASK header."""

    for global_row_index in range(10):
        for grid_data in sheet_data:
            start_row = grid_data.get("startRow", 0)
            start_column = grid_data.get("startColumn", 0)
            relative_row_index = global_row_index - start_row
            row_data = grid_data.get("rowData", [])
            if relative_row_index < 0 or relative_row_index >= len(row_data):
                continue

            row = row_data[relative_row_index]
            for offset, cell in enumerate(row.get("values", [])):
                value = cell.get("formattedValue") or cell.get("userEnteredValue", {}).get("stringValue")
                if isinstance(value, str) and value.strip().upper() == "TASK":
                    return start_column + offset
    return None


def find_placeholder_cells(
    sheet_data: Sequence[dict],
    column_index: int,
    diagnostics: Optional[List[str]] = None,
    sheet_title: str = "",
) -> Tuple[List[int], Set[str]]:
    """Return placeholder rows and any existing placeholder codes in the column."""

    matches: List[int] = []
    existing_codes: Set[str] = set()
    seen_rows: set[int] = set()
    sample_count = 0

    for row_index, cell in _iter_column_cells(sheet_data, column_index):
        text = _extract_cell_text(cell)
        if not text:
            continue

        code = _extract_placeholder_code(text)
        if code:
            existing_codes.add(code)

        if not _starts_with_flag_emoji(text):
            continue

        if row_index in seen_rows:
            continue

        matches.append(row_index)
        seen_rows.add(row_index)

        if diagnostics is not None and sample_count < 5:
            preview = text.splitlines()[0][:40]
            diagnostics.append(
                f"{sheet_title}: Row {row_index + 1} flagged for placeholder replacement with value {preview!r}."
            )
            sample_count += 1

    return matches, existing_codes


def _starts_with_flag_emoji(text: str) -> bool:
    """Return True if *text* begins with an emoji flag after optional whitespace."""

    return bool(FLAG_EMOJI_PATTERN.match(text))


def _extract_placeholder_code(text: str) -> Optional[str]:
    """Extract an existing placeholder code from *text*, if present."""

    match = PLACEHOLDER_CODE_PATTERN.match(text.strip())
    if not match:
        return None
    return match.group(1).upper()


def _highest_placeholder_index(existing_codes: Set[str], prefix: str) -> Optional[int]:
    """Return the highest placeholder index encountered for *prefix*, if any."""

    highest: Optional[int] = None
    for code in existing_codes:
        if len(code) != 3 or not code.startswith(prefix):
            continue

        suffix = code[1:]
        if not suffix.isalpha():
            continue

        index = (ord(suffix[0]) - 65) * 26 + (ord(suffix[1]) - 65)
        if highest is None or index > highest:
            highest = index

    return highest


def _iter_column_cells(
    sheet_data: Sequence[dict], column_index: int
) -> Iterable[Tuple[int, dict]]:
    """Yield ``(row_index, cell)`` pairs for the specified column."""

    for grid_data in sheet_data:
        start_row = grid_data.get("startRow", 0)
        start_column = grid_data.get("startColumn", 0)
        cell_index = column_index - start_column
        if cell_index < 0:
            continue

        for offset, row in enumerate(grid_data.get("rowData", [])):
            row_index = start_row + offset
            values = row.get("values", [])
            if cell_index >= len(values):
                continue
            yield row_index, values[cell_index]


def placeholder_code_iter(prefix: str) -> Iterable[str]:
    """Yield placeholder codes with the given prefix."""

    def letters() -> Iterable[str]:
        for first in range(26):
            for second in range(26):
                yield chr(ord("A") + first) + chr(ord("A") + second)

    for suffix in letters():
        yield prefix + suffix


def column_index_to_letter(index: int) -> str:
    """Convert a zero-based column index to its Excel-style letter."""

    index += 1
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def single_cell_range(sheet_title: str, cell_a1: str) -> str:
    """Return a properly quoted A1 range for a single cell."""

    escaped_title = sheet_title.replace("'", "''")
    return f"'{escaped_title}'!{cell_a1}:{cell_a1}"


def format_report(
    report: Dict[str, List[str]],
    diagnostics: Sequence[str] = (),
    *,
    success_header: str = "Placeholder updates applied:",
    empty_message: str = "No matching placeholders were found.",
) -> str:
    """Create a human-readable report of the updates performed."""

    lines: List[str] = []
    if report:
        if success_header:
            lines.append(success_header)
        for sheet, entries in report.items():
            lines.append(f"• {sheet}:")
            for entry in entries:
                lines.append(f"  - {entry}")
    else:
        lines.append(empty_message)

    if diagnostics:
        if lines:
            lines.append("")
        lines.append("Diagnostics:")
        for entry in diagnostics:
            lines.append(f"• {entry}")

    return "\n".join(lines)


def _method_supports_parameter(method: object, parameter: str) -> bool:
    """Return True if a Google API client method accepts the given parameter."""

    method_desc = getattr(method, "_methodDesc", {})
    if isinstance(method_desc, dict):
        parameters = method_desc.get("parameters", {})
        return isinstance(parameters, dict) and parameter in parameters
    return False


def main() -> None:
    """Run the application."""
    root = create_main_window()
    root.mainloop()


if __name__ == "__main__":
    main()
