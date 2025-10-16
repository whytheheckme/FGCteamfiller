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

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

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
    tools_frame = ttk.Frame(notebook)

    for frame in (team_videos_frame, tools_frame):
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

    notebook.add(team_videos_frame, text="Team Videos")
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
    build_tools_tab(tools_frame, console)

    return root


def build_tools_tab(parent: ttk.Frame, console: ApplicationConsole) -> None:
    """Populate the Tools tab with utility sections."""

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

    ros_document_frame = ttk.Frame(parent, padding=(12, 6, 12, 6))
    ros_document_frame.grid(row=2, column=0, sticky="nsew")
    ros_document_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_document_frame,
        text="Load ROS Document",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ros_document_loader = ROSDocumentLoaderUI(ros_document_frame, console)
    ros_document_loader.render(row=1)

    separator_two = ttk.Separator(parent, orient="horizontal")
    separator_two.grid(row=3, column=0, sticky="ew", padx=12, pady=6)

    ros_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    ros_frame.grid(row=4, column=0, sticky="nsew")
    ros_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_frame,
        text="ROS Placeholder Generator",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ROSPlaceholderGeneratorUI(ros_frame, credentials_manager, ros_document_loader, console).render(
        row=1
    )


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
        self._status_var = tk.StringVar()
        self._credentials: Optional[Credentials] = None
        self._user_poll_after_id: Optional[str] = None
        self._last_user_message: Optional[str] = None

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

        import_button = ttk.Button(
            self.parent,
            text="Import saved credentials...",
            command=self.import_saved_credentials,
        )
        import_button.grid(row=row + 2, column=0, columnspan=2, sticky="w", pady=(0, 8))

        ttk.Label(self.parent, text="Currently Logged-in User:").grid(
            row=row + 3, column=0, sticky="w"
        )

        user_display = ttk.Entry(
            self.parent,
            textvariable=self.logged_in_user_var,
            state="readonly",
        )
        user_display.grid(row=row + 3, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            foreground="#1a73e8",
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 4, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.set_status("No credentials loaded.", log=False)
        self._load_persisted_credentials()
        self._schedule_user_poll(0)

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

            self.parent.after(0, _on_success)

        threading.Thread(target=_run_flow, daemon=True).start()

    def import_saved_credentials(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.parent,
            title="Select authorized credentials file",
            filetypes=[("JSON files", "*.json"), ("All files", "*")],
        )
        if not filename:
            return

        if Credentials is None:
            self.set_status(
                "google-auth-oauthlib is not installed. Install it with 'pip install google-auth-oauthlib'."
            )
            return

        try:
            imported_credentials = Credentials.from_authorized_user_file(
                filename, self.SCOPES
            )
        except Exception as exc:  # pragma: no cover - file parsing
            self.set_status(f"Failed to import credentials: {exc}"[:500])
            return

        if not imported_credentials.valid and getattr(
            imported_credentials, "refresh_token", None
        ):
            if Request is None:
                self.set_status(
                    "Imported credentials have expired but google-auth is missing the Requests transport."
                )
                return
            try:
                imported_credentials.refresh(Request())
            except Exception as exc:  # pragma: no cover - network interaction
                self.set_status(f"Failed to refresh imported credentials: {exc}"[:500])
                return

        if not imported_credentials.valid:
            self.set_status(
                "Imported credentials are invalid. Try authorizing the account again."
            )
            return

        self._credentials = imported_credentials
        self.credentials_path_var.set(filename)
        message = self._persist_credentials(imported_credentials)
        if message is None:
            message = "Credentials imported successfully."
        elif message.startswith("Authorization successful."):
            message = message.replace(
                "Authorization successful.", "Credentials imported successfully."
            )
        self.set_status(message)
        self._schedule_user_poll(0)

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
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
            self.set_status("Loaded saved credentials. You're ready to go.")
            self._persist_credentials(credentials)
            self._schedule_user_poll(0)
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

    def __init__(self, parent: ttk.Frame, console: ApplicationConsole) -> None:
        self.parent = parent
        self.console = console
        self.sheet_url_var = tk.StringVar(value=self._load_saved_url())
        self._status_var = tk.StringVar()
        self._listeners: List[Callable[[str], None]] = []

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Google Sheets URL:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        entry = ttk.Entry(self.parent, textvariable=self.sheet_url_var)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        save_button = ttk.Button(
            self.parent,
            text="Save Document URL",
            command=self.save_document_url,
        )
        save_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 2, column=0, columnspan=2, sticky="w")

        initial_url = self.sheet_url_var.get().strip()
        if initial_url:
            self._set_status("Loaded saved ROS document URL.", log=False)
        else:
            self._set_status("Enter a Google Sheets link to use with ROS tools.", log=False)

    def add_listener(self, callback: Callable[[str], None]) -> None:
        self._listeners.append(callback)

    def get_document_url(self) -> str:
        return self.sheet_url_var.get().strip()

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


class ROSPlaceholderGeneratorUI:
    """User interface wrapper for the ROS placeholder generator tool."""

    # Light Cornflower Blue 3 as exported by Google Sheets, used to flag
    # placeholder cells. The previous value was copied from an external source
    # and did not match the actual cell color (≈#C9DAF8), causing detection to
    # miss legitimately highlighted rows.
    COLOR_TARGET = (0.788, 0.855, 0.973)

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
        self._on_document_url_changed(self.document_loader.get_document_url())

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
        display_value = url if url else "Not set"
        self.current_document_var.set(display_value)
        if not url:
            self.set_status("Save a ROS document URL before generating placeholders.", log=False)

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
                report, diagnostics = generate_placeholders_for_sheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            success_message = format_report(report, diagnostics)
            self.parent.after(0, lambda: self.set_status(success_message))

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[ROS Placeholder Generator] {message}")


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


def generate_placeholders_for_sheet(
    credentials: Credentials, spreadsheet_id: str
) -> Tuple[Dict[str, List[str]], List[str]]:
    """Scan the spreadsheet and replace placeholder cells.

    Returns the applied updates together with diagnostics describing how the
    spreadsheet was inspected (for example, which column contained the ``TASK``
    header and sample colors that were evaluated).
    """

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
    get_kwargs = {
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

    updates: Dict[str, List[Tuple[str, str]]] = {}
    diagnostics: List[str] = [
        (
            "Placeholder detection uses Light Cornflower Blue 3 with RGB "
            f"{ROSPlaceholderGeneratorUI.COLOR_TARGET!r}."
        )
    ]
    data_updates: List[Dict[str, Sequence[Sequence[str]]]] = []

    if not theme_supported:
        diagnostics.append(
            "Spreadsheet theme colors were unavailable; falling back to raw RGB values."
        )

    theme_colors = _build_theme_color_map(spreadsheet.get("spreadsheetTheme", {}))

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

        matching_cells = find_placeholder_cells(
            sheet_data, task_column, diagnostics, title, theme_colors
        )
        if not matching_cells:
            diagnostics.append(
                f"{title}: No cells matched the placeholder color in column {column_letter}."
            )
            continue

        updates[title] = []
        code_iter = placeholder_code_iter(letter_prefix)
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
    )


def _build_theme_color_map(
    spreadsheet_theme: Dict[str, Any]
) -> Dict[str, Dict[str, float]]:
    """Return a mapping of theme color names to resolved RGB components."""

    theme_colors: Dict[str, Dict[str, float]] = {}
    for entry in spreadsheet_theme.get("themeColors", []):
        if not isinstance(entry, dict):
            continue
        name = entry.get("colorType")
        style = entry.get("color", entry.get("colorStyle", {}))
        rgb_color = {}
        if isinstance(style, dict):
            rgb_color = style.get("rgbColor", {})
        if (
            isinstance(name, str)
            and isinstance(rgb_color, dict)
            and rgb_color
        ):
            theme_colors[name] = {
                "red": float(rgb_color.get("red", 1.0)),
                "green": float(rgb_color.get("green", 1.0)),
                "blue": float(rgb_color.get("blue", 1.0)),
            }
    return theme_colors


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
    theme_colors: Optional[Dict[str, Dict[str, float]]] = None,
) -> List[int]:
    """Return the row indices containing the target placeholder color."""

    matches: List[int] = []
    seen_rows: set[int] = set()
    sample_count = 0
    for row_index, cell in _iter_column_cells(sheet_data, column_index):
        color, color_source = _get_cell_color(cell, theme_colors)

        if diagnostics is not None and sample_count < 5:
            red = color.get("red", 1.0)
            green = color.get("green", 1.0)
            blue = color.get("blue", 1.0)
            diagnostics.append(
                (
                    f"{sheet_title}: Row {row_index + 1} color rgb=({red:.3f}, {green:.3f}, {blue:.3f})"
                    f" from {color_source}."
                )
            )
            sample_count += 1

        if is_light_cornflower_blue(color):
            if row_index not in seen_rows:
                matches.append(row_index)
                seen_rows.add(row_index)
    return matches


def _get_cell_color(
    cell: Dict[str, Any], theme_colors: Optional[Dict[str, Dict[str, float]]]
) -> Tuple[Dict[str, float], str]:
    """Extract a color dictionary and describe the source used."""

    # ``userEnteredFormat`` contains the formatting explicitly applied by the
    # user, whereas ``effectiveFormat`` reflects the resolved styling Google
    # Sheets would display.  When a spreadsheet is opened in compatibility
    # mode, ``effectiveFormat`` often reports the default white background even
    # though ``userEnteredFormat`` retains the actual fill color we need to
    # detect placeholders.  Prefer the user-entered data and fall back to the
    # effective values if necessary.
    sources = (
        ("userEnteredFormat", cell.get("userEnteredFormat", {})),
        ("effectiveFormat", cell.get("effectiveFormat", {})),
    )

    for prefix, fmt in sources:
        if not isinstance(fmt, dict):
            continue

        background = fmt.get("backgroundColor")
        if isinstance(background, dict) and background:
            return (
                {
                    "red": float(background.get("red", 1.0)),
                    "green": float(background.get("green", 1.0)),
                    "blue": float(background.get("blue", 1.0)),
                },
                f"{prefix}.backgroundColor",
            )

        style = fmt.get("backgroundColorStyle")
        if not isinstance(style, dict):
            continue

        rgb_color = style.get("rgbColor")
        if isinstance(rgb_color, dict) and rgb_color:
            return (
                {
                    "red": float(rgb_color.get("red", 1.0)),
                    "green": float(rgb_color.get("green", 1.0)),
                    "blue": float(rgb_color.get("blue", 1.0)),
                },
                f"{prefix}.backgroundColorStyle.rgbColor",
            )

        theme_name = style.get("themeColor")
        if (
            isinstance(theme_name, str)
            and theme_colors is not None
            and theme_name in theme_colors
        ):
            theme_color = theme_colors[theme_name]
            return (
                {
                    "red": float(theme_color.get("red", 1.0)),
                    "green": float(theme_color.get("green", 1.0)),
                    "blue": float(theme_color.get("blue", 1.0)),
                },
                f"{prefix}.backgroundColorStyle.themeColor={theme_name}",
            )

    return ({"red": 1.0, "green": 1.0, "blue": 1.0}, "default")


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


def is_light_cornflower_blue(color: Dict[str, float]) -> bool:
    """Return True if the color matches Light Cornflower Blue 3 exactly."""

    target = ROSPlaceholderGeneratorUI.COLOR_TARGET
    red = float(color.get("red", 1.0))
    green = float(color.get("green", 1.0))
    blue = float(color.get("blue", 1.0))
    return red == target[0] and green == target[1] and blue == target[2]


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
    report: Dict[str, List[str]], diagnostics: Sequence[str] = ()
) -> str:
    """Create a human-readable report of the updates performed."""

    lines: List[str] = []
    if report:
        lines.append("Placeholder updates applied:")
        for sheet, entries in report.items():
            lines.append(f"• {sheet}:")
            for entry in entries:
                lines.append(f"  - {entry}")
    else:
        lines.append("No matching placeholders were found.")

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
