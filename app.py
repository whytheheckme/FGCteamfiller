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
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

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
    notebook.grid(row=1, column=0, padx=16, pady=16, sticky="nsew")

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

    build_tools_tab(tools_frame)

    return root


def build_tools_tab(parent: ttk.Frame) -> None:
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

    credentials_manager = GoogleDriveCredentialsManager(load_drive_frame)
    credentials_manager.render(row=1)

    separator = ttk.Separator(parent, orient="horizontal")
    separator.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

    ros_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    ros_frame.grid(row=2, column=0, sticky="nsew")
    ros_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_frame,
        text="ROS Placeholder Generator",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ROSPlaceholderGeneratorUI(ros_frame, credentials_manager).render(row=1)


class GoogleDriveCredentialsManager:
    """Handle selection and loading of Google Drive OAuth credentials."""

    SCOPES = [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    STORAGE_PATH = Path.home() / ".fgc_team_filler" / "google_credentials.json"

    def __init__(self, parent: ttk.Frame) -> None:
        self.parent = parent
        self.credentials_path_var = tk.StringVar()
        self._credentials: Optional[Credentials] = None
        self._status_display: Optional[tk.Text] = None

    @property
    def credentials(self) -> Optional[Credentials]:
        return self._credentials

    def get_valid_credentials(self) -> Tuple[Optional[Credentials], Optional[str]]:
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
                self.set_status(message)
                return None, message
            try:
                credentials.refresh(Request())
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to refresh credentials: {exc}"[:500]
                self.set_status(message)
                return None, message
            else:
                self._credentials = credentials
                self._persist_credentials(credentials)

        if not credentials.valid:
            message = "Stored credentials are invalid. Please re-authorize."
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

        status_display = tk.Text(
            self.parent,
            height=4,
            wrap="word",
            borderwidth=1,
            relief="solid",
            cursor="xterm",
        )
        background = _get_widget_background(self.parent, status_display)
        status_display.configure(background=background)
        status_display.tag_configure("status", foreground="#1a73e8")
        status_display.grid(row=row + 2, column=0, columnspan=2, sticky="nsew", pady=(0, 4))

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

        status_display.bind("<Key>", _on_key)
        self._status_display = status_display
        self.set_status("No credentials loaded.")

        # Attempt to load persisted credentials immediately.
        self._load_persisted_credentials()

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

            self.parent.after(0, _on_success)

        threading.Thread(target=_run_flow, daemon=True).start()

    def set_status(self, message: str) -> None:
        if self._status_display is None:
            return
        widget = self._status_display
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", message, "status")
        widget.configure(state="normal")

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


class ROSPlaceholderGeneratorUI:
    """User interface wrapper for the ROS placeholder generator tool."""

    COLOR_TARGET = (0.698, 0.843, 0.984)
    COLOR_TOLERANCE = 0.05

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: GoogleDriveCredentialsManager,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.sheet_url_var = tk.StringVar()
        self._status_display: Optional[tk.Text] = None
        self._default_status = "Paste a Google Sheets link to begin."

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Google Sheets URL:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        entry = ttk.Entry(
            self.parent,
            textvariable=self.sheet_url_var,
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        generate_button = ttk.Button(
            self.parent,
            text="Generate Placeholders",
            command=self.generate_placeholders,
        )
        generate_button.grid(row=row + 1, column=0, columnspan=2, sticky="e", pady=8)

        status_display = tk.Text(
            self.parent,
            height=6,
            wrap="word",
            borderwidth=1,
            relief="solid",
            cursor="xterm",
        )
        background = _get_widget_background(self.parent, status_display)
        status_display.configure(background=background)
        status_display.tag_configure("status", foreground="black")
        status_display.grid(row=row + 2, column=0, columnspan=2, sticky="nsew", pady=(0, 4))

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

        status_display.bind("<Key>", _on_key)
        self._status_display = status_display
        self.set_status(self._default_status)

    def generate_placeholders(self) -> None:
        sheet_url = self.sheet_url_var.get().strip()
        if not sheet_url:
            self.set_status("Please paste a Google Sheets link.")
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

        spreadsheet_id = extract_spreadsheet_id(sheet_url)
        if not spreadsheet_id:
            self.set_status("Unable to determine spreadsheet ID from the provided link.")
            return

        self.set_status("Contacting Google Sheets API...")

        def _worker() -> None:
            try:
                report = generate_placeholders_for_sheet(credentials, spreadsheet_id)
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"[:500]
                self.parent.after(0, lambda: self.set_status(message))
                return

            success_message = format_report(report)
            self.parent.after(0, lambda: self.set_status(success_message))

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str) -> None:
        if self._status_display is None:
            return
        widget = self._status_display
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", message, "status")
        widget.configure(state="normal")


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


def generate_placeholders_for_sheet(credentials: Credentials, spreadsheet_id: str) -> Dict[str, List[str]]:
    """Scan the spreadsheet and replace placeholder cells.

    Returns a dictionary keyed by sheet title with a list of the cell updates applied.
    """

    service = build("sheets", "v4", credentials=credentials)
    spreadsheets_resource = service.spreadsheets()
    get_kwargs = {
        "spreadsheetId": spreadsheet_id,
        "includeGridData": True,
        "fields": (
            "sheets(properties(title,sheetId,index),"
            "data(rowData(values(userEnteredValue,formattedValue,"
            "effectiveFormat(backgroundColor)))))"
        ),
    }
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
        raise

    updates: Dict[str, List[Tuple[str, str]]] = {}
    data_updates: List[Dict[str, Sequence[Sequence[str]]]] = []

    for sheet in spreadsheet.get("sheets", []):
        properties = sheet.get("properties", {})
        title = properties.get("title", "Untitled")
        sheet_index = properties.get("index", 0)
        letter_prefix = chr(ord("A") + sheet_index)

        sheet_data = sheet.get("data", [])
        if not sheet_data:
            continue

        row_data = sheet_data[0].get("rowData", [])
        if not row_data:
            continue

        task_column = find_task_column(row_data)
        if task_column is None:
            continue

        matching_cells = find_placeholder_cells(row_data, task_column)
        if not matching_cells:
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
                    "range": f"'{title}'!{cell_a1}:{cell_a1}",
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

    return {
        sheet: [f"{cell}: {text}" for cell, text in entries]
        for sheet, entries in updates.items()
    }


def find_task_column(row_data: Sequence[dict]) -> Optional[int]:
    """Locate the zero-based column index containing the TASK header."""

    for row in row_data[:10]:
        for idx, cell in enumerate(row.get("values", [])):
            value = cell.get("formattedValue") or cell.get("userEnteredValue", {}).get("stringValue")
            if isinstance(value, str) and value.strip().upper() == "TASK":
                return idx
    return None


def find_placeholder_cells(row_data: Sequence[dict], column_index: int) -> List[int]:
    """Return the row indices containing the target placeholder color."""

    matches: List[int] = []
    for row_index, row in enumerate(row_data):
        if column_index >= len(row.get("values", [])):
            continue
        cell = row["values"][column_index]
        color = cell.get("effectiveFormat", {}).get("backgroundColor") or {}
        if is_light_cornflower_blue(color):
            matches.append(row_index)
    return matches


def is_light_cornflower_blue(color: Dict[str, float]) -> bool:
    """Return True if the color approximately matches Light Cornflower Blue 3."""

    target = ROSPlaceholderGeneratorUI.COLOR_TARGET
    tolerance = ROSPlaceholderGeneratorUI.COLOR_TOLERANCE
    red = color.get("red", 1.0)
    green = color.get("green", 1.0)
    blue = color.get("blue", 1.0)
    return (
        abs(red - target[0]) <= tolerance
        and abs(green - target[1]) <= tolerance
        and abs(blue - target[2]) <= tolerance
    )


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


def format_report(report: Dict[str, List[str]]) -> str:
    """Create a human-readable report of the updates performed."""

    if not report:
        return "No matching placeholders were found."

    lines = ["Placeholder updates applied:"]
    for sheet, entries in report.items():
        lines.append(f"• {sheet}:")
        for entry in entries:
            lines.append(f"  - {entry}")
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
