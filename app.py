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

import tkinter as tk
from tkinter import ttk


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

    placeholder_tools_label = ttk.Label(
        tools_frame,
        text="Tools content will go here.",
        anchor="center",
    )
    placeholder_tools_label.grid(padx=12, pady=12)

    return root


def main() -> None:
    """Run the application."""
    root = create_main_window()
    root.mainloop()


if __name__ == "__main__":
    main()
