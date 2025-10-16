# FIRST Global 2025 Team Video Slotter

This project provides a desktop GUI for organizing information about the
FIRST Global 2025 robotics event.

## Getting Started

1. Ensure Python 3.10 or newer is installed.
2. Install the Tkinter package if it is not bundled with your Python
   distribution.
   - On macOS using Homebrew's Python 3.13, run `brew install python-tk@3.13`
     to install the missing `_tkinter` extension module.
3. Launch the application:

   ```bash
   python app.py
   ```

The initial application window includes tabs for "Team Videos" and
"Tools" that will be extended in future iterations. Within the Tools tab
you should see three sections, starting with **Load Google Drive
Credentials**, followed by **Load ROS Document** and **ROS Placeholder
Generator**. If a section does not appear, confirm you are running the
latest version of the app and that the window is tall enough to display
all of the sections.
