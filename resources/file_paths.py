import os
import sys


def _app_root() -> str:
    """Return the directory that contains the running exe or main.py."""
    if getattr(sys, 'frozen', False):          # PyInstaller bundle
        return os.path.dirname(sys.executable)
    # Running as plain Python – go up from resources/ to project root
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Export
def get_io_dir(module_name: str) -> str:
    folder = os.path.join(_app_root(), "Data", "Exports", module_name)
    os.makedirs(folder, exist_ok=True)
    return folder

# Import
def get_import_dir(module_name: str) -> str:
    folder = os.path.join(_app_root(), "Data", "Imports")
    os.makedirs(folder, exist_ok=True)
    return folder

# Bundled resources (style.qss, user_manual.html, etc.)
def get_resource(filename: str) -> str:
    if getattr(sys, 'frozen', False):
        # Bundled files are extracted to sys._MEIPASS at runtime
        return os.path.join(sys._MEIPASS, 'resources', filename)
    return os.path.join(_app_root(), 'resources', filename)