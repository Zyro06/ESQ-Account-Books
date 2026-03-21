import os


def load_stylesheet(qss_path: str, dark: bool = False) -> str:
    """Load QSS stylesheet from file, returning light or dark section."""
    if not os.path.exists(qss_path):
        return ""
    with open(qss_path, 'r', encoding='utf-8') as f:
        content = f.read()

    parts = content.split('[DARK]', 1)

    if dark:
        return parts[1] if len(parts) > 1 else ""
    else:
        return parts[0]