"""
ui/widgets/fs/fs_db.py
----------------------
SQLite persistence helpers for the saved_statements table.

Public API:
    ensure_table(db_manager)
    save_statement(db_manager, label, stmt_type, stmt_text, params) -> int
    load_all_statements(db_manager) -> list[dict]
    rename_statement(db_manager, stmt_id, new_label)
    delete_statement(db_manager, stmt_id)
"""

from __future__ import annotations

import json
from database.db_manager import DatabaseManager
from PySide6.QtCore import QDate


# ---------------------------------------------------------------------------
# QDate ↔ JSON serialisation
# ---------------------------------------------------------------------------

def _serialise_params(params: dict) -> dict:
    """Convert QDate values → ISO string dicts for JSON storage."""
    out = {}
    for k, v in params.items():
        if isinstance(v, QDate):
            out[k] = {'__qdate__': v.toString("yyyy-MM-dd")}
        else:
            out[k] = v
    return out


def _deserialise_params(params: dict) -> dict:
    """Restore QDate objects from JSON-loaded params."""
    out = {}
    for k, v in params.items():
        if isinstance(v, dict) and '__qdate__' in v:
            out[k] = QDate.fromString(v['__qdate__'], "yyyy-MM-dd")
        else:
            out[k] = v
    return out


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def ensure_table(db_manager: DatabaseManager) -> None:
    """Create saved_statements table if it doesn't already exist."""
    conn = db_manager.get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS saved_statements (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            label       TEXT    NOT NULL,
            stmt_type   TEXT    NOT NULL,
            stmt_text   TEXT    NOT NULL,
            params_json TEXT    NOT NULL,
            created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()


def save_statement(db_manager: DatabaseManager, label: str,
                   stmt_type: str, stmt_text: str, params: dict) -> int:
    """Insert a new saved statement. Returns the new row id."""
    safe = _serialise_params(params)
    conn = db_manager.get_connection()
    cur  = conn.execute(
        "INSERT INTO saved_statements (label, stmt_type, stmt_text, params_json) "
        "VALUES (?,?,?,?)",
        (label, stmt_type, stmt_text, json.dumps(safe))
    )
    conn.commit()
    return cur.lastrowid


def load_all_statements(db_manager: DatabaseManager) -> list[dict]:
    """Return all saved statements ordered newest-first."""
    conn = db_manager.get_connection()
    cur  = conn.execute(
        "SELECT id, label, stmt_type, stmt_text, params_json, created_at "
        "FROM saved_statements ORDER BY id DESC"
    )
    rows = []
    for row in cur.fetchall():
        params = _deserialise_params(json.loads(row[4]))
        rows.append({
            'id':        row[0],
            'label':     row[1],
            'stmt_type': row[2],
            'text':      row[3],
            'params':    params,
            'timestamp': row[5][:16] if row[5] else '',
        })
    return rows


def rename_statement(db_manager: DatabaseManager, stmt_id: int, new_label: str) -> None:
    conn = db_manager.get_connection()
    conn.execute("UPDATE saved_statements SET label=? WHERE id=?", (new_label, stmt_id))
    conn.commit()


def delete_statement(db_manager: DatabaseManager, stmt_id: int) -> None:
    conn = db_manager.get_connection()
    conn.execute("DELETE FROM saved_statements WHERE id=?", (stmt_id,))
    conn.commit()