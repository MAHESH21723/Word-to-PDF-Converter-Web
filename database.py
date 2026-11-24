"""SQLite persistence for the conversion history.

Every function takes the database path explicitly so the app (and the tests)
can point at any file. The database and its table are created on demand.
"""
import os
import sqlite3
from contextlib import closing

DEFAULT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "database.db")

_COLUMNS = (
    "id, filename, original_size, converted_filename, converted_size, "
    "status, error_message, conversion_time, created_at"
)


def _connect(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path=DEFAULT_DB_PATH):
    """Create the database file and the conversions table if missing."""
    parent = os.path.dirname(os.path.abspath(db_path))
    os.makedirs(parent, exist_ok=True)
    with closing(_connect(db_path)) as conn, conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS conversions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                original_size INTEGER NOT NULL,
                converted_filename TEXT,
                converted_size INTEGER,
                status TEXT NOT NULL,
                error_message TEXT,
                conversion_time REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )


def add_conversion(db_path, filename, original_size, converted_filename,
                   converted_size, status, error_message, conversion_time):
    with closing(_connect(db_path)) as conn, conn:
        conn.execute(
            """
            INSERT INTO conversions (
                filename, original_size, converted_filename, converted_size,
                status, error_message, conversion_time
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (filename, original_size, converted_filename, converted_size,
             status, error_message, conversion_time),
        )


def get_all_conversions(db_path, limit=200):
    with closing(_connect(db_path)) as conn:
        rows = conn.execute(
            f"SELECT {_COLUMNS} FROM conversions ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(row) for row in rows]


def get_conversion_by_stored_name(db_path, stored_name):
    """Return the successful conversion whose PDF is stored under `stored_name`."""
    with closing(_connect(db_path)) as conn:
        row = conn.execute(
            f"SELECT {_COLUMNS} FROM conversions "
            "WHERE converted_filename = ? AND status = 'success'",
            (stored_name,),
        ).fetchone()
    return dict(row) if row else None


def get_metrics(db_path):
    with closing(_connect(db_path)) as conn:
        total, success, failed, processed, avg_time = conn.execute(
            """
            SELECT
                COUNT(*),
                COALESCE(SUM(status = 'success'), 0),
                COALESCE(SUM(status = 'failed'), 0),
                COALESCE(SUM(original_size), 0),
                AVG(CASE WHEN status = 'success' THEN conversion_time END)
            FROM conversions
            """
        ).fetchone()
    return {
        "total": total,
        "success": success,
        "failed": failed,
        "success_rate": round(success / total * 100, 1) if total else 0,
        "bytes_processed": processed,
        "avg_speed": round(avg_time, 2) if avg_time else 0.0,
    }


def clear_history(db_path):
    with closing(_connect(db_path)) as conn, conn:
        conn.execute("DELETE FROM conversions")
