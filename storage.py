import json
import os
import shutil
import sqlite3
from datetime import datetime, timezone


APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, "data")
PHOTO_DIR = os.path.join(DATA_DIR, "scanner_photos")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
DB_FILE = os.path.join(DATA_DIR, "ailyn_house.db")


def _ensure_storage():
    os.makedirs(PHOTO_DIR, exist_ok=True)
    os.makedirs(BACKUP_DIR, exist_ok=True)
    with sqlite3.connect(DB_FILE) as connection:
        connection.execute(
            "CREATE TABLE IF NOT EXISTS app_state (id INTEGER PRIMARY KEY CHECK (id = 1), payload TEXT NOT NULL)"
        )


def load_state():
    _ensure_storage()
    with sqlite3.connect(DB_FILE) as connection:
        row = connection.execute("SELECT payload FROM app_state WHERE id = 1").fetchone()
    if not row:
        return {}
    try:
        return json.loads(row[0])
    except (TypeError, json.JSONDecodeError):
        return {}


def save_state(state):
    _ensure_storage()
    persistent_keys = {
        "records", "labor_records", "payroll_expenses", "planner_tasks", "budget",
        "budget_history", "remaining_money", "view", "receipt_archive", "project",
        "scanner_photos", "dark_mode", "client_notes", "app_settings", "messages",
    }
    payload = {key: state.get(key) for key in persistent_keys if key in state}
    encoded = json.dumps(payload, ensure_ascii=False, default=str)
    with sqlite3.connect(DB_FILE) as connection:
        connection.execute(
            "INSERT INTO app_state(id, payload) VALUES(1, ?) "
            "ON CONFLICT(id) DO UPDATE SET payload=excluded.payload",
            (encoded,),
        )
    return payload


def history_count():
    return len(load_state().get("records", []))


def create_backup():
    _ensure_storage()
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"ailyn-house-{timestamp}.db")
    shutil.copy2(DB_FILE, backup_path)
    return backup_path


def restore_backup(backup_path):
    if not os.path.isfile(backup_path):
        raise ValueError("The selected backup file does not exist.")
    try:
        with sqlite3.connect(backup_path) as connection:
            row = connection.execute("SELECT payload FROM app_state WHERE id = 1").fetchone()
            payload = json.loads(row[0]) if row else {}
    except (sqlite3.Error, TypeError, json.JSONDecodeError) as error:
        raise ValueError("The selected file is not a valid Ailyn House backup.") from error
    _ensure_storage()
    shutil.copy2(backup_path, DB_FILE)
    return payload


def save_scanner_photo(photo_bytes, mime_type, photo_id):
    _ensure_storage()
    extension = ".png" if mime_type == "image/png" else ".jpg"
    filename = f"{photo_id}{extension}"
    absolute_path = os.path.join(PHOTO_DIR, filename)
    with open(absolute_path, "wb") as photo_file:
        photo_file.write(photo_bytes)
    return os.path.relpath(absolute_path, APP_DIR)


def delete_scanner_photo(relative_path):
    if not relative_path:
        return
    absolute_path = os.path.abspath(os.path.join(APP_DIR, relative_path))
    if os.path.commonpath((absolute_path, PHOTO_DIR)) != PHOTO_DIR:
        return
    if os.path.isfile(absolute_path):
        os.remove(absolute_path)