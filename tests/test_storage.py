import storage


def configure_storage(monkeypatch, tmp_path):
    data_dir = tmp_path / "data"
    monkeypatch.setattr(storage, "DATA_DIR", str(data_dir))
    monkeypatch.setattr(storage, "PHOTO_DIR", str(data_dir / "scanner_photos"))
    monkeypatch.setattr(storage, "BACKUP_DIR", str(data_dir / "backups"))
    monkeypatch.setattr(storage, "DB_FILE", str(data_dir / "state.db"))


def test_save_and_load_state_isolated_from_repository(monkeypatch, tmp_path):
    configure_storage(monkeypatch, tmp_path)
    state = {"records": [{"name": "Cement"}], "client_notes": "Keep receipts"}

    storage.save_state(state)

    assert storage.load_state() == state


def test_backup_and_restore_round_trip(monkeypatch, tmp_path):
    configure_storage(monkeypatch, tmp_path)
    original = {"records": [{"name": "Cement"}]}
    replacement = {"records": [{"name": "Paint"}]}
    storage.save_state(original)
    backup_path = storage.create_backup()
    storage.save_state(replacement)

    assert storage.restore_backup(backup_path) == original
    assert storage.load_state() == original