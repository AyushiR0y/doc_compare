import os
import shutil
from pathlib import Path


def get_usage_log_path(reference_file: str) -> Path:
    configured_path = os.getenv("USAGE_LOG_FILE_PATH")
    if configured_path:
        return Path(configured_path).expanduser().resolve()

    project_root = Path(reference_file).resolve().parent
    legacy_path = project_root / "usage_logs.jsonl"

    persistent_dir = Path.home() / ".doc_compare"
    persistent_path = persistent_dir / "usage_logs.jsonl"

    if persistent_path.exists():
        return persistent_path

    if legacy_path.exists():
        try:
            persistent_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(legacy_path, persistent_path)
            return persistent_path
        except OSError:
            return legacy_path

    return persistent_path