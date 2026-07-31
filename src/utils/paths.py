"""Đăng ký đường dẫn tập trung — định nghĩa một lần duy nhất."""
from __future__ import annotations

import os
from pathlib import Path

REPO_ROOT = Path(os.environ.get("TYA_REPO_ROOT", Path(__file__).resolve().parents[2]))

CONFIGS = REPO_ROOT / "configs"
DATA = REPO_ROOT / "data"
NOTEBOOK = DATA / "notebook" / "traffic-yolo-v2-run.ipynb"

RESULTS = REPO_ROOT / "results"
RESULTS_TABLES = RESULTS / "tables"
RESULTS_FIGS = RESULTS / "figs"

REPORTS = REPO_ROOT / "reports"


def ensure_dir(path: Path) -> Path:
    Path(path).mkdir(parents=True, exist_ok=True)
    return Path(path)
