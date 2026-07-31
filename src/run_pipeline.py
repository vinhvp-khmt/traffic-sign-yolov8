"""Chạy toàn bộ pipeline phân tích thứ cấp theo thứ tự — src/run_pipeline.py.

    python -m src.run_pipeline
"""
from __future__ import annotations

import importlib

STAGES = [
    ("extract", "src.extract_notebook_results"),
    ("audit", "src.audit_analysis"),
    ("figures", "src.make_figures"),
]


def main() -> None:
    for label, mod in STAGES:
        print(f"\n===== {label} ({mod}) =====")
        importlib.import_module(mod).run()
    print("\n[pipeline] Hoàn tất. Xem results/tables/ và results/figs/.")


if __name__ == "__main__":
    main()
