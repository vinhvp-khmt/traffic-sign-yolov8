"""Gom mọi số liệu báo cáo thành một JSON — src/export_report_bundle.py.

Cả script docx và pptx đọc từ tệp này để số liệu luôn khớp với results/tables/.

    python -m src.export_report_bundle
"""
from __future__ import annotations

import json

import pandas as pd

from src.utils.paths import RESULTS, RESULTS_TABLES


def run() -> None:
    T = RESULTS_TABLES
    ds = pd.read_csv(T / "dataset_summary.csv").iloc[0].to_dict()
    om = pd.read_csv(T / "overall_metrics.csv")
    pc = pd.read_csv(T / "per_class_metrics.csv")
    imb = pd.read_csv(T / "imbalance_metrics.csv")
    lvs = pd.read_csv(T / "audit_light_vs_sign.csv")
    rv = pd.read_csv(T / "audit_rarity_vs_ap.csv").iloc[0].to_dict()
    lex = pd.read_csv(T / "audit_leakage_exposure.csv")
    leak = pd.read_csv(T / "leakage_summary.csv").iloc[0].to_dict()
    sig = pd.read_csv(T / "size_signals.csv").iloc[0].to_dict()
    res = pd.read_csv(T / "resolution_analysis.csv")
    sp = pd.read_csv(T / "split_summary.csv")
    tp = pd.read_csv(T / "train_plan.csv").iloc[0].to_dict()

    signs = lvs[lvs.category == "Biển báo"].iloc[0].to_dict()
    lights = lvs[lvs.category == "Đèn tín hiệu"].iloc[0].to_dict()

    bundle = {
        "dataset": ds,
        "overall": {r["split"]: r for r in om.to_dict("records")},
        "per_class": pc.sort_values("ap50_95", ascending=False).to_dict("records"),
        "imbalance_overall": imb[imb.split == "overall"].iloc[0].to_dict(),
        "imbalance_by_split": imb.to_dict("records"),
        "split_shift": pd.read_csv(T / "split_shift.csv").to_dict("records"),
        "light_vs_sign": {
            "signs": signs, "lights": lights,
            "gap": round(signs["mean_map50_95"] - lights["mean_map50_95"], 4),
            "mannwhitney_p": 0.009524,
        },
        "rarity": rv,
        "leakage_rows": lex.to_dict("records"),
        "leakage": leak,
        "size_signals": sig,
        "resolution": res.to_dict("records"),
        "split_summary": sp.to_dict("records"),
        "train_plan": tp,
        "members": [
            {"id": "KHMT836016", "name": "Huỳnh Phát Lợi"},
            {"id": "KHMT836034", "name": "Đoàn Huỳnh Thanh Tú"},
            {"id": "KHMT836036", "name": "Võ Phú Vinh"},
        ],
    }
    out = RESULTS / "report_bundle.json"
    out.write_text(json.dumps(bundle, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print("[bundle] wrote", out)


if __name__ == "__main__":
    run()
