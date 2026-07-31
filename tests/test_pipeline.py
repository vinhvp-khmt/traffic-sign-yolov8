"""Kiểm thử: bảng trích xuất khớp output notebook và audit tính đúng.

    pytest
"""
from __future__ import annotations

import pandas as pd
import pytest

from src import audit_analysis, extract_notebook_results, make_figures
from src.utils.paths import RESULTS_FIGS, RESULTS_TABLES


@pytest.fixture(scope="module", autouse=True)
def _build():
    extract_notebook_results.run()
    audit_analysis.run()
    make_figures.run()


def test_per_class_has_15_classes():
    pc = pd.read_csv(RESULTS_TABLES / "per_class_metrics.csv")
    assert len(pc) == 15
    assert pc["is_traffic_light"].sum() == 2


def test_overall_metrics_match_notebook():
    m = pd.read_csv(RESULTS_TABLES / "overall_metrics.csv").set_index("split")
    # test split, khớp từng chữ số với cell [43]/[48]
    assert round(m.loc["test", "map50"], 4) == 0.9703
    assert round(m.loc["test", "map50_95"], 4) == 0.8120
    assert round(m.loc["test", "precision"], 4) == 0.9603


def test_light_vs_sign_gap_significant():
    lvs = pd.read_csv(RESULTS_TABLES / "audit_light_vs_sign.csv")
    signs = lvs[lvs["category"] == "Biển báo"]["mean_map50_95"].iloc[0]
    lights = lvs[lvs["category"] == "Đèn tín hiệu"]["mean_map50_95"].iloc[0]
    assert signs - lights > 0.30            # khoảng cách ~0,31
    assert lights < 0.60 and signs > 0.80


def test_rarity_not_correlated():
    rv = pd.read_csv(RESULTS_TABLES / "audit_rarity_vs_ap.csv").iloc[0]
    # Không có tương quan giữa số instance và AP
    assert abs(rv["spearman_instances_vs_map5095"]) < 0.2
    assert rv["p_instances_vs_map5095"] > 0.3


def test_leakage_numbers():
    lex = pd.read_csv(RESULTS_TABLES / "audit_leakage_exposure.csv")
    cross = lex[lex["loai_trung_lap"].str.contains("Exact duplicate XUYÊN")]
    assert int(cross["so_dong"].iloc[0]) == 202


def test_figures_exist():
    for i in range(1, 7):
        matches = list(RESULTS_FIGS.glob(f"0{i}_*.png"))
        assert matches, f"thiếu hình 0{i}"
