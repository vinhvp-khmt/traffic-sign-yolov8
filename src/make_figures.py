"""Sinh hình cho báo cáo/slide từ bảng đã trích — src/make_figures.py.

Tất cả hình vẽ từ results/tables/*.csv (số liệu thực của notebook), lưu vào results/figs/.

    python -m src.make_figures
"""
from __future__ import annotations

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from src.utils.paths import RESULTS_FIGS, RESULTS_TABLES, ensure_dir

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "figure.dpi": 130,
    "savefig.dpi": 170,
    "font.size": 11,
    "axes.grid": True,
    "grid.alpha": 0.25,
    "axes.axisbelow": True,
})

TEAL = "#1C7293"
RED = "#E63946"
AMBER = "#F4A261"
GREEN = "#2C7A4B"
GREY = "#8895A7"


def _load(name):
    return pd.read_csv(RESULTS_TABLES / name)


def _save(fig, name):
    ensure_dir(RESULTS_FIGS)
    fig.tight_layout()
    fig.savefig(RESULTS_FIGS / name, bbox_inches="tight")
    plt.close(fig)
    print("  saved", name)


def fig_per_class_ap():
    pc = _load("per_class_metrics.csv").sort_values("ap50_95")
    colors = [RED if x else TEAL for x in pc["is_traffic_light"]]
    fig, ax = plt.subplots(figsize=(10, 5.4))
    ax.barh(pc["class_name"], pc["ap50_95"], color=colors)
    ax.set_xlabel("mAP@0.5:0.95 (test)")
    ax.set_title("Hiệu năng theo lớp — hai lớp đèn (đỏ) tách hẳn khỏi 13 lớp biển báo")
    ax.axvline(pc[~pc["is_traffic_light"]]["ap50_95"].mean(), color=TEAL, ls="--", lw=1,
               label="TB biển báo")
    ax.axvline(pc[pc["is_traffic_light"]]["ap50_95"].mean(), color=RED, ls="--", lw=1,
               label="TB đèn tín hiệu")
    for y, (v, r) in enumerate(zip(pc["ap50_95"], pc["is_traffic_light"])):
        ax.text(v + 0.005, y, f"{v:.3f}", va="center", fontsize=8)
    ax.set_xlim(0, 1.05)
    ax.legend(loc="lower right")
    _save(fig, "01_per_class_ap.png")


def fig_light_vs_sign():
    lvs = _load("audit_light_vs_sign.csv")
    metrics = ["mean_precision", "mean_recall", "mean_map50", "mean_map50_95"]
    labels = ["Precision", "Recall", "mAP@0.5", "mAP@0.5:0.95"]
    x = np.arange(len(metrics))
    w = 0.38
    fig, ax = plt.subplots(figsize=(8.6, 4.8))
    signs = lvs[lvs["category"] == "Biển báo"].iloc[0]
    lights = lvs[lvs["category"] == "Đèn tín hiệu"].iloc[0]
    ax.bar(x - w/2, [signs[m] for m in metrics], w, label="Biển báo (13 lớp)", color=TEAL)
    ax.bar(x + w/2, [lights[m] for m in metrics], w, label="Đèn tín hiệu (2 lớp)", color=RED)
    for i, m in enumerate(metrics):
        ax.text(i - w/2, signs[m] + 0.01, f"{signs[m]:.3f}", ha="center", fontsize=8)
        ax.text(i + w/2, lights[m] + 0.01, f"{lights[m]:.3f}", ha="center", fontsize=8)
    ax.set_xticks(x); ax.set_xticklabels(labels)
    ax.set_ylim(0, 1.08); ax.set_ylabel("Giá trị trung bình")
    ax.set_title("Đèn tín hiệu kém hơn biển báo trên MỌI chỉ số (khoảng cách lớn nhất ở mAP@.5:.95)")
    ax.legend()
    _save(fig, "02_light_vs_sign.png")


def fig_rarity_vs_ap():
    pc = _load("per_class_metrics.csv")
    fig, ax = plt.subplots(figsize=(9, 5.4))
    # Chỉ chú thích các điểm đáng chú ý để tránh chồng chữ trong cụm biển báo.
    annotate = {"Red Light", "Green Light", "Speed Limit 10", "Speed Limit 110",
                "Stop", "Speed Limit 80"}
    for _, r in pc.iterrows():
        c = RED if r["is_traffic_light"] else TEAL
        ax.scatter(r["test_instances"], r["ap50_95"], s=80, color=c, zorder=3,
                   edgecolor="white", linewidth=0.8)
        if r["class_name"] in annotate:
            dy = -14 if r["class_name"] == "Green Light" else 6
            ax.annotate(r["class_name"], (r["test_instances"], r["ap50_95"]),
                        fontsize=8, xytext=(6, dy), textcoords="offset points")
    au = _load("audit_rarity_vs_ap.csv").iloc[0]
    ax.set_xlabel("Số instance trên tập test")
    ax.set_ylabel("mAP@0.5:0.95")
    ax.set_title(f"Độ hiếm KHÔNG giải thích được hiệu năng\n"
                 f"Spearman ρ = {au['spearman_instances_vs_map5095']:.3f} "
                 f"(p = {au['p_instances_vs_map5095']:.2f}) — không có tương quan")
    ax.scatter([], [], color=RED, label="Đèn tín hiệu")
    ax.scatter([], [], color=TEAL, label="Biển báo")
    ax.legend(loc="upper left")
    ax.set_xlim(-5, 122)
    ax.set_ylim(0.45, 1.02)
    _save(fig, "03_rarity_vs_ap.png")


def fig_leakage():
    lex = _load("audit_leakage_exposure.csv")
    fig, ax = plt.subplots(figsize=(8.8, 4.6))
    colors = [AMBER if "XUYÊN" in s else GREY for s in lex["loai_trung_lap"]]
    ax.barh(lex["loai_trung_lap"], lex["so_dong"], color=colors)
    for y, v in enumerate(lex["so_dong"]):
        ax.text(v + 6, y, f"{v}", va="center", fontsize=9)
    ax.set_xlabel("Số dòng ảnh")
    ax.set_title("Trùng lặp & rò rỉ dữ liệu — chỉ 91 ảnh bị loại khỏi train,\n"
                 "tập valid/test đánh giá vẫn còn trùng lặp xuyên split")
    ax.invert_yaxis()
    _save(fig, "04_leakage.png")


def fig_imbalance():
    imb = _load("imbalance_metrics.csv")
    o = imb[imb["split"] == "overall"].iloc[0]
    fig, ax = plt.subplots(figsize=(7.8, 4.4))
    labels = ["Imbalance ratio\n(÷10 để hiển thị)", "Entropy chuẩn hoá", "Hệ số Gini"]
    vals = [o["imbalance_ratio"] / 10, o["normalized_entropy"], o["gini"]]
    raw = [f"{o['imbalance_ratio']:.1f}", f"{o['normalized_entropy']:.3f}", f"{o['gini']:.3f}"]
    bars = ax.bar(labels, vals, color=[RED, TEAL, AMBER])
    for b, t in zip(bars, raw):
        ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.03, t, ha="center", fontweight="bold")
    ax.set_ylim(0, 4.2)
    ax.set_title("Mất cân bằng lớp toàn corpus (15 lớp, 6.012 instance)")
    _save(fig, "05_imbalance.png")


def fig_size():
    res = _load("resolution_analysis.csv")
    fig, ax = plt.subplots(figsize=(9, 4.8))
    x = np.arange(len(res))
    w = 0.26
    ax.bar(x - w, res["ratio_min_side_lt_4px"] * 100, w, label="< 4 px", color=RED)
    ax.bar(x, res["ratio_min_side_lt_8px"] * 100, w, label="< 8 px", color=AMBER)
    ax.bar(x + w, res["ratio_min_side_lt_16px"] * 100, w, label="< 16 px", color=TEAL)
    ax.set_xticks(x); ax.set_xticklabels([p.replace(" / ", "\n") for p in res["input_policy"]], fontsize=8)
    ax.set_ylabel("% instance có cạnh ngắn nhỏ hơn ngưỡng")
    ax.set_title("Vật thể nhỏ theo độ phân giải đầu vào — tăng imgsz làm giảm mạnh tỷ lệ vật thể tí hon")
    ax.legend()
    _save(fig, "06_small_objects.png")


def fig_precision_recall():
    pc = _load("per_class_metrics.csv")
    fig, ax = plt.subplots(figsize=(8.4, 5.6))
    for _, r in pc.iterrows():
        c = RED if r["is_traffic_light"] else TEAL
        ax.scatter(r["recall"], r["precision"], s=90, color=c, zorder=3, edgecolor="white", linewidth=0.8)
        if r["is_traffic_light"] or r["recall"] < 0.95:
            ax.annotate(r["class_name"], (r["recall"], r["precision"]), fontsize=7.5,
                        xytext=(5, 4), textcoords="offset points")
    ax.set_xlabel("Recall (test)")
    ax.set_ylabel("Precision (test)")
    ax.set_title("Không gian Precision–Recall theo lớp\nHai lớp đèn (đỏ) lệch về phía Recall thấp — đặc trưng lỗi bỏ sót")
    ax.scatter([], [], color=RED, label="Đèn tín hiệu")
    ax.scatter([], [], color=TEAL, label="Biển báo")
    ax.legend(loc="lower left")
    ax.set_xlim(0.7, 1.02)
    ax.set_ylim(0.85, 1.01)
    _save(fig, "07_precision_recall.png")


def fig_valid_test():
    om = _load("overall_metrics.csv").set_index("split")
    metrics = ["precision", "recall", "map50", "map50_95"]
    labels = ["Precision", "Recall", "mAP@0,5", "mAP@0,5:0,95"]
    x = np.arange(len(metrics)); w = 0.38
    fig, ax = plt.subplots(figsize=(8.6, 4.8))
    ax.bar(x - w / 2, [om.loc["valid", m] for m in metrics], w, label="Valid", color=TEAL)
    ax.bar(x + w / 2, [om.loc["test", m] for m in metrics], w, label="Test", color=AMBER)
    for i, m in enumerate(metrics):
        ax.text(i - w / 2, om.loc["valid", m] + 0.008, f"{om.loc['valid', m]:.3f}", ha="center", fontsize=8)
        ax.text(i + w / 2, om.loc["test", m] + 0.008, f"{om.loc['test', m]:.3f}", ha="center", fontsize=8)
    ax.set_xticks(x); ax.set_xticklabels(labels)
    ax.set_ylim(0, 1.08); ax.set_ylabel("Giá trị")
    ax.set_title("Hiệu năng tổng thể trên tập kiểm định và kiểm thử gần như bằng nhau")
    ax.legend()
    _save(fig, "08_valid_test.png")


def run():
    ensure_dir(RESULTS_FIGS)
    print("[figures] Đang vẽ...")
    fig_per_class_ap()
    fig_light_vs_sign()
    fig_rarity_vs_ap()
    fig_leakage()
    fig_imbalance()
    fig_size()
    fig_precision_recall()
    fig_valid_test()
    print("[figures] Xong — 8 hình trong results/figs/.")


if __name__ == "__main__":
    run()
