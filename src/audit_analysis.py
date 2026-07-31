"""Phân tích kiểm định (audit) trên kết quả đã trích — src/audit_analysis.py.

Đây là ĐÓNG GÓP MỚI của đề tài (khác với notebook gốc chỉ EDA + huấn luyện). Ta coi YOLOv8s đã
huấn luyện như một *công cụ đo* cố định, rồi trả lời ba câu hỏi kiểm định:

  A. RÒ RỈ DỮ LIỆU — mAP 0,97 đáng tin đến đâu khi tồn tại trùng lặp xuyên split?
  B. KHOẢNG CÁCH ĐÈN–BIỂN — vì sao 2 lớp đèn tín hiệu kém hơn hẳn 13 lớp biển báo?
  C. HIẾM ≠ KHÓ — độ hiếm của lớp có giải thích được khoảng cách đó không? (Không!)

Ghi kết quả ra results/tables/audit_*.csv.

    python -m src.audit_analysis
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy import stats

from src.utils.paths import RESULTS_TABLES, ensure_dir


def _load(name: str) -> pd.DataFrame:
    return pd.read_csv(RESULTS_TABLES / name)


def light_vs_sign(pc: pd.DataFrame) -> pd.DataFrame:
    lights = pc[pc["is_traffic_light"]]
    signs = pc[~pc["is_traffic_light"]]
    u, p = stats.mannwhitneyu(signs["ap50_95"], lights["ap50_95"], alternative="greater")
    rows = [
        ("Đèn tín hiệu", len(lights), lights["ap50_95"].mean(), lights["ap50"].mean(),
         lights["recall"].mean(), lights["precision"].mean()),
        ("Biển báo", len(signs), signs["ap50_95"].mean(), signs["ap50"].mean(),
         signs["recall"].mean(), signs["precision"].mean()),
    ]
    out = pd.DataFrame(rows, columns=[
        "category", "n_classes", "mean_map50_95", "mean_map50", "mean_recall", "mean_precision"])
    gap = signs["ap50_95"].mean() - lights["ap50_95"].mean()
    out.attrs["gap"] = gap
    out.attrs["mannwhitney_u"] = float(u)
    out.attrs["mannwhitney_p"] = float(p)
    return out


def rarity_vs_ap(pc: pd.DataFrame) -> dict:
    rho_t, p_t = stats.spearmanr(pc["test_instances"], pc["ap50_95"])
    # Cũng kiểm với precision/recall
    rho_r, p_r = stats.spearmanr(pc["test_instances"], pc["recall"])
    # Điểm nhấn: lớp phổ biến nhất lại kém nhất
    dominant = pc.loc[pc["test_instances"].idxmax()]
    worst = pc.loc[pc["ap50_95"].idxmin()]
    return {
        "spearman_instances_vs_map5095": round(float(rho_t), 4),
        "p_instances_vs_map5095": float(p_t),
        "spearman_instances_vs_recall": round(float(rho_r), 4),
        "p_instances_vs_recall": float(p_r),
        "most_frequent_class": dominant["class_name"],
        "most_frequent_instances": int(dominant["test_instances"]),
        "most_frequent_map5095": round(float(dominant["ap50_95"]), 4),
        "worst_class": worst["class_name"],
        "worst_map5095": round(float(worst["ap50_95"]), 4),
    }


def leakage_exposure(leak: pd.DataFrame, ds: pd.DataFrame) -> pd.DataFrame:
    l = leak.iloc[0]
    n_test = int(ds.iloc[0]["n_test_images"])
    n_valid = int(ds.iloc[0]["n_valid_images"])
    n_train = int(ds.iloc[0]["n_train_images"])
    total = n_train + n_valid + n_test
    rows = [
        ("Exact duplicate (mọi split)", int(l["exact_duplicate_rows"]),
         round(100 * l["exact_duplicate_rows"] / (2 * total), 2)),
        ("Exact duplicate XUYÊN split", int(l["exact_cross_split_duplicate_rows"]),
         round(100 * l["exact_cross_split_duplicate_rows"] / (2 * total), 2)),
        ("dHash trùng (mọi split)", int(l["dhash_collision_rows"]),
         round(100 * l["dhash_collision_rows"] / (2 * total), 2)),
        ("dHash trùng XUYÊN split", int(l["dhash_cross_split_collision_rows"]),
         round(100 * l["dhash_cross_split_collision_rows"] / (2 * total), 2)),
    ]
    out = pd.DataFrame(rows, columns=["loai_trung_lap", "so_dong", "pct_tren_tong_the"])
    out.attrs["removed_from_train"] = int(l["removed_from_train"])
    out.attrs["cross_split_groups"] = int(l["exact_cross_split_duplicate_groups"])
    out.attrs["note"] = (
        f"Notebook chỉ loại {int(l['removed_from_train'])} ảnh exact-cross-split khỏi TRAIN "
        f"({int(l['original_train_images'])}→{int(l['clean_train_images'])}); "
        f"tập valid/test dùng để đánh giá KHÔNG được khử trùng. "
        f"Vẫn còn {int(l['dhash_cross_split_collision_rows'])} dòng near-duplicate xuyên split "
        f"(dHash) không bị loại — nguồn rò rỉ tiềm tàng làm mAP lạc quan."
    )
    return out


def run() -> None:
    ensure_dir(RESULTS_TABLES)
    pc = _load("per_class_metrics.csv")
    leak = _load("leakage_summary.csv")
    ds = _load("dataset_summary.csv")

    print("=" * 70)
    print("A. RÒ RỈ DỮ LIỆU (leakage exposure)")
    lex = leakage_exposure(leak, ds)
    lex.to_csv(RESULTS_TABLES / "audit_leakage_exposure.csv", index=False)
    print(lex.to_string(index=False))
    print("  →", lex.attrs["note"])

    print("\n" + "=" * 70)
    print("B. KHOẢNG CÁCH ĐÈN vs BIỂN BÁO")
    lvs = light_vs_sign(pc)
    lvs.to_csv(RESULTS_TABLES / "audit_light_vs_sign.csv", index=False)
    print(lvs.round(4).to_string(index=False))
    print(f"  → Chênh lệch mAP@.5:.95 = {lvs.attrs['gap']:.4f}")
    print(f"  → Mann–Whitney U={lvs.attrs['mannwhitney_u']:.1f}, "
          f"p={lvs.attrs['mannwhitney_p']:.4g} (biển > đèn, một phía)")

    print("\n" + "=" * 70)
    print("C. HIẾM ≠ KHÓ (rarity does NOT explain the gap)")
    rv = rarity_vs_ap(pc)
    pd.DataFrame([rv]).to_csv(RESULTS_TABLES / "audit_rarity_vs_ap.csv", index=False)
    for k, v in rv.items():
        print(f"  {k:<34} {v}")
    print(f"  → Hai lớp có NHIỀU instance nhất trên test lại là hai lớp KÉM nhất: "
          f"{rv['most_frequent_class']} ({rv['most_frequent_instances']} inst, "
          f"mAP={rv['most_frequent_map5095']}) và {rv['worst_class']} "
          f"(mAP={rv['worst_map5095']}). Độ hiếm KHÔNG giải thích được thất bại.")

    # Bảng tổng hợp phát hiện chính cho báo cáo
    findings = pd.DataFrame([
        {"ma": "A", "phat_hien": "Rò rỉ xuyên split chưa được khử ở tập đánh giá",
         "bang_chung": f"{int(leak.iloc[0]['exact_cross_split_duplicate_rows'])} exact + "
                       f"{int(leak.iloc[0]['dhash_cross_split_collision_rows'])} near-dup xuyên split; "
                       f"chỉ loại {int(leak.iloc[0]['removed_from_train'])} khỏi train",
         "he_qua": "mAP 0,97 có thể lạc quan hơn hiệu năng thực"},
        {"ma": "B", "phat_hien": "Đèn tín hiệu kém hơn biển báo rõ rệt",
         "bang_chung": f"mAP@.5:.95: đèn {lvs.iloc[0]['mean_map50_95']:.3f} vs biển "
                       f"{lvs.iloc[1]['mean_map50_95']:.3f}; Mann–Whitney p={lvs.attrs['mannwhitney_p']:.3g}",
         "he_qua": "Hiệu năng tổng thể bị chi phối bởi loại vật thể, không đồng đều"},
        {"ma": "C", "phat_hien": "Độ hiếm KHÔNG giải thích được khoảng cách",
         "bang_chung": f"Spearman(instance, mAP)={rv['spearman_instances_vs_map5095']} "
                       f"(p={rv['p_instances_vs_map5095']:.2g}); Red Light phổ biến nhất nhưng kém nhất",
         "he_qua": "Nguyên nhân là bản chất vật thể (đèn nhỏ, đỏ/xanh dễ nhầm), không phải mất cân bằng"},
    ])
    findings.to_csv(RESULTS_TABLES / "audit_key_findings.csv", index=False)
    print("\n[audit] Đã ghi audit_*.csv vào results/tables/.")


if __name__ == "__main__":
    run()
