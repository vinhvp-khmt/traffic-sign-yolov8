"""Trích xuất kết quả thực từ notebook Kaggle -> results/tables/*.csv.

Notebook nguồn (`data/notebook/traffic-yolo-v2-run.ipynb`) đã chạy trên Kaggle T4 và huấn
luyện một YOLOv8s hướng-EDA trên bộ `pkdarabi/cardetection`. Vì đề tài này là PHÂN TÍCH THỨ CẤP
(secondary analysis / kiểm định độ tin cậy) ở chế độ *results-only*, ta không chạy lại YOLO mà
đọc lại các con số đã in ra trong notebook và cố định chúng thành bảng CSV chuẩn. Mọi giá trị ở
đây khớp từng chữ số với output của notebook — không có số nào bịa ra.

    python -m src.extract_notebook_results
"""
from __future__ import annotations

import pandas as pd

from src.utils.paths import RESULTS_TABLES, ensure_dir


# ---------------------------------------------------------------------------
# 1) Tổng quan corpus (cell [30], [31], [48])
# ---------------------------------------------------------------------------
DATASET_SUMMARY = {
    "dataset": "pkdarabi/cardetection",
    "task": "traffic sign & traffic light object detection",
    "n_classes": 15,
    "total_images_scanned": 4969,
    "total_valid_objects": 6012,
    "n_train_images": 3530,
    "n_valid_images": 801,
    "n_test_images": 638,
    "background_image_ratio": 0.0008049909438518816,
    "missing_label_files": 0,
    "invalid_objects": 0,
    "malformed_label_lines": 0,
    "dominant_class": "Red Light",
    "dominant_class_instances": 787,
    "rarest_nonzero_class": "Speed Limit 10",
    "rarest_class_instances": 22,
}

# ---------------------------------------------------------------------------
# 2) Thống kê theo split (cell [13])
# ---------------------------------------------------------------------------
SPLIT_SUMMARY = [
    # split, n_images, n_objects, objects_per_image_mean, bbox_area_ratio_median,
    # bbox_area_px_median, background_images
    ("train", 3530, 4298, 1.217564, 0.050865, 8802.500, 3),
    ("valid", 801, 944, 1.178527, 0.132493, 22928.0, 0),
    ("test", 638, 770, 1.206897, 0.049303, 8532.125, 1),
]

# ---------------------------------------------------------------------------
# 3) Mất cân bằng lớp (cell [14], [30])
# ---------------------------------------------------------------------------
IMBALANCE = [
    # split, total_instances, classes_with_instances, max_class_count,
    # min_positive_class_count, imbalance_ratio, coef_variation, normalized_entropy, gini
    ("overall", 6012, 15, 787, 22, 35.772727, 0.474358, 0.952828, 0.246707),
    ("train", 4298, 15, 585, 19, 30.789474, 0.479729, 0.952752, 0.249977),
    ("valid", 944, 14, 122, 17, 7.176471, 0.475120, 0.946619, 0.258898),
    ("test", 770, 15, 110, 3, 36.666667, 0.481063, 0.953365, 0.247273),
]

# Split-shift (Jensen–Shannon divergence, bits) — cell [14]
SPLIT_SHIFT = [
    ("test", "train", 0.002101),
    ("test", "valid", 0.007622),
    ("train", "valid", 0.005816),
]

# ---------------------------------------------------------------------------
# 4) Phân tích kích thước vật thể theo độ phân giải đầu vào (cell [15])
# ---------------------------------------------------------------------------
RESOLUTION = [
    # input_policy, instances, min_side_median_px, min_side_p10_px,
    # ratio_lt_4px, ratio_lt_8px, ratio_lt_16px, ratio_ge_32px
    # Ba chính sách độ phân giải đầu vào phổ biến (letterbox 640, 1024, short-side 800).
    ("Letterbox 640", 6012, 125.384615, 7.692308, 0.033766, 0.108949, 0.207917, 0.707585),
    ("Letterbox 1024", 6012, 200.615385, 12.307308, 0.007651, 0.049734, 0.132568, 0.760978),
    ("Short-side 800 / max 1333", 6012, 156.730769, 9.615385, 0.020792, 0.080838, 0.173819, 0.731371),
]

# Tỷ lệ small-object tổng thể (cell [30], [33])
SIZE_SIGNALS = {
    "coco_small_object_ratio": 0.31370592149035265,   # <32x32 px COCO
    "relative_tiny_object_ratio": 0.1763140385894877,
    "tiny_ratio_area_lt_0_1pct": 0.18683108422522104,
    "small_ratio_area_lt_1pct": 0.3657515123313169,
    "border_touch_ratio": 0.021457085828343315,
    "yolo640_min_side_lt_8px_ratio": 0.10894876912840985,
    "hires800_min_side_lt_8px_ratio": 0.08083832335329341,
}

# ---------------------------------------------------------------------------
# 5) Trùng lặp & rò rỉ dữ liệu (cell [17], [30], [33])
# ---------------------------------------------------------------------------
LEAKAGE = {
    "exact_duplicate_rows": 460,
    "exact_cross_split_duplicate_rows": 202,
    "exact_cross_split_duplicate_groups": 101,
    "dhash_collision_rows": 799,
    "dhash_cross_split_collision_rows": 423,
    # Notebook chỉ loại 91 ảnh (exact cross-split) khỏi TRAIN, giữ nguyên valid/test khi đánh giá.
    "removed_from_train": 91,
    "original_train_images": 3530,
    "clean_train_images": 3439,
}

# ---------------------------------------------------------------------------
# 6) Metric tổng thể trên test (cell [43], [48]) & valid (cell [42])
# ---------------------------------------------------------------------------
OVERALL_METRICS = [
    # split, images, instances, precision, recall, map50, map50_95
    ("valid", 801, 944, 0.967023, 0.950666, 0.976229, 0.823760),
    ("test", 638, 770, 0.960291, 0.951029, 0.970263, 0.811996),
]

# ---------------------------------------------------------------------------
# 7) Metric theo lớp trên TEST (cell [44]) — bằng chứng cốt lõi của đề tài
# ---------------------------------------------------------------------------
# Bảng đầy đủ, giá trị chính xác cao lấy từ cell [44]; test_images / test_instances lấy nguyên
# từ cell [43] (hai cột "Images", "Instances"). KHÔNG suy đoán giá trị nào.
PER_CLASS = [
    # class_id, class_name, test_images, test_instances, precision, recall, ap50, ap50_95
    (14, "Stop", 50, 50, 0.998511, 1.000000, 0.995000, 0.899068),
    (12, "Speed Limit 80", 60, 61, 0.976662, 1.000000, 0.995000, 0.877975),
    (3, "Speed Limit 100", 45, 46, 0.967628, 1.000000, 0.994149, 0.870534),
    (11, "Speed Limit 70", 52, 53, 0.921754, 0.943396, 0.973431, 0.867052),
    (6, "Speed Limit 20", 46, 46, 0.971896, 0.978261, 0.973261, 0.866319),
    (5, "Speed Limit 120", 40, 44, 0.930134, 1.000000, 0.985494, 0.863016),
    (7, "Speed Limit 30", 60, 60, 0.983039, 0.966005, 0.983833, 0.860135),
    (10, "Speed Limit 60", 45, 45, 1.000000, 0.932761, 0.991676, 0.858572),
    (9, "Speed Limit 50", 47, 50, 0.927570, 0.960000, 0.987289, 0.856111),
    (8, "Speed Limit 40", 51, 53, 0.976522, 0.962264, 0.981858, 0.838153),
    (2, "Speed Limit 10", 2, 3, 1.000000, 0.970734, 0.995000, 0.829500),
    (13, "Speed Limit 90", 33, 34, 0.995382, 0.970588, 0.992895, 0.827129),
    (4, "Speed Limit 110", 21, 21, 0.906076, 0.904762, 0.929527, 0.784466),
    (0, "Green Light", 77, 110, 0.962030, 0.921345, 0.950300, 0.575613),
    (1, "Red Light", 71, 94, 0.887156, 0.755319, 0.825229, 0.506292),
]

# Cấu hình huấn luyện suy ra từ EDA (cell [33]) — để mô tả "instrument"
TRAIN_PLAN = {
    "base_model": "yolov8s.pt", "fallback_model": "yolov8n.pt",
    "stage1_imgsz": 640, "stage1_batch": 16, "stage1_epochs": 60, "stage1_lr0": 0.001,
    "stage2_imgsz": 768, "stage2_batch": 10, "stage2_epochs": 35, "stage2_lr0": 0.0003,
    "optimizer": "AdamW", "cos_lr": True, "fliplr": 0.0, "flipud": 0.0,
    "box": 7.5, "cls": 0.7, "dfl": 1.5, "patience": 15,
    "params_million": 11.13, "gflops": 28.5, "ultralytics": "8.4.95", "device": "Tesla T4",
    "inference_ms_per_image": 14.7,
}


def _write(df: pd.DataFrame, name: str) -> None:
    df.to_csv(RESULTS_TABLES / name, index=False)
    print(f"  wrote {name:<34} ({len(df)} rows)")


def run() -> None:
    ensure_dir(RESULTS_TABLES)
    print("[extract] Ghi bảng CSV chuẩn từ output notebook...")

    _write(pd.DataFrame([DATASET_SUMMARY]), "dataset_summary.csv")

    _write(pd.DataFrame(SPLIT_SUMMARY, columns=[
        "split", "n_images", "n_objects", "objects_per_image_mean",
        "bbox_area_ratio_median", "bbox_area_px_median", "n_background_images"]),
        "split_summary.csv")

    _write(pd.DataFrame(IMBALANCE, columns=[
        "split", "total_instances", "classes_with_instances", "max_class_count",
        "min_positive_class_count", "imbalance_ratio", "coef_variation",
        "normalized_entropy", "gini"]), "imbalance_metrics.csv")

    _write(pd.DataFrame(SPLIT_SHIFT, columns=["split_a", "split_b", "js_divergence_bits"]),
           "split_shift.csv")

    _write(pd.DataFrame(RESOLUTION, columns=[
        "input_policy", "instances", "min_side_median_px", "min_side_p10_px",
        "ratio_min_side_lt_4px", "ratio_min_side_lt_8px", "ratio_min_side_lt_16px",
        "ratio_min_side_ge_32px"]), "resolution_analysis.csv")

    _write(pd.DataFrame([SIZE_SIGNALS]), "size_signals.csv")
    _write(pd.DataFrame([LEAKAGE]), "leakage_summary.csv")

    _write(pd.DataFrame(OVERALL_METRICS, columns=[
        "split", "images", "instances", "precision", "recall", "map50", "map50_95"]),
        "overall_metrics.csv")

    pc = pd.DataFrame(PER_CLASS, columns=[
        "class_id", "class_name", "test_images", "test_instances",
        "precision", "recall", "ap50", "ap50_95"])
    pc["is_traffic_light"] = pc["class_name"].isin(["Red Light", "Green Light"])
    pc["category"] = pc["is_traffic_light"].map({True: "Đèn tín hiệu", False: "Biển báo"})
    _write(pc, "per_class_metrics.csv")

    _write(pd.DataFrame([TRAIN_PLAN]), "train_plan.csv")

    print("[extract] Xong. Tất cả bảng nằm trong results/tables/.")


if __name__ == "__main__":
    run()
