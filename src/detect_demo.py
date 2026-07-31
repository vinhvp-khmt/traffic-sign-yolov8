"""Suy luận YOLO cho tab demo của app — src/detect_demo.py.

Tải trọng số YOLOv8 đã huấn luyện từ Hugging Face và chạy phát hiện trên một ảnh, rồi GẮN mỗi
phát hiện với độ tin cậy đã kiểm định của lớp tương ứng (từ results/tables/per_class_metrics.csv).
Đây là điểm nối giữa demo trực quan và câu chuyện kiểm định: nếu mô hình phát hiện một lớp có
mAP thấp (ví dụ Red Light 0,506), demo sẽ cảnh báo độ tin cậy thấp.

Thiết kế phòng thủ: nếu thiếu ultralytics/torch hoặc không tải được best.pt, hàm trả về trạng
thái rõ ràng để app hiển thị hướng dẫn thay vì lỗi.
"""
from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlretrieve

import numpy as np
import pandas as pd

from src.utils.paths import REPO_ROOT, RESULTS_TABLES

WEIGHTS_URL = (
    "https://huggingface.co/datasets/vancevo/traffic-sign-yolov8/"
    "resolve/main/best.pt"
)
WEIGHTS_CACHE = REPO_ROOT / ".cache" / "weights" / "best.pt"

# Bảng màu cố định theo class_id để ảnh tái lập.
def class_color(cid: int):
    import random
    rng = random.Random(1000 + int(cid))
    return tuple(rng.randint(40, 230) for _ in range(3))


def download_weights() -> Path:
    """Tải best.pt từ Hugging Face vào cache local nếu chưa có."""
    override = os.environ.get("TYA_WEIGHTS")
    if override:
        path = Path(override).expanduser()
        if path.exists():
            return path
        raise FileNotFoundError(f"TYA_WEIGHTS không tồn tại: {path}")

    if WEIGHTS_CACHE.exists() and WEIGHTS_CACHE.stat().st_size > 0:
        return WEIGHTS_CACHE

    WEIGHTS_CACHE.parent.mkdir(parents=True, exist_ok=True)
    tmp = WEIGHTS_CACHE.with_suffix(".pt.tmp")
    try:
        urlretrieve(WEIGHTS_URL, tmp)
        tmp.replace(WEIGHTS_CACHE)
    except (OSError, URLError) as e:
        if tmp.exists():
            tmp.unlink()
        raise RuntimeError(f"Không tải được best.pt từ Hugging Face: {e}") from e
    return WEIGHTS_CACHE


def backend_available() -> tuple[bool, str]:
    try:
        import torch  # noqa: F401
        import ultralytics  # noqa: F401
    except Exception as e:
        return False, f"Thiếu thư viện: {e}. Cài: pip install ultralytics"
    try:
        weights = download_weights()
    except Exception as e:
        return False, str(e)
    return True, f"sẵn sàng ({weights})"


@lru_cache(maxsize=1)
def _load_model():
    from ultralytics import YOLO
    return YOLO(download_weights())


@lru_cache(maxsize=1)
def reliability_map() -> dict:
    """{class_name: (ap50_95, verdict, is_light)} từ bảng kiểm định."""
    pc = pd.read_csv(RESULTS_TABLES / "per_class_metrics.csv")
    out = {}
    for _, r in pc.iterrows():
        ap = float(r["ap50_95"])
        verdict = ("Thấp" if ap < 0.65 else "Trung bình" if ap < 0.80 else "Cao")
        out[r["class_name"]] = (ap, verdict, bool(r["is_traffic_light"]))
    return out


def detect(image_bgr: np.ndarray, conf: float = 0.25, imgsz: int = 640) -> dict:
    """Chạy phát hiện. Trả về dict: {ok, image (RGB có hộp), detections[list], note}."""
    import cv2

    ok, msg = backend_available()
    if not ok:
        return {"ok": False, "note": msg, "detections": [], "image": None}

    model = _load_model()
    res = model.predict(image_bgr, imgsz=imgsz, conf=conf, verbose=False)[0]
    names = model.names
    rel = reliability_map()

    rgb = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2RGB).copy()
    h, w = rgb.shape[:2]
    lw = max(2, round(max(h, w) / 400))
    dets = []
    for b in res.boxes:
        cid = int(b.cls)
        name = names[cid]
        cf = float(b.conf)
        x1, y1, x2, y2 = [int(v) for v in b.xyxy[0].tolist()]
        color = class_color(cid)
        cv2.rectangle(rgb, (x1, y1), (x2, y2), color, lw)
        label = f"{name} {cf:.2f}"
        (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.5, 1)
        cv2.rectangle(rgb, (x1, y1 - th - 6), (x1 + tw + 4, y1), color, -1)
        cv2.putText(rgb, label, (x1 + 2, y1 - 4), cv2.FONT_HERSHEY_SIMPLEX, 0.5,
                    (255, 255, 255), 1, cv2.LINE_AA)
        ap, verdict, is_light = rel.get(name, (float("nan"), "?", False))
        dets.append({
            "Lớp": name, "Độ tin YOLO": round(cf, 3),
            "mAP@.5:.95 (kiểm định)": round(ap, 3),
            "Độ tin cậy lớp": verdict,
            "Nhóm": "Đèn tín hiệu" if is_light else "Biển báo",
        })
    note = (f"Phát hiện {len(dets)} đối tượng." if dets
            else "Không phát hiện đối tượng nào ở ngưỡng conf hiện tại.")
    return {"ok": True, "note": note, "detections": dets, "image": rgb}


def run() -> None:
    """Tự kiểm tra nhanh trên một ảnh mẫu."""
    import cv2
    ok, msg = backend_available()
    print("[detect_demo] backend:", msg)
    if not ok:
        return
    samples = sorted((REPO_ROOT / "app" / "samples").glob("*.jpg"))
    if not samples:
        print("[detect_demo] không có ảnh mẫu.")
        return
    img = cv2.imread(str(samples[0]))
    out = detect(img)
    print("[detect_demo]", out["note"])
    for d in out["detections"]:
        print("   ", d)


if __name__ == "__main__":
    run()
