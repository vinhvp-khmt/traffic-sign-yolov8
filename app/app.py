"""Ứng dụng web — Kiểm định độ tin cậy bộ phát hiện biển báo YOLOv8.

    streamlit run app/app.py

Dashboard trình bày phân tích thứ cấp trên kết quả notebook Kaggle: rò rỉ dữ liệu, khoảng cách
đèn–biển, và bằng chứng "hiếm ≠ khó". Mọi số liệu đọc từ results/tables/ (số thực của notebook).
"""
from __future__ import annotations

import sys
import uuid
from pathlib import Path

import pandas as pd
import streamlit as st

# Cho phép chạy `streamlit run app/app.py` từ gốc repo.
REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from src.utils.paths import RESULTS_FIGS, RESULTS_TABLES  # noqa: E402

st.set_page_config(page_title="Kiểm định YOLOv8 biển báo", layout="wide")

DEFAULT_DEMO_CONF = 0.25
FIGURE_WIDTH = 760
DEMO_IMAGE_MAX_SIZE = (720, 460)
VIDEO_MAX_SIDE = 720


def table(name, **kw):
    p = RESULTS_TABLES / name
    if not p.exists():
        st.warning(f"Thiếu {name} — chạy `python -m src.run_pipeline` trước.")
        return None
    df = pd.read_csv(p)
    st.dataframe(df, width="stretch", hide_index=True, **kw)
    return df


def fig(pattern, caption=None, width=FIGURE_WIDTH):
    matches = sorted(RESULTS_FIGS.glob(pattern))
    if matches:
        st.image(str(matches[0]), caption=caption, width=width)
    else:
        st.info(f"Thiếu hình {pattern}")


def fit_demo_image(image):
    from PIL import Image

    out = Image.fromarray(image)
    out.thumbnail(DEMO_IMAGE_MAX_SIZE)
    return out


def resize_frame(frame_bgr, cv2):
    h, w = frame_bgr.shape[:2]
    scale = min(1.0, VIDEO_MAX_SIDE / max(h, w))
    if scale == 1.0:
        return frame_bgr
    size = (int(w * scale), int(h * scale))
    return cv2.resize(frame_bgr, size, interpolation=cv2.INTER_AREA)


def process_video(uploaded_video, detect, cv2):
    cache_dir = REPO / ".cache" / "video_outputs"
    cache_dir.mkdir(parents=True, exist_ok=True)
    token = uuid.uuid4().hex
    suffix = Path(uploaded_video.name).suffix.lower() or ".mp4"
    input_path = cache_dir / f"{token}_input{suffix}"
    output_path = cache_dir / f"{token}_output.mp4"
    input_path.write_bytes(uploaded_video.getbuffer())

    cap = cv2.VideoCapture(str(input_path))
    if not cap.isOpened():
        raise RuntimeError("Không đọc được video đã tải lên.")

    fps = cap.get(cv2.CAP_PROP_FPS)
    fps = fps if fps and fps > 0 else 24
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    ok, frame = cap.read()
    if not ok:
        cap.release()
        raise RuntimeError("Video không có khung hình hợp lệ.")

    frame = resize_frame(frame, cv2)
    h, w = frame.shape[:2]
    writer = cv2.VideoWriter(
        str(output_path),
        cv2.VideoWriter_fourcc(*"mp4v"),
        fps,
        (w, h),
    )
    if not writer.isOpened():
        cap.release()
        raise RuntimeError("Không tạo được video output.")

    progress = st.progress(0, text="Đang xử lý video...")
    status = st.empty()
    frame_index = 0
    total_detections = 0
    light_frames = 0

    while ok:
        frame = resize_frame(frame, cv2)
        out = detect(frame, conf=DEFAULT_DEMO_CONF)
        annotated = out["image"]
        detections = out["detections"]
        total_detections += len(detections)
        if any(d["Nhóm"] == "Đèn tín hiệu" for d in detections):
            light_frames += 1

        writer.write(cv2.cvtColor(annotated, cv2.COLOR_RGB2BGR))
        frame_index += 1
        if total_frames:
            progress.progress(min(frame_index / total_frames, 1.0), text="Đang xử lý video...")
        status.caption(f"Đã xử lý {frame_index:,} khung hình")
        ok, frame = cap.read()

    cap.release()
    writer.release()
    progress.empty()
    status.empty()
    return output_path, {
        "Số khung hình": frame_index,
        "Tổng phát hiện": total_detections,
        "Khung hình có đèn tín hiệu": light_frames,
    }


st.title("🚦 Kiểm định độ tin cậy — YOLOv8 phát hiện biển báo (pkdarabi/cardetection)")

# KPI hàng đầu
ds = pd.read_csv(RESULTS_TABLES / "dataset_summary.csv").iloc[0]
om = pd.read_csv(RESULTS_TABLES / "overall_metrics.csv").set_index("split")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Ảnh / Object", f"{ds['total_images_scanned']:,} / {ds['total_valid_objects']:,}")
c2.metric("Số lớp", int(ds["n_classes"]))
c3.metric("mAP@0.5 (test)", f"{om.loc['test','map50']:.3f}")
c4.metric("mAP@0.5:0.95 (test)", f"{om.loc['test','map50_95']:.3f}")
c5.metric("Imbalance ratio", f"{ds['dominant_class_instances']/ds['rarest_class_instances']:.1f}")

(
    tab_intro,
    tab_analysis,
    tab_demo,
) = st.tabs([
    "📌 Giới thiệu đề tài", "📊 Phân tích dữ liệu", "🎬 Mô phỏng ứng dụng",
])

with tab_intro:
    st.subheader("Giới thiệu đề tài")
    st.markdown(
        "**Tên đề tài:** Kiểm định độ tin cậy bộ phát hiện biển báo giao thông YOLOv8 "
        "trên bộ dữ liệu `pkdarabi/cardetection`."
    )
    st.markdown(
        "Đề tài phân tích lại kết quả từ notebook Kaggle đã huấn luyện YOLOv8s, xem mô hình "
        "như một công cụ đo cố định. Trọng tâm không phải thay mô hình để tăng điểm, mà là "
        "kiểm tra xem chỉ số mAP cao có đáng tin, có bị ảnh hưởng bởi rò rỉ dữ liệu, và có "
        "đồng đều giữa các nhóm lớp biển báo/đèn tín hiệu hay không."
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Bộ dữ liệu", "4.969 ảnh")
    c2.metric("Đối tượng hợp lệ", "6.012")
    c3.metric("Số lớp", "15")

    st.subheader("Thông tin học viên")
    students = pd.DataFrame([
        {"Họ và tên": "Huỳnh Phát Lợi", "Mã số học viên": "KHMT836016"},
        {"Họ và tên": "Đoàn Huỳnh Thanh Tú", "Mã số học viên": "KHMT836034"},
        {"Họ và tên": "Võ Phú Vinh", "Mã số học viên": "KHMT836036"},
    ])
    st.dataframe(students, width="stretch", hide_index=True)

with tab_analysis:
    st.subheader("Ba phát hiện kiểm định")
    kf = pd.read_csv(RESULTS_TABLES / "audit_key_findings.csv")
    for _, r in kf.iterrows():
        with st.container(border=True):
            st.markdown(f"**[{r['ma']}] {r['phat_hien']}**")
            st.markdown(f"- *Bằng chứng:* {r['bang_chung']}")
            st.markdown(f"- *Hệ quả:* {r['he_qua']}")

    st.subheader("Hiệu năng theo lớp")
    fig("01_*.png")

    with st.expander("A · Rò rỉ dữ liệu xuyên split"):
        st.markdown(
            "Notebook tự phát hiện trùng lặp nhưng **chỉ loại 91 ảnh khỏi train**; "
            "tập valid/test dùng để đánh giá **không** được khử trùng. Đây là nguồn khiến "
            "mAP 0,97 có thể lạc quan hơn hiệu năng thực.")
        fig("04_*.png")
        table("audit_leakage_exposure.csv")

    with st.expander("B · Khoảng cách đèn tín hiệu vs biển báo"):
        lvs = pd.read_csv(RESULTS_TABLES / "audit_light_vs_sign.csv")
        gap = (lvs[lvs.category == "Biển báo"]["mean_map50_95"].iloc[0]
               - lvs[lvs.category == "Đèn tín hiệu"]["mean_map50_95"].iloc[0])
        st.metric("Chênh lệch mAP@0.5:0.95", f"{gap:.3f}",
                  help="Mann–Whitney p = 0,0095 (biển > đèn)")
        fig("02_*.png")
        table("audit_light_vs_sign.csv")

    with st.expander("C · Độ hiếm KHÔNG giải thích được thất bại"):
        rv = pd.read_csv(RESULTS_TABLES / "audit_rarity_vs_ap.csv").iloc[0]
        st.markdown(
            f"Spearman(số instance, mAP) = **{rv['spearman_instances_vs_map5095']:.3f}** "
            f"(p = {rv['p_instances_vs_map5095']:.2f}) — không có tương quan. "
            f"Hai lớp **nhiều instance nhất** trên test (Green Light, Red Light) lại **kém nhất**. "
            f"Nguyên nhân là bản chất vật thể, không phải mất cân bằng.")
        fig("03_*.png")

    st.subheader("Bảng số liệu chi tiết")
    st.markdown("Các bảng dưới đây là dữ liệu nền dùng để sinh biểu đồ và kiểm định trong phần trên.")
    st.markdown("**Metric theo lớp trên tập test**")
    table("per_class_metrics.csv")
    st.markdown("**Mất cân bằng lớp**")
    fig("05_*.png")
    table("imbalance_metrics.csv")
    st.markdown("**Vật thể nhỏ theo độ phân giải**")
    fig("06_*.png")
    table("resolution_analysis.csv")

with tab_demo:
    st.subheader("🎬 Mô phỏng ứng dụng — YOLOv8 gắn với độ tin cậy đã kiểm định")
    st.markdown(
        "Tải một ảnh giao thông (hoặc chọn ảnh mẫu). Mô hình YOLOv8 đã huấn luyện sẽ vẽ hộp; "
        "với **mỗi** phát hiện, demo đối chiếu **độ tin cậy đã kiểm định** của lớp đó. Nếu là lớp "
        "đèn tín hiệu (mAP thấp), demo **cảnh báo** — đúng câu chuyện kiểm định của đề tài.")

    import numpy as np
    try:
        import cv2
        from src.detect_demo import backend_available, detect
        ok_backend, msg = backend_available()
    except Exception as e:  # pragma: no cover
        ok_backend, msg = False, f"Thiếu thư viện: {e}"

    if not ok_backend:
        st.warning(f"Không chạy được suy luận trực tiếp: {msg}")
        st.markdown("Cài `pip install ultralytics opencv-python-headless`; app sẽ tự tải "
                    "`best.pt` từ Hugging Face khi chạy suy luận. Dưới đây là hai ảnh demo "
                    "đã kết xuất sẵn:")
        c1, c2 = st.columns(2)
        with c1:
            fig("demo_sign.png", "Biển báo — độ tin cậy CAO", width=520)
        with c2:
            fig("demo_light.png", "Đèn đỏ — độ tin cậy THẤP (bị gắn cờ)", width=520)
    else:
        image_tab, video_tab = st.tabs(["Ảnh", "Video"])
        sample_dir = REPO / "app" / "samples"
        samples = sorted(sample_dir.glob("*.jpg")) if sample_dir.exists() else []

        with image_tab:
            up = st.file_uploader("Tải ảnh (JPG/PNG)", type=["jpg", "jpeg", "png"])
            pick = None
            if samples:
                names = ["(không chọn)"] + [p.name for p in samples]
                sel = st.selectbox("…hoặc chọn ảnh mẫu", names)
                if sel != "(không chọn)":
                    pick = sample_dir / sel

            img_bgr = None
            if up is not None:
                arr = np.frombuffer(up.read(), np.uint8)
                img_bgr = cv2.imdecode(arr, cv2.IMREAD_COLOR)
            elif pick is not None:
                img_bgr = cv2.imread(str(pick))

            if img_bgr is None:
                st.info("Chọn ảnh mẫu hoặc tải ảnh lên để chạy phát hiện.")
            else:
                with st.spinner("Đang chạy YOLOv8…"):
                    out = detect(img_bgr, conf=DEFAULT_DEMO_CONF)
                c1, c2 = st.columns([3, 2])
                with c1:
                    if out["image"] is not None:
                        st.image(fit_demo_image(out["image"]), caption=out["note"], width="content")
                with c2:
                    if out["detections"]:
                        dfd = pd.DataFrame(out["detections"])
                        st.dataframe(dfd, width="stretch", hide_index=True)
                        lights = [d for d in out["detections"] if d["Nhóm"] == "Đèn tín hiệu"]
                        if lights:
                            st.error(
                                f"⚠️ Phát hiện {len(lights)} đối tượng thuộc nhóm **đèn tín hiệu** — "
                                f"nhóm chỉ đạt mAP@.5:.95 ≈ 0,54 trong kiểm định. Kết quả cần thận trọng.")
                        else:
                            st.success("✓ Toàn bộ phát hiện thuộc nhóm biển báo — nhóm có độ tin cậy cao (mAP ≈ 0,85).")
                    else:
                        st.info(out["note"])

        with video_tab:
            video = st.file_uploader("Tải video (MP4/MOV/AVI)", type=["mp4", "mov", "avi"])
            if video is None:
                st.info("Tải video lên để chạy mô phỏng phát hiện.")
            else:
                with st.spinner("Đang chạy YOLOv8 trên video…"):
                    try:
                        output_path, summary = process_video(video, detect, cv2)
                    except Exception as e:
                        st.error(str(e))
                    else:
                        st.video(output_path.read_bytes(), format="video/mp4")
                        st.dataframe(pd.DataFrame([summary]), width="stretch", hide_index=True)

st.divider()
st.caption("Nhóm: Huỳnh Phát Lợi · Đoàn Huỳnh Thanh Tú · Võ Phú Vinh.")
