# traffic-sign-yolov8

## Kiểm định độ tin cậy bộ phát hiện biển báo YOLOv8 (data-centric)

Đề tài **môn Phân tích dữ liệu** — một phân tích **kiểm định độ tin cậy** (reliability audit)
đối với một bộ phát hiện biển báo giao thông **YOLOv8** trên bộ dữ liệu
`pkdarabi/cardetection` (15 lớp, 4.969 ảnh, 6.012 đối tượng).

Ý tưởng cốt lõi: khi độ chính xác tổng thể (mAP@0,5 = 0,970) đã rất cao, câu hỏi có giá trị
không còn là *“mô hình nào tốt hơn”* mà là *“con số đó có đáng tin và đồng đều không”*. Đề tài
coi mô hình YOLOv8 đã huấn luyện là **công cụ đo cố định** và phân tích lại kết quả nó sinh ra
(chế độ *results-only* — không huấn luyện lại).

## Ba phát hiện chính

| | Phát hiện | Bằng chứng |
|---|---|---|
| **A** | Rò rỉ dữ liệu chưa khử ở tập đánh giá | 202 exact + 423 near-dup xuyên split; chỉ 91 ảnh bị loại khỏi train |
| **B** | Đèn tín hiệu kém hơn biển báo rõ rệt | mAP@.5:.95 đèn 0,541 vs biển 0,854 (chênh 0,313; Mann–Whitney p = 0,0095) |
| **C** | Độ hiếm KHÔNG giải thích được thất bại | Spearman ρ = 0,032 (p = 0,91); hai lớp nhiều đối tượng nhất lại kém nhất |

## Cấu trúc

```
traffic-yolo-analysis/
├── data/notebook/         notebook Kaggle nguồn (.ipynb) — số liệu gốc
├── src/
│   ├── extract_notebook_results.py   trích output notebook -> results/tables/*.csv
│   ├── audit_analysis.py             kiểm định A/B/C -> results/tables/audit_*.csv
│   ├── make_figures.py               vẽ hình -> results/figs/
│   ├── detect_demo.py                suy luận YOLO + gắn độ tin cậy kiểm định (cho app)
│   ├── export_report_bundle.py       gom số liệu -> results/report_bundle.json
│   └── run_pipeline.py               chạy cả ba bước
├── app/app.py             ứng dụng web Streamlit (có tab DEMO phát hiện trực tiếp)
├── app/samples/           ảnh mẫu cho demo (có ảnh chứa đèn tín hiệu)
├── models/                thư mục local cho trọng số nếu cần override thủ công
├── tests/                 pytest (6/6 đạt)
├── results/tables|figs/   sinh tự động
└── reports/
    ├── BaoCao_KiemDinh_YOLOv8_BienBao.docx   báo cáo Word (chuẩn luận văn)
    ├── latex/ (main.pdf)                     báo cáo LaTeX (XeLaTeX)
    ├── Slide_KiemDinh_YOLOv8_BienBao.pptx    slide trình bày (có văn thuyết trình)
    ├── build_docx.js · build_slides.js       mã sinh báo cáo/slide
    └── slides/figures · latex/figures        bản sao hình
```

## Chạy

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

python -m src.run_pipeline          # trích xuất -> kiểm định -> vẽ hình
python -m src.export_report_bundle  # gom số liệu cho báo cáo
pytest                              # kiểm thử
streamlit run app/app.py            # ứng dụng web (mở tab "Demo phát hiện")
```

**Tab Demo phát hiện:** tải một ảnh giao thông (hoặc chọn ảnh mẫu trong `app/samples/`) →
YOLOv8 tự tải trọng số từ
`https://huggingface.co/datasets/vancevo/traffic-sign-yolov8/resolve/main/best.pt` và vẽ hộp →
mỗi phát hiện được gắn **độ tin cậy đã kiểm định** của lớp.
Nếu phát hiện thuộc nhóm đèn tín hiệu (mAP thấp), app **cảnh báo** kết quả cần thận trọng — nối
demo trực quan với câu chuyện kiểm định. Chạy trên CPU; model tải về được cache tại
`.cache/weights/best.pt`. Nếu thiếu `ultralytics` hoặc không tải được model, app tự hiện hai ảnh
demo kết xuất sẵn thay vì báo lỗi. Có thể override đường dẫn bằng biến môi trường `TYA_WEIGHTS`.

**Nhóm:** Huỳnh Phát Lợi (KHMT836016) · Đoàn Huỳnh Thanh Tú (KHMT836034) · Võ Phú Vinh (KHMT836036)
