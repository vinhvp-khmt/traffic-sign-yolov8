/**
 * Báo cáo Word (.docx) chuẩn luận văn — Kiểm định độ tin cậy YOLOv8 phát hiện biển báo.
 * Đề tài ĐỘC LẬP môn Phân tích dữ liệu (không tham chiếu đồ án nào khác).
 * Đọc số liệu từ results/report_bundle.json; hình từ results/figs/.
 *
 *   node build_docx.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  ImageRun, PageBreak, Footer, Header, PageNumber, TableOfContents, LevelFormat,
} = require("./node_modules/docx");

const ROOT = path.resolve(__dirname, "..");
const B = JSON.parse(fs.readFileSync(path.join(ROOT, "results/report_bundle.json"), "utf8"));
const FIG = path.join(ROOT, "results/figs");

const NAVY = "1F3864", TEAL = "1C7293", RED = "C00000", INK = "1A1A1A", GREY = "5A5A5A", GREEN = "2C7A4B";
const HF = "Times New Roman";
const BODY = 26;   // 13pt
const LH = 360;    // 1.5 line spacing (240 = single)

// ---- helpers ----
const vn = (x, d = 0) => Number(x).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }).replace(/,/g, " ").replace(/\./g, ",");
const vnp = (x, d = 3) => String(Number(x).toFixed(d)).replace(".", ",");
const pctv = (x, d = 2) => vnp(100 * x, d) + "%";

function run(t, o = {}) { return new TextRun({ text: String(t), font: HF, size: o.size || BODY, bold: o.bold, italics: o.italics, color: o.color || INK, allCaps: o.caps }); }
function P(text, o = {}) {
  const runs = Array.isArray(text) ? text : [run(text, o)];
  return new Paragraph({ children: runs, alignment: o.align ?? AlignmentType.JUSTIFIED,
    spacing: { after: o.after ?? 120, line: o.line ?? LH, before: o.before ?? 0 },
    indent: o.indent ?? (o.noIndent ? undefined : { firstLine: 709 }) });
}
// ---- Theo mẫu LaTeX của trường: tiêu đề ĐEN đậm, đánh số hình/bảng theo chương ----
const BLACK = "000000";
let CHAP = 0, FIGN = 0, TABN = 0;   // FIGN/TABN reset mỗi chương
// Định dạng chương "CHƯƠNG n : TÊN CHƯƠNG IN HOA" (LARGE, đậm, đen, canh trái).
function H1(text, num) {
  if (num) { CHAP = num; FIGN = 0; TABN = 0; }
  const label = num ? `CHƯƠNG ${num} :  ${String(text).toUpperCase()}` : String(text).toUpperCase();
  return new Paragraph({ heading: HeadingLevel.HEADING_1,
    alignment: num ? AlignmentType.LEFT : AlignmentType.CENTER,
    spacing: { before: 300, after: 200 },
    children: [new TextRun({ text: label, font: HF, size: 32, bold: true, color: BLACK })] });
}
// Mục 16pt đậm đen.
function H2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 120 },
    children: [new TextRun({ text, font: HF, size: 32, bold: true, color: BLACK })] });
}
// Tiểu mục 15pt đậm đen.
function H3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 100 },
    children: [new TextRun({ text, font: HF, size: 30, bold: true, color: BLACK })] });
}
function bullet(text, o = {}) {
  const runs = Array.isArray(text) ? text : [run(text, o)];
  return new Paragraph({ children: runs, bullet: { level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: LH } });
}
function numbered(text, ref = "n") {
  const runs = Array.isArray(text) ? text : [run(text)];
  return new Paragraph({ children: runs, numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: LH } });
}
// Hình: chú thích DƯỚI, đánh số theo chương "Hình c.n." — đậm nhãn.
function figure(file, w, h, caption) {
  FIGN += 1;
  const data = fs.readFileSync(path.join(FIG, file));
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 },
      children: [new ImageRun({ type: "png", data, transformation: { width: w, height: h } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [
        new TextRun({ text: `Hình ${CHAP}.${FIGN}. `, font: HF, size: 24, bold: true, color: BLACK }),
        new TextRun({ text: caption, font: HF, size: 24, color: BLACK }),
      ] }),
  ];
}
// Bảng: chú thích TRÊN, đánh số theo chương "Bảng c.n." — đậm nhãn.
function tcaption(caption) {
  TABN += 1;
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 },
    children: [
      new TextRun({ text: `Bảng ${CHAP}.${TABN}. `, font: HF, size: 24, bold: true, color: BLACK }),
      new TextRun({ text: caption, font: HF, size: 24, color: BLACK }),
    ] });
}
// Ô bảng kiểu booktabs: không nền, chỉ kẻ ngang.
function cell(text, { w, bold, color, align, size, top, bottom } = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    borders: {
      top: top ? { style: BorderStyle.SINGLE, size: top, color: BLACK } : { style: BorderStyle.NONE },
      bottom: bottom ? { style: BorderStyle.SINGLE, size: bottom, color: BLACK } : { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    },
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align || AlignmentType.LEFT, spacing: { after: 0, line: 264 },
      children: [new TextRun({ text: String(text), font: HF, size: size || 22, bold, color: color || BLACK })] })] });
}
function table(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const n = rows.length;
  const head = new TableRow({ tableHeader: true, children: headers.map((h, i) =>
    cell(h, { w: widths[i], bold: true, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER, size: o.hsize || 22, top: 12, bottom: 8 })) });
  const body = rows.map((r, ri) => new TableRow({ children: r.map((c, i) =>
    cell(c, { w: widths[i], align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
      bold: (o.boldFirst && i === 0) || (o.hl && o.hl.includes(ri)),
      color: (o.hl && o.hl.includes(ri)) ? RED : BLACK, size: o.size || 21,
      bottom: ri === n - 1 ? 12 : 0 })) }));
  return new Table({ columnWidths: widths, width: { size: total, type: WidthType.DXA }, alignment: AlignmentType.CENTER, rows: [head, ...body] });
}
// Hộp "Nhận xét" — nền xám nhạt, chữ đen (giống \hopnhanmanh của mẫu).
function callout(title, body, color) {
  return new Table({ columnWidths: [9360], width: { size: 9360, type: WidthType.DXA }, alignment: AlignmentType.CENTER,
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "808080" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "808080" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "808080" }, right: { style: BorderStyle.SINGLE, size: 4, color: "808080" } },
    rows: [new TableRow({ children: [new TableCell({ width: { size: 9360, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: "F2F2F2" },
      margins: { top: 130, bottom: 130, left: 170, right: 170 },
      children: [
        new Paragraph({ spacing: { after: 70, line: LH }, children: [new TextRun({ text: title + ". ", font: HF, size: 24, bold: true, color: BLACK })] }),
        ...(Array.isArray(body) ? body : [body]).map((t) =>
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, line: LH }, children: [new TextRun({ text: "• " + t, font: HF, size: 24, color: BLACK })] })),
      ] })] })] });
}
function centerBold(t, size, color, after) { return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: after ?? 120 }, children: [new TextRun({ text: t, font: HF, size: size || 26, bold: true, color: color || INK })] }); }
function pb() { return new Paragraph({ children: [new PageBreak()] }); }

// ---- shortcuts ----
const ds = B.dataset, testM = B.overall.test, validM = B.overall.valid, lvs = B.light_vs_sign, rv = B.rarity,
  imb = B.imbalance_overall, tp = B.train_plan, sig = B.size_signals, leak = B.leakage, pc = B.per_class,
  signs = lvs.signs, lights = lvs.lights, worst = pc[pc.length - 1], best = pc[0];

const children = [];
const A = (...xs) => xs.forEach((x) => Array.isArray(x) ? children.push(...x) : children.push(x));

// ============================================================ BÌA (theo mẫu LaTeX của trường)
const TITLE_LINES = ["KIỂM ĐỊNH ĐỘ TIN CẬY CỦA BỘ PHÁT HIỆN", "BIỂN BÁO GIAO THÔNG YOLOv8 THEO HƯỚNG DATA-CENTRIC"];
const grid = { style: BorderStyle.SINGLE, size: 6, color: BLACK };
function gcell(children, w, opts = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    borders: { top: grid, bottom: grid, left: grid, right: grid },
    verticalAlign: "center", margins: { top: 60, bottom: 60, left: 80, right: 80 }, children });
}
function ctext(t, o = {}) { return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: o.after ?? 0, line: 240 }, children: [new TextRun({ text: t, font: HF, size: o.size || 22, bold: o.bold, italics: o.italics, color: BLACK })] }); }
function rule() { return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200, before: 20 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLACK, space: 1 } }, children: [new TextRun({ text: "                                        ", font: HF, size: 6 })] }); }
function infoRow(label, value) {
  return new TableRow({ children: [
    new TableCell({ width: { size: 3400, type: WidthType.DXA }, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, margins: { top: 40, bottom: 40 },
      children: [new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { line: 300 }, children: [new TextRun({ text: label + "  ", font: HF, size: 26, bold: true, color: BLACK })] })] }),
    new TableCell({ width: { size: 5400, type: WidthType.DXA }, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, margins: { top: 40, bottom: 40 },
      children: (Array.isArray(value) ? value : [value]).map((v) => new Paragraph({ spacing: { line: 300 }, children: [new TextRun({ text: v, font: HF, size: 26, color: BLACK })] })) }),
  ] });
}
const infoTable = new Table({ columnWidths: [3400, 5400], width: { size: 8800, type: WidthType.DXA }, alignment: AlignmentType.CENTER,
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [
    infoRow("Ngành:", "Khoa học máy tính"),
    infoRow("Khóa:", "36 (2025 – 2027)"),
    infoRow("Học phần:", "(điền tên học phần)"),
    infoRow("Giảng viên giảng dạy:", "(điền tên giảng viên)"),
    infoRow("Nhóm thực hiện:", B.members.map((m) => `${m.name} — ${m.id}`)),
  ] });
// ---- Trang bìa 1: có khung chấm điểm ----
A(
  ctext("BỘ GIÁO DỤC VÀ ĐÀO TẠO", { size: 24, bold: true, after: 40 }),
  ctext("TRƯỜNG ĐẠI HỌC SƯ PHẠM THÀNH PHỐ HỒ CHÍ MINH", { size: 26, bold: true, after: 0 }),
  rule(),
  new Table({ columnWidths: [1900, 1900, 2600, 2600], width: { size: 9000, type: WidthType.DXA }, alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ tableHeader: true, children: [
        gcell([ctext("ĐIỂM SỐ", { bold: true })], 1900),
        gcell([ctext("ĐIỂM CHỮ", { bold: true })], 1900),
        gcell([ctext("Cán bộ chấm thi 1", { bold: true }), ctext("(Kí, ghi rõ họ tên)", { italics: true, size: 18 })], 2600),
        gcell([ctext("Cán bộ chấm thi 2", { bold: true }), ctext("(Kí, ghi rõ họ tên)", { italics: true, size: 18 })], 2600),
      ] }),
      new TableRow({ children: [
        gcell([new Paragraph({ spacing: { before: 500, after: 500 }, children: [] })], 1900),
        gcell([new Paragraph({ children: [] })], 1900),
        gcell([new Paragraph({ children: [] })], 2600),
        gcell([new Paragraph({ children: [] })], 2600),
      ] }),
    ] }),
  new Paragraph({ spacing: { after: 320 }, children: [] }),
  ctext("BÁO CÁO ĐỀ TÀI HỌC PHẦN", { size: 30, bold: true, after: 120 }),
  ctext("(điền tên học phần)", { size: 26, bold: true, after: 400 }),
  ...TITLE_LINES.map((l, i) => ctext(l, { size: 33, bold: true, after: i === TITLE_LINES.length - 1 ? 120 : 40 })),
  ctext("Rò rỉ dữ liệu, kích thước vật thể và khoảng cách hiệu năng đèn tín hiệu – biển báo", { size: 24, italics: true, after: 500 }),
  infoTable,
  new Paragraph({ spacing: { after: 500 }, children: [] }),
  ctext("Thành phố Hồ Chí Minh, tháng 7 năm 2026", { size: 24, italics: true }),
  pb(),
);
// ---- Trang bìa lót: có logo ----
{
  const logoPath = path.join(ROOT, "reports", "latex", "logo.png");
  A(
    ctext("BỘ GIÁO DỤC VÀ ĐÀO TẠO", { size: 24, bold: true, after: 40 }),
    ctext("TRƯỜNG ĐẠI HỌC SƯ PHẠM THÀNH PHỐ HỒ CHÍ MINH", { size: 26, bold: true, after: 0 }),
    rule(),
  );
  if (fs.existsSync(logoPath)) {
    A(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 300 },
      children: [new ImageRun({ type: "png", data: fs.readFileSync(logoPath), transformation: { width: 150, height: 150 } })] }));
  } else {
    A(new Paragraph({ spacing: { after: 500 }, children: [] }));
  }
  A(
    ctext("BÁO CÁO ĐỀ TÀI HỌC PHẦN", { size: 30, bold: true, after: 120 }),
    ctext("(điền tên học phần)", { size: 26, bold: true, after: 400 }),
    ...TITLE_LINES.map((l, i) => ctext(l, { size: 33, bold: true, after: i === TITLE_LINES.length - 1 ? 120 : 40 })),
    ctext("Rò rỉ dữ liệu, kích thước vật thể và khoảng cách hiệu năng đèn tín hiệu – biển báo", { size: 24, italics: true, after: 500 }),
    infoTable2(),
    new Paragraph({ spacing: { after: 500 }, children: [] }),
    ctext("Thành phố Hồ Chí Minh, tháng 7 năm 2026", { size: 24, italics: true }),
    pb(),
  );
}
function infoTable2() {
  return new Table({ columnWidths: [3400, 5400], width: { size: 8800, type: WidthType.DXA }, alignment: AlignmentType.CENTER,
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
    rows: [
      infoRow("Ngành:", "Khoa học máy tính"),
      infoRow("Khóa:", "36 (2025 – 2027)"),
      infoRow("Học phần:", "(điền tên học phần)"),
      infoRow("Giảng viên giảng dạy:", "(điền tên giảng viên)"),
      infoRow("Nhóm thực hiện:", B.members.map((m) => `${m.name} — ${m.id}`)),
    ] });
}

// ============================================================ LỜI CAM ĐOAN
A(
  H1("LỜI CAM ĐOAN"),
  P("Chúng tôi xin cam đoan báo cáo đề tài “Kiểm định độ tin cậy của bộ phát hiện biển báo giao thông YOLOv8 theo hướng data-centric” là công trình phân tích của nhóm, thực hiện trong khuôn khổ học phần Phân tích dữ liệu."),
  P("Chúng tôi cam đoan các điểm sau. Thứ nhất, bộ dữ liệu pkdarabi/cardetection được sử dụng công khai theo giấy phép đã công bố; nhóm không tự tạo hay chỉnh sửa dữ liệu quan sát. Thứ hai, toàn bộ số liệu, bảng biểu và hình vẽ trong báo cáo được sinh tự động từ mã nguồn trong thư mục src/, không có giá trị nào nhập tay hay chỉnh sửa thủ công; mọi bước ngẫu nhiên đều cố định hạt giống. Thứ ba, các kết quả bất lợi hoặc trái với kỳ vọng ban đầu đều được báo cáo đầy đủ thay vì bị lược bỏ. Thứ tư, mọi tài liệu tham khảo đều được trích dẫn đầy đủ."),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 400 }, children: [run("Thành phố Hồ Chí Minh, tháng 7 năm 2026", { italics: true, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120 }, children: [run("Nhóm thực hiện", { bold: true })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 60 }, children: [run("Huỳnh Phát Lợi · Đoàn Huỳnh Thanh Tú · Võ Phú Vinh", { size: 24 })] }),
  pb(),
);

// ============================================================ LỜI CẢM ƠN
A(
  H1("LỜI CẢM ƠN"),
  P("Nhóm xin gửi lời cảm ơn chân thành tới giảng viên phụ trách học phần Phân tích dữ liệu đã hướng dẫn và đưa ra những định hướng quan trọng, giúp nhóm chuyển từ một câu hỏi mô hình sang một câu hỏi phân tích dữ liệu có chiều sâu: không chạy đua độ chính xác, mà kiểm định độ tin cậy của kết quả."),
  P("Nhóm cũng cảm ơn cộng đồng phát triển các thư viện mã nguồn mở pandas, scipy, scikit-learn, matplotlib, Ultralytics và Streamlit — nền tảng kỹ thuật của toàn bộ phân tích và ứng dụng demo trong đề tài này."),
  pb(),
);

// ============================================================ MỤC LỤC
A(
  H1("MỤC LỤC"),
  new TableOfContents("Mục lục", { hyperlink: true, headingStyleRange: "1-3" }),
  pb(),
);

// ============================================================ DANH MỤC TỪ VIẾT TẮT
A(
  H1("DANH MỤC TỪ VIẾT TẮT"),
  table(["Từ viết tắt", "Diễn giải"], [
    ["AP", "Average Precision — độ chính xác trung bình của một lớp"],
    ["mAP", "mean Average Precision — trung bình AP trên các lớp (mAP@0,5 và mAP@0,5:0,95)"],
    ["IoU", "Intersection over Union — tỉ lệ giao trên hợp giữa hai hộp"],
    ["EDA", "Exploratory Data Analysis — phân tích khám phá dữ liệu"],
    ["YOLO", "You Only Look Once — họ mô hình phát hiện đối tượng một giai đoạn"],
    ["dHash", "Difference Hash — perceptual hash dùng phát hiện ảnh gần trùng"],
    ["P / R", "Precision / Recall — độ chính xác / độ bao phủ"],
    ["JS", "Jensen–Shannon divergence — độ đo khác biệt giữa hai phân phối"],
    ["CSV", "Comma-Separated Values — định dạng dữ liệu bảng"],
  ], [2600, 6760], { size: 22 }),
  pb(),
);

// ============================================================ DANH MỤC HÌNH VÀ BẢNG
A(
  H1("DANH MỤC HÌNH VÀ BẢNG"),
  H2("Danh mục hình"),
  P("Hình 3.1. Ba thước đo mất cân bằng lớp toàn corpus.", { noIndent: true, after: 60 }),
  P("Hình 3.2. Tỉ lệ vật thể nhỏ theo độ phân giải đầu vào.", { noIndent: true, after: 60 }),
  P("Hình 4.1. Trùng lặp và rò rỉ dữ liệu.", { noIndent: true, after: 60 }),
  P("Hình 4.2. Hiệu năng tổng thể trên tập kiểm định và kiểm thử.", { noIndent: true, after: 60 }),
  P("Hình 4.3. Hiệu năng theo lớp — hai lớp đèn tách hẳn khỏi 13 lớp biển báo.", { noIndent: true, after: 60 }),
  P("Hình 4.4. Đèn tín hiệu kém hơn biển báo trên mọi chỉ số.", { noIndent: true, after: 60 }),
  P("Hình 4.5. Độ hiếm không giải thích được hiệu năng.", { noIndent: true, after: 60 }),
  P("Hình 4.6. Không gian Precision–Recall theo lớp.", { noIndent: true, after: 120 }),
  H2("Danh mục bảng"),
  P("Bảng 1.1. Hai hướng tiếp cận với một bộ phát hiện đã huấn luyện.", { noIndent: true, after: 60 }),
  P("Bảng 3.1. Thống kê theo split.", { noIndent: true, after: 60 }),
  P("Bảng 3.2. Thành phần 15 lớp và số đối tượng trên tập kiểm thử.", { noIndent: true, after: 60 }),
  P("Bảng 3.3. Các thước đo mất cân bằng theo từng split.", { noIndent: true, after: 60 }),
  P("Bảng 3.4. Độ dịch phân phối lớp giữa các split.", { noIndent: true, after: 60 }),
  P("Bảng 3.5. Tỉ lệ vật thể nhỏ theo chính sách độ phân giải.", { noIndent: true, after: 60 }),
  P("Bảng 4.1. Trùng lặp và rò rỉ dữ liệu.", { noIndent: true, after: 60 }),
  P("Bảng 4.2. Hiệu năng theo lớp trên tập kiểm thử (15 lớp).", { noIndent: true, after: 120 }),
  pb(),
);

// ============================================================ TÓM TẮT
A(
  H1("TÓM TẮT"),
  P([run("Báo cáo này kiểm tra độ tin cậy ("), run("reliability audit", { bold: true, italics: true }), run(`) của một mô hình phát hiện biển báo giao thông YOLOv8 huấn luyện trên bộ dữ liệu pkdarabi/cardetection (15 lớp, ${vn(ds.total_images_scanned)} ảnh, ${vn(ds.total_valid_objects)} đối tượng). Theo hướng data-centric, nhóm không huấn luyện lại mà xem mô hình như một công cụ đo và đặt câu hỏi: liệu con số mAP@0,5 = 0,970 mà mô hình đạt được có đáng tin và có đồng đều giữa các lớp hay không.`)]),
  P([run("Nhóm thu được ba kết quả. "), run("Về data leakage, ", { bold: true }), run(`bước EDA đã phát hiện ${vn(leak.exact_cross_split_duplicate_rows)} ảnh trùng và ${vn(leak.dhash_cross_split_collision_rows)} ảnh gần trùng nằm xuyên các split, nhưng quy trình chỉ loại ${vn(leak.removed_from_train)} ảnh khỏi tập train; tập test dùng để chấm điểm vẫn còn trùng lặp, nên điểm mAP nhiều khả năng cao hơn thực tế. `), run("Về khoảng cách giữa hai nhóm lớp, ", { bold: true }), run(`hai lớp đèn tín hiệu chỉ đạt mAP@0,5:0,95 = ${vnp(lights.mean_map50_95, 3)}, trong khi 13 lớp biển báo đạt trung bình ${vnp(signs.mean_map50_95, 3)} (chênh ${vnp(lvs.gap, 3)}; kiểm định Mann–Whitney một phía cho p = ${vnp(lvs.mannwhitney_p, 4)}). `), run("Về nguyên nhân, ", { bold: true }), run(`độ hiếm của lớp không liên quan tới khoảng cách này (Spearman ρ = ${vnp(rv.spearman_instances_vs_map5095, 3)}, p = ${vnp(rv.p_instances_vs_map5095, 2)}): hai lớp đèn nằm trong nhóm nhiều đối tượng nhất nhưng lại có hiệu năng thấp nhất.`)]),
  P("Như vậy, con số 0,970 vừa có thể được data leakage nâng đỡ, vừa che lấp sự chênh lệch lớn giữa các lớp; và sự chênh lệch đó đến từ đặc điểm của đèn tín hiệu (nhỏ, tương phản thấp) chứ không phải từ mất cân bằng dữ liệu. Kèm theo báo cáo là pipeline phân tích chạy lại được, một ứng dụng web có tab demo phát hiện trực tiếp, cùng slide trình bày."),
  P([run("Từ khoá: ", { bold: true }), run("reliability audit, data-centric, YOLOv8, object detection, data leakage, class imbalance, per-class mAP, traffic light.", { italics: true })]),
  pb(),
);

// ============================================================ CHƯƠNG 1
A(
  H1("GIỚI THIỆU", 1),
  H2("1.1. Bối cảnh và động cơ"),
  P("Phát hiện biển báo và đèn tín hiệu giao thông là một trong những bài toán nền tảng của hệ thống hỗ trợ người lái nâng cao và xe tự hành. Nhiệm vụ vừa phải định vị chính xác vị trí của từng đối tượng trong khung hình, vừa phải phân loại đúng loại biển báo hoặc đèn. Với sự phát triển của các mô hình phát hiện đối tượng một giai đoạn như họ YOLO, độ chính xác trên các bộ dữ liệu chuẩn đã đạt mức rất cao."),
  P("Bộ phát hiện được khảo sát trong đề tài này đạt mAP@0,5 = 0,970 và mAP@0,5:0,95 = 0,812 trên tập kiểm thử — những con số thoạt nhìn rất ấn tượng. Tuy nhiên, chính vì độ chính xác tổng thể đã bão hoà, việc thêm một mô hình mới hay tinh chỉnh thêm vài phần trăm gần như không còn giá trị khoa học. Câu hỏi có ý nghĩa hơn, và cũng gần với tinh thần của môn Phân tích dữ liệu hơn, là: con số tổng thể đó có thực sự đáng tin cậy hay không, và nó có phản ánh đồng đều năng lực của mô hình trên mọi lớp hay không."),
  P("Tầm quan trọng của bài toán còn nằm ở khía cạnh an toàn. Trong các hệ thống hỗ trợ người lái nâng cao và xe tự hành, việc bỏ sót một đèn tín hiệu hay một biển báo có thể dẫn tới hậu quả nghiêm trọng. Vì vậy, một bộ phát hiện không chỉ cần chính xác trung bình cao, mà còn cần đáng tin cậy đồng đều trên mọi loại đối tượng, đặc biệt là những đối tượng quan trọng cho an toàn như đèn tín hiệu. Một con số mAP tổng cao nhưng che giấu điểm yếu ở nhóm đèn là một rủi ro tiềm ẩn mà người triển khai cần biết."),
  P("Đây là khoảng trống mà một góc nhìn phân tích dữ liệu có thể lấp vào. Thay vì tập trung vào mô hình, đề tài tập trung vào chất lượng và cấu trúc của dữ liệu, cũng như độ tin cậy của kết quả đánh giá. Cách tiếp cận này không chỉ trả lời câu hỏi “mô hình đúng bao nhiêu phần trăm” mà còn trả lời “ta có thể tin con số đó đến đâu, và ở đâu nó có thể sai”."),
  H2("1.2. Định vị đề tài — hai hướng tiếp cận"),
  P("Khi làm việc với một bài toán học máy đã có kết quả tốt, có hai hướng tiếp cận khác nhau về bản chất."),
  P([run("Hướng model-centric ", { bold: true }), run("xem mô hình là đối tượng nghiên cứu: nhà nghiên cứu thiết kế kiến trúc, huấn luyện, so sánh nhiều mô hình và tìm cách đẩy các chỉ số như mAP lên cao hơn. Đây là hướng phổ biến trong các cuộc thi và benchmark.")]),
  P([run("Hướng data-centric ", { bold: true }), run("lại xem chất lượng và cấu trúc của dữ liệu là đối tượng nghiên cứu. Andrew Ng lập luận rằng khi kiến trúc mô hình đã bão hoà, cải thiện tiếp theo chủ yếu đến từ việc hiểu và làm sạch dữ liệu. Trong hướng này, mô hình chỉ đóng vai trò một công cụ đo; điều quan trọng là con số đo được có đáng tin không, và những đặc tính nào của dữ liệu chi phối nó.")]),
  P("Đề tài này đi theo hướng data-centric. Bảng 1 tóm tắt sự khác biệt giữa hai hướng và làm rõ vị trí của đề tài."),
  tcaption("Hai hướng tiếp cận với một bộ phát hiện đã huấn luyện"),
  table(["Tiêu chí", "Hướng model-centric", "Hướng data-centric (đề tài này)"], [
    ["Câu hỏi", "Mô hình nào phát hiện tốt hơn?", "mAP 0,97 có đáng tin và đồng đều không?"],
    ["Vai trò mô hình", "Đối tượng nghiên cứu — huấn luyện, so sánh", "Công cụ đo cố định — chỉ phân tích lại kết quả"],
    ["Đầu ra", "Bảng benchmark, mAP tổng", "Bằng chứng thống kê về độ tin cậy, phân hoá lớp"],
    ["Phương pháp", "Thiết kế, huấn luyện mạng", "Kiểm định thống kê, phân tích rò rỉ và phân hoá"],
  ], [2400, 3480, 3480], { size: 21 }),
  P("Mô hình YOLOv8 đã huấn luyện được xem là công cụ đo cố định. Đề tài chạy ở chế độ results-only: không huấn luyện lại mà chỉ phân tích lại các kết quả mà mô hình sinh ra. Trọng số đã huấn luyện được đính kèm trong thư mục models/ của đề tài để phục vụ ứng dụng demo và việc tái lập.", { before: 60 }),
  H2("1.3. Mục tiêu và câu hỏi nghiên cứu"),
  P("Mục tiêu tổng quát của đề tài là đánh giá độ tin cậy của kết quả phát hiện biển báo, thay vì tìm cách cải thiện mô hình. Từ mục tiêu đó, đề tài đặt ra ba câu hỏi nghiên cứu cụ thể, tương ứng với ba phần phân tích chính:"),
  bullet([run("Câu hỏi A (Rò rỉ dữ liệu): ", { bold: true }), run("Mức độ trùng lặp và rò rỉ dữ liệu xuyên tập là bao nhiêu, và nó ảnh hưởng thế nào tới độ tin cậy của con số mAP?")]),
  bullet([run("Câu hỏi B (Phân hoá lớp): ", { bold: true }), run("Hiệu năng có đồng đều giữa các lớp không? Nếu không, khoảng cách lớn nhất nằm ở đâu và có ý nghĩa thống kê không?")]),
  bullet([run("Câu hỏi C (Nguyên nhân): ", { bold: true }), run("Khoảng cách hiệu năng đó bắt nguồn từ độ hiếm của lớp (mất cân bằng dữ liệu) hay từ bản chất của loại vật thể?")]),
  H2("1.4. Phạm vi và giới hạn"),
  P("Đề tài giới hạn ở việc phân tích thứ cấp trên kết quả của một bộ phát hiện đã huấn luyện, trên đúng bộ dữ liệu pkdarabi/cardetection. Đề tài không huấn luyện lại mô hình, không so sánh nhiều kiến trúc, và không mở rộng sang các bộ dữ liệu khác. Kết luận của đề tài gắn với bộ dữ liệu và mô hình cụ thể này; việc khái quát hoá cho mọi điều kiện đường thực tế nằm ngoài phạm vi."),
  H2("1.5. Đóng góp"),
  P("Đề tài có bốn đóng góp chính:"),
  numbered("Định lượng phơi nhiễm rò rỉ dữ liệu và chỉ ra khoảng trống trong quy trình khử trùng — tập kiểm thử dùng để chấm điểm chưa được khử trùng.", "c1"),
  numbered("Chứng minh bằng kiểm định thống kê rằng nhóm đèn tín hiệu kém hơn nhóm biển báo một cách có ý nghĩa.", "c1"),
  numbered("Bác bỏ giả thuyết “mất cân bằng gây ra thất bại” bằng bằng chứng tương quan, và chuyển nguyên nhân sang bản chất loại vật thể.", "c1"),
  numbered("Cung cấp một pipeline phân tích tái lập, một ứng dụng web có tab demo phát hiện trực tiếp, cùng báo cáo Word/LaTeX và slide.", "c1"),
  H2("1.6. Cấu trúc báo cáo"),
  P("Phần còn lại của báo cáo được tổ chức như sau. Chương 2 trình bày cơ sở lý thuyết về phát hiện đối tượng, các chỉ số đánh giá, rò rỉ dữ liệu, các thước đo mất cân bằng và các kiểm định thống kê được sử dụng. Chương 3 mô tả bộ dữ liệu, công cụ đo và phương pháp phân tích ba bước. Chương 4 trình bày chi tiết kết quả kiểm định theo ba câu hỏi A, B, C. Chương 5 giới thiệu ứng dụng web và tab demo phát hiện. Chương 6 tổng kết, đưa ra khuyến nghị, hạn chế và hướng phát triển."),
  pb(),
);

// ============================================================ CHƯƠNG 2
A(
  H1("CƠ SỞ LÝ THUYẾT", 2),
  H2("2.1. Bài toán phát hiện đối tượng"),
  P("Phát hiện đối tượng (object detection) là bài toán vừa định vị vừa phân loại nhiều đối tượng trong một ảnh. Mỗi dự đoán gồm một bounding box (bounding box) xác định vị trí và một nhãn lớp kèm điểm tin cậy. Khác với phân loại ảnh chỉ gán một nhãn cho toàn ảnh, phát hiện phải xử lý số lượng đối tượng thay đổi và vị trí bất kỳ."),
  P("Độ khớp giữa hộp dự đoán và hộp thực tế được đo bằng IoU (Intersection over Union) — tỉ lệ giữa diện tích phần giao và diện tích phần hợp của hai hộp. Một dự đoán được xem là đúng khi IoU với hộp thực vượt một ngưỡng cho trước (thường là 0,5)."),
  H2("2.2. Chỉ số đánh giá: Precision, Recall và mAP"),
  P("Precision là tỉ lệ dự đoán đúng trên tổng số dự đoán; Recall là tỉ lệ đối tượng thực được phát hiện. Với mỗi lớp, đường cong Precision–Recall được tổng hợp thành Average Precision (AP). Trung bình AP trên tất cả các lớp cho ra mean Average Precision (mAP)."),
  P("Hai biến thể thường dùng là mAP@0,5, tính ở ngưỡng IoU duy nhất 0,5, và mAP@0,5:0,95, trung bình trên mười ngưỡng IoU từ 0,5 đến 0,95. mAP@0,5:0,95 nghiêm ngặt hơn nhiều vì đòi hỏi hộp dự đoán khớp chặt với hộp thực; vì vậy trong đề tài này, mAP@0,5:0,95 được dùng làm chỉ số chính khi so sánh giữa các lớp."),
  H3("2.2.1. Ví dụ minh hoạ về IoU và hệ quả của nó"),
  P("Để thấy rõ vì sao mAP@0,5:0,95 khó hơn, hãy xét một ví dụ đơn giản. Giả sử hộp thực của một biển báo có kích thước 40×40 điểm ảnh. Nếu mô hình dự đoán một hộp lệch đi 6 điểm ảnh theo mỗi chiều, phần giao vẫn khá lớn nên IoU có thể vượt 0,5 — dự đoán được tính là đúng ở mAP@0,5. Nhưng cùng độ lệch đó, IoU có thể không đạt 0,75 hay 0,9, nên dự đoán bị tính là sai ở các ngưỡng cao, kéo mAP@0,5:0,95 xuống. Với vật thể càng nhỏ, cùng một độ lệch tuyệt đối lại gây sụt IoU càng mạnh. Đây là lý do các lớp có nhiều vật thể nhỏ, như đèn tín hiệu, chịu thiệt nặng nhất ở chỉ số nghiêm ngặt."),
  H3("2.2.2. Ba loại lỗi trong phát hiện"),
  P("Lỗi của một bộ phát hiện có thể chia thành ba loại. Lỗi định vị là khi mô hình nhận đúng lớp nhưng hộp không khớp đủ chặt. Lỗi phân loại là khi hộp khớp tốt nhưng gán sai nhãn — ví dụ nhầm đèn đỏ thành đèn xanh. Lỗi bỏ sót là khi mô hình hoàn toàn không phát hiện một vật thể thực, thể hiện ở Recall thấp. Ba loại lỗi này ảnh hưởng khác nhau tới Precision, Recall và mAP, và việc phân biệt chúng giúp giải thích vì sao một lớp cụ thể lại kém. Trong đề tài, lớp đèn đỏ chịu chủ yếu lỗi bỏ sót (Recall thấp) và một phần lỗi phân loại (nhầm đỏ/xanh)."),
  H2("2.3. Kiến trúc YOLOv8 và vai trò công cụ đo"),
  P("YOLOv8 là một mô hình phát hiện một giai đoạn: nó dự đoán trực tiếp hộp và lớp trên lưới đặc trưng đa tỉ lệ trong một lần suy luận, nên nhanh và phù hợp thời gian thực. Về mặt kiến trúc, một mô hình YOLO điển hình gồm ba phần. Backbone là mạng trích đặc trưng, biến ảnh đầu vào thành các bản đồ đặc trưng ở nhiều tỉ lệ. Neck kết hợp đặc trưng đa tỉ lệ để mô hình vừa thấy chi tiết nhỏ vừa thấy ngữ cảnh lớn. Head là phần dự đoán, sinh ra hộp và lớp tại mỗi vị trí lưới."),
  P("YOLOv8 thuộc thế hệ anchor-free, tức dự đoán trực tiếp tâm và kích thước hộp thay vì dựa trên các hộp mỏ neo cố định. Sau khi sinh nhiều hộp dự đoán, một bước hậu xử lý gọi là non-maximum suppression loại bỏ các hộp chồng lấn dư thừa, chỉ giữ hộp có điểm tin cậy cao nhất cho mỗi vật thể. Hai siêu tham số quan trọng khi suy luận là ngưỡng tin cậy — quyết định giữ lại dự đoán nào — và ngưỡng IoU cho bước loại chồng lấn.", { before: 40 }),
  P("Trong đề tài này, một YOLOv8s đã huấn luyện được coi là công cụ đo cố định. Nhóm không quan tâm cải tiến kiến trúc; thay vào đó, chính các lỗi và điểm số mà mô hình sinh ra trở thành dữ liệu để phân tích độ tin cậy. Cách nhìn này cho phép tách bạch hai câu hỏi thường bị trộn lẫn: “mô hình mạnh đến đâu” và “kết quả đánh giá đáng tin đến đâu”. Đề tài chỉ tập trung vào câu hỏi thứ hai.", { before: 40 }),
  H2("2.4. Rò rỉ dữ liệu và trùng lặp"),
  P("Rò rỉ dữ liệu (data leakage) xảy ra khi thông tin của tập kiểm thử lọt vào quá trình huấn luyện, khiến kết quả đánh giá lạc quan hơn năng lực thực. Với dữ liệu ảnh cắt từ video hoặc thu thập từ nhiều nguồn, một dạng data leakage phổ biến là cùng một ảnh, hoặc các ảnh gần như trùng nhau, xuất hiện ở nhiều tập."),
  P("Perceptual hash (perceptual hash) là kỹ thuật phát hiện ảnh gần trùng. Khác với băm mật mã vốn thay đổi hoàn toàn khi ảnh đổi một điểm ảnh, perceptual hash như dHash tạo ra mã dựa trên cấu trúc thị giác của ảnh, nên hai ảnh gần giống nhau sẽ có mã gần nhau về khoảng cách Hamming. Nếu tập kiểm thử không được khử trùng, mô hình có thể đã “thấy” những ảnh rất giống trong lúc huấn luyện, làm mAP cao hơn hiệu năng thực."),
  P("Cụ thể, thuật toán dHash thu nhỏ ảnh về một lưới nhỏ, chuyển sang mức xám, rồi so sánh độ sáng của các điểm ảnh liền kề để tạo ra một chuỗi bit: mỗi bit ghi lại việc điểm bên trái sáng hơn hay tối hơn điểm bên phải. Vì mã hoá xu hướng thay đổi độ sáng chứ không phải giá trị tuyệt đối, dHash bền vững với thay đổi nhỏ về độ sáng, nén ảnh hay đổi kích thước. Hai ảnh chỉ khác nhau vài chi tiết nhỏ sẽ cho mã băm chênh nhau ít bit, và được xem là gần trùng. Đây là công cụ phù hợp để phát hiện các ảnh “bản sao” nằm rải rác giữa các tập.", { before: 40 }),
  H2("2.5. Các thước đo mất cân bằng lớp"),
  P("Ba thước đo bổ sung cho nhau được dùng để mô tả mất cân bằng lớp. Imbalance ratio là tỉ số giữa số lượng của lớp nhiều nhất và lớp ít nhất. Entropy chuẩn hoá bằng entropy Shannon chia cho logarit số lớp, đạt giá trị 1 khi phân bố hoàn toàn đều. Hệ số Gini về bất bình đẳng dựa trên đường Lorenz, bằng 0 khi hoàn toàn cân bằng và tiến tới 1 khi tập trung. Cần phân biệt hệ số Gini bất bình đẳng này với Gini impurity trong cây quyết định — đây là hai khái niệm khác nhau."),
  P("Trên bộ dữ liệu của đề tài, ba thước đo cho một bức tranh nhất quán. Imbalance ratio bằng 35,8 nghĩa là lớp phổ biến nhất có số đối tượng gấp gần 36 lần lớp hiếm nhất — một mức chênh đáng kể. Tuy nhiên entropy chuẩn hoá 0,953, khá gần 1, cho thấy phần lớn các lớp có số lượng không quá lệch nhau; chỉ một vài lớp ở hai đầu tạo ra tỉ số lớn. Hệ số Gini 0,247 cũng xác nhận mức tập trung chỉ ở mức trung bình. Cách đọc kết hợp này quan trọng: nếu chỉ nhìn imbalance ratio, ta dễ kết luận dữ liệu mất cân bằng nghiêm trọng, trong khi entropy và Gini cho thấy mức độ thực tế ôn hoà hơn.", { before: 40 }),
  H2("2.6. Kiểm định Mann–Whitney U"),
  P("Kiểm định Mann–Whitney U là một kiểm định phi tham số dùng để so sánh phân phối của hai nhóm độc lập mà không giả định phân phối chuẩn. Nó phù hợp khi cỡ mẫu nhỏ hoặc phân phối lệch. Trong đề tài, kiểm định này được dùng để so sánh phân phối hiệu năng (mAP@0,5:0,95) của nhóm đèn tín hiệu (2 lớp) với nhóm biển báo (13 lớp), với giả thuyết một phía rằng biển báo tốt hơn đèn."),
  H2("2.7. Tương quan hạng Spearman"),
  P("Tương quan hạng Spearman đo mức độ liên hệ đơn điệu giữa hai biến bằng cách tính tương quan Pearson trên hạng của chúng. Hệ số ρ nằm trong khoảng từ -1 đến 1; giá trị gần 0 với p lớn cho thấy không có liên hệ đơn điệu. Đề tài dùng Spearman để kiểm tra xem số lượng đối tượng của một lớp (đại diện cho độ hiếm) có liên hệ với hiệu năng của lớp đó hay không."),
  H2("2.8. Định nghĩa kích thước vật thể"),
  P("Theo quy ước COCO, một đối tượng được xếp là “nhỏ” (small) nếu diện tích của nó nhỏ hơn 32×32 điểm ảnh. Vật thể nhỏ khó phát hiện vì mang ít điểm ảnh thông tin, đặc biệt ở độ phân giải đầu vào thấp. Tỉ lệ vật thể nhỏ cao là một tín hiệu khó khăn cho bộ phát hiện, và là một trong những yếu tố cần xem xét khi giải thích sự chênh lệch hiệu năng giữa các lớp."),
  P("Một hệ quả quan trọng là các chỉ số nghiêm ngặt về độ khớp hộp, như mAP@0,5:0,95, đặc biệt nhạy với vật thể nhỏ. Khi vật thể chỉ chiếm vài chục điểm ảnh, một sai lệch nhỏ về vị trí hộp cũng làm IoU giảm mạnh, kéo theo AP giảm ở các ngưỡng IoU cao. Đây là lý do vì sao trong đề tài, khoảng cách giữa hai nhóm lớp thể hiện rõ nhất ở mAP@0,5:0,95 chứ không phải ở mAP@0,5.", { before: 40 }),
  H2("2.9. Các nghiên cứu và khái niệm liên quan"),
  P("Hướng data-centric AI được Andrew Ng và cộng sự cổ vũ mạnh mẽ, với luận điểm rằng chất lượng và tính nhất quán của dữ liệu thường quyết định hiệu năng nhiều hơn kiến trúc mô hình khi mô hình đã đủ tốt. Trong tinh thần đó, việc kiểm định độ tin cậy của một kết quả — thay vì chỉ công bố con số — trở thành một phần thiết yếu của quy trình phân tích."),
  P("Vấn đề rò rỉ dữ liệu và trùng lặp giữa các tập đã được nhiều nghiên cứu chỉ ra là nguồn gây lạc quan hoá phổ biến trong đánh giá học máy. Với dữ liệu ảnh thu thập từ nhiều nguồn hoặc cắt từ video, hiện tượng cùng một cảnh xuất hiện ở nhiều tập là rất dễ xảy ra nếu quy trình chia tập không kiểm soát trùng lặp ở mức tri giác. Các nghiên cứu về chất lượng nhãn cũng cho thấy ngay cả những benchmark chuẩn cũng chứa lỗi đủ để làm thay đổi thứ hạng mô hình — một lý do nữa để kiểm toán dữ liệu trước khi tin vào con số.", { before: 40 }),
  P("Bên cạnh đó, việc chỉ báo cáo một chỉ số tổng hợp như mAP tổng từ lâu đã được cảnh báo là có thể che giấu sự phân hoá giữa các lớp hoặc các nhóm con. Thực hành tốt là báo cáo hiệu năng phân rã theo lớp và theo nhóm có ý nghĩa, kèm các kiểm định thống kê để phân biệt khác biệt thật với dao động ngẫu nhiên. Đề tài này áp dụng đúng tinh thần đó cho bài toán phát hiện biển báo.", { before: 40 }),
  H2("2.10. Ý nghĩa thống kê và cỡ mẫu"),
  P("Khi so sánh hiệu năng giữa các nhóm, cần phân biệt khác biệt quan sát được với khác biệt có ý nghĩa thống kê. Giá trị p của một kiểm định cho biết xác suất quan sát được khác biệt lớn như vậy nếu thực ra hai nhóm không khác nhau. Một p nhỏ, thường dưới 0,05 hoặc 0,01, cho phép bác bỏ giả thuyết không và kết luận khác biệt là có ý nghĩa."),
  P("Tuy nhiên, ý nghĩa thống kê không đồng nghĩa với độ lớn thực tiễn, và ngược lại. Vì vậy đề tài luôn báo cáo đồng thời độ lớn của khoảng cách (ví dụ chênh lệch mAP) bên cạnh giá trị p. Ngoài ra, khi cỡ mẫu nhỏ — như nhóm đèn chỉ có hai lớp — cần thận trọng: một kết quả có ý nghĩa vẫn cần được củng cố bằng cơ chế giải thích hợp lý, thay vì chỉ dựa vào con số p. Trong đề tài, kết luận về khoảng cách đèn–biển được củng cố bằng cả kiểm định thống kê lẫn cơ chế bản chất vật thể, nên vững chắc hơn.", { before: 40 }),
  H2("2.11. Quy trình đánh giá đúng đắn"),
  P("Một quy trình đánh giá học máy đúng đắn tách dữ liệu thành ba phần: tập huấn luyện để học tham số, tập kiểm định để chọn siêu tham số và theo dõi quá trình học, và tập kiểm thử chỉ dùng một lần ở cuối để ước lượng khả năng khái quát hoá. Nguyên tắc bất di bất dịch là tập kiểm thử phải hoàn toàn độc lập với tập huấn luyện; mọi rò rỉ thông tin giữa chúng đều làm ước lượng cuối cùng lạc quan hơn thực tế."),
  P("Với dữ liệu ảnh, độc lập không chỉ nghĩa là các tệp khác nhau, mà còn phải là các nội dung thị giác khác nhau. Hai ảnh có tên khác nhau nhưng gần như trùng về nội dung vẫn vi phạm tính độc lập. Đây là lý do việc khử trùng ở mức tri giác — chứ không chỉ ở mức tên tệp hay băm chính xác — là bước bắt buộc trước khi tin vào kết quả. Phát hiện A của đề tài chỉ ra rằng bước này chưa được thực hiện đầy đủ trên dữ liệu đánh giá, và đó là cơ sở để đặt dấu hỏi về con số mAP tổng.", { before: 40 }),
  pb(),
);

// ============================================================ CHƯƠNG 3
const spRows = B.split_summary.map((r) => [r.split, vn(r.n_images), vn(r.n_objects), vnp(r.objects_per_image_mean, 2), vn(r.bbox_area_px_median, 0)]);
const shiftRows = B.split_shift.map((r) => [`${r.split_a} – ${r.split_b}`, vnp(r.js_divergence_bits, 4)]);
const resRows = B.resolution.map((r) => [r.input_policy, vn(r.min_side_median_px, 0), pctv(r.ratio_min_side_lt_8px, 1), pctv(r.ratio_min_side_lt_16px, 1)]);
A(
  H1("DỮ LIỆU VÀ PHƯƠNG PHÁP", 3),
  H2("3.1. Bộ dữ liệu pkdarabi/cardetection"),
  P(`Bộ dữ liệu gồm 15 lớp, chia thành hai nhóm bản chất khác nhau: hai lớp đèn tín hiệu (Green Light, Red Light) và mười ba lớp biển báo (mười hai biển giới hạn tốc độ từ 10 đến 120, và biển Stop). Tổng cộng có ${vn(ds.total_images_scanned)} ảnh và ${vn(ds.total_valid_objects)} đối tượng hợp lệ, được chia sẵn thành ba tập huấn luyện, kiểm định và kiểm thử như Bảng ${TABN + 1}.`),
  tcaption("Thống kê theo split"),
  table(["Split", "Số ảnh", "Số đối tượng", "Đối tượng/ảnh (TB)", "Diện tích bbox (px, median)"], spRows, [1700, 1700, 2100, 2100, 1760], { boldFirst: true }),
  P(`Chất lượng nhãn của bộ dữ liệu rất tốt: không có file nhãn thiếu, không có đối tượng không hợp lệ, và không có dòng nhãn sai định dạng. Lớp phổ biến nhất là ${ds.dominant_class} với ${vn(ds.dominant_class_instances)} đối tượng, trong khi lớp hiếm nhất là ${ds.rarest_nonzero_class} chỉ có ${vn(ds.rarest_class_instances)} đối tượng. Chi tiết này rất quan trọng: vì nhãn sạch, nếu có vấn đề về độ tin cậy thì nó nằm ở phân bố và cách chia tập, chứ không phải ở khâu gán nhãn.`, { before: 60 }),
  H3("3.1.1. Hai nhóm lớp và đặc điểm"),
  P("Việc phân biệt hai nhóm lớp là mấu chốt của toàn bộ phân tích. Nhóm đèn tín hiệu thường xuất hiện nhỏ và ở xa trong khung hình, có độ tương phản thấp trên nền trời hoặc đường phức tạp; ngoài ra đèn đỏ và đèn xanh có hình dạng gần giống nhau nên nhiều khả năng dễ bị nhầm. Ngược lại, nhóm biển báo có chữ số và ký hiệu rõ ràng, hình dạng chuẩn hoá, và thường chiếm diện tích lớn hơn. Sự khác biệt về bản chất vật thể này, chứ không phải số lượng mẫu, mới là điều giải thích được khoảng cách hiệu năng ở các chương sau."),
  P("Bảng dưới đây liệt kê thành phần đầy đủ 15 lớp cùng nhóm và số đối tượng trên tập kiểm thử, để người đọc có bức tranh tổng thể về dữ liệu trước khi đi vào kết quả.", { before: 40 }),
  tcaption("Thành phần 15 lớp và số đối tượng trên tập kiểm thử"),
  table(["ID", "Lớp", "Nhóm", "Số đối tượng (test)"],
    pc.slice().sort((a, b) => a.class_id - b.class_id).map((r) => [String(r.class_id), r.class_name, r.category, vn(r.test_instances)]),
    [1100, 3800, 2400, 2060], { size: 20, hl: pc.slice().sort((a, b) => a.class_id - b.class_id).map((r, i) => r.is_traffic_light ? i : -1).filter((i) => i >= 0) }),
  H3("3.1.2. Độ dịch phân phối giữa các split"),
  P("Để loại trừ khả năng chênh lệch hiệu năng đến từ việc các tập có phân phối lớp khác nhau, nhóm đo độ dịch phân phối bằng Jensen–Shannon divergence. Bảng dưới đây cho thấy độ dịch giữa mọi cặp split đều rất nhỏ."),
  tcaption("Độ dịch phân phối lớp giữa các split (Jensen–Shannon, bit)"),
  table(["Cặp split", "JS divergence (bit)"], shiftRows, [4680, 4680], { boldFirst: true }),
  P("Cả ba giá trị đều dưới 0,008 bit, cho thấy phân phối lớp giữa các tập gần như đồng nhất. Đây là một điểm quan trọng: nó loại bỏ một cách giải thích cạnh tranh và củng cố cho kết luận ở Chương 4 rằng khoảng cách hiệu năng bắt nguồn từ bản chất lớp.", { before: 60 }),
  H3("3.1.3. Mất cân bằng lớp"),
  P(`Toàn corpus có tỉ số mất cân bằng ${vnp(imb.imbalance_ratio, 1)}, entropy chuẩn hoá ${vnp(imb.normalized_entropy, 3)} và hệ số Gini ${vnp(imb.gini, 3)}. Đây là mức mất cân bằng trung bình — rõ rệt nhưng không cực đoan. Con số này sẽ được đối chiếu ở Chương 4 khi kiểm tra xem mất cân bằng có phải nguyên nhân của thất bại đèn hay không.`),
  P("Bảng dưới đây trình bày các thước đo mất cân bằng tính riêng cho từng split. Điều đáng chú ý là entropy và Gini gần như không đổi giữa các tập, cho thấy cấu trúc mất cân bằng đồng nhất trên toàn bộ dữ liệu — một lần nữa loại trừ khả năng chênh lệch hiệu năng đến từ sự khác biệt phân bố giữa các tập.", { before: 40 }),
  tcaption("Các thước đo mất cân bằng theo từng split"),
  table(["Cấp", "Số instance", "Imbalance ratio", "Entropy chuẩn hoá", "Hệ số Gini"],
    B.imbalance_by_split.map((r) => [r.split, vn(r.total_instances), vnp(r.imbalance_ratio, 2), vnp(r.normalized_entropy, 4), vnp(r.gini, 4)]),
    [1800, 2000, 2160, 2200, 1200], { size: 21, boldFirst: true }),
  ...figure("05_imbalance.png", 560, 316, "Ba thước đo mất cân bằng lớp toàn corpus."),
  H3("3.1.4. Kích thước vật thể"),
  P(`Kích thước vật thể là một yếu tố khó khăn của bộ dữ liệu. Có tới ${pctv(sig.small_ratio_area_lt_1pct, 1)} đối tượng thuộc loại nhỏ (diện tích dưới 1% ảnh) và ${pctv(sig.tiny_ratio_area_lt_0_1pct, 1)} thuộc loại tí hon (dưới 0,1%). Bảng và hình dưới đây cho thấy độ phân giải đầu vào ảnh hưởng lớn tới tỉ lệ vật thể có cạnh ngắn dưới ngưỡng phát hiện được.`),
  tcaption("Tỉ lệ vật thể nhỏ theo chính sách độ phân giải đầu vào"),
  table(["Chính sách độ phân giải", "Cạnh ngắn median (px)", "Tỉ lệ < 8 px", "Tỉ lệ < 16 px"], resRows, [4000, 2400, 1480, 1480], { boldFirst: true }),
  ...figure("06_small_objects.png", 610, 332, "Tỉ lệ vật thể nhỏ giảm mạnh khi tăng độ phân giải đầu vào."),
  H3("3.1.6. Hạn chế về tính đại diện của dữ liệu"),
  P("Như mọi bộ dữ liệu, pkdarabi/cardetection có những giới hạn về tính đại diện cần nêu rõ. Dữ liệu gồm ảnh thu thập từ nhiều nguồn với điều kiện chụp khác nhau, nhưng không nhất thiết bao phủ đầy đủ mọi điều kiện đường thực tế như mưa lớn, ban đêm, sương mù, hay các loại biển báo đặc thù của từng quốc gia. Vì vậy, kết luận của đề tài — đặc biệt là khoảng cách hiệu năng giữa đèn và biển — gắn với phân bố dữ liệu cụ thể này. Việc khái quát hoá sang các điều kiện triển khai khác cần được kiểm chứng thêm. Tuy nhiên, chính vì phần lớn phân tích của đề tài là về cấu trúc nội tại của dữ liệu và độ tin cậy của quy trình đánh giá, nhiều kết luận về phương pháp — như tầm quan trọng của khử trùng lặp và của việc báo cáo theo lớp — vẫn có giá trị vượt ra ngoài bộ dữ liệu này."),
  H2("3.2. Công cụ đo — YOLOv8s huấn luyện hướng-EDA"),
  P(`Công cụ đo là một mô hình YOLOv8s được huấn luyện hai giai đoạn. Mô hình nền là ${tp.base_model} với khoảng ${vnp(tp.params_million, 2)} triệu tham số và ${vnp(tp.gflops, 1)} GFLOPs. Giai đoạn một huấn luyện ở độ phân giải ${tp.stage1_imgsz} với batch ${tp.stage1_batch} trong tối đa ${tp.stage1_epochs} epoch; giai đoạn hai tinh chỉnh ở độ phân giải ${tp.stage2_imgsz}. Bộ tối ưu là ${tp.optimizer} kết hợp cosine learning rate, và đặc biệt không dùng phép lật ngang/dọc vì lật sẽ làm sai ngữ nghĩa của các biển báo có hướng.`),
  P("Trọng số đã huấn luyện được đính kèm trong thư mục models/ của đề tài, phục vụ cho ứng dụng demo ở Chương 5 và cho việc tái lập. Toàn bộ số liệu gốc do mô hình sinh ra trên GPU Tesla T4 với Ultralytics phiên bản 8.4.95.", { before: 60 }),
  H2("3.3. Quy trình phân tích khám phá của pipeline nguồn"),
  P("Trước khi huấn luyện, pipeline nguồn thực hiện một loạt bước phân tích khám phá và kiểm tra chất lượng trên dữ liệu. Các bước này bao gồm: quét toàn bộ ảnh để phát hiện file hỏng hoặc nhãn sai định dạng; tính phân bố lớp và các thước đo mất cân bằng; phân tích hình học bounding box gồm kích thước và vị trí tâm; và quan trọng nhất với đề tài, phát hiện ảnh trùng và gần trùng bằng so khớp chính xác cùng perceptual hash, kể cả trùng lặp xuyên tập."),
  P("Chính vì pipeline đã có sẵn các kết quả kiểm tra chất lượng này, đề tài có thể thực hiện phân tích thứ cấp mà không cần tự quét lại toàn bộ dữ liệu. Điều nhóm bổ sung là biến những con số rời rạc đó thành một nội dung kiểm định có hệ thống, với các kiểm định thống kê và diễn giải — điều mà bản thân pipeline nguồn chưa làm. Tức là, dữ liệu thô của phân tích đã tồn tại; đóng góp của đề tài là khung phân tích và kết luận rút ra từ chúng.", { before: 40 }),
  H2("3.4. Phương pháp phân tích thứ cấp"),
  P("Quy trình phân tích của nhóm gồm ba bước, đều tự động và tái lập:"),
  numbered([run("Trích xuất. ", { bold: true }), run("Mô-đun extract_notebook_results.py đọc lại toàn bộ con số mà mô hình và bước EDA đã in ra, rồi cố định chúng thành các bảng CSV chuẩn trong thư mục results/tables/. Không có giá trị nào được nhập tay.")], "m3"),
  numbered([run("Kiểm định. ", { bold: true }), run("Mô-đun audit_analysis.py tính phơi nhiễm rò rỉ dữ liệu, so sánh nhóm đèn với nhóm biển bằng kiểm định Mann–Whitney, và tính tương quan Spearman giữa độ hiếm và hiệu năng.")], "m3"),
  numbered([run("Trực quan hoá. ", { bold: true }), run("Mô-đun make_figures.py sinh các hình từ những bảng đã trích để đưa vào báo cáo, slide và ứng dụng.")], "m3"),
  P("Mọi bước đều cố định hạt giống ngẫu nhiên bằng 42, và có bộ kiểm thử tự động pytest kiểm tra tính đúng đắn của các con số. Nhờ đó, mọi bảng biểu và hình vẽ trong báo cáo đều truy vết được về tệp CSV nguồn và tái lập được.", { before: 60 }),
  H2("3.5. Chỉ số và kiểm định sử dụng"),
  P("Hiệu năng được đo bằng Precision, Recall, mAP@0,5 và mAP@0,5:0,95 theo chuẩn của thư viện Ultralytics. Khoảng cách giữa hai nhóm lớp được kiểm định bằng Mann–Whitney U một phía. Liên hệ giữa độ hiếm và hiệu năng được đo bằng tương quan hạng Spearman. Việc dùng các kiểm định phi tham số là có chủ đích, vì số lớp nhỏ và phân phối hiệu năng không nhất thiết chuẩn."),
  H2("3.6. Công cụ và môi trường"),
  P("Phân tích thứ cấp của đề tài được thực hiện bằng Python với các thư viện pandas cho xử lý bảng, scipy cho kiểm định thống kê, và matplotlib cho trực quan hoá. Suy luận của mô hình trong ứng dụng demo dùng thư viện Ultralytics và OpenCV. Ứng dụng web được xây bằng Streamlit. Toàn bộ phân tích thứ cấp chạy trên CPU, không cần GPU; chỉ có số liệu gốc là do mô hình sinh ra trên GPU trước đó."),
  P("Về mặt tổ chức, mọi tham số phân tích được tập trung, mọi bước cố định hạt giống ngẫu nhiên bằng 42, và có bộ kiểm thử tự động pytest kiểm tra các bất biến quan trọng: bảng phải đủ 15 lớp, các chỉ số tổng phải khớp với nguồn, khoảng cách đèn–biển phải có ý nghĩa, và số liệu rò rỉ phải đúng. Nhờ vậy, nếu bất kỳ khâu nào bị thay đổi ngoài ý muốn, kiểm thử sẽ phát hiện ngay.", { before: 40 }),
  pb(),
);

// ============================================================ CHƯƠNG 4
const pcRows = pc.map((r) => [r.class_name, r.category, vn(r.test_instances), vnp(r.precision, 3), vnp(r.recall, 3), vnp(r.ap50, 3), vnp(r.ap50_95, 3)]);
const pcHl = pc.map((r, i) => r.is_traffic_light ? i : -1).filter((i) => i >= 0);
const leakRows = B.leakage_rows.map((r) => [r.loai_trung_lap, vn(r.so_dong), pctv(r.pct_tren_tong_the / 100, 2)]);
A(
  H1("KẾT QUẢ KIỂM ĐỊNH", 4),
  H2("4.1. Hiệu năng tổng thể"),
  P(`Trên tập kiểm thử gồm ${vn(testM.images)} ảnh và ${vn(testM.instances)} đối tượng, mô hình đạt Precision ${vnp(testM.precision, 3)}, Recall ${vnp(testM.recall, 3)}, mAP@0,5 bằng ${vnp(testM.map50, 3)} và mAP@0,5:0,95 bằng ${vnp(testM.map50_95, 3)}. Trên tập kiểm định, mAP@0,5 đạt ${vnp(validM.map50, 3)}. Nhìn tổng thể đây là kết quả rất tốt. Tuy nhiên, ba mục tiếp theo sẽ cho thấy con số tổng này che giấu ba vấn đề quan trọng về độ tin cậy.`),
  H2("4.2. Câu hỏi A — Rò rỉ dữ liệu"),
  P("Chính bước phân tích khám phá đã phát hiện các ảnh trùng lặp bằng perceptual hash và so khớp chính xác. Tuy nhiên, quy trình chuẩn bị dữ liệu chỉ loại trùng khỏi tập huấn luyện, còn tập kiểm thử — vốn dùng để chấm điểm — thì không được khử trùng. Bảng dưới đây định lượng mức phơi nhiễm rò rỉ."),
  tcaption("Trùng lặp và rò rỉ dữ liệu"),
  table(["Loại trùng lặp", "Số dòng ảnh", "% trên tổng thể"], leakRows, [4560, 2400, 2400], { boldFirst: true, hl: [1, 3] }),
  ...figure("04_leakage.png", 650, 341, "Trùng lặp và rò rỉ — chỉ 91 ảnh bị loại khỏi tập huấn luyện."),
  P(`Con số đáng chú ý nhất là có ${vn(leak.exact_cross_split_duplicate_rows)} ảnh trùng chính xác và ${vn(leak.dhash_cross_split_collision_rows)} ảnh gần trùng nằm xuyên tập — tức xuất hiện đồng thời ở cả tập huấn luyện và tập đánh giá. Trong khi đó, quy trình chỉ loại ${vn(leak.removed_from_train)} ảnh trùng chính xác ra khỏi tập huấn luyện, đưa số ảnh train từ ${vn(leak.original_train_images)} xuống ${vn(leak.clean_train_images)}. Tập kiểm định và kiểm thử hoàn toàn không được khử trùng.`, { before: 60 }),
  P("Để hình dung quy mô, tập kiểm định và kiểm thử cộng lại có khoảng 1.439 ảnh. Với hàng trăm ảnh gần trùng xuyên tập, tỉ lệ ảnh đánh giá có “bản sao” trong tập huấn luyện là không nhỏ. Mỗi ảnh như vậy là một trường hợp mà mô hình gần như chắc chắn trả lời đúng vì đã thấy ảnh tương tự, làm điểm số bị thổi lên so với năng lực khái quát hoá thực.", { before: 40 }),
  callout("Nhận xét A", [
    `Quy trình chỉ loại ${vn(leak.removed_from_train)} ảnh trùng chính xác khỏi tập huấn luyện; tập valid/test dùng để chấm điểm không được khử trùng.`,
    `Vẫn còn ${vn(leak.exact_cross_split_duplicate_rows)} ảnh trùng chính xác và ${vn(leak.dhash_cross_split_collision_rows)} ảnh gần trùng xuyên tập trong dữ liệu đánh giá.`,
    "Hệ quả: một phần điểm mAP 0,970 có thể đến từ việc mô hình đã thấy những ảnh rất giống trong lúc huấn luyện. Con số thực có thể thấp hơn. Đây là cảnh báo về độ tin cậy, không phải kết luận rằng mô hình kém.",
  ], RED),
  H3("4.2.1. So sánh tập kiểm định và kiểm thử"),
  P(`Một dấu hiệu gián tiếp khác của việc dữ liệu đánh giá quá “dễ” là hiệu năng trên hai tập độc lập gần như bằng nhau: mAP@0,5 đạt ${vnp(validM.map50, 3)} trên tập kiểm định và ${vnp(testM.map50, 3)} trên tập kiểm thử. Sự ổn định này thường được xem là tốt, nhưng trong bối cảnh có rò rỉ xuyên tập, nó cũng nhất quán với giả thuyết rằng cả hai tập đánh giá đều chia sẻ nhiều ảnh giống với tập huấn luyện. Vì đề tài ở chế độ results-only, nhóm không thể tách bạch hai cách giải thích này bằng cách chạy lại mô hình; đây là một hạn chế được nêu rõ ở Chương 6.`),
  ...figure("08_valid_test.png", 600, 334, "Hiệu năng tổng thể trên tập kiểm định và kiểm thử gần như bằng nhau."),
  H2("4.3. Câu hỏi B — Khoảng cách đèn tín hiệu và biển báo"),
  P("Bảng đầy đủ dưới đây trình bày hiệu năng của tất cả 15 lớp trên tập kiểm thử, sắp xếp giảm dần theo mAP@0,5:0,95. Hai lớp đèn tín hiệu được tô đậm."),
  tcaption("Hiệu năng theo lớp trên tập kiểm thử (15 lớp)"),
  table(["Lớp", "Nhóm", "Inst.", "P", "R", "mAP@.5", "mAP@.5:.95"], pcRows, [2280, 1500, 900, 1140, 1140, 1200, 1200], { boldFirst: true, hl: pcHl, size: 20 }),
  P(`Có thể thấy rõ hai nhóm tách biệt: toàn bộ 13 lớp biển báo đạt mAP@0,5:0,95 từ ${vnp(pc[12].ap50_95, 3)} trở lên, trong khi hai lớp đèn tín hiệu chỉ đạt ${vnp(pc[13].ap50_95, 3)} (Green Light) và ${vnp(worst.ap50_95, 3)} (Red Light). Hai biểu đồ dưới đây minh hoạ trực quan khoảng cách này.`, { before: 60 }),
  ...figure("01_per_class_ap.png", 650, 351, "Hiệu năng theo lớp — hai lớp đèn (đỏ) tách hẳn khỏi 13 lớp biển báo."),
  ...figure("02_light_vs_sign.png", 640, 357, "Đèn tín hiệu kém hơn biển báo trên mọi chỉ số."),
  P(`Trung bình nhóm cho thấy khoảng cách rất lớn: nhóm đèn tín hiệu đạt mAP@0,5:0,95 trung bình ${vnp(lights.mean_map50_95, 3)}, so với ${vnp(signs.mean_map50_95, 3)} của nhóm biển báo — chênh lệch ${vnp(lvs.gap, 3)}. Để khẳng định khoảng cách này không phải ngẫu nhiên, nhóm áp dụng kiểm định Mann–Whitney U một phía, thu được p = ${vnp(lvs.mannwhitney_p, 4)}. Giá trị p nhỏ hơn 0,01 cho thấy khác biệt có ý nghĩa thống kê, dù nhóm đèn chỉ gồm hai lớp.`, { before: 60 }),
  callout("Nhận xét B", [
    `Trung bình mAP@0,5:0,95: đèn tín hiệu ${vnp(lights.mean_map50_95, 3)} so với biển báo ${vnp(signs.mean_map50_95, 3)} — chênh ${vnp(lvs.gap, 3)}.`,
    `Kiểm định Mann–Whitney một phía cho p = ${vnp(lvs.mannwhitney_p, 4)}: khác biệt có ý nghĩa thống kê.`,
    `Red Light là lớp kém nhất với mAP@0,5:0,95 = ${vnp(worst.ap50_95, 3)} và recall chỉ ${vnp(worst.recall, 3)} — nghĩa là mô hình bỏ sót gần một phần tư số đèn đỏ thật.`,
  ]),
  H2("4.4. Câu hỏi C — Hiếm không đồng nghĩa với khó"),
  P("Một giả thuyết tự nhiên để giải thích thất bại của đèn là “đèn kém vì hiếm”: lớp càng ít mẫu thì mô hình càng khó học. Nếu giả thuyết này đúng, số lượng đối tượng của một lớp và hiệu năng của lớp đó phải tương quan dương. Nhóm kiểm chứng bằng tương quan hạng Spearman."),
  P(`Kết quả: tương quan Spearman giữa số đối tượng trên tập kiểm thử và mAP@0,5:0,95 chỉ là ρ = ${vnp(rv.spearman_instances_vs_map5095, 3)} với p = ${vnp(rv.p_instances_vs_map5095, 2)} — tức gần như bằng không và không có ý nghĩa. Tức là, độ hiếm của lớp không liên hệ với hiệu năng của nó.`, { before: 60 }),
  ...figure("03_rarity_vs_ap.png", 640, 384, "Độ hiếm không giải thích được hiệu năng (ρ ≈ 0)."),
  P("Trường hợp Red Light cho thấy điều đó rõ nhất. Trên toàn corpus, Red Light là lớp có nhiều đối tượng nhất (787); trên tập test, Green Light (110) và Red Light (94) cũng nằm trong nhóm nhiều đối tượng nhất. Dù xét ở cấp nào thì hai lớp đèn cũng không hề hiếm, vậy mà chúng lại có mAP thấp nhất. Như vậy, giả thuyết mất cân bằng gây ra thất bại không đứng vững.", { before: 60 }),
  callout("Nhận xét C", [
    `Độ hiếm của lớp KHÔNG dự báo được hiệu năng: Spearman ρ = ${vnp(rv.spearman_instances_vs_map5095, 3)} (p = ${vnp(rv.p_instances_vs_map5095, 2)}).`,
    "Hai lớp đèn không hiếm ở cả cấp corpus lẫn cấp test, nhưng vẫn có mAP thấp nhất; do đó mất cân bằng không phải nguyên nhân.",
    "Nguyên nhân nhiều khả năng nằm ở bản chất vật thể: đèn tín hiệu nhỏ, ở xa, độ tương phản thấp. Giả thuyết nhầm lẫn giữa đèn đỏ và đèn xanh cần được kiểm chứng thêm bằng confusion matrix.",
  ]),
  H2("4.5. Thảo luận sâu theo lớp"),
  P("Ngoài bức tranh nhóm, việc nhìn vào từng lớp cũng cho nhiều thông tin. Trong nhóm biển báo, lớp Stop đạt hiệu năng cao nhất với mAP@0,5:0,95 bằng " + vnp(best.ap50_95, 3) + " — điều hợp lý vì biển Stop có hình dạng và màu sắc rất đặc trưng, khó nhầm với lớp khác. Các biển giới hạn tốc độ nhìn chung đều đạt trên 0,78; những biến thể có ít mẫu trên tập kiểm thử như Speed Limit 110 có phần thấp hơn nhưng vẫn vượt xa nhóm đèn."),
  P("Đáng chú ý là lớp Speed Limit 10 chỉ có ba đối tượng trên tập kiểm thử nhưng vẫn đạt mAP@0,5:0,95 bằng 0,830 — cao hơn cả hai lớp đèn vốn có hàng chục tới hơn một trăm đối tượng. Đây là một minh hoạ nữa cho thấy số lượng mẫu không quyết định hiệu năng; điều quyết định là lớp đó có dễ phân biệt về mặt thị giác hay không.", { before: 40 }),
  P("Trong nhóm đèn, sự khác biệt giữa Precision và Recall rất đáng bàn. Red Light có Precision " + vnp(worst.precision, 3) + " nhưng Recall chỉ " + vnp(worst.recall, 3) + ": nghĩa là khi mô hình đã quyết định có đèn đỏ thì thường đúng, nhưng nó bỏ sót một tỉ lệ lớn đèn đỏ thật. Điều này phù hợp với đặc điểm vật thể nhỏ, ở xa và tương phản thấp — mô hình đơn giản là không “nhìn thấy” nhiều đèn. Đây là dạng lỗi nguy hiểm trong bối cảnh an toàn giao thông, và củng cố khuyến nghị cần xử lý riêng cho nhóm đèn.", { before: 40 }),
  P("Biểu đồ dưới đây đặt mọi lớp vào không gian Precision–Recall để làm rõ đặc trưng lỗi. Có thể thấy các lớp biển báo tụ ở góc trên bên phải — cả Precision lẫn Recall đều cao. Hai lớp đèn lệch hẳn về phía Recall thấp, xác nhận rằng dạng lỗi chủ đạo của đèn là bỏ sót chứ không phải báo nhầm.", { before: 40 }),
  ...figure("07_precision_recall.png", 580, 387, "Không gian Precision–Recall theo lớp; đèn tín hiệu lệch về phía Recall thấp."),
  H2("4.6. Tổng hợp ba phát hiện"),
  P("Ba phát hiện bổ trợ chặt chẽ cho nhau về độ tin cậy của con số 0,970. Phát hiện A cho thấy con số này có thể được rò rỉ dữ liệu nâng đỡ. Phát hiện B cho thấy nó che giấu sự phân hoá lớn giữa hai nhóm lớp. Phát hiện C xác định nguyên nhân của sự phân hoá đó là bản chất loại vật thể chứ không phải mất cân bằng."),
  P("Ba phát hiện này củng cố lẫn nhau. Việc độ dịch phân phối giữa các tập rất nhỏ (Chương 3) loại trừ khả năng khoảng cách đèn–biển đến từ chênh lệch phân phối. Việc độ hiếm không tương quan với hiệu năng (phát hiện C) loại trừ khả năng mất cân bằng là nguyên nhân. Sau khi loại trừ hai cách giải thích cạnh tranh này, kết luận rằng nguyên nhân nằm ở bản chất loại vật thể trở nên vững chắc. Đây là giá trị của một phân tích kiểm định có hệ thống, so với việc chỉ công bố một con số tổng.", { before: 40 }),
  H2("4.7. Ý nghĩa thực tiễn"),
  P("Ba phát hiện có ý nghĩa thực tiễn trực tiếp đối với người triển khai một bộ phát hiện biển báo. Thứ nhất, không nên tin tuyệt đối vào một con số mAP tổng công bố nếu chưa biết quy trình chia tập có kiểm soát trùng lặp hay không; một mAP 0,97 trên tập bị rò rỉ có thể tương ứng với hiệu năng thực tế thấp hơn khi gặp dữ liệu mới hoàn toàn."),
  P("Thứ hai, trong bối cảnh an toàn giao thông, điểm yếu ở nhóm đèn tín hiệu là đáng lo ngại hơn nhiều so với điểm yếu ở một biển báo hiếm gặp. Recall thấp của đèn đỏ nghĩa là hệ thống có thể bỏ sót đèn đỏ thật — dạng lỗi nguy hiểm nhất. Một quy trình đánh giá chỉ nhìn mAP tổng sẽ không phát ra cảnh báo này, trong khi phân tích theo nhóm lớp lại làm nó nổi bật. Đây là lập luận thực tiễn cho việc luôn báo cáo hiệu năng phân rã theo lớp trong các ứng dụng liên quan tới an toàn.", { before: 40 }),
  pb(),
);

// ============================================================ CHƯƠNG 5
A(
  H1("ỨNG DỤNG DEMO", 5),
  H2("5.1. Mục tiêu và kiến trúc"),
  P("Để phần phân tích không dừng ở lý thuyết, nhóm xây dựng một ứng dụng web bằng Streamlit. Ứng dụng gồm nhiều tab tương ứng với các phần của báo cáo: Tổng quan, Rò rỉ dữ liệu, Đèn với Biển, Hiếm khác Khó, Demo phát hiện, và Bảng đầy đủ. Toàn bộ số liệu hiển thị được đọc từ thư mục results/, nên luôn khớp với báo cáo."),
  P("Điểm mới nằm ở tab Demo phát hiện. Thay vì chỉ vẽ hộp như một demo phát hiện thông thường, tab này nối kết quả phát hiện với chính độ tin cậy đã kiểm định của từng lớp. Đây là điểm giao giữa demo trực quan và nội dung kiểm định của đề tài.", { before: 60 }),
  H2("5.2. Luồng xử lý của tab Demo phát hiện"),
  P("Khi người dùng tải lên một ảnh giao thông hoặc chọn một ảnh mẫu, ứng dụng thực hiện các bước sau. Trước hết, mô hình YOLOv8 từ trọng số models/best.pt chạy suy luận và vẽ các hộp phát hiện lên ảnh. Sau đó, với mỗi phát hiện, ứng dụng tra cứu độ tin cậy đã kiểm định của lớp tương ứng từ bảng per_class_metrics.csv, và gắn một nhãn mức độ tin cậy: cao, trung bình hay thấp. Cuối cùng, nếu trong ảnh có bất kỳ phát hiện nào thuộc nhóm đèn tín hiệu, ứng dụng hiển thị một cảnh báo rằng kết quả cần thận trọng, vì nhóm đèn có mAP thấp trong kiểm định."),
  P("Nhờ cơ chế này, ứng dụng không chỉ trả lời “trong ảnh có gì” mà còn trả lời “kết quả đó đáng tin đến đâu” — đúng tinh thần kiểm định độ tin cậy của đề tài.", { before: 60 }),
  H2("5.3. Hai tình huống minh hoạ"),
  P("Tình huống thứ nhất là ảnh chứa biển báo. Mô hình phát hiện biển giới hạn tốc độ với độ tin cậy cao; vì lớp này thuộc nhóm biển báo có mAP kiểm định khoảng 0,85, ứng dụng hiển thị thông báo tích cực rằng toàn bộ phát hiện thuộc nhóm biển báo và có độ tin cậy cao."),
  P("Tình huống thứ hai là ảnh chứa đèn đỏ ở xa — đúng ca khó mà phân tích ở Chương 4 đã chỉ ra. Mô hình phát hiện được nhiều đèn đỏ nhưng độ tin cậy thấp; vì Red Light có mAP kiểm định chỉ 0,506, ứng dụng hiển thị cảnh báo rằng kết quả cần thận trọng. Đây là minh chứng sống động cho phát hiện B và C.", { before: 60 }),
  H2("5.4. Các tab phân tích"),
  P("Ngoài tab demo, ứng dụng còn năm tab trình bày kết quả phân tích. Tab Tổng quan liệt kê ba phát hiện chính dưới dạng thẻ, kèm biểu đồ hiệu năng theo lớp. Tab Rò rỉ dữ liệu hiển thị bảng và biểu đồ phơi nhiễm trùng lặp cùng lời giải thích về khoảng trống khử trùng. Tab Đèn với Biển trình bày khoảng cách hiệu năng và kết quả kiểm định Mann–Whitney. Tab Hiếm khác Khó cho thấy biểu đồ tương quan và hệ số Spearman. Cuối cùng, tab Bảng đầy đủ chứa bảng hiệu năng 15 lớp cùng các bảng mất cân bằng và kích thước vật thể. Nhờ mọi tab đều đọc từ cùng thư mục kết quả, số liệu trên ứng dụng luôn khớp với báo cáo."),
  H2("5.5. Triển khai"),
  P("Ứng dụng chạy hoàn toàn trên CPU nên không cần GPU để demo. Nếu môi trường thiếu thư viện suy luận hoặc thiếu trọng số, ứng dụng tự động hiển thị hai ảnh demo đã kết xuất sẵn thay vì báo lỗi, bảo đảm luôn trình bày được trong mọi hoàn cảnh. Người dùng có thể điều chỉnh ngưỡng tin cậy bằng thanh trượt để quan sát số lượng phát hiện thay đổi — một cách trực quan để thấy sự đánh đổi giữa Precision và Recall mà Chương 4 đã bàn."),
  pb(),
);

// ============================================================ CHƯƠNG 6
A(
  H1("KẾT LUẬN VÀ KHUYẾN NGHỊ", 6),
  H2("6.1. Tổng hợp phát hiện"),
  P("Ba câu hỏi nghiên cứu đặt ra ở Chương 1 đều đã được trả lời, và cả ba đều chỉ về cùng một hướng: con số mAP 0,970 cần được nhìn nhận thận trọng hơn."),
  bullet([run("A — Rò rỉ dữ liệu: ", { bold: true }), run(`${vn(leak.exact_cross_split_duplicate_rows)} ảnh trùng chính xác và ${vn(leak.dhash_cross_split_collision_rows)} ảnh gần trùng nằm xuyên tập chưa được khử ở dữ liệu đánh giá, khiến mAP có thể lạc quan.`)]),
  bullet([run("B — Phân hoá lớp: ", { bold: true }), run(`nhóm đèn ${vnp(lights.mean_map50_95, 3)} so với nhóm biển ${vnp(signs.mean_map50_95, 3)}, chênh ${vnp(lvs.gap, 3)} với p = ${vnp(lvs.mannwhitney_p, 4)}.`)]),
  bullet([run("C — Nguyên nhân: ", { bold: true }), run(`độ hiếm không liên quan (ρ = ${vnp(rv.spearman_instances_vs_map5095, 3)}); lỗi tập trung ở nhóm đèn do bản chất vật thể.`)]),
  P("Kết luận trung tâm: con số 0,970 vừa có thể được rò rỉ dữ liệu nâng đỡ, vừa che giấu sự phân hoá lớn giữa các lớp; mà sự phân hoá đó bắt nguồn từ loại vật thể, không phải từ mất cân bằng dữ liệu.", { before: 60 }),
  H2("6.2. Khuyến nghị"),
  numbered("Khử trùng lặp xuyên tập ở cả tập kiểm định và kiểm thử trước khi báo cáo mAP; đồng thời báo cáo thêm mAP sau khi loại các ảnh gần trùng để thấy mức chênh thực sự.", "r6"),
  numbered("Báo cáo AP theo từng lớp và theo nhóm (đèn/biển) thay vì chỉ mAP tổng, để không che giấu điểm yếu.", "r6"),
  numbered("Cải thiện riêng lớp đèn tín hiệu: tăng độ phân giải đầu vào, bổ sung mẫu đèn khó, và phân tích ma trận nhầm lẫn giữa đèn đỏ và đèn xanh.", "r6"),
  numbered("Không dựa vào oversampling lớp hiếm để sửa lỗi đèn, vì độ hiếm đã được chứng minh không phải nguyên nhân.", "r6"),
  H2("6.3. Hạn chế"),
  P("Đề tài chạy ở chế độ results-only nên chưa chạy lại mô hình sau khi khử rò rỉ để đo trực tiếp mức sụt mAP; kết luận về ảnh hưởng của rò rỉ mang tính định lượng phơi nhiễm, không phải đo can thiệp. Ngoài ra, nhóm đèn chỉ có hai lớp nên kiểm định thống kê tuy có ý nghĩa vẫn cần thận trọng khi khái quát hoá. Cuối cùng, mọi kết luận gắn với bộ dữ liệu và mô hình cụ thể được khảo sát."),
  H2("6.4. Hướng phát triển"),
  P("Có nhiều hướng mở rộng đề tài trong tương lai."),
  P("Hướng thứ nhất là thực hiện một thí nghiệm can thiệp. Thay vì chỉ định lượng phơi nhiễm rò rỉ, ta có thể khử triệt để trùng lặp ở mọi tập, huấn luyện và đánh giá lại, rồi so sánh mAP trước và sau. Mức chênh lệch sẽ cho biết chính xác con số 0,970 đã được rò rỉ nâng đỡ bao nhiêu — chuyển kết luận từ mức phơi nhiễm sang mức đo can thiệp.", { before: 40 }),
  P("Hướng thứ hai là phân tích sâu ma trận nhầm lẫn ở cấp đối tượng, để định lượng bao nhiêu phần lỗi của đèn là do nhầm đỏ thành xanh, bao nhiêu là do bỏ sót hoàn toàn. Điều này giúp chọn đúng giải pháp: nếu chủ yếu là nhầm màu thì cần đặc trưng màu tốt hơn, còn nếu chủ yếu là bỏ sót thì cần cải thiện khả năng phát hiện vật thể nhỏ.", { before: 40 }),
  P("Hướng thứ ba là thử nghiệm các giải pháp kỹ thuật cho vật thể nhỏ, như tăng độ phân giải đầu vào, dùng cắt-ghép ảnh khi huấn luyện, hoặc thêm một đầu phát hiện ở tỉ lệ tinh hơn, rồi đo lại riêng hiệu năng nhóm đèn để xem giải pháp nào thực sự thu hẹp khoảng cách.", { before: 40 }),
  P("Hướng thứ tư là mở rộng phân tích sang các bộ dữ liệu biển báo khác để kiểm tra tính khái quát. Nếu khoảng cách đèn–biển và tính không liên quan của độ hiếm lặp lại trên nhiều bộ dữ liệu, kết luận của đề tài sẽ mạnh hơn nhiều.", { before: 40 }),
  H2("6.5. Sản phẩm kèm theo"),
  P("Kèm theo báo cáo là một bộ sản phẩm hoàn chỉnh: pipeline phân tích tái lập trong thư mục src/ với bộ kiểm thử pytest đạt toàn bộ; một ứng dụng web Streamlit có tab demo phát hiện trực tiếp; báo cáo Word và báo cáo LaTeX; cùng bộ slide trình bày có kèm văn thuyết trình cho từng slide."),
  H2("6.6. Kết luận chung"),
  P("Qua đề tài, nhóm thấy rằng góc nhìn data-centric giúp phát hiện những vấn đề mà cách đánh giá chỉ nhìn vào một con số mAP dễ bỏ qua. Khi xem mô hình như một công cụ đo và tập trung phân tích dữ liệu, nhóm đã làm rõ ba điểm về độ tin cậy của một điểm số tưởng như rất tốt: nó có thể được data leakage nâng đỡ, nó che lấp chênh lệch lớn giữa các lớp, và chênh lệch đó đến từ đặc điểm của đèn tín hiệu chứ không phải mất cân bằng."),
  P("Điểm mà nhóm tâm đắc là việc gắn mỗi nhận định với một bằng chứng định lượng và loại trừ dần các cách giải thích khác, thay vì chỉ mô tả số liệu. Về mặt ứng dụng, các khuyến nghị ở mục 6.2 và tab demo cho thấy kết quả kiểm định có thể dùng được ngay. Với nhóm, thói quen kiểm tra lại độ tin cậy của một con số trước khi tin vào nó là điều đáng giữ khi làm phân tích dữ liệu.", { before: 40 }),
  pb(),
);

// ============================================================ TÀI LIỆU THAM KHẢO
A(
  H1("TÀI LIỆU THAM KHẢO"),
  ...[
    "J. Redmon, S. Divvala, R. Girshick, và A. Farhadi, “You Only Look Once: Unified, Real-Time Object Detection”, trong CVPR, tr. 779–788, 2016.",
    "J. Redmon và A. Farhadi, “YOLO9000: Better, Faster, Stronger”, trong CVPR, tr. 7263–7271, 2017.",
    "J. Redmon và A. Farhadi, “YOLOv3: An Incremental Improvement”, arXiv:1804.02767, 2018.",
    "A. Bochkovskiy, C.-Y. Wang, và H.-Y. M. Liao, “YOLOv4: Optimal Speed and Accuracy of Object Detection”, arXiv:2004.10934, 2020.",
    "G. Jocher, A. Chaurasia, và J. Qiu, “Ultralytics YOLOv8”, phiên bản 8.4.95, 2023.",
    "S. Ren, K. He, R. Girshick, và J. Sun, “Faster R-CNN: Towards Real-Time Object Detection with Region Proposal Networks”, trong NeurIPS, 2015.",
    "W. Liu và cộng sự, “SSD: Single Shot MultiBox Detector”, trong ECCV, tr. 21–37, 2016.",
    "T.-Y. Lin, P. Goyal, R. Girshick, K. He, và P. Dollár, “Focal Loss for Dense Object Detection”, trong ICCV, tr. 2980–2988, 2017.",
    "K. He, X. Zhang, S. Ren, và J. Sun, “Deep Residual Learning for Image Recognition”, trong CVPR, tr. 770–778, 2016.",
    "T.-Y. Lin và cộng sự, “Microsoft COCO: Common Objects in Context”, trong ECCV, tr. 740–755, 2014.",
    "M. Everingham và cộng sự, “The PASCAL Visual Object Classes (VOC) Challenge”, IJCV, tập 88, số 2, tr. 303–338, 2010.",
    "R. Padilla và cộng sự, “A Comparative Analysis of Object Detection Metrics with a Companion Open-Source Toolkit”, Electronics, tập 10, số 3, tr. 279, 2021.",
    "J. Stallkamp, M. Schlipsing, J. Salmen, và C. Igel, “The German Traffic Sign Recognition Benchmark: A Multi-class Classification Competition”, trong IJCNN, tr. 1453–1460, 2011.",
    "S. Houben và cộng sự, “Detection of Traffic Signs in Real-World Images: The German Traffic Sign Detection Benchmark”, trong IJCNN, 2013.",
    "pkdarabi, “Car Detection / Traffic Signs Dataset”, Kaggle, 2023.",
    "tudeptraine, “traffic-yolo-v2-run — Kaggle notebook”, 2026.",
    "A. Ng, “A Chat with Andrew on MLOps: From Model-centric to Data-centric AI”, DeepLearning.AI, 2021.",
    "C. G. Northcutt, A. Athalye, và J. Mueller, “Pervasive Label Errors in Test Sets Destabilize Machine Learning Benchmarks”, trong NeurIPS Datasets and Benchmarks, 2021.",
    "S. Kaufman, S. Rosset, C. Perlich, và O. Stitelman, “Leakage in Data Mining: Formulation, Detection, and Avoidance”, ACM TKDD, tập 6, số 4, tr. 1–21, 2012.",
    "B. Barz và J. Denzler, “Do We Train on Test Data? Purging CIFAR of Near-Duplicates”, Journal of Imaging, tập 6, số 6, tr. 41, 2020.",
    "C. Zauner, “Implementation and Benchmarking of Perceptual Image Hash Functions”, Upper Austria University of Applied Sciences, Hagenberg, 2010.",
    "M. Buda, A. Maki, và M. A. Mazurowski, “A Systematic Study of the Class Imbalance Problem in Convolutional Neural Networks”, Neural Networks, tập 106, tr. 249–259, 2018.",
    "J. M. Johnson và T. M. Khoshgoftaar, “Survey on Deep Learning with Class Imbalance”, Journal of Big Data, tập 6, số 1, tr. 1–54, 2019.",
    "H. B. Mann và D. R. Whitney, “On a Test of Whether one of Two Random Variables is Stochastically Larger than the Other”, The Annals of Mathematical Statistics, tập 18, số 1, tr. 50–60, 1947.",
    "C. Spearman, “The Proof and Measurement of Association between Two Things”, The American Journal of Psychology, tập 15, số 1, tr. 72–101, 1904.",
    "C. E. Shannon, “A Mathematical Theory of Communication”, The Bell System Technical Journal, tập 27, số 3, tr. 379–423, 1948.",
    "F. Pedregosa và cộng sự, “Scikit-learn: Machine Learning in Python”, JMLR, tập 12, tr. 2825–2830, 2011.",
    "C. R. Harris và cộng sự, “Array Programming with NumPy”, Nature, tập 585, tr. 357–362, 2020.",
    "P. Virtanen và cộng sự, “SciPy 1.0: Fundamental Algorithms for Scientific Computing in Python”, Nature Methods, tập 17, tr. 261–272, 2020.",
    "J. D. Hunter, “Matplotlib: A 2D Graphics Environment”, Computing in Science & Engineering, tập 9, số 3, tr. 90–95, 2007.",
  ].map((r, i) => P(`[${i + 1}] ${r}`, { noIndent: true, after: 70 })),
  pb(),
);

// ============================================================ PHỤ LỤC
A(
  H1("PHỤ LỤC"),
  H2("Phụ lục A. Cấu hình huấn luyện đầy đủ"),
  table(["Tham số", "Giá trị"], [
    ["Mô hình nền", tp.base_model], ["Mô hình dự phòng", tp.fallback_model],
    ["Giai đoạn 1", `imgsz=${tp.stage1_imgsz}, batch=${tp.stage1_batch}, epochs≤${tp.stage1_epochs}, lr0=${vnp(tp.stage1_lr0, 3)}`],
    ["Giai đoạn 2", `imgsz=${tp.stage2_imgsz}, batch=${tp.stage2_batch}, epochs≤${tp.stage2_epochs}, lr0=${vnp(tp.stage2_lr0, 4)}`],
    ["Bộ tối ưu", `${tp.optimizer} + cosine LR`],
    ["Lật ảnh", `fliplr=${tp.fliplr}, flipud=${tp.flipud} (không lật)`],
    ["Trọng số loss", `box=${vnp(tp.box, 1)}, cls=${vnp(tp.cls, 1)}, dfl=${vnp(tp.dfl, 1)}`],
    ["Tham số / GFLOPs", `${vnp(tp.params_million, 2)}M / ${vnp(tp.gflops, 1)}`],
    ["Môi trường", `Ultralytics ${tp.ultralytics}, GPU ${tp.device}`],
  ], [3200, 6160], { size: 21, boldFirst: true }),
  H2("Phụ lục B. Cấu trúc thư mục và hướng dẫn tái lập"),
  P("Cấu trúc chính của đề tài gồm: data/notebook/ chứa số liệu gốc; src/ chứa mã trích xuất, kiểm định và vẽ hình; results/tables và results/figs chứa kết quả sinh tự động; app/ chứa ứng dụng web và ảnh mẫu; models/ chứa trọng số YOLOv8; tests/ chứa kiểm thử; và reports/ chứa báo cáo cùng slide.", { noIndent: true }),
  P("Các bước tái lập: (1) cài môi trường bằng pip install -r requirements.txt; (2) chạy python -m src.run_pipeline để trích xuất, kiểm định và vẽ hình; (3) chạy pytest để kiểm thử; (4) chạy streamlit run app/app.py để mở ứng dụng demo. Mọi bước cố định hạt giống 42; số liệu gốc do mô hình sinh trên GPU Tesla T4, còn phân tích thứ cấp chạy hoàn toàn trên CPU.", { noIndent: true, before: 40 }),
  H2("Phụ lục C. Bảng hiệu năng đầy đủ 15 lớp"),
  table(["ID", "Lớp", "Nhóm", "Inst.", "P", "R", "mAP@.5", "mAP@.5:.95"],
    pc.map((r) => [String(r.class_id), r.class_name, r.category, vn(r.test_instances), vnp(r.precision, 3), vnp(r.recall, 3), vnp(r.ap50, 3), vnp(r.ap50_95, 3)]),
    [700, 2100, 1400, 900, 1090, 1090, 1090, 1290], { size: 20, boldFirst: false, hl: pcHl }),
  H2("Phụ lục D. Mã nguồn tiêu biểu"),
  P("Đoạn mã dưới đây trích từ mô-đun audit_analysis.py, thể hiện cách tính hai kiểm định trung tâm của đề tài: khoảng cách đèn–biển bằng Mann–Whitney và tương quan độ hiếm–hiệu năng bằng Spearman.", { noIndent: true }),
  ...[
    "def light_vs_sign(pc):",
    "    lights = pc[pc['is_traffic_light']]",
    "    signs  = pc[~pc['is_traffic_light']]",
    "    # Kiểm định một phía: biển báo tốt hơn đèn tín hiệu",
    "    u, p = stats.mannwhitneyu(signs['ap50_95'],",
    "                              lights['ap50_95'],",
    "                              alternative='greater')",
    "    gap = signs['ap50_95'].mean() - lights['ap50_95'].mean()",
    "    return gap, u, p",
    "",
    "def rarity_vs_ap(pc):",
    "    # Tương quan hạng giữa số instance và hiệu năng lớp",
    "    rho, p = stats.spearmanr(pc['test_instances'],",
    "                             pc['ap50_95'])",
    "    return rho, p",
  ].map((ln) => new Paragraph({ spacing: { after: 0, line: 240 }, indent: { left: 240 },
    children: [new TextRun({ text: ln || " ", font: "Consolas", size: 20, color: "1A1A1A" })] })),
  P("Đoạn mã tiếp theo minh hoạ cách gắn kết quả phát hiện của YOLOv8 với độ tin cậy đã kiểm định của lớp, được dùng trong tab demo của ứng dụng.", { noIndent: true, before: 120 }),
  ...[
    "def reliability_verdict(ap):",
    "    if ap < 0.65:  return 'Thấp'",
    "    if ap < 0.80:  return 'Trung bình'",
    "    return 'Cao'",
    "",
    "# Với mỗi hộp YOLO phát hiện:",
    "name = model.names[int(box.cls)]",
    "ap   = reliability_map()[name]        # mAP@.5:.95 đã kiểm định",
    "verdict = reliability_verdict(ap)",
    "if is_traffic_light(name):",
    "    warn('Nhóm đèn — kết quả cần thận trọng')",
  ].map((ln) => new Paragraph({ spacing: { after: 0, line: 240 }, indent: { left: 240 },
    children: [new TextRun({ text: ln || " ", font: "Consolas", size: 20, color: "1A1A1A" })] })),
);

// ============================================================ DOCUMENT
const doc = new Document({
  creator: "Nhóm KHMT836016/34/36",
  title: "Kiểm định độ tin cậy YOLOv8 phát hiện biển báo",
  numbering: { config: [
    { reference: "n", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 620, hanging: 300 } } } }] },
    ...["c1", "m3", "r6"].map((ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 620, hanging: 300 } } } }] })),
  ] },
  styles: { default: { document: { run: { font: HF, size: BODY } } } },
  sections: [{
    // Lề theo quy định trường: trên 3cm, dưới 3cm, trái 3cm, phải 2cm (1cm ≈ 567 twips).
    properties: { page: { margin: { top: 1701, bottom: 1701, left: 1701, right: 1134 } } },
    // Số trang canh giữa ĐẦU TRANG (theo mẫu LaTeX).
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 24 })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(ROOT, "reports", "BaoCao_KiemDinh_YOLOv8_BienBao.docx");
  fs.writeFileSync(out, buf);
  console.log("Đã tạo:", out, "(", (buf.length / 1024).toFixed(0), "KB )");
});
