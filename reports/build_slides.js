/**
 * Slide trình bày (tiếng Việt) — Kiểm định độ tin cậy YOLOv8 phát hiện biển báo.
 * Phiên bản đầy đủ (>= 40 slide) có mục demo phát hiện.
 * Đọc số liệu từ results/report_bundle.json và hình từ results/figs/.
 *
 *   PPTX_LIB=/path/to/pptxgenjs node build_slides.js
 */
const fs = require("fs");
const path = require("path");
const PptxGenJS = require(process.env.PPTX_LIB || "pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const B = JSON.parse(fs.readFileSync(path.join(ROOT, "results/report_bundle.json"), "utf8"));
const FIG = path.join(ROOT, "results/figs");
const OUT = path.join(__dirname, "Slide_KiemDinh_YOLOv8_BienBao.pptx");

const NAVY = "152238", NAVY2 = "1E3050", TEAL = "1C7293", RED = "E63946",
      AMBER = "F4A261", WHITE = "FFFFFF", OFF = "F4F7FA", INK = "1A2430", MUTED = "5A6B7D", GREEN2 = "2C7A4B";
const HF = "Cambria", BF = "Calibri";

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Nhóm KHMT836016/34/36";
pptx.title = "Kiểm định độ tin cậy YOLOv8 phát hiện biển báo";
const W = 13.3, H = 7.5, M = 0.6;
let SN = 0; // đếm slide để đánh số

const vn = (x, d = 0) => Number(x).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }).replace(/,/g, ".").replace(/\.(\d+)$/, ",$1");
const vnp = (x, d = 3) => String(Number(x).toFixed(d)).replace(".", ",");

// ---- Văn thuyết trình (speaker notes) cho từng slide, theo đúng thứ tự tạo slide ----
const NOTES = [
  // 1 Bìa
  "Kính chào thầy/cô và các bạn. Nhóm em gồm ba thành viên xin trình bày đề tài môn Phân tích dữ liệu: Kiểm định độ tin cậy của một bộ phát hiện biển báo giao thông YOLOv8. Điểm khác biệt của đề tài là không chạy đua độ chính xác, mà đặt câu hỏi con số mAP 0,97 rất cao kia có thực sự đáng tin và đồng đều hay không. Bộ dữ liệu gồm gần 5.000 ảnh, 15 lớp, và khoảng cách hiệu năng giữa hai nhóm lớp lên tới 0,31 — đó là điều nhóm sẽ làm rõ hôm nay.",
  // 2 Nhóm
  "Đề tài do ba thành viên cùng thực hiện: Huỳnh Phát Lợi, Đoàn Huỳnh Thanh Tú và Võ Phú Vinh, đều thuộc chuyên ngành Khoa học máy tính khoá 36. Công việc được chia đều cho ba khâu chính: trích xuất và xử lý số liệu, phân tích thống kê kiểm định, và xây dựng ứng dụng cùng báo cáo.",
  // 3 Agenda
  "Bài trình bày gồm sáu phần. Đầu tiên là giới thiệu và định vị đề tài. Thứ hai là dữ liệu và phương pháp. Sau đó là ba phát hiện kiểm định, em ký hiệu là A, B, C — đây là phần trọng tâm. Tiếp theo là phần demo ứng dụng phát hiện trực tiếp, và cuối cùng là kết luận với khuyến nghị. Em sẽ đi lần lượt để mạch câu chuyện được liền lạc.",
  // 4 Bối cảnh
  "Trước hết là bối cảnh. Phát hiện biển báo là bài toán nền tảng của hỗ trợ lái. Với YOLOv8, độ chính xác trên các bộ chuẩn đã rất cao — như notebook nguồn của đề tài đạt mAP 0,970. Vấn đề là: khi con số tổng đã bão hoà thì thêm một mô hình nữa gần như không còn giá trị. Câu hỏi đáng giá chuyển thành: con số đó có đáng tin và đồng đều không. Đây là khoảng trống mà góc nhìn phân tích dữ liệu có thể lấp vào.",
  // 5 Vấn đề
  "Cụ thể, notebook đạt mAP 0,970 nhìn tổng thể rất tốt. Nhưng nhóm đặt ba nghi vấn: con số đó có bị rò rỉ dữ liệu nâng đỡ không; nó có đồng đều giữa các lớp không; và điểm yếu thực sự nằm ở đâu. Cách tiếp cận của nhóm là xem mô hình đã huấn luyện như một công cụ đo cố định — không huấn luyện lại, chỉ phân tích lại kết quả mà nó sinh ra, tức chế độ results-only.",
  // 6 Định vị
  "Slide này định vị đề tài. Có hai hướng tiếp cận: model-centric là chạy đua cải tiến mô hình để đẩy chỉ số lên; còn data-centric là xem dữ liệu và độ tin cậy của kết quả là đối tượng nghiên cứu. Đề tài của nhóm đi theo hướng data-centric: mô hình chỉ là công cụ đo, đầu ra là bằng chứng thống kê về độ tin cậy và sự phân hoá giữa các lớp, phương pháp là kiểm định thống kê chứ không phải huấn luyện mạng.",
  // 7 Data-centric
  "Vì sao chọn hướng data-centric? Andrew Ng đã lập luận rằng khi kiến trúc mô hình bão hoà, cải thiện tiếp theo đến từ việc hiểu và làm sạch dữ liệu. Một con số mAP cao chưa chắc đáng tin nếu tập đánh giá bị rò rỉ, và một con số trung bình có thể che giấu vài lớp rất yếu. Hướng data-centric cho phép nhóm phát hiện đúng những rủi ro đó — điều mà một báo cáo chỉ nêu mAP tổng sẽ bỏ sót.",
  // 8 Câu hỏi
  "Từ đó nhóm đặt ba câu hỏi nghiên cứu. Câu A về rò rỉ dữ liệu: mức độ trùng lặp xuyên tập là bao nhiêu và ảnh hưởng thế nào tới độ tin cậy của mAP. Câu B về phân hoá lớp: hiệu năng có đồng đều không, nếu không thì khoảng cách lớn nhất ở đâu. Câu C về nguyên nhân: khoảng cách đó do độ hiếm của lớp hay do bản chất vật thể. Ba câu này tương ứng ba phần chính của báo cáo.",
  // 9 Đóng góp
  "Đề tài có bốn đóng góp. Một là định lượng phơi nhiễm rò rỉ và chỉ ra khoảng trống trong quy trình khử trùng. Hai là chứng minh bằng kiểm định thống kê rằng đèn tín hiệu kém hơn biển báo có ý nghĩa. Ba là bác bỏ giả thuyết mất cân bằng gây ra thất bại, bằng bằng chứng tương quan. Bốn là một bộ sản phẩm hoàn chỉnh: pipeline tái lập, ứng dụng demo, báo cáo và slide.",
  // 10 Divider 2
  "Chuyển sang phần hai: dữ liệu và phương pháp. Em sẽ mô tả bộ dữ liệu, công cụ đo là YOLOv8, và quy trình phân tích ba bước mà nhóm dùng để bảo đảm mọi số liệu đều tái lập được.",
  // 11 Dữ liệu KPI
  "Bộ dữ liệu pkdarabi/cardetection gồm 15 lớp, gần 4.969 ảnh và 6.012 đối tượng, chia sẵn ba tập train, valid, test. Một điểm em muốn lưu ý là nhãn rất sạch: không có file nhãn thiếu, không có đối tượng lỗi, không có dòng sai định dạng. Vì vậy nếu có vấn đề thì nó nằm ở phân bố và cách chia tập, chứ không phải ở chất lượng gán nhãn.",
  // 12 15 lớp
  "15 lớp chia hai nhóm bản chất rất khác nhau. Nhóm đèn tín hiệu chỉ có 2 lớp: đèn xanh và đèn đỏ; và hơi bất ngờ là đèn đỏ lại là lớp phổ biến nhất toàn corpus với 787 đối tượng. Nhóm biển báo có 13 lớp: 12 biển giới hạn tốc độ và biển Stop. Đặc điểm quan trọng: đèn thường nhỏ, ở xa, tương phản thấp và hai màu đỏ-xanh dễ nhầm; còn biển báo có chữ số và hình dạng rõ ràng. Sự khác biệt bản chất này là mấu chốt của phần sau.",
  // 13 Split
  "Bảng thống kê theo split. Điều nhóm muốn nhấn mạnh là độ dịch phân phối lớp giữa ba tập rất nhỏ — Jensen–Shannon dưới 0,008 bit. Điều này quan trọng vì nó loại trừ một cách giải thích: chênh lệch hiệu năng KHÔNG đến từ việc các tập có phân phối lớp khác nhau. Vậy nguyên nhân phải nằm ở chỗ khác.",
  // 14 Mất cân bằng
  "Về mất cân bằng lớp: tỷ số mất cân bằng là 35,8 — lớp nhiều nhất gấp khoảng 36 lần lớp ít nhất. Entropy chuẩn hoá 0,953 và hệ số Gini 0,247 cho thấy mức mất cân bằng chỉ ở mức trung bình, không cực đoan. Con số này sẽ được dùng lại ở phần C khi ta kiểm tra xem mất cân bằng có phải nguyên nhân gây ra thất bại của đèn hay không.",
  // 15 Kích thước
  "Về kích thước vật thể: có tới 36,6% đối tượng là nhỏ, dưới 1% diện tích ảnh, và 18,7% là tí hon dưới 0,1%. Vật thể nhỏ khó phát hiện vì ít điểm ảnh mang thông tin. Biểu đồ cho thấy tăng độ phân giải đầu vào từ 640 lên 1024 làm giảm mạnh tỷ lệ vật thể có cạnh dưới 8 pixel — đây là cơ sở cho khuyến nghị tăng imgsz ở phần cuối.",
  // 16 Công cụ đo
  "Công cụ đo là một YOLOv8s huấn luyện hai giai đoạn: giai đoạn một ở độ phân giải 640, giai đoạn hai tinh chỉnh ở 768. Dùng AdamW với cosine learning rate, và cố ý không lật ảnh để giữ ngữ nghĩa biển báo có hướng. Trọng số best.pt được đính kèm ngay trong thư mục models của đề tài, cũng chính là trọng số dùng cho tab demo mà em sẽ trình bày sau.",
  // 17 Phương pháp 3 bước
  "Quy trình phân tích gồm ba bước tự động. Bước một, trích xuất: đọc lại toàn bộ con số notebook đã in và cố định thành bảng CSV chuẩn, không nhập tay số nào. Bước hai, kiểm định: tính phơi nhiễm rò rỉ, so sánh nhóm đèn với nhóm biển bằng Mann–Whitney, và tương quan độ hiếm với hiệu năng bằng Spearman. Bước ba, trực quan hoá: sinh hình từ các bảng đó. Mọi bước đều cố định seed và có kiểm thử tự động.",
  // 18 Nguyên tắc
  "Một nguyên tắc xuyên suốt là minh bạch và tái lập. Mọi bảng trong báo cáo đều ghi rõ tệp CSV nguồn, và mỗi CSV là đầu ra của một mô-đun. Không có con số nào nhập tay — chạy lại pipeline sẽ tạo ra y hệt. Nhóm đã đối chiếu 15 chỉ số chính khớp nhau giữa báo cáo Word, LaTeX và slide, và có bộ kiểm thử pytest để bảo vệ tính đúng đắn.",
  // 19 Divider A
  "Bây giờ vào phần A — phát hiện đầu tiên và cũng là phần đặt dấu hỏi cho con số 0,97: rò rỉ dữ liệu.",
  // 20 Rò rỉ là gì
  "Rò rỉ dữ liệu là khi thông tin của tập kiểm thử lọt vào quá trình huấn luyện. Với ảnh cắt từ video, dạng phổ biến nhất là ảnh trùng hoặc gần trùng xuất hiện ở nhiều tập. Ta dùng perceptual hash dHash để phát hiện ảnh gần trùng. Nếu không khử trùng ở tập đánh giá, mô hình có thể đã thấy ảnh rất giống lúc huấn luyện, khiến mAP cao hơn hiệu năng thực. Vì vậy bước đầu tiên là đo mức trùng lặp này.",
  // 21 Bảng+fig leakage
  "Kết quả đo phơi nhiễm rò rỉ: có 202 ảnh trùng chính xác và 423 ảnh gần trùng nằm xuyên tập — tức xuất hiện ở cả tập huấn luyện lẫn tập đánh giá. Đây không phải con số nhỏ. Đáng chú ý là chính notebook đã phát hiện ra các trùng lặp này trong bước EDA, nhưng cách xử lý lại chưa triệt để, như slide sau sẽ chỉ rõ.",
  // 22 Khoảng trống
  "Đây là khoảng trống quan trọng. Notebook chỉ loại 91 ảnh trùng chính xác ra khỏi tập huấn luyện, đưa train từ 3.530 xuống 3.439. Nhưng tập valid và test — vốn là tập dùng để chấm điểm — thì không được khử trùng. Nghĩa là dữ liệu đánh giá vẫn còn 202 ảnh trùng chính xác và 423 ảnh gần trùng với train. Mô hình chấm điểm trên những ảnh nó gần như đã thấy.",
  // 23 Nhận xét A
  "Tóm lại phần A: chỉ 91 ảnh bị loại khỏi train, tập đánh giá không được khử trùng, nên mAP 0,97 có thể lạc quan hơn hiệu năng thực. Đây không phải kết luận rằng mô hình kém, mà là một cảnh báo về độ tin cậy của con số. Khuyến nghị của nhóm là khử trùng ở cả valid và test, rồi báo cáo thêm mAP sau khi loại near-duplicate để thấy mức chênh thực sự.",
  // 24 Divider B
  "Sang phần B: ngay cả khi tạm gác chuyện rò rỉ, con số tổng 0,97 còn che giấu một vấn đề khác — sự phân hoá rất lớn giữa các lớp.",
  // 25 Bảng per-class
  "Đây là bảng hiệu năng đầy đủ 15 lớp, sắp theo mAP giảm dần. Nhìn hai cột: toàn bộ 13 lớp biển báo đều đạt từ 0,78 trở lên. Nhưng hai lớp đèn tín hiệu, tô đỏ ở cột phải, rơi hẳn xuống 0,576 và 0,506 — tách biệt rõ rệt khỏi phần còn lại. Con số trung bình 0,81 đã che giấu chính sự phân đôi này.",
  // 26 Fig per-class
  "Biểu đồ thanh cho thấy rõ hơn: 13 thanh xanh của biển báo nằm sát nhau trong khoảng 0,78 đến 0,90, còn hai thanh đỏ của đèn tín hiệu tụt hẳn xuống dưới. Hai đường đứt là trung bình mỗi nhóm — khoảng cách giữa chúng rất lớn. Đây là hình ảnh trực quan nhất cho phát hiện B.",
  // 27 Fig light vs sign
  "So sánh trực tiếp hai nhóm trên cả bốn chỉ số: precision, recall, mAP 0,5 và mAP 0,5:0,95. Nhóm đèn thua trên mọi chỉ số, và khoảng cách lớn nhất nằm ở mAP 0,5:0,95 — chỉ số nghiêm ngặt nhất về độ khớp hộp. Trung bình đèn chỉ 0,541 so với 0,854 của biển, chênh 0,313.",
  // 28 Mann-Whitney
  "Để khẳng định khoảng cách này không phải ngẫu nhiên, nhóm dùng kiểm định Mann–Whitney U — một kiểm định phi tham số, không giả định phân phối chuẩn, phù hợp khi so hai nhóm với số lớp nhỏ. Kết quả p bằng 0,0095, nhỏ hơn 0,01, tức khác biệt có ý nghĩa thống kê dù nhóm đèn chỉ có hai lớp. Khoảng cách 0,313 là thật, không phải nhiễu.",
  // 29 Nhận xét B
  "Tóm lại phần B: đèn 0,541 so với biển 0,854, khác biệt có ý nghĩa với p bằng 0,0095. Red Light là lớp kém nhất với mAP 0,506 và recall chỉ 0,755 — nghĩa là bỏ sót tới một phần tư số đèn đỏ thật. Và vì biển báo dễ chiếm đa số, mAP tổng 0,97 bị chúng kéo lên, che giấu điểm yếu ở đèn.",
  // 30 Divider C
  "Phần C trả lời câu hỏi quan trọng nhất: vì sao đèn kém? Một giả thuyết tự nhiên là đèn kém vì hiếm. Nhóm sẽ kiểm chứng và bác bỏ giả thuyết này.",
  // 31 Giả thuyết
  "Giả thuyết cần kiểm: lớp càng ít mẫu thì mô hình càng khó học, nên đèn kém vì thuộc nhóm hiếm. Nếu giả thuyết đúng, thì số lượng đối tượng của một lớp và hiệu năng của lớp đó phải tương quan dương. Cách kiểm chứng là tính tương quan hạng Spearman giữa số instance và mAP của từng lớp. Nếu hệ số gần 0 và p lớn thì độ hiếm không liên quan, và giả thuyết bị bác bỏ.",
  // 32 Fig rarity
  "Và đây là kết quả: hệ số Spearman chỉ 0,032, với p bằng 0,91 — tức hoàn toàn không có tương quan giữa số lượng đối tượng và hiệu năng. Nhìn biểu đồ, hai điểm đỏ của đèn nằm ở phía có nhiều đối tượng nhất nhưng lại thấp nhất về hiệu năng. Giả thuyết đèn kém vì hiếm bị bác bỏ.",
  // 33 Nghịch lý Red Light
  "Bằng chứng trực tiếp nhất là nghịch lý Red Light. Đèn đỏ là lớp phổ biến nhất toàn corpus với 787 đối tượng, vậy mà lại là lớp có hiệu năng kém nhất với mAP chỉ 0,506. Phổ biến nhất nhưng kém nhất — điều này chỉ ra rõ ràng rằng độ hiếm không phải nguyên nhân. Vậy nguyên nhân thật là gì?",
  // 34 Vì sao đèn khó
  "Nguyên nhân nằm ở bản chất vật thể, không phải ở số lượng mẫu. Đèn tín hiệu thường nhỏ và ở xa nên chiếm ít điểm ảnh; độ tương phản thấp trên nền trời hoặc đường phức tạp; và quan trọng nhất, đèn đỏ với đèn xanh cùng hình dạng chỉ khác màu nên dễ nhầm về mặt hình học. Recall của Red Light chỉ 0,755 cho thấy mô hình bỏ sót nhiều đèn thật. Đây là vấn đề của loại vật thể, cần giải pháp riêng.",
  // 35 Nhận xét C
  "Tóm lại phần C: Spearman 0,032 với p 0,91 cho thấy độ hiếm không liên quan tới hiệu năng; nghịch lý Red Light phổ biến nhất nhưng kém nhất bác bỏ giả thuyết mất cân bằng; nguyên nhân thực là bản chất vật thể. Hệ quả thực tiễn: không nên dùng oversampling lớp hiếm để sửa lỗi đèn, vì độ hiếm không phải nguyên nhân.",
  // 36 Divider demo
  "Để phần phân tích không dừng ở lý thuyết, nhóm xây dựng một ứng dụng web có tab demo phát hiện trực tiếp, gắn kết quả phát hiện với chính độ tin cậy mà nhóm đã kiểm định.",
  // 37 Demo kiến trúc
  "Ứng dụng Streamlit gồm nhiều tab tương ứng các phần: tổng quan, rò rỉ, đèn với biển, hiếm khác khó, demo phát hiện, và bảng đầy đủ. Điểm mới ở tab demo: người dùng tải một ảnh, YOLOv8 vẽ hộp phát hiện, và với mỗi phát hiện app đối chiếu độ tin cậy đã kiểm định của lớp đó. Nếu là nhóm đèn có mAP thấp, app cảnh báo kết quả cần thận trọng. Demo biến nội dung kiểm định thành công cụ dùng được.",
  // 38 Demo biển
  "Ví dụ thứ nhất, ảnh có biển báo. Mô hình phát hiện biển Speed Limit với độ tin cậy YOLO rất cao. Lớp này thuộc nhóm biển báo, mAP kiểm định khoảng 0,85, nên app hiển thị xanh: toàn bộ phát hiện thuộc nhóm biển báo, độ tin cậy cao. Đây là trường hợp ta có thể yên tâm với kết quả.",
  // 39 Demo đèn
  "Ví dụ thứ hai, ảnh có đèn đỏ ở xa — đúng ca khó mà phần B và C đã chỉ ra. Mô hình phát hiện được nhiều đèn đỏ nhưng độ tin cậy chỉ 0,29 đến 0,61. Vì Red Light có mAP kiểm định chỉ 0,506, app hiển thị cảnh báo đỏ: phát hiện thuộc nhóm đèn tín hiệu, kết quả cần thận trọng. Demo nói cho người dùng biết đúng lúc nào nên tin và lúc nào nên nghi ngờ.",
  // 40 Demo cách chạy
  "Về cách chạy: chỉ cần cài các thư viện trong requirements và gõ streamlit run app. App chạy hoàn toàn trên CPU, không cần GPU để demo. Nếu máy thiếu ultralytics hoặc trọng số, app tự hiển thị hai ảnh demo kết xuất sẵn thay vì báo lỗi. Người dùng có thể chỉnh ngưỡng tin cậy bằng thanh trượt để xem số lượng phát hiện thay đổi.",
  // 41 Kết luận
  "Kết luận, ba phát hiện bổ trợ chặt chẽ cho nhau. Con số 0,97 vừa có thể được rò rỉ dữ liệu nâng đỡ ở phần A, vừa che giấu sự phân hoá lớn giữa các lớp ở phần B, mà sự phân hoá đó bắt nguồn từ loại vật thể chứ không phải mất cân bằng ở phần C. Đó là góc nhìn kiểm định mà một báo cáo chỉ nêu mAP tổng sẽ bỏ sót hoàn toàn.",
  // 42 Khuyến nghị
  "Từ đó nhóm đưa ra bốn khuyến nghị. Một, khử trùng lặp ở cả valid và test trước khi báo mAP. Hai, báo cáo hiệu năng theo từng lớp và theo nhóm, không chỉ mAP tổng. Ba, cải thiện riêng lớp đèn: tăng độ phân giải, thêm mẫu đèn khó, phân tích nhầm lẫn đỏ-xanh. Bốn, không dùng oversampling lớp hiếm để sửa lỗi đèn vì độ hiếm không phải nguyên nhân.",
  // 43 Hạn chế + sản phẩm
  "Về hạn chế: đề tài ở chế độ results-only nên chưa đo trực tiếp mức sụt mAP sau khi khử rò rỉ, mới chỉ định lượng phơi nhiễm; nhóm đèn cũng chỉ có hai lớp nên cần thận trọng khi khái quát. Kèm theo báo cáo là bộ sản phẩm đầy đủ: pipeline tái lập với kiểm thử, ứng dụng web có tab demo, báo cáo Word và LaTeX, cùng bộ slide này.",
  // 44 Cảm ơn
  "Phần trình bày của nhóm đến đây là hết. Nhóm em xin cảm ơn thầy/cô và các bạn đã lắng nghe, và rất mong nhận được câu hỏi cùng góp ý để hoàn thiện đề tài.",
];

// Vùng nội dung: y từ ~1.55 tới ~6.85 (chân trang ở 7.0).
const CY0 = 1.62, CY1 = 6.72, FOOT = 6.98;
const KICK = "Kiểm định độ tin cậy YOLOv8 · Nhóm KHMT836016/34/36";

function footer(s) {
  // Đường kẻ mảnh + nhãn trái + số trang phải.
  s.addShape(pptx.ShapeType.line, { x: M, y: FOOT, w: W - 2 * M, h: 0, line: { color: "DDE3EA", width: 1 } });
  s.addText(KICK, { x: M, y: FOOT + 0.04, w: 9.5, h: 0.3, fontFace: BF, fontSize: 9, color: "9AA7B4", margin: 0, valign: "middle" });
  s.addText(String(SN), { x: W - M - 0.6, y: FOOT + 0.04, w: 0.6, h: 0.3, fontFace: BF, fontSize: 10, bold: true, color: TEAL, align: "right", margin: 0, valign: "middle" });
}
function decor(s) {
  // Vòng tròn nhạt góc trên-phải tạo điểm nhấn thị giác nhẹ.
  s.addShape(pptx.ShapeType.ellipse, { x: W - 1.5, y: -1.35, w: 3.0, h: 3.0, fill: { color: "EAF1F6" }, line: { width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: W - 0.75, y: 0.2, w: 0.9, h: 0.9, fill: { color: "DCE8F0" }, line: { width: 0 } });
}
function pageNum(s) {
  SN += 1;
  const note = NOTES[SN - 1];
  if (note) s.addNotes(note);
}
function light() { const s = pptx.addSlide(); s.background = { color: WHITE }; pageNum(s); decor(s); footer(s); return s; }
function dark() { const s = pptx.addSlide(); s.background = { color: NAVY }; pageNum(s); return s; }
function title(s, t, sub) {
  const fs2 = t.length <= 42 ? 29 : t.length <= 54 ? 25 : 22;
  // Ô vuông nhấn màu teal trước tiêu đề.
  s.addShape(pptx.ShapeType.roundRect, { x: M, y: 0.5, w: 0.17, h: 0.5, rectRadius: 0.03, fill: { color: TEAL }, line: { width: 0 } });
  s.addText(t, { x: M + 0.36, y: 0.4, w: W - 2 * M - 0.36, h: 0.78, fontFace: HF, fontSize: fs2, bold: true, color: INK, margin: 0, valign: "middle" });
  if (sub) s.addText(sub, { x: M + 0.36, y: 1.12, w: W - 2 * M - 0.36, h: 0.4, fontFace: BF, fontSize: 13.5, italics: true, color: MUTED, margin: 0 });
  return sub ? CY0 + 0.16 : CY0;
}
function card(s, x, y, w, h, fill) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.1, fill: { color: fill || OFF }, line: { color: "E6ECF2", width: 1 },
    shadow: { type: "outer", angle: 90, blur: 9, offset: 2, color: "9FB0C0", opacity: 0.22 } });
}
function badge(s, x, y, txt, fill, d) {
  const dia = d || 0.52;
  s.addShape(pptx.ShapeType.ellipse, { x, y, w: dia, h: dia, fill: { color: fill || TEAL }, line: { width: 0 },
    shadow: { type: "outer", angle: 90, blur: 5, offset: 1, color: fill || TEAL, opacity: 0.35 } });
  s.addText(String(txt), { x, y, w: dia, h: dia, align: "center", valign: "middle", fontFace: BF, fontSize: dia > 0.6 ? 18 : 15, bold: true, color: WHITE, margin: 0 });
}
function stat(s, x, y, w, v, l, color) {
  s.addText(String(v), { x, y, w, h: 1.0, align: "center", valign: "middle", fontFace: HF, fontSize: 42, bold: true, color: color || TEAL, margin: 0 });
  s.addText(l, { x, y: y + 0.98, w, h: 0.6, align: "center", valign: "top", fontFace: BF, fontSize: 12.5, color: MUTED, margin: 0 });
}
function finding(s, y, text, color) {
  const c = color || NAVY;
  s.addShape(pptx.ShapeType.roundRect, { x: M, y, w: W - 2 * M, h: 0.98, rectRadius: 0.1, fill: { color: c }, line: { width: 0 },
    shadow: { type: "outer", angle: 90, blur: 8, offset: 2, color: "6E7C8A", opacity: 0.3 } });
  // chấm nhấn amber
  s.addShape(pptx.ShapeType.ellipse, { x: M + 0.28, y: y + 0.4, w: 0.18, h: 0.18, fill: { color: AMBER }, line: { width: 0 } });
  s.addText(text, { x: M + 0.64, y: y + 0.08, w: W - 2 * M - 0.92, h: 0.82, fontFace: BF, fontSize: 14, color: WHITE, margin: 0, valign: "middle" });
}
function bullets(s, items, o) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: i !== items.length - 1 } })),
    { x: o.x, y: o.y, w: o.w, h: o.h, fontFace: BF, fontSize: o.fontSize || 15, color: o.color || INK, margin: 0, valign: o.valign || "top", paraSpaceAfter: o.gap ?? 9, lineSpacingMultiple: o.lh || 1.08 });
}
// Panel bullet lớn lấp đầy chiều dọc: thẻ bo góc + tiêu đề nhỏ + gạch đầu dòng căn giữa theo chiều dọc.
function bulletPanel(s, yTop, items, o = {}) {
  const yEnd = o.yEnd || (CY1 - 0.02);
  const h = yEnd - yTop;
  card(s, M, yTop, W - 2 * M, h, o.fill || "F7FAFC");
  // dải nhấn dọc mảnh bên trái panel
  s.addShape(pptx.ShapeType.roundRect, { x: M + 0.18, y: yTop + 0.22, w: 0.09, h: h - 0.44, rectRadius: 0.04, fill: { color: o.accent || TEAL }, line: { width: 0 } });
  bullets(s, items, { x: M + 0.62, y: yTop + 0.1, w: W - 2 * M - 0.95, h: h - 0.2, fontSize: o.fontSize || 16.5, gap: o.gap ?? 14, valign: "middle", color: INK });
}
function img(s, file, o) { s.addImage({ path: path.join(FIG, file), x: o.x, y: o.y, w: o.w, h: o.h }); }
function divider(num, heading, subtitle, notes) {
  const s = dark(notes);
  badge(s, M, 2.5, num, TEAL, 0.82);
  s.addText(heading, { x: M, y: 3.5, w: W - 2 * M, h: 0.9, fontFace: HF, fontSize: 38, bold: true, color: WHITE, margin: 0 });
  s.addText(subtitle, { x: M, y: 4.42, w: W - 2.6, h: 0.6, fontFace: BF, fontSize: 16, color: "9DB2C6", margin: 0 });
  return s;
}
function simpleTable(s, header, rows, colW, o = {}) {
  const y = o.y, fontSize = o.fontSize || 12.5;
  const head = header.map((h, i) => ({ text: h, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: fontSize + 0.5, align: i === 0 ? "left" : "center" } }));
  const body = rows.map((r, ri) => r.map((c, i) => ({ text: String(c), options: { fontSize, fill: { color: ri % 2 ? "EEF3F7" : "FFFFFF" }, align: i === 0 ? "left" : "center", color: o.hlRows && o.hlRows.includes(ri) ? RED : INK, bold: o.hlRows && o.hlRows.includes(ri) } })));
  s.addTable([head, ...body], { x: o.x ?? M, y, w: colW.reduce((a, b) => a + b, 0), colW, border: { type: "solid", color: "DCE3EB", pt: 1 }, fontFace: BF, rowH: o.rowH || 0.3, valign: "middle", margin: 0.05 });
}

const ds = B.dataset, testM = B.overall.test, validM = B.overall.valid, lvs = B.light_vs_sign, rv = B.rarity, leak = B.leakage, imb = B.imbalance_overall, sig = B.size_signals, tp = B.train_plan;
const worst = B.per_class[B.per_class.length - 1], best = B.per_class[0];
const members = B.members;

// ============================================================ 1. BÌA
{
  const s = dark("Xin chào thầy/cô và các bạn. Nhóm em trình bày đề tài kiểm định độ tin cậy của bộ phát hiện biển báo YOLOv8.");
  s.addShape(pptx.ShapeType.ellipse, { x: 9.9, y: -1.5, w: 5.4, h: 5.4, fill: { color: NAVY2 }, line: { width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 11.2, y: 4.6, w: 3.2, h: 3.2, fill: { color: TEAL }, line: { width: 0 }, transparency: 72 });
  s.addText("BÁO CÁO ĐỀ TÀI · MÔN PHÂN TÍCH DỮ LIỆU", { x: M, y: 1.2, w: 9.4, h: 0.34, fontFace: BF, fontSize: 13, bold: true, color: AMBER, charSpacing: 2, margin: 0 });
  s.addText("Kiểm định độ tin cậy của bộ phát hiện\nbiển báo giao thông YOLOv8", { x: M, y: 1.72, w: 9.4, h: 1.7, fontFace: HF, fontSize: 34, bold: true, color: WHITE, margin: 0, lineSpacingMultiple: 1.06 });
  s.addText("Rò rỉ dữ liệu · Kích thước vật thể · Khoảng cách hiệu năng đèn tín hiệu – biển báo", { x: M, y: 3.66, w: 9.4, h: 0.5, fontFace: BF, fontSize: 15, color: "9DB2C6", margin: 0 });
  const sv = [[vn(ds.total_images_scanned), "ảnh"], ["15", "lớp"], [vnp(testM.map50, 3), "mAP@0,5"], [vnp(lvs.gap, 3), "chênh đèn–biển"]];
  sv.forEach((v, i) => { const x = M + i * 2.5;
    s.addText(v[0], { x, y: 4.5, w: 2.3, h: 0.6, fontFace: HF, fontSize: 26, bold: true, color: WHITE, margin: 0 });
    s.addText(v[1], { x, y: 5.08, w: 2.3, h: 0.32, fontFace: BF, fontSize: 11.5, color: "8FA6BC", margin: 0 }); });
  s.addText("Nhóm: Huỳnh Phát Lợi (KHMT836016) · Đoàn Huỳnh Thanh Tú (KHMT836034) · Võ Phú Vinh (KHMT836036)",
    { x: M, y: 6.4, w: 11.5, h: 0.4, fontFace: BF, fontSize: 12, color: "7C93AA", margin: 0 });
}

// ============================================================ 2. NHÓM
{
  const s = light("Nhóm gồm ba thành viên. Đề tài là kết quả làm việc chung.");
  const y = title(s, "Nhóm thực hiện");
  const cardTop = y + 0.35, cardH = 3.5, roleCol = [TEAL, "1E5F8C", RED], roles = ["Trích xuất & xử lý số liệu", "Phân tích thống kê kiểm định", "Ứng dụng & báo cáo"];
  members.forEach((m, i) => { const x = M + i * 4.1;
    card(s, x, cardTop, 3.85, cardH);
    s.addShape(pptx.ShapeType.roundRect, { x, y: cardTop, w: 3.85, h: 0.14, rectRadius: 0.03, fill: { color: roleCol[i] }, line: { width: 0 } });
    s.addShape(pptx.ShapeType.ellipse, { x: x + 1.3, y: cardTop + 0.5, w: 1.25, h: 1.25, fill: { color: roleCol[i] }, line: { width: 0 },
      shadow: { type: "outer", angle: 90, blur: 6, offset: 2, color: roleCol[i], opacity: 0.35 } });
    s.addText(m.name.split(" ").pop().substring(0, 1), { x: x + 1.3, y: cardTop + 0.5, w: 1.25, h: 1.25, align: "center", valign: "middle", fontFace: HF, fontSize: 38, bold: true, color: WHITE, margin: 0 });
    s.addText(m.name, { x: x + 0.2, y: cardTop + 1.95, w: 3.45, h: 0.5, align: "center", fontFace: HF, fontSize: 18, bold: true, color: INK, margin: 0 });
    s.addText(m.id, { x: x + 0.2, y: cardTop + 2.42, w: 3.45, h: 0.35, align: "center", fontFace: BF, fontSize: 13, color: MUTED, margin: 0 });
    s.addText(roles[i], { x: x + 0.2, y: cardTop + 2.82, w: 3.45, h: 0.5, align: "center", fontFace: BF, fontSize: 12.5, italics: true, color: roleCol[i], margin: 0 });
  });
  s.addText("Chuyên ngành Khoa học máy tính · Khoá 36 (2025–2027) · Trường ĐH Sư phạm TP.HCM",
    { x: M, y: cardTop + cardH + 0.35, w: W - 2 * M, h: 0.5, align: "center", fontFace: BF, fontSize: 14, color: MUTED, margin: 0 });
}

// ============================================================ 3. AGENDA
{
  const s = light("Bài trình bày gồm sáu phần: giới thiệu, dữ liệu & phương pháp, ba phát hiện kiểm định, demo và kết luận.");
  const y = title(s, "Nội dung trình bày");
  const items = [
    ["1", "Giới thiệu & định vị đề tài", "Vấn đề, câu hỏi nghiên cứu, đóng góp", TEAL],
    ["2", "Dữ liệu & phương pháp", "Bộ dữ liệu, công cụ đo YOLOv8, quy trình 3 bước", "1E5F8C"],
    ["A", "Rò rỉ dữ liệu", "mAP 0,97 đáng tin đến đâu?", AMBER],
    ["B", "Khoảng cách đèn – biển", "Hiệu năng có đồng đều giữa các lớp?", RED],
    ["C", "Hiếm ≠ khó", "Nguyên nhân là mất cân bằng hay loại vật thể?", GREEN2],
    ["6", "Demo & kết luận", "Demo phát hiện trực tiếp, khuyến nghị", NAVY],
  ];
  const rowH = 1.5, rowGap = (CY1 - y - 3 * rowH) / 2;
  items.forEach((it, i) => { const x = M + (i % 2) * 6.2; const yy = y + Math.floor(i / 2) * (rowH + rowGap);
    card(s, x, yy, 5.9, rowH);
    s.addShape(pptx.ShapeType.roundRect, { x, y: yy, w: 0.15, h: rowH, rectRadius: 0.05, fill: { color: it[3] }, line: { width: 0 } });
    badge(s, x + 0.36, yy + 0.47, it[0], it[3], 0.56);
    s.addText(it[1], { x: x + 1.18, y: yy + 0.26, w: 4.6, h: 0.5, fontFace: HF, fontSize: 17, bold: true, color: it[3] === AMBER ? "B5651D" : it[3], margin: 0 });
    s.addText(it[2], { x: x + 1.18, y: yy + 0.8, w: 4.6, h: 0.55, fontFace: BF, fontSize: 13, color: MUTED, margin: 0, valign: "top" });
  });
}

// ============================================================ 4. BỐI CẢNH
{
  const s = light();
  const y = title(s, "Bối cảnh: bài toán phát hiện biển báo");
  bulletPanel(s, y, [
    "Phát hiện biển báo và đèn tín hiệu là bài toán nền tảng của hệ hỗ trợ lái và xe tự hành.",
    "Với các mô hình hiện đại như YOLOv8, độ chính xác trên bộ dữ liệu chuẩn đã rất cao.",
    "Notebook nguồn đạt mAP@0,5 = 0,970 trên tập test — nhìn tổng thể rất tốt.",
    "Khi con số tổng đã bão hoà, câu hỏi có giá trị chuyển từ “mô hình nào tốt hơn” sang “con số đó có đáng tin và đồng đều không”.",
  ], { yEnd: 5.55 });
  finding(s, 5.72, "Đây là khoảng trống mà một góc nhìn phân tích dữ liệu (data-centric) có thể lấp vào: kiểm định độ tin cậy của con số, thay vì chạy đua thêm một mô hình nữa.");
}

// ============================================================ 5. VẤN ĐỀ
{
  const s = light();
  const y = title(s, "Vấn đề: con số 0,97 có đáng tin và đồng đều không?");
  const ch = 3.35;
  card(s, M, y, 5.9, ch);
  s.addShape(pptx.ShapeType.roundRect, { x: M + 0.32, y: y + 0.3, w: 1.0, h: 0.44, rectRadius: 0.06, fill: { color: "E3E9EF" }, line: { width: 0 } });
  s.addText("BỐI CẢNH", { x: M + 0.32, y: y + 0.3, w: 1.0, h: 0.44, align: "center", valign: "middle", fontFace: BF, fontSize: 10, bold: true, color: MUTED, margin: 0 });
  s.addText("Điều đã có", { x: M + 0.32, y: y + 0.92, w: 5.2, h: 0.44, fontFace: HF, fontSize: 20, bold: true, color: INK, margin: 0 });
  s.addText("Một notebook Kaggle huấn luyện YOLOv8s trên pkdarabi/cardetection đạt mAP@0,5 = 0,970 trên tập kiểm thử — một con số rất cao.", { x: M + 0.32, y: y + 1.5, w: 5.3, h: 1.6, fontFace: BF, fontSize: 15.5, color: INK, margin: 0, valign: "top", lineSpacingMultiple: 1.15 });
  card(s, M + 6.2, y, 5.9, ch, NAVY);
  s.addShape(pptx.ShapeType.roundRect, { x: M + 6.52, y: y + 0.3, w: 1.3, h: 0.44, rectRadius: 0.06, fill: { color: "2C4160" }, line: { width: 0 } });
  s.addText("CÂU HỎI", { x: M + 6.52, y: y + 0.3, w: 1.3, h: 0.44, align: "center", valign: "middle", fontFace: BF, fontSize: 10, bold: true, color: AMBER, margin: 0 });
  s.addText("Câu hỏi của đề tài", { x: M + 6.52, y: y + 0.92, w: 5.2, h: 0.44, fontFace: HF, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  s.addText("Con số đó có bị rò rỉ dữ liệu nâng đỡ không? Có đồng đều giữa các lớp không? Điểm yếu thực sự nằm ở đâu?", { x: M + 6.52, y: y + 1.5, w: 5.3, h: 1.6, fontFace: BF, fontSize: 15.5, color: "D6E0EA", margin: 0, valign: "top", lineSpacingMultiple: 1.15 });
  finding(s, y + ch + 0.35, "Cách tiếp cận: xem mô hình đã huấn luyện là CÔNG CỤ ĐO cố định — không huấn luyện lại, chỉ phân tích lại kết quả (results-only).", TEAL);
}

// ============================================================ 6. ĐỊNH VỊ
{
  const s = light();
  const y = title(s, "Định vị đề tài — hai hướng tiếp cận");
  const rows = [
    ["Câu hỏi", "Mô hình nào tốt hơn?", "mAP 0,97 có đáng tin & đồng đều?"],
    ["Vai trò mô hình", "Đối tượng — huấn luyện, so sánh", "Công cụ đo cố định — phân tích lại"],
    ["Đầu ra", "Benchmark, mAP tổng", "Bằng chứng thống kê về độ tin cậy"],
    ["Phương pháp", "Thiết kế, huấn luyện mạng", "Kiểm định thống kê, phân tích rò rỉ"],
  ];
  s.addTable([[
    { text: "", options: { fill: { color: WHITE } } },
    { text: "Hướng model-centric", options: { bold: true, color: WHITE, fill: { color: MUTED }, fontSize: 14, align: "center" } },
    { text: "Hướng data-centric (đề tài này)", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 14, align: "center" } },
  ], ...rows.map((r) => [
    { text: r[0], options: { bold: true, fontSize: 13, fill: { color: OFF } } },
    { text: r[1], options: { fontSize: 13 } },
    { text: r[2], options: { fontSize: 13, fill: { color: "E8F4F0" } } },
  ])], { x: M, y, w: W - 2 * M, colW: [2.6, 4.55, 4.95], border: { type: "solid", color: "DCE3EB", pt: 1 }, fontFace: BF, rowH: 0.62, valign: "middle", margin: 0.06 });
  finding(s, y + 3.5, "Không chạy đua mAP, mà kiểm định độ tin cậy của mAP — đúng tinh thần data-centric của môn Phân tích dữ liệu.");
}

// ============================================================ 7. TÁI SỬ DỤNG
{
  const s = light();
  const y = title(s, "Vì sao chọn hướng data-centric?");
  bulletPanel(s, y, [
    "Khi kiến trúc mô hình đã bão hoà, cải thiện tiếp theo đến từ việc hiểu và làm sạch DỮ LIỆU (Andrew Ng).",
    "Một mAP cao chưa chắc đáng tin nếu tập đánh giá bị rò rỉ dữ liệu.",
    "Một mAP trung bình có thể che giấu vài lớp rất yếu — cần nhìn theo từng lớp.",
    "Hướng data-centric giúp phát hiện đúng những rủi ro này — điều mà một báo cáo chỉ nêu mAP tổng sẽ bỏ sót.",
  ], { yEnd: 5.55 });
  finding(s, 5.72, "Vì vậy đề tài xem chính DỮ LIỆU và độ tin cậy của kết quả là đối tượng nghiên cứu, thay vì chạy đua thêm một mô hình.");
}

// ============================================================ 8. CÂU HỎI NGHIÊN CỨU
{
  const s = light("Ba câu hỏi kiểm định, tương ứng ba phần chính của báo cáo.");
  const y = title(s, "Ba câu hỏi nghiên cứu");
  const qs = [
    ["A", "Rò rỉ dữ liệu", "Mức độ rò rỉ xuyên split là bao nhiêu, và nó ảnh hưởng thế nào tới độ tin cậy của mAP?", AMBER],
    ["B", "Phân hoá lớp", "Hiệu năng có đồng đều giữa các lớp không? Nếu không, khoảng cách lớn nhất nằm ở đâu?", RED],
    ["C", "Nguyên nhân", "Khoảng cách đó do độ hiếm của lớp (mất cân bằng) hay do bản chất vật thể?", GREEN2],
  ];
  const gap = (CY1 - y - 3 * 1.4) / 2;  // dàn đều 3 thẻ khắp vùng nội dung
  qs.forEach((q, i) => { const yy = y + i * (1.4 + gap);
    card(s, M, yy, W - 2 * M, 1.4);
    // dải màu dọc bên trái theo màu câu hỏi
    s.addShape(pptx.ShapeType.roundRect, { x: M, y: yy, w: 0.16, h: 1.4, rectRadius: 0.06, fill: { color: q[3] }, line: { width: 0 } });
    badge(s, M + 0.42, yy + 0.44, q[0], q[3], 0.56);
    s.addText(q[1], { x: M + 1.25, y: yy + 0.18, w: 3.1, h: 1.04, valign: "middle", fontFace: HF, fontSize: 19, bold: true, color: q[3] === AMBER ? "B5651D" : q[3], margin: 0 });
    s.addText(q[2], { x: M + 4.5, y: yy + 0.18, w: W - 2 * M - 4.8, h: 1.04, valign: "middle", fontFace: BF, fontSize: 15, color: INK, margin: 0, lineSpacingMultiple: 1.1 });
  });
}

// ============================================================ 9. ĐÓNG GÓP
{
  const s = light("Bốn đóng góp của đề tài.");
  const y = title(s, "Đóng góp của đề tài");
  const cs = [
    ["Định lượng rò rỉ", "Đo phơi nhiễm rò rỉ và chỉ ra khoảng trống khử trùng của notebook.", TEAL],
    ["Kiểm định phân hoá", "Chứng minh đèn kém hơn biển có ý nghĩa thống kê (Mann–Whitney).", RED],
    ["Bác bỏ giả thuyết", "Chứng minh độ hiếm KHÔNG gây ra thất bại (Spearman ≈ 0).", GREEN2],
    ["Sản phẩm hoàn chỉnh", "Pipeline tái lập, app demo, báo cáo Word/LaTeX và slide.", "1E5F8C"],
  ];
  cs.forEach((c, i) => { const x = M + (i % 2) * 6.2; const yy = y + Math.floor(i / 2) * 2.5;
    card(s, x, yy, 5.9, 2.2);
    badge(s, x + 0.34, yy + 0.36, i + 1, c[2], 0.6);
    s.addText(c[0], { x: x + 1.15, y: yy + 0.34, w: 4.5, h: 0.55, valign: "middle", fontFace: HF, fontSize: 18, bold: true, color: c[2], margin: 0 });
    s.addText(c[1], { x: x + 0.36, y: yy + 1.05, w: 5.2, h: 1.0, fontFace: BF, fontSize: 14.5, color: INK, margin: 0, valign: "top", lineSpacingMultiple: 1.15 });
  });
}

// ============================================================ 10. DIVIDER 2
divider("2", "Dữ liệu và phương pháp", "Bộ dữ liệu, công cụ đo YOLOv8, và quy trình phân tích 3 bước", "Phần 2: dữ liệu và cách phân tích.");

// ============================================================ 11. DỮ LIỆU KPI
{
  const s = light("Bộ dữ liệu 15 lớp, nhãn rất sạch.");
  const y = title(s, "Bộ dữ liệu pkdarabi/cardetection", "15 lớp · nhãn rất sạch (0 lỗi) · chia sẵn train/valid/test");
  const cols = [[vn(ds.total_images_scanned), "ảnh"], [vn(ds.total_valid_objects), "đối tượng"], ["15", "lớp"], [vn(imb.imbalance_ratio, 1), "imbalance ratio"]];
  cols.forEach((c, i) => { const x = M + i * 3.05; card(s, x, y, 2.85, 1.7); stat(s, x, y + 0.15, 2.85, c[0], c[1]); });
  finding(s, y + 2.1, `Chất lượng nhãn rất sạch: ${ds.missing_label_files} file thiếu · ${ds.invalid_objects} đối tượng lỗi · ${ds.malformed_label_lines} dòng sai định dạng. Vấn đề (nếu có) nằm ở phân bố và cách chia tập, không ở nhãn.`);
}

// ============================================================ 12. 15 LỚP
{
  const s = light("15 lớp chia hai nhóm: 2 đèn tín hiệu và 13 biển báo.");
  const y = title(s, "15 lớp — hai nhóm bản chất khác nhau");
  card(s, M, y, 5.9, 3.4, "FDEEEE");
  s.addText("Đèn tín hiệu (2 lớp)", { x: M + 0.3, y: y + 0.24, w: 5.2, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: RED, margin: 0 });
  bullets(s, ["Green Light", "Red Light — lớp PHỔ BIẾN NHẤT toàn corpus (787 đối tượng)",
    "Đặc điểm: thường nhỏ, ở xa, độ tương phản thấp", "Hai màu đỏ/xanh dễ nhầm về hình học"],
    { x: M + 0.34, y: y + 0.72, w: 5.25, h: 2.5, fontSize: 14, gap: 8 });
  card(s, M + 6.2, y, 5.9, 3.4, "E8F4F0");
  s.addText("Biển báo (13 lớp)", { x: M + 6.5, y: y + 0.24, w: 5.2, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: TEAL, margin: 0 });
  bullets(s, ["12 biển giới hạn tốc độ (10–120) + biển Stop",
    "Đặc điểm: chữ số/ký hiệu rõ, hình dạng chuẩn", "Kích thước thường lớn hơn đèn",
    "Rarest: Speed Limit 10 (chỉ 22 đối tượng)"],
    { x: M + 6.54, y: y + 0.72, w: 5.25, h: 2.5, fontSize: 14, gap: 8 });
}

// ============================================================ 13. SPLIT TABLE
{
  const s = light("Thống kê theo split. Độ dịch phân phối giữa các tập rất nhỏ.");
  const y = title(s, "Thống kê theo split");
  simpleTable(s, ["Split", "Số ảnh", "Số đối tượng", "Đối tượng/ảnh", "Diện tích bbox (px)"],
    B.split_summary.map((r) => [r.split, vn(r.n_images), vn(r.n_objects), vnp(r.objects_per_image_mean, 2), vn(r.bbox_area_px_median, 0)]),
    [2.2, 2.2, 2.6, 2.4, 2.7], { y, rowH: 0.5, fontSize: 14 });
  finding(s, y + 2.5, "Độ dịch phân phối lớp giữa các split rất nhỏ (Jensen–Shannon ≤ 0,008 bit) → chênh lệch hiệu năng KHÔNG đến từ việc các tập có phân phối khác nhau.");
}

// ============================================================ 14. MẤT CÂN BẰNG
{
  const s = light("Mất cân bằng mức trung bình.");
  const y = title(s, "Mất cân bằng lớp — mức trung bình");
  img(s, "05_imbalance.png", { x: M, y: y + 0.1, w: 6.0, h: 3.4 });
  const stt = [[vn(imb.imbalance_ratio, 1), "Imbalance ratio"], [vnp(imb.normalized_entropy, 3), "Entropy chuẩn hoá"], [vnp(imb.gini, 3), "Hệ số Gini"]];
  stt.forEach((v, i) => { const yy = y + 0.2 + i * 1.1; card(s, 7.0, yy, 5.6, 0.95, OFF);
    s.addText(v[0], { x: 7.1, y: yy, w: 1.9, h: 0.95, align: "center", valign: "middle", fontFace: HF, fontSize: 26, bold: true, color: TEAL, margin: 0 });
    s.addText(v[1], { x: 9.0, y: yy, w: 3.5, h: 0.95, valign: "middle", fontFace: BF, fontSize: 14, color: INK, margin: 0 }); });
  s.addText("Lớp phổ biến nhất (Red Light, 787) gấp ~36 lần lớp hiếm nhất (Speed Limit 10, 22).",
    { x: 7.0, y: y + 3.5, w: 5.6, h: 0.6, fontFace: BF, fontSize: 12.5, italics: true, color: MUTED, margin: 0 });
}

// ============================================================ 15. KÍCH THƯỚC
{
  const s = light("Nhiều vật thể nhỏ; tăng độ phân giải giúp giảm mạnh vật thể tí hon.");
  const y = title(s, "Kích thước vật thể — nhiều vật thể nhỏ");
  img(s, "06_small_objects.png", { x: M, y: y + 0.1, w: 6.6, h: 3.5 });
  bullets(s, [
    `${vnp(sig.small_ratio_area_lt_1pct * 100, 1)}% đối tượng là “nhỏ” (<1% diện tích ảnh).`,
    `${vnp(sig.tiny_ratio_area_lt_0_1pct * 100, 1)}% là “tí hon” (<0,1% diện tích ảnh).`,
    "Vật thể nhỏ khó phát hiện vì ít điểm ảnh mang thông tin.",
    "Tăng imgsz (640 → 1024) làm giảm mạnh tỉ lệ vật thể cạnh < 8 px.",
  ], { x: 7.4, y: y + 0.3, w: 5.3, h: 3.0, fontSize: 14, gap: 10 });
}

// ============================================================ 16. CÔNG CỤ ĐO
{
  const s = light("Công cụ đo là YOLOv8s huấn luyện hai giai đoạn.");
  const y = title(s, "Công cụ đo — YOLOv8s huấn luyện hướng-EDA");
  const specs = [
    ["Mô hình nền", `${tp.base_model} · ${vnp(tp.params_million, 2)}M tham số · ${vnp(tp.gflops, 1)} GFLOPs`],
    ["Giai đoạn 1", `imgsz=${tp.stage1_imgsz}, batch ${tp.stage1_batch}, ≤${tp.stage1_epochs} epoch, lr0=${vnp(tp.stage1_lr0, 3)}`],
    ["Giai đoạn 2", `imgsz=${tp.stage2_imgsz}, batch ${tp.stage2_batch}, ≤${tp.stage2_epochs} epoch, lr0=${vnp(tp.stage2_lr0, 4)}`],
    ["Tối ưu", `${tp.optimizer} + cosine LR · KHÔNG lật ngang/dọc (giữ ngữ nghĩa biển có hướng)`],
    ["Phần cứng", `GPU ${tp.device} · Ultralytics ${tp.ultralytics}`],
  ];
  specs.forEach((r, i) => { const yy = y + i * 0.72; card(s, M, yy, W - 2 * M, 0.62, i % 2 ? OFF : WHITE);
    s.addText(r[0], { x: M + 0.25, y: yy, w: 2.7, h: 0.62, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: TEAL, margin: 0 });
    s.addText(r[1], { x: M + 3.1, y: yy, w: W - 2 * M - 3.3, h: 0.62, valign: "middle", fontFace: BF, fontSize: 13.5, color: INK, margin: 0 }); });
  s.addText("Trọng số best.pt được đính kèm trong thư mục models/ của đề tài — cũng dùng cho tab demo của app.",
    { x: M, y: y + 3.8, w: W - 2 * M, h: 0.5, fontFace: BF, fontSize: 13, italics: true, color: MUTED, margin: 0 });
}

// ============================================================ 17. PHƯƠNG PHÁP 3 BƯỚC
{
  const s = light("Quy trình ba bước, tự động và tái lập.");
  const y = title(s, "Phương pháp phân tích thứ cấp — 3 bước");
  const steps = [
    ["Trích xuất", "Đọc lại toàn bộ con số notebook đã in, cố định thành bảng CSV chuẩn. Không nhập tay số nào.", TEAL],
    ["Kiểm định", "Phơi nhiễm rò rỉ; so đèn–biển bằng Mann–Whitney; tương quan độ hiếm–hiệu năng bằng Spearman.", "1E5F8C"],
    ["Trực quan hoá", "Sinh hình từ bảng đã trích; đưa vào app, báo cáo và slide.", RED],
  ];
  const ch = 4.1;
  steps.forEach((st, i) => { const x = M + i * 4.1; card(s, x, y, 3.85, ch);
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: 3.85, h: 0.14, rectRadius: 0.03, fill: { color: st[2] }, line: { width: 0 } });
    badge(s, x + 0.3, y + 0.42, i + 1, st[2], 0.62);
    s.addText(st[0], { x: x + 0.3, y: y + 1.32, w: 3.3, h: 0.5, fontFace: HF, fontSize: 19, bold: true, color: st[2], margin: 0 });
    s.addText(st[1], { x: x + 0.3, y: y + 1.94, w: 3.3, h: 1.9, fontFace: BF, fontSize: 15, color: INK, margin: 0, valign: "top", lineSpacingMultiple: 1.18 }); });
  finding(s, y + ch + 0.3, "Cố định seed = 42 · kiểm thử tự động pytest 6/6 đạt · số liệu gốc từ GPU Tesla T4 — mọi số liệu tái lập và truy vết được.", TEAL);
}

// ============================================================ 18. NGUYÊN TẮC
{
  const s = light("Nguyên tắc: mọi số liệu truy vết được, không nhập tay.");
  const y = title(s, "Nguyên tắc: minh bạch và tái lập");
  bulletPanel(s, y, [
    "Mọi bảng trong báo cáo ghi rõ tệp CSV nguồn; mọi CSV là đầu ra của một mô-đun src/.",
    "Không có con số nào nhập tay — chạy lại pipeline sẽ tạo lại y hệt.",
    "Đã đối chiếu 15 chỉ số chính khớp nhau giữa Word, LaTeX và slide.",
    "Kiểm thử tự động (pytest) kiểm tra: bảng đúng 15 lớp, mAP khớp notebook, khoảng cách đèn–biển có ý nghĩa, rò rỉ đúng số.",
  ], { yEnd: 5.55 });
  finding(s, 5.72, "Tinh thần khoa học: một con số chỉ đáng tin khi tái lập được và truy vết được về nguồn.");
}

// ============================================================ 19. DIVIDER A
divider("A", "Rò rỉ dữ liệu", "mAP 0,97 đáng tin đến đâu khi tập kiểm thử còn trùng lặp?", "Phần A: phơi nhiễm rò rỉ dữ liệu.");

// ============================================================ 20. RÒ RỈ LÀ GÌ
{
  const s = light("Rò rỉ là khi thông tin tập kiểm thử lọt vào huấn luyện.");
  const y = title(s, "A · Rò rỉ dữ liệu là gì?");
  bulletPanel(s, y, [
    "Rò rỉ dữ liệu (data leakage): thông tin của tập kiểm thử “lọt” vào quá trình huấn luyện.",
    "Với ảnh cắt từ video, dạng phổ biến là ẢNH TRÙNG hoặc GẦN TRÙNG xuất hiện ở nhiều split.",
    "Perceptual hash (dHash) phát hiện ảnh gần trùng qua khoảng cách Hamming của mã băm.",
    "Nếu không khử trùng ở tập đánh giá, mô hình có thể “ghi nhớ” ảnh đã thấy → mAP CAO HƠN hiệu năng thực.",
  ], { yEnd: 5.55, accent: AMBER });
  finding(s, 5.72, "Vì vậy bước đầu tiên của kiểm định là đo xem tập valid/test có bị trùng với train hay không.");
}

// ============================================================ 21. BẢNG + FIG LEAKAGE
{
  const s = light("Notebook tự phát hiện trùng lặp nhưng chỉ loại 91 ảnh khỏi train.");
  const y = title(s, "A · Đo phơi nhiễm rò rỉ");
  img(s, "04_leakage.png", { x: M, y: y + 0.05, w: 6.7, h: 3.5 });
  s.addText([{ text: "Bằng chứng\n", options: { bold: true, fontSize: 17, color: NAVY, breakLine: true } },
    ...B.leakage_rows.map((r) => ({ text: `• ${r.loai_trung_lap}: ${vn(r.so_dong)} dòng (${vnp(r.pct_tren_tong_the, 2)}%)\n`, options: { fontSize: 13.5, breakLine: true, color: INK } }))],
    { x: 7.5, y: y + 0.2, w: 5.2, h: 3.2, fontFace: BF, margin: 0, valign: "top", lineSpacingMultiple: 1.15 });
}

// ============================================================ 22. KHOẢNG TRỐNG
{
  const s = light("Notebook chỉ khử trùng ở train, bỏ quên valid/test.");
  const y = title(s, "A · Khoảng trống: chỉ khử trùng ở tập huấn luyện");
  const flow = [[vn(leak.original_train_images), "train gốc"], ["− " + vn(leak.removed_from_train), "loại exact-cross"], [vn(leak.clean_train_images), "train sạch"]];
  flow.forEach((f, i) => { const x = M + i * 3.05; card(s, x, y, 2.6, 1.5, i === 2 ? "E8F4F0" : OFF);
    s.addText(f[0], { x, y: y + 0.22, w: 2.6, h: 0.62, align: "center", valign: "middle", fontFace: HF, fontSize: 26, bold: true, color: i === 1 ? RED : INK, margin: 0 });
    s.addText(f[1], { x, y: y + 0.86, w: 2.6, h: 0.4, align: "center", valign: "top", fontFace: BF, fontSize: 12.5, color: MUTED, margin: 0 });
    if (i < 2) s.addText("→", { x: x + 2.62, y: y + 0.42, w: 0.42, h: 0.6, align: "center", valign: "middle", fontFace: BF, fontSize: 22, bold: true, color: TEAL, margin: 0 }); });
  card(s, M, y + 1.9, W - 2 * M, 1.6, "FDEEEE");
  s.addText("Vấn đề", { x: M + 0.3, y: y + 2.08, w: 5, h: 0.36, fontFace: HF, fontSize: 16, bold: true, color: RED, margin: 0 });
  s.addText(`Tập valid/test — vốn dùng để CHẤM ĐIỂM — không được khử trùng. Vẫn còn ${vn(leak.exact_cross_split_duplicate_rows)} ảnh trùng chính xác và ${vn(leak.dhash_cross_split_collision_rows)} ảnh gần trùng xuyên split trong dữ liệu đánh giá.`,
    { x: M + 0.3, y: y + 2.5, w: W - 2 * M - 0.6, h: 0.9, fontFace: BF, fontSize: 14.5, color: INK, margin: 0, valign: "top" });
}

// ============================================================ 23. NHẬN XÉT A
{
  const s = dark("Kết luận phần A.");
  s.addText("Nhận xét A — Rò rỉ dữ liệu", { x: M, y: 0.6, w: W - 2 * M, h: 0.8, fontFace: HF, fontSize: 30, bold: true, color: WHITE, margin: 0 });
  const pts = [
    [`Chỉ ${vn(leak.removed_from_train)} ảnh bị loại khỏi train`, "notebook không khử trùng ở tập đánh giá"],
    [`${vn(leak.exact_cross_split_duplicate_rows)} + ${vn(leak.dhash_cross_split_collision_rows)} trùng xuyên split`, "vẫn còn trong valid/test"],
    ["mAP 0,97 có thể LẠC QUAN", "một phần điểm có thể do mô hình đã thấy ảnh rất giống"],
  ];
  pts.forEach((p, i) => { const yy = 1.8 + i * 1.35; card(s, M, yy, W - 2 * M, 1.15, NAVY2);
    s.addText(p[0], { x: M + 0.35, y: yy + 0.12, w: 5.3, h: 0.9, valign: "middle", fontFace: HF, fontSize: 18, bold: true, color: AMBER, margin: 0 });
    s.addText(p[1], { x: M + 5.9, y: yy + 0.12, w: W - 2 * M - 6.2, h: 0.9, valign: "middle", fontFace: BF, fontSize: 15, color: "C6D4E2", margin: 0 }); });
  s.addText("Khuyến nghị: khử trùng ở CẢ valid/test và báo cáo thêm mAP sau khi loại near-duplicate.",
    { x: M, y: 6.0, w: W - 2 * M, h: 0.6, fontFace: BF, fontSize: 15, italics: true, color: "9DB2C6", margin: 0 });
}

// ============================================================ 24. DIVIDER B
divider("B", "Khoảng cách đèn – biển", "Hiệu năng có đồng đều giữa các lớp không?", "Phần B: hai lớp đèn kém hơn hẳn 13 lớp biển.");

// ============================================================ 25. BẢNG PER-CLASS
{
  const s = light("Bảng đầy đủ 15 lớp. Hai lớp đèn ở đáy bảng.");
  const y = title(s, "B · Hiệu năng theo lớp (test) — đầy đủ 15 lớp");
  const pc = B.per_class;
  const half = Math.ceil(pc.length / 2);
  const mk = (arr) => arr.map((r) => [r.class_name, vnp(r.ap50_95, 3), vnp(r.recall, 3)]);
  simpleTable(s, ["Lớp", "mAP@.5:.95", "Recall"], mk(pc.slice(0, half)), [3.0, 1.7, 1.4], { x: M, y, rowH: 0.34, fontSize: 12 });
  const rows2 = mk(pc.slice(half));
  const hl = pc.slice(half).map((r, i) => r.is_traffic_light ? i : -1).filter((i) => i >= 0);
  simpleTable(s, ["Lớp", "mAP@.5:.95", "Recall"], rows2, [3.0, 1.7, 1.4], { x: 6.9, y, rowH: 0.34, fontSize: 12, hlRows: hl });
  finding(s, y + 3.05, "13 lớp biển báo đều ≥ 0,78. Hai lớp đèn (đỏ, cột phải) rơi xuống 0,58 và 0,51 — tách hẳn khỏi phần còn lại.");
}

// ============================================================ 26. FIG PER-CLASS
{
  const s = light("Biểu đồ per-class: hai lớp đèn tách hẳn.");
  const y = title(s, "B · Hai lớp đèn tách hẳn khỏi 13 lớp biển");
  img(s, "01_per_class_ap.png", { x: M, y: y + 0.05, w: 8.0, h: 4.3 });
  s.addText([
    { text: "Đọc biểu đồ\n\n", options: { bold: true, fontSize: 17, color: NAVY, breakLine: true } },
    { text: "13 biển báo: 0,78–0,90\n\n", options: { fontSize: 14, color: TEAL, bold: true, breakLine: true } },
    { text: "2 lớp đèn (đỏ):\n", options: { fontSize: 14, color: RED, bold: true, breakLine: true } },
    { text: "Green Light 0,576\nRed Light 0,506\n\n", options: { fontSize: 14, color: RED, breakLine: true } },
    { text: "Đường đứt = trung bình mỗi nhóm.", options: { fontSize: 12.5, color: MUTED, italics: true } },
  ], { x: 8.9, y: y + 0.4, w: 3.8, h: 3.6, fontFace: BF, margin: 0, valign: "top", lineSpacingMultiple: 1.05 });
}

// ============================================================ 27. FIG LIGHT VS SIGN + STAT
{
  const s = light("Đèn kém hơn trên mọi chỉ số; kiểm định có ý nghĩa.");
  const y = title(s, "B · Đèn kém hơn biển trên MỌI chỉ số");
  card(s, M, y, 3.6, 1.5, "E8F4F0"); stat(s, M, y + 0.12, 3.6, vnp(lvs.signs.mean_map50_95, 3), "Biển báo (13 lớp)", TEAL);
  card(s, M + 3.75, y, 3.6, 1.5, "FDEEEE"); stat(s, M + 3.75, y + 0.12, 3.6, vnp(lvs.lights.mean_map50_95, 3), "Đèn tín hiệu (2 lớp)", RED);
  card(s, M + 7.5, y, 4.0, 1.5, NAVY);
  s.addText(vnp(lvs.gap, 3), { x: M + 7.5, y: y + 0.2, w: 4.0, h: 0.85, align: "center", valign: "middle", fontFace: HF, fontSize: 38, bold: true, color: AMBER, margin: 0 });
  s.addText("chênh lệch mAP@0,5:0,95", { x: M + 7.5, y: y + 1.02, w: 4.0, h: 0.4, align: "center", valign: "top", fontFace: BF, fontSize: 12, color: "9DB2C6", margin: 0 });
  img(s, "02_light_vs_sign.png", { x: 2.7, y: y + 1.75, w: 7.9, h: 3.05 });
}

// ============================================================ 28. MANN-WHITNEY
{
  const s = light("Kiểm định Mann–Whitney xác nhận khác biệt có ý nghĩa.");
  const y = title(s, "B · Kiểm định thống kê Mann–Whitney U");
  card(s, M, y, 5.9, 3.3, "E8F4F0");
  s.addText("Vì sao dùng Mann–Whitney?", { x: M + 0.3, y: y + 0.24, w: 5.3, h: 0.4, fontFace: HF, fontSize: 17, bold: true, color: TEAL, margin: 0 });
  bullets(s, [
    "Kiểm định phi tham số — không giả định phân phối chuẩn.",
    "So phân phối hiệu năng của hai nhóm độc lập: đèn (2 lớp) vs biển (13 lớp).",
    "Giả thuyết H₁ (một phía): biển > đèn.",
  ], { x: M + 0.34, y: y + 0.74, w: 5.25, h: 2.4, fontSize: 14, gap: 10 });
  card(s, M + 6.2, y, 5.9, 3.3, NAVY);
  s.addText("Kết quả", { x: M + 6.5, y: y + 0.24, w: 5.3, h: 0.4, fontFace: HF, fontSize: 17, bold: true, color: AMBER, margin: 0 });
  s.addText("p = " + vnp(lvs.mannwhitney_p, 4), { x: M + 6.5, y: y + 0.8, w: 5.3, h: 0.9, fontFace: HF, fontSize: 40, bold: true, color: WHITE, margin: 0 });
  s.addText("Khác biệt có ý nghĩa thống kê (p < 0,01) dù nhóm đèn chỉ có 2 lớp. Khoảng cách 0,313 không phải ngẫu nhiên.",
    { x: M + 6.5, y: y + 1.9, w: 5.3, h: 1.2, fontFace: BF, fontSize: 14.5, color: "C6D4E2", margin: 0, valign: "top" });
}

// ============================================================ 29. NHẬN XÉT B
{
  const s = dark("Kết luận phần B.");
  s.addText("Nhận xét B — Khoảng cách đèn – biển", { x: M, y: 0.6, w: W - 2 * M, h: 0.8, fontFace: HF, fontSize: 28, bold: true, color: WHITE, margin: 0 });
  const pts = [
    ["Đèn 0,541 vs biển 0,854", "chênh 0,313 trên mAP@0,5:0,95"],
    ["Mann–Whitney p = 0,0095", "khác biệt có ý nghĩa thống kê"],
    ["Red Light kém nhất (0,506)", "recall chỉ 0,755 — bỏ sót nhiều"],
    ["mAP tổng bị chi phối bởi biển dễ", "con số 0,97 che giấu phân hoá này"],
  ];
  pts.forEach((p, i) => { const yy = 1.7 + i * 1.15; card(s, M, yy, W - 2 * M, 0.98, NAVY2);
    s.addText(p[0], { x: M + 0.35, y: yy + 0.05, w: 5.3, h: 0.88, valign: "middle", fontFace: HF, fontSize: 17, bold: true, color: RED, margin: 0 });
    s.addText(p[1], { x: M + 5.9, y: yy + 0.05, w: W - 2 * M - 6.2, h: 0.88, valign: "middle", fontFace: BF, fontSize: 14.5, color: "C6D4E2", margin: 0 }); });
}

// ============================================================ 30. DIVIDER C
divider("C", "Hiếm ≠ Khó", "Đèn kém vì hiếm, hay vì bản chất vật thể?", "Phần C: bác bỏ giả thuyết mất cân bằng.");

// ============================================================ 31. GIẢ THUYẾT
{
  const s = light("Giả thuyết tự nhiên: đèn kém vì hiếm. Ta kiểm chứng.");
  const y = title(s, "C · Giả thuyết: “đèn kém vì hiếm”?");
  card(s, M, y, W - 2 * M, 1.5, OFF);
  s.addText("Giả thuyết", { x: M + 0.3, y: y + 0.2, w: 4, h: 0.36, fontFace: HF, fontSize: 16, bold: true, color: MUTED, margin: 0 });
  s.addText("Lớp càng ít mẫu thì mô hình càng khó học → đèn kém vì thuộc nhóm hiếm. Nếu đúng, số instance và mAP phải TƯƠNG QUAN DƯƠNG.",
    { x: M + 0.3, y: y + 0.62, w: W - 2 * M - 0.6, h: 0.8, fontFace: BF, fontSize: 15, color: INK, margin: 0, valign: "top" });
  s.addText("Cách kiểm chứng: tính tương quan hạng Spearman giữa số instance mỗi lớp và mAP@0,5:0,95 của lớp đó.",
    { x: M, y: y + 1.8, w: W - 2 * M, h: 0.5, fontFace: BF, fontSize: 15, bold: true, color: TEAL, margin: 0 });
  s.addText("Nếu ρ ≈ 0 và p lớn → độ hiếm KHÔNG liên quan tới hiệu năng → giả thuyết bị bác bỏ.",
    { x: M, y: y + 2.4, w: W - 2 * M, h: 0.5, fontFace: BF, fontSize: 15, color: MUTED, italics: true, margin: 0 });
}

// ============================================================ 32. FIG RARITY
{
  const s = light("Kết quả: tương quan gần 0. Giả thuyết bị bác bỏ.");
  const y = title(s, "C · Độ hiếm KHÔNG giải thích được hiệu năng");
  img(s, "03_rarity_vs_ap.png", { x: M, y: y + 0.05, w: 7.4, h: 4.35 });
  s.addText([
    { text: "Spearman ρ = " + vnp(rv.spearman_instances_vs_map5095, 3) + "\n", options: { fontSize: 22, bold: true, color: RED, breakLine: true } },
    { text: "p = " + vnp(rv.p_instances_vs_map5095, 2) + "\n\n", options: { fontSize: 15, color: MUTED, breakLine: true } },
    { text: "→ Không có tương quan giữa số instance và hiệu năng.\n\n", options: { fontSize: 14, color: INK, breakLine: true } },
    { text: "Giả thuyết “đèn kém vì hiếm” BỊ BÁC BỎ.", options: { fontSize: 14.5, color: TEAL, bold: true } },
  ], { x: 8.2, y: y + 0.4, w: 4.5, h: 3.8, fontFace: BF, margin: 0, valign: "top", lineSpacingMultiple: 1.08 });
}

// ============================================================ 33. NGHỊCH LÝ RED LIGHT
{
  const s = light("Red Light là lớp phổ biến nhất nhưng kém nhất — nghịch lý bác bỏ giả thuyết.");
  const y = title(s, "C · Nghịch lý Red Light");
  card(s, M, y, 5.9, 2.4, "E8F4F0");
  s.addText("787", { x: M + 0.3, y: y + 0.3, w: 5.3, h: 1.0, fontFace: HF, fontSize: 46, bold: true, color: TEAL, margin: 0 });
  s.addText("đối tượng Red Light toàn corpus — lớp PHỔ BIẾN NHẤT trong 15 lớp.", { x: M + 0.3, y: y + 1.35, w: 5.3, h: 0.9, fontFace: BF, fontSize: 14.5, color: INK, margin: 0, valign: "top" });
  card(s, M + 6.2, y, 5.9, 2.4, "FDEEEE");
  s.addText("0,506", { x: M + 6.5, y: y + 0.3, w: 5.3, h: 1.0, fontFace: HF, fontSize: 46, bold: true, color: RED, margin: 0 });
  s.addText("mAP@0,5:0,95 của Red Light — lớp KÉM NHẤT trong 15 lớp.", { x: M + 6.5, y: y + 1.35, w: 5.3, h: 0.9, fontFace: BF, fontSize: 14.5, color: INK, margin: 0, valign: "top" });
  finding(s, y + 2.7, "Phổ biến nhất NHƯNG kém nhất — bằng chứng trực tiếp cho thấy độ hiếm không phải nguyên nhân. Vậy nguyên nhân là gì?", RED);
}

// ============================================================ 34. VÌ SAO ĐÈN KHÓ
{
  const s = light("Nguyên nhân nằm ở bản chất vật thể, không phải mất cân bằng.");
  const y = title(s, "C · Vì sao đèn tín hiệu khó? — bản chất vật thể");
  const rs = [
    ["Kích thước nhỏ", "Đèn thường ở xa, chiếm ít điểm ảnh — thuộc nhóm vật thể nhỏ khó phát hiện."],
    ["Độ tương phản thấp", "Đèn trên nền trời/đường phức tạp, ánh sáng thay đổi."],
    ["Đỏ ↔ xanh dễ nhầm", "Hai lớp cùng hình dạng, chỉ khác màu — dễ nhầm về hình học."],
    ["Recall thấp", "Red Light recall chỉ 0,755 — mô hình BỎ SÓT nhiều đèn thật."],
  ];
  rs.forEach((r, i) => { const yy = y + i * 0.92; card(s, M, yy, W - 2 * M, 0.8, i % 2 ? OFF : WHITE);
    s.addText(r[0], { x: M + 0.25, y: yy, w: 3.2, h: 0.8, valign: "middle", fontFace: HF, fontSize: 15, bold: true, color: RED, margin: 0 });
    s.addText(r[1], { x: M + 3.6, y: yy, w: W - 2 * M - 3.8, h: 0.8, valign: "middle", fontFace: BF, fontSize: 13.5, color: INK, margin: 0 }); });
  s.addText("→ Đây là vấn đề của LOẠI VẬT THỂ, cần giải pháp riêng cho đèn — không phải cân bằng lại dữ liệu.",
    { x: M, y: y + 3.8, w: W - 2 * M, h: 0.5, fontFace: BF, fontSize: 14.5, bold: true, color: TEAL, margin: 0 });
}

// ============================================================ 35. NHẬN XÉT C
{
  const s = dark("Kết luận phần C.");
  s.addText("Nhận xét C — Hiếm ≠ Khó", { x: M, y: 0.6, w: W - 2 * M, h: 0.8, fontFace: HF, fontSize: 30, bold: true, color: WHITE, margin: 0 });
  const pts = [
    ["Spearman ρ = 0,032 (p = 0,91)", "độ hiếm và hiệu năng KHÔNG tương quan"],
    ["Red Light phổ biến nhất → kém nhất", "nghịch lý bác bỏ giả thuyết mất cân bằng"],
    ["Nguyên nhân: bản chất vật thể", "đèn nhỏ, tương phản thấp, đỏ/xanh dễ nhầm"],
    ["Không oversampling để sửa", "vì độ hiếm không phải nguyên nhân"],
  ];
  pts.forEach((p, i) => { const yy = 1.7 + i * 1.15; card(s, M, yy, W - 2 * M, 0.98, NAVY2);
    s.addText(p[0], { x: M + 0.35, y: yy + 0.05, w: 5.6, h: 0.88, valign: "middle", fontFace: HF, fontSize: 16, bold: true, color: GREEN2 === "2C7A4B" ? "7FC99B" : GREEN2, margin: 0 });
    s.addText(p[1], { x: M + 6.1, y: yy + 0.05, w: W - 2 * M - 6.4, h: 0.88, valign: "middle", fontFace: BF, fontSize: 14.5, color: "C6D4E2", margin: 0 }); });
}

// ============================================================ 36. DIVIDER DEMO
divider("D", "Demo phát hiện trực tiếp", "Upload ảnh → YOLOv8 vẽ hộp → gắn độ tin cậy đã kiểm định", "Phần demo: ứng dụng web.");

// ============================================================ 37. DEMO KIẾN TRÚC
{
  const s = light("App gồm nhiều tab; tab demo nối mô hình với kết quả kiểm định.");
  const y = title(s, "Ứng dụng web — kiến trúc demo");
  bulletPanel(s, y, [
    "App Streamlit gồm các tab: Tổng quan · Rò rỉ (A) · Đèn vs Biển (B) · Hiếm ≠ Khó (C) · Demo phát hiện · Bảng đầy đủ.",
    "Tab Demo: tải ảnh (hoặc chọn ảnh mẫu) → YOLOv8 best.pt vẽ hộp phát hiện.",
    "Điểm mới: mỗi phát hiện được GẮN với độ tin cậy đã kiểm định của lớp (mAP@0,5:0,95).",
    "Nếu phát hiện thuộc nhóm đèn (mAP thấp) → app CẢNH BÁO “kết quả cần thận trọng”.",
  ], { yEnd: 5.55 });
  finding(s, 5.72, "Demo biến nội dung kiểm định thành công cụ dùng được: không chỉ phát hiện, mà còn nói cho người dùng biết KẾT QUẢ ĐÓ ĐÁNG TIN đến đâu.");
}

// ============================================================ 38. DEMO BIỂN BÁO
{
  const s = light("Demo trên ảnh có biển báo: độ tin cậy cao.");
  const y = title(s, "Demo · Biển báo — độ tin cậy CAO", "YOLOv8 phát hiện Speed Limit; lớp thuộc nhóm biển (mAP ≈ 0,85)");
  img(s, "demo_sign.png", { x: M, y: y + 0.05, w: 4.3, h: 4.3 });
  card(s, 5.3, y + 0.3, 7.4, 3.0, "E8F4F0");
  s.addText("Nhóm biển báo — ĐỘ TIN CẬY CAO", { x: 5.6, y: y + 0.5, w: 6.8, h: 0.5, fontFace: HF, fontSize: 18, bold: true, color: GREEN2, margin: 0 });
  bullets(s, [
    "Mô hình phát hiện biển Speed Limit với độ tin cậy YOLO cao.",
    "Lớp này thuộc nhóm biển báo — mAP@0,5:0,95 đã kiểm định ≈ 0,85.",
    "App hiển thị: “Toàn bộ phát hiện thuộc nhóm biển báo — độ tin cậy cao”.",
  ], { x: 5.64, y: y + 1.1, w: 6.9, h: 2.0, fontSize: 14.5, gap: 9 });
}

// ============================================================ 39. DEMO ĐÈN
{
  const s = light("Demo trên ảnh có đèn đỏ ở xa: app gắn cờ độ tin cậy thấp.");
  const y = title(s, "Demo · Đèn đỏ — bị GẮN CỜ độ tin cậy thấp", "Đèn nhỏ, ở xa — đúng ca khó mà kiểm định đã chỉ ra");
  img(s, "demo_light.png", { x: M, y: y + 0.05, w: 4.3, h: 4.3 });
  card(s, 5.3, y + 0.3, 7.4, 3.0, "FDEEEE");
  s.addText("Nhóm đèn tín hiệu — CẢNH BÁO", { x: 5.6, y: y + 0.5, w: 6.8, h: 0.5, fontFace: HF, fontSize: 18, bold: true, color: RED, margin: 0 });
  bullets(s, [
    "Mô hình phát hiện nhiều Red Light nhỏ ở xa, độ tin cậy YOLO thấp (0,29–0,61).",
    "Lớp Red Light có mAP@0,5:0,95 kiểm định chỉ 0,506 — nhóm kém nhất.",
    "App cảnh báo: “Phát hiện đối tượng thuộc nhóm đèn tín hiệu — kết quả cần thận trọng”.",
  ], { x: 5.64, y: y + 1.1, w: 6.9, h: 2.0, fontSize: 14.5, gap: 9 });
}

// ============================================================ 40. DEMO CÁCH CHẠY
{
  const s = light("Cách chạy app.");
  const y = title(s, "Demo · Cách chạy ứng dụng");
  card(s, M, y, W - 2 * M, 1.5, NAVY);
  s.addText("streamlit run app/app.py", { x: M + 0.35, y: y + 0.2, w: W - 2 * M - 0.7, h: 0.5, fontFace: "Consolas", fontSize: 20, bold: true, color: "7FE0C0", margin: 0 });
  s.addText("Cần: pip install -r requirements.txt (đã gồm streamlit, ultralytics, opencv). Trọng số đặt ở models/best.pt.",
    { x: M + 0.35, y: y + 0.8, w: W - 2 * M - 0.7, h: 0.6, fontFace: BF, fontSize: 13.5, color: "C6D4E2", margin: 0, valign: "top" });
  bullets(s, [
    "Chạy hoàn toàn trên CPU — không cần GPU để demo.",
    "Nếu thiếu ultralytics/best.pt, app tự hiện hai ảnh demo kết xuất sẵn (không lỗi).",
    "Điều chỉnh ngưỡng tin cậy (conf) bằng thanh trượt để xem thay đổi.",
  ], { x: M + 0.1, y: y + 1.75, w: W - 2 * M - 0.2, h: 2.0, fontSize: 15, gap: 10 });
}

// ============================================================ 41. KẾT LUẬN
{
  const s = dark("Ba phát hiện, cùng một kết luận về con số mAP.");
  s.addText("Kết luận", { x: M, y: 0.55, w: W - 2 * M, h: 0.8, fontFace: HF, fontSize: 32, bold: true, color: WHITE, margin: 0 });
  const chain = [
    ["A · Rò rỉ", `${vn(leak.exact_cross_split_duplicate_rows)} + ${vn(leak.dhash_cross_split_collision_rows)} trùng xuyên split chưa khử → mAP có thể lạc quan`, AMBER],
    ["B · Phân hoá", `đèn ${vnp(lvs.lights.mean_map50_95, 3)} vs biển ${vnp(lvs.signs.mean_map50_95, 3)} (p=${vnp(lvs.mannwhitney_p, 4)})`, RED],
    ["C · Nguyên nhân", `độ hiếm không liên quan (ρ=${vnp(rv.spearman_instances_vs_map5095, 3)}); lỗi ở loại vật thể (đèn)`, GREEN2],
  ];
  chain.forEach((c, i) => { const yy = 1.6 + i * 1.35;
    s.addShape(pptx.ShapeType.roundRect, { x: M, y: yy, w: W - 2 * M, h: 1.15, rectRadius: 0.08, fill: { color: NAVY2 }, line: { color: c[2], width: 2 } });
    s.addText(c[0], { x: M + 0.3, y: yy + 0.1, w: 3.0, h: 0.95, valign: "middle", fontFace: HF, fontSize: 20, bold: true, color: c[2], margin: 0 });
    s.addText(c[1], { x: M + 3.4, y: yy + 0.1, w: W - 2 * M - 3.7, h: 0.95, valign: "middle", fontFace: BF, fontSize: 15, color: WHITE, margin: 0 }); });
  s.addText("Con số 0,97 vừa có thể được rò rỉ nâng đỡ, vừa che giấu phân hoá lớp — mà phân hoá đó do LOẠI VẬT THỂ, không phải mất cân bằng.",
    { x: M, y: 5.75, w: W - 2 * M, h: 0.8, fontFace: BF, fontSize: 15, italics: true, color: "C6D4E2", margin: 0 });
}

// ============================================================ 42. KHUYẾN NGHỊ
{
  const s = light("Bốn khuyến nghị hành động.");
  const y = title(s, "Khuyến nghị");
  const recs = [
    ["Khử trùng ở valid/test", "Loại near-duplicate xuyên split trước khi báo mAP; báo thêm mAP sau khi khử."],
    ["Báo cáo AP theo lớp/nhóm", "Không chỉ mAP tổng — để không che giấu điểm yếu đèn."],
    ["Cải thiện lớp đèn", "Tăng imgsz 896/1024, thêm mẫu đèn khó, phân tích nhầm lẫn đỏ↔xanh."],
    ["Không oversampling lớp hiếm", "Vì độ hiếm không phải nguyên nhân của thất bại đèn."],
  ];
  const rh = 1.12, rgap = (CY1 - y - 4 * rh) / 3;
  recs.forEach((r, i) => { const yy = y + i * (rh + rgap); card(s, M, yy, W - 2 * M, rh);
    s.addShape(pptx.ShapeType.roundRect, { x: M, y: yy, w: 0.15, h: rh, rectRadius: 0.05, fill: { color: TEAL }, line: { width: 0 } });
    badge(s, M + 0.36, yy + (rh - 0.58) / 2, i + 1, TEAL, 0.58);
    s.addText(r[0], { x: M + 1.2, y: yy, w: 3.5, h: rh, valign: "middle", fontFace: HF, fontSize: 16.5, bold: true, color: TEAL, margin: 0 });
    s.addText(r[1], { x: M + 4.9, y: yy, w: W - 2 * M - 5.1, h: rh, valign: "middle", fontFace: BF, fontSize: 14.5, color: INK, margin: 0, lineSpacingMultiple: 1.1 }); });
}

// ============================================================ 43. HẠN CHẾ + SẢN PHẨM
{
  const s = light("Hạn chế và sản phẩm kèm theo.");
  const y = title(s, "Hạn chế và sản phẩm kèm theo");
  const ch = CY1 - y;
  card(s, M, y, 5.9, ch, "FDF3E7");
  s.addShape(pptx.ShapeType.roundRect, { x: M, y, w: 5.9, h: 0.14, rectRadius: 0.03, fill: { color: "B5651D" }, line: { width: 0 } });
  s.addText("Hạn chế", { x: M + 0.34, y: y + 0.36, w: 5.2, h: 0.4, fontFace: HF, fontSize: 19, bold: true, color: "B5651D", margin: 0 });
  bullets(s, [
    "Chế độ results-only: chưa đo trực tiếp mức sụt mAP sau khi khử rò rỉ (mới định lượng phơi nhiễm).",
    "Nhóm đèn chỉ 2 lớp → kiểm định có ý nghĩa nhưng cần thận trọng khi khái quát.",
    "Kết quả gắn với bộ dữ liệu và mô hình cụ thể này.",
  ], { x: M + 0.4, y: y + 1.0, w: 5.2, h: ch - 1.2, fontSize: 15, gap: 12, valign: "middle" });
  card(s, M + 6.2, y, 5.9, ch, NAVY);
  s.addShape(pptx.ShapeType.roundRect, { x: M + 6.2, y, w: 5.9, h: 0.14, rectRadius: 0.03, fill: { color: AMBER }, line: { width: 0 } });
  s.addText("Sản phẩm kèm theo", { x: M + 6.54, y: y + 0.36, w: 5.2, h: 0.4, fontFace: HF, fontSize: 19, bold: true, color: AMBER, margin: 0 });
  bullets(s, [
    "Pipeline phân tích tái lập (src/) + pytest 6/6 đạt.",
    "Ứng dụng web Streamlit có tab DEMO phát hiện trực tiếp.",
    "Báo cáo Word (.docx) và LaTeX (mẫu trường).",
    "Bộ slide trình bày này.",
  ], { x: M + 6.6, y: y + 1.0, w: 5.2, h: ch - 1.2, fontSize: 15, color: "E8EEF4", gap: 12, valign: "middle" });
}

// ============================================================ 44. CẢM ƠN
{
  const s = dark("Cảm ơn thầy/cô đã lắng nghe. Nhóm em sẵn sàng nhận câu hỏi.");
  s.addShape(pptx.ShapeType.ellipse, { x: 10.2, y: -1.2, w: 4.8, h: 4.8, fill: { color: NAVY2 }, line: { width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 11.4, y: 4.8, w: 2.8, h: 2.8, fill: { color: TEAL }, line: { width: 0 }, transparency: 74 });
  s.addText("Xin cảm ơn thầy/cô đã lắng nghe", { x: M, y: 2.4, w: 9.6, h: 0.9, fontFace: HF, fontSize: 34, bold: true, color: WHITE, margin: 0 });
  s.addText("Nhóm em sẵn sàng nhận câu hỏi và góp ý", { x: M, y: 3.34, w: 9.6, h: 0.46, fontFace: BF, fontSize: 16, color: "9DB2C6", margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: M, y: 4.3, w: 10.0, h: 1.5, rectRadius: 0.1, fill: { color: NAVY2 }, line: { width: 0 } });
  s.addText("Huỳnh Phát Lợi · Đoàn Huỳnh Thanh Tú · Võ Phú Vinh", { x: M + 0.35, y: 4.52, w: 9.3, h: 0.4, fontFace: BF, fontSize: 15, bold: true, color: WHITE, margin: 0 });
  s.addText("KHMT836016 · KHMT836034 · KHMT836036\nNguồn: notebook Kaggle traffic-yolo-v2-run · YOLOv8s · Ultralytics 8.4.95", { x: M + 0.35, y: 4.96, w: 9.3, h: 0.7, fontFace: BF, fontSize: 12.5, color: "9DB2C6", margin: 0, valign: "top" });
}

pptx.writeFile({ fileName: OUT }).then(() => console.log("Đã tạo:", OUT, "· số slide:", SN));
