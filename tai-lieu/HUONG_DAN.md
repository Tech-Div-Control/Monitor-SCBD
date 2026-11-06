# Hệ thống quản lý yêu cầu SC / BD / HC / CT (v2)
Trung tâm Bảo dưỡng Nội trường Hà Nội – Google Apps Script Web App + Google Sheets

## 1. Các quyết định nghiệp vụ đã chốt
| Nội dung | Quy định trong hệ thống |
|---|---|
| Mã yêu cầu | Duy nhất, không đổi: `YC-năm-00001` (việc mới); `L-VP0001` / `L-PV0001` (dữ liệu cũ). |
| Dữ liệu cũ | 464 dòng nhập từ 2 file đã làm sạch, chỉ để theo dõi (`legacy=1`, mặc định chỉ đọc, không gửi email nhắc, không tính vào số liệu "tiếp nhận mới"). Việc mới nhập thẳng trên app. Admin có thể bật `LEGACY_EDITABLE=1` ở tab Quản trị nếu cần đóng/sửa việc cũ. |
| "Số Barvo" | Là **mã hồ sơ trên Bravo** (cột `ma_bravo`), không phải mã tài sản. App không quản lý mã tài sản. |
| Quy trình | 4 bước chỉ để điền thông tin quản lý, làm song song, không chặn việc đóng: Xác nhận thông tin → Cập nhật Bravo (ghi mã hồ sơ) → In tài liệu và trình ký → Cập nhật Vaeco (chỉ luồng Trung tâm Phục vụ Bảo dưỡng). Đóng khi thiếu bước: hệ thống cảnh báo, người đóng xác nhận và việc đó được ghi nhật ký. Không có bước mã tài sản, lưu hồ sơ bản cứng, cập nhật file theo dõi 2026. |
| Trạng thái | Open / Inprogress / Pending / Closed. **Pending bắt buộc lý do + ngày xin gia hạn đến** → tự tạo đề nghị gia hạn "Chờ duyệt". |
| Hạn | `deadline_goc` không đổi; `deadline` (hạn hiện tại) chỉ đổi khi gia hạn được duyệt. Quá hạn mà chưa xong phải xin gia hạn kèm lý do (nút "Xin gia hạn"). Pending vẫn tính quá hạn nếu chưa được duyệt gia hạn. |
| Duyệt gia hạn | Cấp duyệt gán theo người dùng: **Tổ** hoặc **Phòng**. Mặc định CB tổ duyệt. Nếu đặt `EXT_PHONG_NGAY = N` (>0), gia hạn dài hơn N ngày so với hạn hiện tại phải do CB phòng duyệt (CB tổ không duyệt được). Không ai được tự duyệt đề nghị do mình lập. Từ chối phải ghi lý do. |
| Vai trò | **Khách** (ai có link, không đăng nhập): chỉ xem, email người thao tác bị che, không xem nhật ký thay đổi, không xuất báo cáo. **Admin** (bạn): quản lý người dùng/cấu hình, duyệt gia hạn (nếu có cấp duyệt), có quyền như Editor nhưng việc đóng nên do Editor làm (Admin đóng sẽ có dòng ghi riêng trong nhật ký). **Editor**: nhập liệu, sửa, đóng việc, trao đổi, đính kèm. Admin/Editor đăng nhập bằng mã 6 số gửi qua email. Quyền được kiểm tra ở phía server. |
| Nhập liệu | Editor nhận thông tin qua Viber/điện thoại và nhập lên app (form không còn ô "Nguồn tiếp nhận"). |
| Đơn vị gửi | Đội 1–Đội 6, Đội SCCT, Đội PVBD, Phòng KT. |
| Đơn vị tiếp nhận | Văn phòng; Trung tâm Phục vụ Bảo dưỡng. **Văn phòng không nhận Hiệu chuẩn và Chế tạo** (bị chặn). |
| Biểu mẫu tự hiện | Bảo dưỡng: VP → 9101, PV → 7001. Sửa chữa: VP → 9101 + 9110, PV → 7001. Hiệu chuẩn: PV → 7011. Chế tạo: PV → 7001. Sửa trong sheet `FORM_RULES`; đặt tên biểu mẫu trong `FORMS`. |
| Thiết bị gộp | Bảo dưỡng có thể gộp: ghi danh sách thiết bị và mã tài sản (nếu có) ngay trong ô "Nội dung yêu cầu", không có ô riêng. Sửa chữa: 1 yêu cầu = 1 mục. |
| Chuẩn hóa đơn vị gửi (dữ liệu cũ) | Lấy cột "Đvi yc" làm chuẩn (PKT→Phòng KT, SCCT→Đội SCCT, PVBD→Đội PVBD). 34 dòng "Doi BD": lấy Đội 1–6 từ cột "Phân loại đơn vị". Dòng hai cột mâu thuẫn hoặc "Doi BD" không có Đội cụ thể (19 dòng), dòng hạn trước ngày làm yêu cầu và mã Bravo trùng: gắn cờ **Cần rà soát** (tổng 26 dòng), giữ nguyên thông tin gốc trong `pl_cu`, bấm "Đã rà soát" sau khi kiểm tra. |
| Loại yêu cầu của dữ liệu cũ | Để trống (file cũ không có phân loại SC/BD/HC/CT); phân loại gốc lưu ở `pl_cu`. |

## 2. Cài đặt (làm một lần, khoảng 15 phút)
1. Tạo Google Sheet mới bằng Gmail của bạn → **Tệp > Nhập > Tải lên** `DuLieu_Khoi_Tao.xlsx` → **Thay thế bảng tính**.
2. **Tiện ích mở rộng > Apps Script**. Dán `Code.gs` vào tệp `Code.gs`; tạo tệp HTML tên `Index` và dán `Index.html`; bật **Cài đặt dự án > Hiển thị tệp kê khai appsscript.json** rồi dán `appsscript.json`. Nhấn Ctrl+S.
3. Chọn hàm `setup` → **Chạy**, cấp quyền (chỉ mình bạn cần cấp). Hàm này ghi ID Sheet, tạo thư mục Drive đính kèm, tạo sheet còn thiếu, thêm **Gmail của bạn làm Admin đầu tiên**, đặt lịch nhắc việc 7h hằng ngày.
4. **Triển khai > Quản lý bản triển khai > bút chì** (hoặc *Tùy chọn triển khai mới > Ứng dụng web*): **Thực thi với tư cách: Tôi (Me)**; **Ai có quyền truy cập: Bất kỳ ai (Anyone)**; Phiên bản mới → Triển khai. Nếu bản triển khai cũ không cho sửa hai mục này, hãy tạo bản triển khai mới và dùng URL mới.
5. Gửi URL (kết thúc bằng `/exec`) cho mọi người: ai có link đều xem được ngay, không cần Google đăng nhập hay cấp quyền.
6. Bạn bấm **Đăng nhập (Admin/Editor)** ở góc trên, nhập Gmail của bạn, nhập mã 6 số nhận được qua email. Vào tab **Quản trị** để thêm Editor/Admin theo email. Mỗi người tự đăng nhập bằng cùng cách.

**Không chia sẻ Google Sheet và thư mục Drive cho nhân viên.** App chạy bằng quyền của bạn nên họ không cần quyền gì, và nhờ vậy không ai sửa được dữ liệu ngoài app. Mã đăng nhập gửi từ Gmail của bạn (giới hạn khoảng 100 email/ngày); phiên đăng nhập giữ 14 ngày (`SESSION_DAYS`). Khóa một người ở tab Quản trị là mất quyền ngay. Nếu muốn bắt buộc đăng nhập mới được xem, đặt `KHACH_XEM=0` trong tab Quản trị.

Lưu ý bảo mật: ai có link đều xem được toàn bộ danh sách yêu cầu (đã che email người thao tác). Đừng đăng link nơi công khai; nếu bị lộ link, tạo bản triển khai mới (URL mới) và xóa bản cũ.

## 2b. Đưa giao diện lên website riêng (không có thanh cảnh báo của Google, địa chỉ gọn)
Địa chỉ `script.google.com/...` luôn có thanh "Ứng dụng này do một người dùng Google Apps Script tạo". Để có một trang web đúng nghĩa, đưa thư mục `web/` (gồm `index.html` và `config.js`) lên hosting tĩnh miễn phí. Apps Script lúc này chỉ làm phần dữ liệu phía sau (API); mọi phần phân quyền ở mục 1 giữ nguyên.
1. Làm xong mục 2 (đã triển khai Apps Script, Thực thi: Tôi, Truy cập: Bất kỳ ai) và sao chép **URL ứng dụng web** (kết thúc bằng `/exec`).
2. Mở `web/config.js` bằng Notepad, dán URL đó vào giữa hai dấu nháy: `window.APP_API_URL = "https://script.google.com/macros/s/.../exec";` rồi lưu.
3. Đưa thư mục `web/` lên hosting (chọn một cách):
   - **Netlify** (dễ nhất): tạo tài khoản miễn phí tại netlify.com → *Add new site → Deploy manually* → kéo thả thư mục `web/` vào. Nhận ngay địa chỉ dạng `https://ten-ban-chon.netlify.app` (đổi tên trong Site settings → Change site name).
   - **Cloudflare Pages**: *Workers & Pages → Create → Pages → Upload assets* → kéo thả thư mục `web/`.
   - **GitHub Pages**: tạo repository, tải `index.html` và `config.js` lên, bật Pages ở Settings.
4. Gửi địa chỉ website cho mọi người: bấm vào là xem như đọc báo. Admin/Editor bấm **Đăng nhập** ở góc trên (nhập email, nhập mã 6 số).
- Muốn tên miền riêng (ví dụ `yeucau.tendonvi.vn`), thêm trong phần Domain của Netlify/Cloudflare (cần mua tên miền).
- Khi tôi gửi bản `Code.gs` mới: dán vào Apps Script rồi **Triển khai > Quản lý bản triển khai > Phiên bản mới**. Khi có bản `Index.html` mới: chạy lại bước tạo `web/index.html` (tôi sẽ gửi sẵn) và kéo thả lại lên hosting. URL không đổi.
- Trang đầu tiên mở có thể chậm 1-3 giây vì dữ liệu lấy từ Google; đó là bình thường.

## 3. Chức năng
- Tổng quan: các ô số liệu, biểu đồ và danh sách yêu cầu chi tiết ở bên dưới (lọc, phân trang); số đang mở, quá hạn, sắp đến hạn (mặc định ≤3 ngày), Pending kéo dài (>7 ngày), chưa cập nhật (>7 ngày), gia hạn chờ duyệt, đóng trong tháng, cần rà soát; biểu đồ theo trạng thái/đơn vị/nhóm. Bấm ô để lọc. Có công tắc gồm/không gồm dữ liệu cũ.
- Danh sách: tìm kiếm, lọc nhiều tiêu chí, sắp xếp, phân trang, xuất CSV (mở được bằng Excel).
- Chi tiết: sửa thông tin, quy trình 4 bước, đổi trạng thái, xin gia hạn, trao đổi theo từng yêu cầu, ảnh/tài liệu đính kèm (ảnh tự thu nhỏ ≤1600px), nhật ký thay đổi.
- Gia hạn: danh sách chờ duyệt, duyệt/từ chối.
- Báo cáo tuần/tháng xuất Word hoặc Excel.
- Email (MailApp, gửi từ tài khoản Admin): một thư tổng hợp/người/ngày (EMAIL_LOG chống gửi trùng) cho người có `nhan_mail=1`; thư tức thời cho người duyệt khi có đề nghị gia hạn và cho người đề nghị khi có kết quả. Không gồm dữ liệu cũ.
- Không tích hợp tự động với Viber, Bravo, Vaeco.


## 3b. Cập nhật lại dữ liệu cũ khi có file Excel mới
1. Mở file Excel đã làm sạch → chọn toàn bộ sheet → Copy.
2. Trong Google Sheet của hệ thống, dán vào ô **A1** của sheet `NHAP_CU_VP` (file Văn phòng) hoặc `NHAP_CU_PV` (file TT PVBD), giữ nguyên bố cục gốc (có dòng tiêu đề "STT"). Nếu chưa có hai sheet này, bấm "Kiểm tra" một lần, app sẽ tự tạo.
3. Tab **Quản trị → Cập nhật lại dữ liệu cũ từ Excel → Kiểm tra (không ghi)** để xem số dòng thêm mới / cập nhật / giữ nguyên / không còn trong file và các lỗi (ví dụ ngày không đọc được). Hết lỗi thì bấm **Áp dụng cập nhật**.
- Ghép theo mã cũ (VP0021, PV0005…): dòng đổi thì cập nhật, dòng mới thì thêm, dòng không còn trong file thì **giữ lại và gắn cờ rà soát** (không xóa). Trao đổi, đính kèm, nhật ký, quy trình của yêu cầu được giữ nguyên.
- Không ghi đè: việc đã đóng trong app, hạn đã được duyệt gia hạn, dòng đã bấm "Đã rà soát" mà nội dung cảnh báo không đổi.
- Giữ nguyên số thứ tự (STT) của các dòng cũ trong file mới; nếu đổi STT, dòng đó sẽ bị coi là dòng mới và dòng cũ bị gắn cờ "không còn trong file".
- Mọi lần áp dụng đều ghi vào nhật ký thay đổi.

## 4. Kiểm thử và giới hạn đã biết
- Đã chạy 91 kiểm tra logic phía server (kể cả cập nhật dữ liệu cũ, đăng nhập bằng mã, chế độ khách) và 25 kiểm tra giao diện trên bộ giả lập Google (phân quyền, chặn tổ hợp loại/đơn vị, Pending, gia hạn Tổ/Phòng, tự duyệt, đóng khi thiếu bước, dữ liệu cũ chỉ đọc, chống gửi trùng email, chuỗi băm nhật ký, báo cáo tuần/tháng). Bộ giả lập đã phát hiện và sửa 1 lỗi giao diện.
- **Chưa kiểm chứng trên Google thật**: gửi email (kể cả mã đăng nhập), tải ảnh lên Drive, xuất tệp Word/Excel, trigger, ghi nhớ đăng nhập trong trình duyệt. Hãy thử đủ các bước này sau khi cài và báo lại nếu có lỗi.
- Gmail cá nhân giới hạn khoảng 100 email/ngày; mỗi thao tác ghi đọc cả sheet nên phù hợp quy mô vài nghìn dòng.

## 5. Điểm cần bạn xác nhận
1. Dữ liệu cũ mặc định chỉ đọc – có cần đóng/cập nhật các việc cũ đang Inprogress (128 việc) ngay trên app không? Nếu có, bật `LEGACY_EDITABLE=1`.
2. Có áp quy tắc "gia hạn dài cần CB phòng duyệt" không, và ngưỡng bao nhiêu ngày (`EXT_PHONG_NGAY`)? Hiện tắt, CB tổ duyệt tất cả.
3. Tên các biểu mẫu 9101, 9110, 7001, 7011 (điền ở sheet `FORMS`).


## Bản cập nhật v3 – Ưu tiên, Thông báo đơn vị, Báo cáo định kỳ, Dashboard mới

**Cách cập nhật**
1. Dán toàn bộ `Code.gs` vào `Mã.gs` trong Apps Script → Triển khai → Quản lý triển khai → ✏️ → Phiên bản mới (Thực thi: Tôi; Truy cập: Bất kỳ ai).
2. Thay `index.html` trong thư mục `web` trên Netlify, **giữ nguyên `config.js` có địa chỉ /exec của bạn**.
3. Lần gọi đầu tiên hệ thống tự thêm cột `uu_tien`, sheet `UNIT_CONTACTS` và cài trigger.

**Tính năng**
- **Thứ tự ưu tiên** (Khẩn / Cao / Bình thường / Thấp): cột riêng trong danh sách, lọc và sắp xếp theo ưu tiên; Editor đổi trong form chi tiết.
- **Dashboard mới**: thẻ KPI, bảng "Cần xử lý ngay", biểu đồ ưu tiên, xu hướng 6 tháng, tuổi quá hạn, tải theo đơn vị.
- **Thông báo tới đơn vị**: nút "📣 Thông báo đơn vị" trong chi tiết yêu cầu; thông báo tự động khi có yêu cầu mới, hoàn thành, Pending, gia hạn (bật/tắt bằng NOTIFY_AUTO). Admin nhập email từng đơn vị ở Quản trị → Liên hệ đơn vị.
- **Báo cáo định kỳ**: tuần (sáng thứ Hai), tháng (sáng ngày 1), gửi Admin + cán bộ Tổ/Phòng có bật nhận mail, kèm CSV. Nút gửi thử ở Quản trị → Báo cáo định kỳ.
- Điền **APP_URL** (địa chỉ web Netlify) để email có nút "Mở hệ thống".
- Lưu ý: Gmail thường giới hạn ~100 người nhận/ngày.

## Bản cập nhật v3.1
- Sửa lỗi xuất báo cáo Excel ("Số cột trong dữ liệu không khớp...").
- Editor và Admin có nút **Xóa yêu cầu** trong chi tiết (bắt buộc nhập lý do; bản chụp nội dung + lý do lưu trong Nhật ký thay đổi; đề nghị gia hạn đang chờ của yêu cầu đó bị hủy). Viewer không có quyền này.
- Giao diện kỹ thuật VAECO: nền lưới bản vẽ, màu xanh navy + vàng, thẻ số liệu phẳng chữ đơn cách, bảng gọn.
- Cập nhật: dán `Code.gs` vào Mã.gs + phiên bản triển khai mới; thay `index.html` trên Netlify (giữ `config.js`).

## Bản v3.2
- Thanh menu trái luôn hiện, có logo VAECO; nút **Tạo yêu cầu mới** ở góc trên phải (mọi tab); nút đổi giao diện sáng/tối (nhớ lựa chọn theo trình duyệt).
- Dashboard bỏ 3 ô: Pending kéo dài, Chưa cập nhật, Cần rà soát dữ liệu.
