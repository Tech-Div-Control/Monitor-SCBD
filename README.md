# Quản lý yêu cầu sửa chữa – bảo dưỡng (HAN BASE – Technical Division, VAECO)

Hệ thống theo dõi yêu cầu sửa chữa / bảo dưỡng / hiệu chuẩn / chế tạo.
Dữ liệu: Google Sheets · Máy chủ: Google Apps Script (web app) · Giao diện: trang tĩnh trên Netlify.

```
apps-script/   Code.gs, Index.html (nguồn giao diện), appsscript.json  → dán/đẩy lên Apps Script
docs/          index.html (SINH TỰ ĐỘNG), config.js (địa chỉ /exec)   → GitHub Pages xuất bản thư mục này
tools/         build_web.py: sinh docs/index.html từ apps-script/Index.html
tai-lieu/      HUONG_DAN.md, DuLieu_Khoi_Tao.xlsx
```

## Quy trình làm việc
1. Sửa `apps-script/Index.html` và/hoặc `apps-script/Code.gs`.
2. Chạy `python3 tools/build_web.py` để cập nhật `docs/index.html`.
3. `git add -A && git commit -m "..." && git push`.
4. **Giao diện (GitHub Pages):** tự triển khai sau vài chục giây khi push lên nhánh `main`.
5. **Máy chủ (Apps Script):** dán `apps-script/Code.gs` vào Mã.gs → Triển khai → Quản lý triển khai → ✏️ → Phiên bản mới
   (Thực thi: Tôi · Truy cập: Bất kỳ ai). Hoặc dùng `clasp push` (xem dưới).

## Cấu hình lần đầu
- Sửa `docs/config.js`: dán URL `/exec` của web app. Tệp này nằm trong repo nên các lần triển khai sau **không bị ghi đè**.
- GitHub: *Settings → Pages → Source: Deploy from a branch → Branch `main`, thư mục `/docs`*.

## Tuỳ chọn: đẩy Code.gs bằng clasp
```
npm i -g @google/clasp && clasp login
clasp clone <SCRIPT_ID> --rootDir apps-script   # hoặc tạo .clasp.json có rootDir = apps-script
clasp push
```
Sau `clasp push` vẫn cần tạo phiên bản triển khai mới trong giao diện Apps Script (hoặc `clasp deploy`).

## Lưu ý
- `config.js` chỉ chứa URL web app (công khai), không chứa mật khẩu/khóa.
- Dữ liệu thật nằm trong Google Sheet, **không** đưa vào repo.
