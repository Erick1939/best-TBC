# Điều khiển từ xa — config.json

Sửa trực tiếp file `config.json` trong repo này (bấm vào file trên GitHub → nút ✏️ Edit →
sửa → Commit changes). Mọi người dùng tool sẽ nhận thay đổi trong tối đa 15 phút
(hoặc ngay lần tiếp theo họ mở lại trang), **không cần đẩy bản code mới**.

## 1. Khoá tạm toàn bộ tool

```json
{
  "locked": true,
  "lockMsg": "Đang bảo trì hệ thống, sẽ mở lại trong khoảng 30 phút nữa.",
  "notice": "",
  "noticeId": ""
}
```

- `locked: true` → tool hiện màn khoá (🔒), không dùng được, chỉ có nút Đóng.
- `lockMsg` → lý do hiển thị cho người dùng. Có thể để nhiều dòng bằng `\n`.
- Mở lại: đổi `"locked": false`.

## 2. Gửi thông báo (không khoá tool)

```json
{
  "locked": false,
  "lockMsg": "",
  "notice": "Từ ngày 25/10, đổi giờ chốt đơn mới sang 22:00. Anh/chị lưu ý nhé!",
  "noticeId": "2026-10-23-doi-gio-chot"
}
```

- `notice` → nội dung banner cam hiện trong tool, có nút "Đã đọc".
- `noticeId` → mỗi lần đổi nội dung thông báo, đổi luôn `noticeId` (ví dụ thêm ngày)
  để banner hiện lại cho cả người đã từng bấm "Đã đọc" thông báo cũ.
  Nếu giữ `noticeId` cũ mà chỉ sửa `notice`, người đã bấm "Đã đọc" sẽ không thấy lại.
- Muốn tắt thông báo: để `"notice": ""`.

## 3. Vừa khoá vừa kèm thông báo

Hai cái độc lập nhau, có thể bật cùng lúc: `locked:true` sẽ chặn hẳn tool (banner
thông báo sẽ không hiện vì màn khoá thay chỗ nó), còn `locked:false` + có `notice`
thì tool vẫn dùng được bình thường, chỉ thêm 1 banner nhắc.
