# ViFormat (bản thử 0.1, tác giả Phạm Văn Huy) - add-in Word chạy offline, Windows + Mac

Một file `ViFormat.dotm` (template Word có VBA và Ribbon). Không cần internet, không cần runtime, copy là chạy.

## Tính năng bản 0.1

| Nút | Việc làm |
|---|---|
| Quy chuẩn (dropdown) | Chọn NĐ 30/2020/NĐ-CP hoặc HD 05-HD/VPTW |
| Xem quy chuẩn | Hộp thoại cuộn hiển thị chi tiết quy chuẩn đang chọn |
| Chuẩn hóa | Căn lề A4, phông, cỡ chữ, giãn dòng, căn đều, khối đầu (Quốc hiệu/Tiêu ngữ/Đảng), "Nơi nhận", số trang. Hoàn tác bằng Undo |
| Rà soát thể thức | Kiểm tra và ghi Comment (tác giả ViFormat), không tự sửa. Tối đa 40 comment |
| Xóa nhận xét ViFormat | Xóa các comment của ViFormat |
| Căn lề trang A4 / Đánh số trang | Theo quy chuẩn đang chọn |
| Chèn trang ngang / dọc | Ngắt section kèm đổi hướng giấy |
| Xóa trang trắng | Ngắt trang kép và đoạn trống cuối văn bản |

**Chưa làm:** chuyển mã TCVN3/VNI sang Unicode, kiểm tra chính tả, thư viện biểu mẫu, chèn nhanh thành phần, nhận diện loại văn bản, bốn quy chuẩn còn lại.

## Cấu trúc

```
src/         mã nguồn VBA (UTF-8, có tiếng Việt, để đọc và sửa)
dist/        sinh tự động từ src/: thuần ASCII, tiếng Việt đã đổi thành U("...\uXXXX...")
customUI/    Ribbon (customUI14.xml)
rules/       tài liệu số liệu và nguồn (rules.json)
build/       build_dotm.py (tạo khung .dotm), ImportModules.vbs (nạp VBA, Windows)
tools/       make_dist.py, lint_vba.py
install/     install.cmd (Windows), install.command (Mac)
```

Vì sao có `dist/`: VBA lưu mã nguồn theo bảng mã ANSI của máy, nên tiếng Việt viết thẳng trong code sẽ hỏng. **Luôn sửa trong `src/`, không sửa `dist/`.**

## Build (Python 3, chạy được trên Mac và Windows, không cần Word)

```
python tools\make_dist.py        # src -> dist (ASCII)
python tools\lint_vba.py         # kiểm tra tĩnh
python tools\make_vbaproject.py  # dist/*.bas -> build/vbaProject.bin (tự sinh VBA project)
python build\build_dotm.py --vba build\vbaProject.bin   # tạo ViFormat.dotm đầy đủ
```

`ImportModules.vbs` cần bật trước: Word > File > Options > Trust Center > Trust Center Settings > Macro Settings > tick **Trust access to the VBA project object model**.

Không dùng script được thì nạp tay: mở `ViFormat.dotm` trong Word, Alt+F11 (Mac: Tools > Macro > Visual Basic Editor), File > Import File, chọn từng file trong `dist/`, rồi Save.

## Cài lên máy khác

- **Windows:** chạy `install\install.cmd` (copy vào `%AppData%\Microsoft\Word\STARTUP` và Unblock file).
- **Mac:** `bash install/install.command` (copy vào `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/`).
- Hoặc tự copy file vào đúng thư mục trên rồi mở lại Word.

## Cần test trên Word thật (chưa kiểm chứng được ở môi trường dựng)

Bản này mới được kiểm tra tĩnh (cú pháp, cặp khối lệnh, callback Ribbon, XML, đóng gói .dotm). **Chưa chạy trong Word lần nào.** Các điểm rủi ro, nhất là trên Mac:

1. Ribbon `customUI14.xml` trong template ở thư mục Startup của Word Mac. Dropdown động có hiện không.
2. `Application.UndoRecord` trên Mac (code đã bọc để thiếu thì chỉ mất Undo gộp).
3. `Information(wdWithInTable)` và `Selection.InsertBreak` trên Mac.
4. Hiệu năng với văn bản dài (vài trăm trang).
5. Icon `imageMso` chỉ để trang trí, sai tên chỉ mất icon.

## Số liệu

- **HD 05:** đọc từ bản PDF gốc.
- **NĐ 30:** lề, phông, cỡ chữ từ trang tóm tắt. **Khoảng cách dòng, khoảng cách đoạn, lùi đầu dòng chưa xác minh**, nên Rà soát không kiểm tra các mục này.
- Mọi số liệu cần đối chiếu lại với phụ lục gốc trước khi dùng thật. Xem `rules/rules.json`.

## Đóng gói

Thư mục `release/ViFormat-0.1/` (và file zip) chứa `ViFormat.dotm`, trình cài/gỡ cho Mac và Windows, hướng dẫn. Windows chưa được kiểm chứng.
