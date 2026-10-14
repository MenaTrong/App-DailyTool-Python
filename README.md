# DailyTool

Ứng dụng xuất ngày nghỉ bác sĩ từ file chấm công tháng. Bản dùng hằng ngày là `dist/DailyTool.exe`.

## Build trên Windows

Đóng DailyTool nếu đang mở, rồi chạy trong thư mục dự án:

```powershell
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean DailyTool.spec
```

File tạo ra: `dist\DailyTool.exe`.

`build_windows.sh` dùng khi build trên macOS hoặc Linux qua Docker. `DailyTool.command` là cách mở app trên Mac.

Danh mục mã khi chạy từ mã nguồn nằm ở `data\ma_bac_si.json`. Bản exe đọc `%APPDATA%\DailyTool\ma_bac_si.json`.
