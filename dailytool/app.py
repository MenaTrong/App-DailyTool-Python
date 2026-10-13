"""DailyTool: ứng dụng cục bộ, tác vụ đầu tiên là xuất ngày nghỉ bác sĩ."""

from __future__ import annotations

import html
import threading
import uuid
import webbrowser
from email.parser import BytesParser
from email.policy import HTTP
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import parse_qs, quote, urlparse

from dailytool.doctors import (
    add_doctor,
    delete_doctor,
    import_attendance,
    list_doctors,
    summary,
    update_doctor,
)
from xuat_ngay_nghi import export_leave_workbook, norm_name

APP_NAME = "DailyTool"
HOST = "127.0.0.1"
PORT = 8765

_RESULTS: dict[str, tuple[Path, str]] = {}
_TEMP = TemporaryDirectory(prefix="dailytool-")


def _page(title: str, body: str) -> bytes:
    document = f"""<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #1f2933;
      --muted: #52606d;
      --line: #e4e7eb;
      --paper: #f7f6f3;
      --card: #ffffff;
      --green: #1b4332;
      --green-2: #2d6a4f;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Avenir Next", "Segoe UI", sans-serif;
      background: var(--paper);
      color: var(--ink);
    }}
    main {{
      max-width: 760px;
      margin: 0 auto;
      padding: 40px 20px 64px;
    }}
    a {{ color: var(--green-2); }}
    h1 {{
      margin: 0;
      font-size: 32px;
      letter-spacing: -0.03em;
    }}
    .sub {{ margin: 8px 0 28px; color: var(--muted); }}
    .card {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 22px;
    }}
    .card h2 {{ margin: 0 0 8px; font-size: 20px; }}
    .card p {{ margin: 0 0 16px; color: var(--muted); line-height: 1.45; }}
    label {{ display: block; font-weight: 600; margin: 14px 0 6px; }}
    input[type="text"], input[type="file"] {{
      width: 100%;
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-radius: 10px;
      font: inherit;
      background: #fff;
    }}
    button, .button {{
      display: inline-block;
      margin-top: 18px;
      margin-right: 8px;
      background: var(--green);
      color: white;
      text-decoration: none;
      border: 0;
      border-radius: 10px;
      padding: 11px 16px;
      font: inherit;
      font-weight: 600;
      cursor: pointer;
    }}
    .button.secondary, button.secondary {{
      background: #e7f5ee;
      color: var(--green);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
    }}
    th, td {{
      text-align: left;
      padding: 8px 6px;
      border-bottom: 1px solid var(--line);
      vertical-align: middle;
    }}
    td input[type="text"] {{ width: 100%; margin: 0; }}
    td button {{ margin-top: 0; }}
    .row-actions {{ display: flex; gap: 6px; }}
    .back {{ color: var(--muted); text-decoration: none; font-size: 14px; }}
    .note {{ margin-top: 14px; color: var(--muted); font-size: 14px; }}
    .stats {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 10px;
      margin: 16px 0;
    }}
    .stat {{
      background: #f8faf8;
      border-radius: 12px;
      padding: 12px;
    }}
    .stat b {{ display: block; font-size: 22px; }}
    .stat span {{ color: var(--muted); font-size: 13px; }}
    .error {{ color: #9f1239; }}
  </style>
</head>
<body>
  <main>
    {body}
  </main>
</body>
</html>"""
    return document.encode("utf-8")


def _home() -> bytes:
    return _page(
        APP_NAME,
        f"""
        <h1>{APP_NAME}</h1>
        <p class="sub">Chọn một tác vụ.</p>
        <section class="card">
          <h2>Ngày nghỉ bác sĩ</h2>
          <p>Đọc file chấm công tháng và xuất danh sách ngày bác sĩ nghỉ. Ngày đã đánh + hoặc T sẽ không có trong file kết quả. Cột mã bác sĩ lấy theo danh mục trong tác vụ này.</p>
          <a class="button" href="/ngay-nghi">Mở tác vụ</a>
        </section>
        """,
    )


def _leave_form(message: str = "", path: str = "") -> bytes:
    notice = f'<p class="error">{html.escape(message)}</p>' if message else ""
    return _page(
        "Ngày nghỉ bác sĩ",
        f"""
        <a class="back" href="/">← Tác vụ</a>
        <h1>Ngày nghỉ bác sĩ</h1>
        <p class="sub">Chọn file Excel chấm công. Kết quả tải về máy.</p>
        <form class="card" action="/ngay-nghi" method="post" enctype="multipart/form-data" accept-charset="utf-8">
          {notice}
          <label for="path">Đường dẫn file trên máy</label>
          <input id="path" name="path" type="text" placeholder="C:\\Users\\...\\cham-cong.xlsx" value="{html.escape(path)}">
          <label for="file">Hoặc tải file lên</label>
          <input id="file" name="file" type="file" accept=".xlsx">
          <button type="submit">Xuất Excel</button>
          <a class="button secondary" href="/ma-bac-si">Mã bác sĩ</a>
          <p class="note">Nếu điền cả hai, DailyTool dùng file được tải lên. Mã bác sĩ thêm, sửa, xóa trong danh mục của tác vụ này.</p>
        </form>
        """,
    )


def _result_page(result: dict, token: str) -> bytes:
    return _page(
        "Đã xuất ngày nghỉ",
        f"""
        <a class="back" href="/ngay-nghi">← Làm lại</a>
        <h1>Tháng {result["month"]}/{result["year"]}</h1>
        <p class="sub">{result["doctors"]} bác sĩ, {result["rows"]} dòng ngày nghỉ.</p>
        <section class="card">
          <div class="stats">
            <div class="stat"><b>{result["full"]}</b><span>Cả ngày</span></div>
            <div class="stat"><b>{result["morning"]}</b><span>Nghỉ sáng</span></div>
            <div class="stat"><b>{result["afternoon"]}</b><span>Nghỉ chiều</span></div>
          </div>
          <a class="button" href="/tai/{token}">Tải Ngày Nghỉ BS T{result["month"]}.xlsx</a>
          <p class="note">{_code_note(result)}</p>
        </section>
        """,
    )


def _code_note(result: dict) -> str:
    parts = [f"Đã gắn mã cho {result['coded']} bác sĩ."]
    uncoded = result.get("uncoded") or []
    ambiguous = result.get("ambiguous") or []
    if uncoded:
        parts.append("Chưa có mã: " + ", ".join(uncoded) + ".")
    if ambiguous:
        parts.append("Cần chọn một mã: " + ", ".join(ambiguous) + ".")
    if result.get("codes_added"):
        parts.append(f"Đã thêm {result['codes_added']} bác sĩ mới vào danh mục.")
    return html.escape(" ".join(parts))


def _doctors_page(message: str = "", error: bool = False, query: str = "") -> bytes:
    info = summary()
    needle = norm_name(query)
    rows = []
    for doctor in list_doctors():
        haystack = f"{norm_name(doctor['name'])} {doctor['code'].casefold()}"
        if needle and needle not in haystack:
            continue
        doctor_id = html.escape(doctor["id"])
        name = html.escape(doctor["name"])
        code = html.escape(doctor["code"])
        missing = ' style="color:#9f1239"' if not doctor["code"] else ""
        rows.append(
            f"""
            <tr>
              <td><input form="save-{doctor_id}" name="name" type="text" value="{name}" required></td>
              <td><input form="save-{doctor_id}" name="code" type="text" value="{code}"{missing}></td>
              <td>
                <div class="row-actions">
                  <form id="save-{doctor_id}" action="/ma-bac-si" method="post">
                    <input type="hidden" name="action" value="save">
                    <input type="hidden" name="id" value="{doctor_id}">
                    <button type="submit">Lưu</button>
                  </form>
                  <form action="/ma-bac-si" method="post">
                    <input type="hidden" name="action" value="delete">
                    <input type="hidden" name="id" value="{doctor_id}">
                    <button class="secondary" type="submit">Xóa</button>
                  </form>
                </div>
              </td>
            </tr>
            """
        )
    notice_class = "error" if error else "note"
    notice = f'<p class="{notice_class}">{html.escape(message)}</p>' if message else ""
    if rows:
        body_rows = "".join(rows)
    elif query:
        body_rows = "<tr><td colspan='3'>Không thấy bác sĩ khớp.</td></tr>"
    else:
        body_rows = "<tr><td colspan='3'>Chưa có bác sĩ.</td></tr>"
    return _page(
        "Mã bác sĩ",
        f"""
        <a class="back" href="/ngay-nghi">← Ngày nghỉ bác sĩ</a>
        <h1>Mã bác sĩ</h1>
        <p class="sub">{info["with_code"]} có mã, {info["without_code"]} chưa có mã. Khi xuất, mã được ghép theo họ tên.</p>
        {notice}
        <section class="card">
          <h2>Thêm mã</h2>
          <form action="/ma-bac-si" method="post">
            <input type="hidden" name="action" value="add">
            <label for="name">Họ và tên</label>
            <input id="name" name="name" type="text" required>
            <label for="code">Mã bác sĩ</label>
            <input id="code" name="code" type="text" placeholder="vi_du_kb">
            <button type="submit">Thêm</button>
          </form>
        </section>
        <section class="card" style="margin-top:16px">
          <h2>Lấy bác sĩ từ file chấm công</h2>
          <p>Thêm người có trong file mà chưa có trong danh mục. Mã đã nhập được giữ nguyên.</p>
          <form action="/ma-bac-si" method="post" enctype="multipart/form-data" accept-charset="utf-8">
            <input type="hidden" name="action" value="import">
            <label for="path">Đường dẫn file trên máy</label>
            <input id="path" name="path" type="text" placeholder="C:\\Users\\...\\cham-cong.xlsx">
            <label for="file">Hoặc tải file lên</label>
            <input id="file" name="file" type="file" accept=".xlsx">
            <button type="submit">Lấy bác sĩ</button>
          </form>
        </section>
        <section class="card" style="margin-top:16px">
          <h2>Danh mục</h2>
          <form action="/ma-bac-si" method="get">
            <label for="q">Tìm</label>
            <input id="q" name="q" type="text" value="{html.escape(query)}">
            <button class="secondary" type="submit">Lọc</button>
          </form>
          <table>
            <thead><tr><th>Họ và tên</th><th>Mã bác sĩ</th><th></th></tr></thead>
            <tbody>{body_rows}</tbody>
          </table>
        </section>
        """,
    )


def _parse_form(handler: BaseHTTPRequestHandler) -> tuple[str, bytes, str, dict[str, str]]:
    length = int(handler.headers.get("Content-Length", "0"))
    body = handler.rfile.read(length)
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        parsed = parse_qs(body.decode("utf-8", errors="replace"))
        fields = {key: values[0] for key, values in parsed.items()}
        return fields.get("path", ""), b"", "", fields
    raw = b"Content-Type: " + content_type.encode("utf-8") + b"\r\nMIME-Version: 1.0\r\n\r\n" + body
    message = BytesParser(policy=HTTP).parsebytes(raw)
    fields: dict[str, str] = {}
    payload = b""
    filename = ""
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        if name == "file":
            filename = part.get_filename() or ""
            payload = part.get_payload(decode=True) or b""
            continue
        data = part.get_payload(decode=True) or b""
        fields[name] = data.decode("utf-8").strip()
    return fields.get("path", ""), payload, filename, fields


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/":
            self._html(_home())
            return
        if path == "/ngay-nghi":
            self._html(_leave_form())
            return
        if path == "/ma-bac-si":
            query = parse_qs(urlparse(self.path).query).get("q", [""])[0]
            self._html(_doctors_page(query=query))
            return
        if path.startswith("/tai/"):
            token = path.removeprefix("/tai/")
            stored = _RESULTS.get(token)
            if stored is None or not stored[0].exists():
                self._html(_leave_form("File tải về không còn. Hãy xuất lại."), status=404)
                return
            target, download_name = stored
            data = target.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(download_name)}")
            self.end_headers()
            self.wfile.write(data)
            return
        self._html(_page(APP_NAME, "<h1>Không thấy trang</h1><p><a href='/'>Về trang chủ</a></p>"), status=404)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/ma-bac-si":
            self._save_doctor()
            return
        if path != "/ngay-nghi":
            self._html(_page(APP_NAME, "<h1>Không thấy trang</h1>"), status=404)
            return
        path_text = ""
        try:
            path_text, payload, filename, _fields = _parse_form(self)
            source = self._source_file(path_text, payload, filename)
            output = Path(_TEMP.name) / f"{uuid.uuid4().hex}.xlsx"
            result = export_leave_workbook(source, output)
            download_name = f"Ngày Nghỉ BS T{result['month']}.xlsx"
            token = uuid.uuid4().hex
            _RESULTS[token] = (output, download_name)
            self._html(_result_page(result, token))
        except Exception as error:
            self._html(_leave_form(str(error), path_text), status=400)

    def _save_doctor(self) -> None:
        try:
            path_text, payload, filename, fields = _parse_form(self)
            action = fields.get("action", "")
            if action == "add":
                add_doctor(fields.get("name", ""), fields.get("code", ""))
                self._html(_doctors_page("Đã thêm mã bác sĩ."))
                return
            if action == "save":
                update_doctor(fields.get("id", ""), fields.get("name", ""), fields.get("code", ""))
                self._html(_doctors_page("Đã lưu mã bác sĩ."))
                return
            if action == "delete":
                delete_doctor(fields.get("id", ""))
                self._html(_doctors_page("Đã xóa mã bác sĩ."))
                return
            if action == "import":
                source = self._source_file(path_text, payload, filename)
                result = import_attendance(source)
                added = result["added"]
                if added:
                    self._html(_doctors_page(f"Đã thêm {len(added)} bác sĩ từ file chấm công."))
                else:
                    self._html(_doctors_page("File chấm công không có bác sĩ mới."))
                return
            raise ValueError("Không rõ thao tác.")
        except Exception as error:
            self._html(_doctors_page(str(error), error=True), status=400)

    def _source_file(self, path_text: str, payload: bytes, filename: str) -> Path:
        if payload and filename:
            if not filename.lower().endswith(".xlsx"):
                raise ValueError("Hãy chọn file Excel .xlsx.")
            target = Path(_TEMP.name) / f"upload-{uuid.uuid4().hex}.xlsx"
            target.write_bytes(payload)
            return target
        if path_text:
            source = Path(path_text).expanduser()
            if not source.exists():
                raise FileNotFoundError(f"Không thấy file: {source}")
            return source
        raise ValueError("Hãy chọn file chấm công.")

    def _html(self, payload: bytes, status: int = 200) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args) -> None:
        print(f"[{self.log_date_time_string()}] {format % args}")


def _open_browser(url: str) -> None:
    try:
        webbrowser.open(url)
    except Exception:
        pass


def _bind_server() -> tuple[ThreadingHTTPServer, int]:
    last_error: Exception | None = None
    for port in range(PORT, PORT + 20):
        try:
            return ThreadingHTTPServer((HOST, port), Handler), port
        except OSError as error:
            last_error = error
    raise RuntimeError(f"Không mở được cổng từ {PORT}. {last_error}")


def main() -> None:
    server, port = _bind_server()
    url = f"http://{HOST}:{port}"
    print(f"{APP_NAME} đang chạy tại {url}", flush=True)
    print("Đóng cửa sổ này để thoát.", flush=True)
    threading.Timer(0.6, lambda: _open_browser(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print(flush=True)
    finally:
        server.server_close()
        _TEMP.cleanup()


if __name__ == "__main__":
    main()
