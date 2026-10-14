"""Cửa sổ DailyTool. Tác vụ đầu tiên là xuất ngày nghỉ bác sĩ."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from dailytool import __version__
from dailytool.doctors import (
    add_doctor,
    delete_doctor,
    import_attendance,
    list_doctors,
    summary,
    update_doctor,
)
from xuat_ngay_nghi import build_leave_order_scan_from_file, export_leave_workbook, norm_name

SIDEBAR = "#1b4332"
SIDEBAR_ACTIVE = "#2d6a4f"
INK = "#1f2933"
MUTED = "#52606d"


def _enable_windows_dpi() -> None:
    if sys.platform != "win32":
        return
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DailyTool")
    except Exception:
        pass


class DoctorCodes(tk.Toplevel):
    """Thêm, sửa, xóa mã bác sĩ của tác vụ ngày nghỉ."""

    def __init__(self, master: tk.Misc, source_path: str) -> None:
        super().__init__(master)
        self.title("Mã bác sĩ")
        self.geometry("740x540")
        self.minsize(640, 440)
        self.transient(master)
        self.source_path = source_path
        self.name_var = tk.StringVar()
        self.code_var = tk.StringVar()
        self.query_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self._selected = ""
        self._busy = False
        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.query_var.trace_add("write", lambda *_args: self.refresh())

        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Mã bác sĩ", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text="Dùng khi xuất ngày nghỉ. Mã ghép theo họ tên trong file chấm công.\nBác sĩ chưa có mã được thêm vào đây để điền.",
            style="Muted.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(4, 10))

        search = ttk.Frame(frame)
        search.pack(fill="x")
        ttk.Label(search, text="Tìm", style="Body.TLabel").pack(side="left")
        ttk.Entry(search, textvariable=self.query_var).pack(side="left", fill="x", expand=True, padx=(8, 0))

        table = ttk.Frame(frame)
        table.pack(fill="both", expand=True, pady=(10, 8))
        self.tree = ttk.Treeview(table, columns=("name", "code"), show="headings", selectmode="browse")
        self.tree.heading("name", text="Họ và tên")
        self.tree.heading("code", text="Mã bác sĩ")
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("code", width=180, anchor="w")
        self.tree.tag_configure("empty", foreground="#9f1239")
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = ttk.Frame(frame)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)
        ttk.Label(form, text="Họ và tên", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.name_var).grid(row=0, column=1, sticky="ew", padx=(8, 16))
        ttk.Label(form, text="Mã bác sĩ", style="Body.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(form, textvariable=self.code_var).grid(row=0, column=3, sticky="ew", padx=(8, 0))

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Thêm", command=self.add).pack(side="left")
        self.save_button = ttk.Button(actions, text="Lưu sửa", command=self.save)
        self.save_button.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Xóa", command=self.remove).pack(side="left", padx=(8, 0))
        self.import_button = ttk.Button(actions, text="Lấy từ file chấm công", command=self.import_file)
        self.import_button.pack(side="left", padx=(16, 0))

        ttk.Label(frame, textvariable=self.status_var, style="Muted.TLabel").pack(anchor="w", pady=(12, 0))
        self.refresh()
        self.grab_set()

    def refresh(self) -> None:
        query = norm_name(self.query_var.get())
        selected = self._selected
        children = self.tree.get_children()
        if children:
            self.tree.delete(*children)
        shown = 0
        for doctor in list_doctors():
            haystack = f"{norm_name(doctor['name'])} {doctor['code'].casefold()}"
            if query and query not in haystack:
                continue
            tags = ("empty",) if not doctor["code"] else ()
            self.tree.insert("", "end", iid=doctor["id"], values=(doctor["name"], doctor["code"]), tags=tags)
            shown += 1
        if selected and self.tree.exists(selected):
            self.tree.selection_set(selected)
            self.tree.see(selected)
        else:
            self._selected = ""
        info = summary()
        self.status_var.set(
            f"{shown} dòng đang hiện. Danh mục: {info['with_code']} có mã, {info['without_code']} chưa có mã."
        )

    def _on_select(self, _event: object) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        self._selected = selected[0]
        values = self.tree.item(self._selected, "values")
        self.name_var.set(values[0] if values else "")
        self.code_var.set(values[1] if len(values) > 1 else "")

    def _fields(self) -> tuple[str, str]:
        return self.name_var.get(), self.code_var.get()

    def add(self) -> None:
        name, code = self._fields()
        try:
            doctor = add_doctor(name, code)
        except ValueError as error:
            messagebox.showwarning("Mã bác sĩ", str(error), parent=self)
            return
        self._selected = doctor["id"]
        self.refresh()

    def save(self) -> None:
        if not self._selected:
            messagebox.showwarning("Mã bác sĩ", "Hãy chọn một dòng để sửa.", parent=self)
            return
        name, code = self._fields()
        try:
            update_doctor(self._selected, name, code)
        except ValueError as error:
            messagebox.showwarning("Mã bác sĩ", str(error), parent=self)
            return
        self.refresh()

    def remove(self) -> None:
        if not self._selected:
            messagebox.showwarning("Mã bác sĩ", "Hãy chọn một dòng để xóa.", parent=self)
            return
        name = self.name_var.get().strip() or "dòng này"
        if not messagebox.askyesno("Mã bác sĩ", f"Xóa mã của {name}?", parent=self):
            return
        try:
            delete_doctor(self._selected)
        except ValueError as error:
            messagebox.showwarning("Mã bác sĩ", str(error), parent=self)
            return
        self._selected = ""
        self.name_var.set("")
        self.code_var.set("")
        self.refresh()

    def import_file(self) -> None:
        if self._busy:
            return
        path = self.source_path
        if not path or not Path(path).expanduser().exists():
            path = filedialog.askopenfilename(
                parent=self,
                title="Chọn file chấm công",
                filetypes=[("Excel", "*.xlsx"), ("Tất cả", "*.*")],
            )
        if not path:
            return
        self.source_path = path
        self._busy = True
        self.import_button.state(["disabled"])
        self.status_var.set("Đang đọc bác sĩ trong file chấm công…")

        def work() -> None:
            try:
                result = import_attendance(Path(path))
            except Exception as error:
                self._queue.put(("err", error))
                return
            self._queue.put(("ok", result))

        threading.Thread(target=work, daemon=True).start()
        self.after(150, self._poll_import)

    def _poll_import(self) -> None:
        if not self.winfo_exists():
            return
        try:
            kind, payload = self._queue.get_nowait()
        except queue.Empty:
            if self._busy:
                self.after(150, self._poll_import)
            return
        self._busy = False
        self.import_button.state(["!disabled"])
        if kind == "err":
            messagebox.showerror("Mã bác sĩ", str(payload), parent=self)
            self.refresh()
            return
        added = payload["added"]  # type: ignore[index]
        self.refresh()
        if added:
            self.status_var.set(f"Đã thêm {len(added)} bác sĩ từ file chấm công. Hãy điền mã cho người chưa có.")
            return
        self.status_var.set("File chấm công không có bác sĩ mới. Danh mục giữ nguyên.")


class DailyTool(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"DailyTool {__version__}")
        self.geometry("900x640")
        self.minsize(780, 560)
        self.configure(bg="#f4f6f5")
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.save_excel_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Chọn file chấm công, rồi xuất.")
        self._busy = False
        self._last_output: Path | None = None
        self._last_sql_text = ""
        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()

        style = ttk.Style(self)
        if sys.platform == "win32" and "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground=INK)
        style.configure("Body.TLabel", font=("Segoe UI", 11), foreground=INK)
        style.configure("Muted.TLabel", font=("Segoe UI", 10), foreground=MUTED)
        style.configure("Result.TLabel", font=("Segoe UI", 12), foreground=INK)
        style.configure("TButton", font=("Segoe UI", 11))

        sidebar = tk.Frame(self, bg=SIDEBAR, width=210)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="DailyTool", bg=SIDEBAR, fg="white", font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=18, pady=(22, 0))
        tk.Label(sidebar, text=f"v{__version__}", bg=SIDEBAR, fg="#b7e4c7", font=("Segoe UI", 9)).pack(anchor="w", padx=18, pady=(0, 16))
        tk.Label(sidebar, text="TÁC VỤ", bg=SIDEBAR, fg="#b7e4c7", font=("Segoe UI", 9)).pack(anchor="w", padx=18)
        tk.Label(
            sidebar,
            text="Ngày nghỉ bác sĩ",
            bg=SIDEBAR_ACTIVE,
            fg="white",
            font=("Segoe UI", 11),
            anchor="w",
            padx=18,
            pady=10,
        ).pack(fill="x", pady=(8, 0))

        content = ttk.Frame(self, padding=(28, 24, 28, 18))
        content.pack(side="left", fill="both", expand=True)

        ttk.Label(content, text="Ngày nghỉ bác sĩ", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            content,
            text="Đọc file chấm công tháng và tạo danh sách ngày bác sĩ nghỉ.\nNgày đã đánh + hoặc T không được đưa vào kết quả.\nCột Mã bác sĩ lấy theo danh mục trong tác vụ này. SQL hiện trong app để copy.\n«SQL quét y lệnh»: SELECT từ lịch nghỉ trong file (không cần INSERT DB).",
            style="Muted.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(6, 18))

        form = ttk.Frame(content)
        form.pack(fill="x")
        form.columnconfigure(0, weight=1)

        ttk.Label(form, text="File chấm công", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.source_var).grid(row=1, column=0, sticky="ew", pady=(4, 8))
        ttk.Button(form, text="Chọn file…", command=self.choose_source).grid(row=1, column=1, padx=(8, 0), pady=(4, 8))

        self.save_excel_check = ttk.Checkbutton(
            form,
            text="Lưu file Excel",
            variable=self.save_excel_var,
            command=self._toggle_save_excel,
        )
        self.save_excel_check.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 4))

        self.output_label = ttk.Label(form, text="File kết quả", style="Body.TLabel")
        self.output_label.grid(row=3, column=0, sticky="w")
        self.output_entry = ttk.Entry(form, textvariable=self.output_var)
        self.output_entry.grid(row=4, column=0, sticky="ew", pady=(4, 0))
        self.output_browse = ttk.Button(form, text="Đổi chỗ lưu…", command=self.choose_output)
        self.output_browse.grid(row=4, column=1, padx=(8, 0), pady=(4, 0))

        actions = ttk.Frame(content)
        actions.pack(fill="x", pady=(18, 8))
        self.export_button = ttk.Button(actions, text="Xuất", command=self.export)
        self.export_button.pack(side="left")
        self.scan_sql_button = ttk.Button(actions, text="SQL quét y lệnh", command=self.scan_leave_orders)
        self.scan_sql_button.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Mã bác sĩ…", command=self.open_codes).pack(side="left", padx=(8, 0))
        self.open_button = ttk.Button(actions, text="Mở file Excel", command=self.open_result, state="disabled")
        self.open_button.pack(side="left", padx=(8, 0))
        self.copy_sql_button = ttk.Button(actions, text="Copy SQL", command=self.copy_sql, state="disabled")
        self.copy_sql_button.pack(side="left", padx=(8, 0))

        self.result_label = ttk.Label(content, text="", style="Result.TLabel", justify="left", wraplength=620)
        self.result_label.pack(anchor="w", pady=(12, 4))

        ttk.Label(content, text="SQL Oracle", style="Body.TLabel").pack(anchor="w", pady=(4, 0))
        self.sql_view = scrolledtext.ScrolledText(
            content,
            height=12,
            wrap="none",
            font=("Consolas", 10),
            state="disabled",
        )
        self.sql_view.pack(fill="both", expand=True, pady=(4, 8))

        ttk.Label(content, textvariable=self.status_var, style="Muted.TLabel").pack(side="bottom", anchor="w")

    def choose_source(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file chấm công",
            filetypes=[("Excel", "*.xlsx"), ("Tất cả", "*.*")],
        )
        if not path:
            return
        self.source_var.set(path)
        if not self.output_var.get().strip():
            self.output_var.set(str(Path(path).with_name("Ngày Nghỉ BS.xlsx")))
        self.status_var.set("Sẵn sàng xuất.")

    def choose_output(self) -> None:
        initial = self.output_var.get().strip()
        initial_dir = str(Path(initial).parent) if initial else ""
        path = filedialog.asksaveasfilename(
            title="Chọn chỗ lưu file ngày nghỉ",
            defaultextension=".xlsx",
            initialfile="Ngày Nghỉ BS.xlsx",
            initialdir=initial_dir or None,
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _toggle_save_excel(self) -> None:
        save = self.save_excel_var.get()
        state = "!disabled" if save else "disabled"
        self.output_entry.state([state])
        self.output_browse.state([state])
        if save:
            self.export_button.configure(text="Xuất")
            self.status_var.set("Sẵn sàng xuất (có lưu Excel).")
        else:
            self.export_button.configure(text="Tạo SQL")
            self.status_var.set("Sẵn sàng tạo SQL (không lưu Excel).")

    def _set_sql_view(self, text: str) -> None:
        self.sql_view.configure(state="normal")
        self.sql_view.delete("1.0", "end")
        if text:
            self.sql_view.insert("1.0", text)
        self.sql_view.configure(state="disabled")

    def export(self) -> None:
        if self._busy:
            return
        source_text = self.source_var.get().strip()
        if not source_text:
            messagebox.showwarning("DailyTool", "Hãy chọn file chấm công.")
            return
        source = Path(source_text).expanduser()
        if not source.exists():
            messagebox.showerror("DailyTool", f"Không thấy file:\n{source}")
            return

        write_excel = self.save_excel_var.get()
        output: Path | None = None
        if write_excel:
            output_text = self.output_var.get().strip()
            output = Path(output_text).expanduser() if output_text else None
            if output is not None and output.resolve() == source.resolve():
                messagebox.showerror("DailyTool", "File kết quả phải khác file chấm công.")
                return
            if output is None or output.resolve() == source.with_name("Ngày Nghỉ BS.xlsx").resolve():
                output = None

        self._busy = True
        self.export_button.state(["disabled"])
        self.scan_sql_button.state(["disabled"])
        self.open_button.state(["disabled"])
        self.copy_sql_button.state(["disabled"])
        self.result_label.configure(text="")
        self._set_sql_view("")
        self.status_var.set("Đang đọc file chấm công…")

        def work() -> None:
            try:
                result = export_leave_workbook(source, output, write_excel=write_excel)
            except Exception as error:
                self._queue.put(("err", error))
                return
            self._queue.put(("ok", result))

        threading.Thread(target=work, daemon=True).start()
        self.after(150, self._poll_export)

    def scan_leave_orders(self) -> None:
        if self._busy:
            return
        source_text = self.source_var.get().strip()
        if not source_text:
            messagebox.showwarning("DailyTool", "Hãy chọn file chấm công.")
            return
        source = Path(source_text).expanduser()
        if not source.exists():
            messagebox.showerror("DailyTool", f"Không thấy file:\n{source}")
            return

        self._busy = True
        self.export_button.state(["disabled"])
        self.scan_sql_button.state(["disabled"])
        self.copy_sql_button.state(["disabled"])
        self.result_label.configure(text="")
        self._set_sql_view("")
        self.status_var.set("Đang tạo SQL quét y lệnh ngày nghỉ…")

        def work() -> None:
            try:
                result = build_leave_order_scan_from_file(source)
            except Exception as error:
                self._queue.put(("scan_err", error))
                return
            self._queue.put(("scan_ok", result))

        threading.Thread(target=work, daemon=True).start()
        self.after(150, self._poll_export)

    def _poll_export(self) -> None:
        try:
            kind, payload = self._queue.get_nowait()
        except queue.Empty:
            if self._busy:
                self.after(150, self._poll_export)
            return
        if kind == "ok":
            self._export_done(payload)  # type: ignore[arg-type]
            return
        if kind == "scan_ok":
            self._scan_done(payload)  # type: ignore[arg-type]
            return
        if kind == "scan_err":
            self._scan_failed(payload)  # type: ignore[arg-type]
            return
        self._export_failed(payload)  # type: ignore[arg-type]

    def _export_done(self, result: dict) -> None:
        self._busy = False
        self.export_button.state(["!disabled"])
        self.scan_sql_button.state(["!disabled"])
        self._last_sql_text = str(result.get("sql_text") or "")
        self._set_sql_view(self._last_sql_text)
        if self._last_sql_text:
            self.copy_sql_button.state(["!disabled"])

        output = result.get("output")
        self._last_output = Path(output) if output else None
        if self._last_output is not None:
            self.output_var.set(str(self._last_output))
            self.open_button.state(["!disabled"])
        else:
            self.open_button.state(["disabled"])

        uncoded = result["uncoded"]
        ambiguous = result["ambiguous"]
        extra = ""
        if uncoded:
            preview = ", ".join(uncoded[:6])
            if len(uncoded) > 6:
                preview += f", và {len(uncoded) - 6} người nữa"
            extra += f"\nChưa có mã ({len(uncoded)}): {preview}"
        if ambiguous:
            extra += "\nCần chọn một mã: " + ", ".join(ambiguous)
        self.result_label.configure(
            text=(
                f"Tháng {result['month']}/{result['year']}\n"
                f"{result['doctors']} bác sĩ, {result['rows']} dòng ngày nghỉ\n"
                f"Cả ngày: {result['full']}    Nghỉ sáng: {result['morning']}    Nghỉ chiều: {result['afternoon']}\n"
                f"Đã gắn mã: {result['coded']}"
                f"{extra}"
            )
        )
        if self._last_output is not None:
            sql_name = Path(result["sql"]).name if result.get("sql") else "SQL trong app"
            self.status_var.set(f"Đã ghi {self._last_output.name} và {sql_name}")
        else:
            self.status_var.set("Đã tạo SQL (không lưu Excel). Có thể Copy SQL.")

    def _scan_done(self, result: dict) -> None:
        self._busy = False
        self.export_button.state(["!disabled"])
        self.scan_sql_button.state(["!disabled"])
        self._last_sql_text = str(result.get("sql_text") or "")
        self._set_sql_view(self._last_sql_text)
        if self._last_sql_text:
            self.copy_sql_button.state(["!disabled"])

        uncoded = result["uncoded"]
        ambiguous = result["ambiguous"]
        extra = ""
        if uncoded:
            preview = ", ".join(uncoded[:6])
            if len(uncoded) > 6:
                preview += f", và {len(uncoded) - 6} người nữa"
            extra += f"\nChưa có mã ({len(uncoded)}): {preview}"
        if ambiguous:
            extra += "\nCần chọn một mã: " + ", ".join(ambiguous)
        self.result_label.configure(
            text=(
                f"SQL quét y lệnh — Tháng {result['month']}/{result['year']}\n"
                f"{result['doctors']} bác sĩ, {result['rows']} dòng ngày nghỉ\n"
                f"Đã nhúng {result['coded_rows']} dòng có mã (JSON_TABLE, lọc HIS cả tháng)\n"
                f"Đã gắn mã: {result['coded']}"
                f"{extra}"
            )
        )
        self.status_var.set("Đã tạo SQL quét y lệnh từ file chấm công. Có thể Copy SQL.")

    def _export_failed(self, error: Exception) -> None:
        self._busy = False
        self.export_button.state(["!disabled"])
        self.scan_sql_button.state(["!disabled"])
        self.status_var.set("Xuất không thành công.")
        messagebox.showerror("DailyTool", str(error))

    def _scan_failed(self, error: Exception) -> None:
        self._busy = False
        self.export_button.state(["!disabled"])
        self.scan_sql_button.state(["!disabled"])
        self.status_var.set("Tạo SQL quét y lệnh không thành công.")
        messagebox.showerror("DailyTool", str(error))

    def open_codes(self) -> None:
        DoctorCodes(self, self.source_var.get().strip())

    def copy_sql(self) -> None:
        if not self._last_sql_text:
            messagebox.showwarning("DailyTool", "Chưa có SQL để copy.")
            return
        self.clipboard_clear()
        self.clipboard_append(self._last_sql_text)
        self.update_idletasks()
        self.status_var.set("Đã copy SQL vào clipboard.")

    def open_result(self) -> None:
        if self._last_output is None or not self._last_output.exists():
            messagebox.showwarning("DailyTool", "Chưa có file Excel.")
            return
        self._open_path(self._last_output)

    def _open_path(self, target: Path) -> None:
        path = str(target)
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
            return
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([opener, path])


def main() -> None:
    _enable_windows_dpi()
    app = DailyTool()
    app.mainloop()
