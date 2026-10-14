#!/usr/bin/env python3
"""Xuất danh sách ngày nghỉ bác sĩ từ file chấm công.

Đọc các sheet khoa trong file chấm công tháng, lấy bác sĩ, và ghi một file
giống mẫu «Ngày Nghỉ BS»: mỗi dòng là một ngày nghỉ. Cột Mã Bác Sĩ
ghép theo họ tên từ danh mục mã.

Quy ước:
- Dấu + hoặc T (cả khối Tăng cường) là đang làm. Ngày đó không đưa vào danh sách.
- Ô có hai dòng: dòng trên là buổi sáng, dòng dưới là buổi chiều.
- Chỉ nghỉ một buổi thì cột Nghỉ sáng hoặc Nghỉ chiều ghi X.
- Cả ngày không có +/T thì vẫn là một dòng ngày nghỉ, hai cột buổi để trống.
- Khối Ngoài giờ (số giờ) không tính.

Cách dùng:
    python3 xuat_ngay_nghi.py "file-cham-cong.xlsx"
    python3 xuat_ngay_nghi.py "file-cham-cong.xlsx" -o "Ngày Nghỉ BS T7.xlsx"
"""

from __future__ import annotations

import argparse
import calendar
import json
import re
import sys
import unicodedata
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
BS_ROLE = re.compile(r"(^|[^A-Z])BS([^A-Z]|$)")
ROMAN = {"I", "II", "III", "IV", "V"}
SECTION_NAMES = {"CHÍNH THỨC", "THỬ VIỆC", "BÁN THỜI GIAN", "TĂNG CƯỜNG", "NGOÀI GIỜ"}
APOSTROPHES = "'’ʻʼ`´′‛"
SKIP_SHEETS = {"CÁC NGÀY LẼ"}


def norm_name(value: str) -> str:
    text = unicodedata.normalize("NFC", str(value)).strip().lower()
    for char in APOSTROPHES:
        text = text.replace(char, "'")
    text = text.replace("\n", " ")
    return re.sub(r"\s+", " ", text)


def is_doctor(role: str) -> bool:
    compact = role.replace("\n", " ").upper()
    return BS_ROLE.search(compact) is not None


def is_work_line(text: str) -> bool:
    cleaned = text.replace("\u00a0", "").strip()
    if not cleaned:
        return False
    if "+" in cleaned:
        return True
    token = cleaned.upper().replace(" ", "")
    return token in {"T", "Т"}


def session_work(raw) -> tuple[bool, bool] | None:
    """Trả về (nghỉ? không) buổi sáng, buổi chiều đang làm.

    None nghĩa là ô không mang ký hiệu chấm công (trống hoặc chỉ là số giờ).
    """
    if raw is None or isinstance(raw, (int, float)):
        return None
    text = str(raw).replace("\u00a0", " ")
    if not text.strip():
        return None
    lines = text.splitlines() or [text]
    if len(lines) == 1:
        working = is_work_line(lines[0])
        return (working, working)
    return (is_work_line(lines[0]), is_work_line(lines[1]))


def find_month_year(rows: list[tuple]) -> tuple[int | None, int | None]:
    month = year = None
    for row in rows[:8]:
        for index, value in enumerate(row):
            if not isinstance(value, str):
                continue
            label = value.strip().lower()
            if label not in {"tháng", "năm"}:
                continue
            for nxt in row[index + 1 : index + 8]:
                if isinstance(nxt, (int, float)) and nxt:
                    number = int(nxt)
                    if label == "tháng" and month is None:
                        month = number
                    elif label == "năm" and year is None:
                        year = number
                    break
    return month, year


def find_day_columns(rows: list[tuple]) -> dict[int, int] | None:
    for row in rows[:15]:
        mapping: dict[int, int] = {}
        for index, value in enumerate(row):
            if isinstance(value, (int, float)) and float(value).is_integer():
                day = int(value)
                if 1 <= day <= 31:
                    mapping[day] = index
        if 1 in mapping and 2 in mapping and len(mapping) >= 28:
            return mapping
    return None


def parse_attendance(path: Path) -> tuple[int, int, list[dict]]:
    workbook = load_workbook(path, read_only=True, data_only=False)
    month = year = None
    people: list[dict] = []
    try:
        for sheet_name in workbook.sheetnames:
            if sheet_name in SKIP_SHEETS or sheet_name.upper().startswith("TỔNG"):
                continue
            sheet = workbook[sheet_name]
            preview = []
            for row in sheet.iter_rows(min_row=1, max_row=12, max_col=40, values_only=True):
                preview.append(tuple(row))
            if not any(isinstance(cell, str) and cell.strip() == "Họ và tên" for row in preview for cell in row):
                continue
            day_columns = find_day_columns(preview)
            if not day_columns:
                continue
            found_month, found_year = find_month_year(preview)
            month = month or found_month
            year = year or found_year

            section = ""
            grouped: dict[str, dict] = {}
            order: list[str] = []
            for row in sheet.iter_rows(min_row=9, max_col=40, values_only=True):
                values = tuple(row)
                marker = "" if values[0] is None else str(values[0]).replace("\n", " ").strip()
                name = "" if len(values) < 2 or values[1] is None else str(values[1]).strip()
                role = "" if len(values) < 3 or values[2] is None else str(values[2]).replace("\n", " ").strip()
                identity = "" if len(values) < 4 or values[3] is None else str(values[3]).strip()
                name_key = name.upper()
                if marker in ROMAN or name_key in SECTION_NAMES:
                    section = marker if marker in ROMAN else name_key
                    continue
                if not name or not role or not is_doctor(role):
                    continue
                if section == "V" or "NGOÀI GIỜ" in section:
                    continue
                key = identity or norm_name(name)
                if key not in grouped:
                    grouped[key] = {"name": name, "role": role, "sheet": sheet_name, "days": {}}
                    order.append(key)
                person = grouped[key]
                for day, column in day_columns.items():
                    if column >= len(values):
                        continue
                    raw = values[column]
                    if raw is None or (isinstance(raw, str) and not raw.strip()):
                        continue
                    person["days"].setdefault(day, []).append(raw)
            people.extend(grouped[key] for key in order)
    finally:
        workbook.close()
    if month is None or year is None:
        raise ValueError("Không đọc được Tháng/Năm trong file chấm công.")
    return month, year, people


def off_mark(cells: list) -> tuple[str, str] | None:
    """None nếu cả hai buổi đều có + hoặc T. Ngược lại ('X' hoặc '', 'X' hoặc '')."""
    morning = afternoon = False
    for raw in cells:
        flags = session_work(raw)
        if flags is None:
            continue
        morning = morning or flags[0]
        afternoon = afternoon or flags[1]
    if morning and afternoon:
        return None
    if morning and not afternoon:
        return ("", "X")
    if afternoon and not morning:
        return ("X", "")
    return ("", "")


def build_rows(month: int, year: int, people: list[dict]) -> list[dict]:
    last_day = calendar.monthrange(year, month)[1]
    rows = []
    for person in people:
        for day in range(1, last_day + 1):
            mark = off_mark(person["days"].get(day, []))
            if mark is None:
                continue
            rows.append(
                {
                    "code": person.get("code") or "",
                    "name": person["name"],
                    "year": year,
                    "month": month,
                    "day": day,
                    "morning": mark[0],
                    "afternoon": mark[1],
                }
            )
    return rows


def write_workbook(path: Path, rows: list[dict]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ngày nghỉ"

    headers = ["Mã Bác Sĩ", "Tên Bác Sĩ", "Năm nghỉ", "Tháng nghỉ", "Ngày Nghỉ", "Nghỉ sáng", "Nghỉ chiều"]
    header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1B4332")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )
    widths = {"A": 18, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12, "G": 14}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for column, title in enumerate(headers, 1):
        cell = sheet.cell(1, column, title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    sheet.row_dimensions[1].height = 22
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:G{max(1, len(rows) + 1)}"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.sheet_properties.tabColor = "9F1239"
    sheet.oddFooter.center.text = "Ngày nghỉ bác sĩ"

    alt = PatternFill("solid", fgColor="F8FAFC")
    for index, item in enumerate(rows, 2):
        fill = alt if index % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        values = [
            item["code"],
            item["name"],
            item["year"],
            item["month"],
            item["day"],
            item["morning"],
            item["afternoon"],
        ]
        for column, value in enumerate(values, 1):
            cell = sheet.cell(index, column, value if value != "" else None)
            cell.border = thin
            cell.fill = fill
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = left if column in (1, 2) else center
        for column in (6, 7):
            if sheet.cell(index, column).value == "X":
                sheet.cell(index, column).font = Font(name="Calibri", size=11, bold=True, color="9F1239")

    for column in range(3, 6):
        for row in range(2, len(rows) + 2):
            sheet.cell(row, column).number_format = "0"

    workbook.save(path)


def _sql_text(value: str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def insert_statement(row: dict) -> str:
    """Một dòng Oracle đúng cấu trúc bvxa_lichnghi."""
    return (
        "insert into bvxa_lichnghi(mabs,nam,thang,ngay,sang,chieu) values ("
        f"{_sql_text(row['code'])},{int(row['year'])},{int(row['month'])},{int(row['day'])},"
        f"{_sql_text(row['morning'])},{_sql_text(row['afternoon'])});"
    )


def build_oracle_script(rows: list[dict]) -> str:
    lines = ["set define off;"]
    lines.extend(insert_statement(row) for row in rows)
    lines.append("commit;")
    return "\n".join(lines) + "\n"


def write_oracle_script(path: Path, rows: list[dict]) -> None:
    path.write_text(build_oracle_script(rows), encoding="utf-8")


def _oracle_clob_literal(text: str, chunk: int = 3000) -> str:
    """Ghép chuỗi dài thành TO_CLOB(...) || ... để vượt giới hạn literal 4000."""
    if not text:
        return "TO_CLOB('')"
    parts = []
    for index in range(0, len(text), chunk):
        piece = text[index : index + chunk].replace("'", "''")
        parts.append(f"TO_CLOB('{piece}')")
    return " ||\n    ".join(parts)


def build_leave_order_scan_sql(rows: list[dict]) -> str:
    """SQL SELECT: ngày nghỉ có mã BS mà vẫn có chỉ định y lệnh HIS.

    Không cần INSERT bvxa_lichnghi: nhúng lịch nghỉ từ file chấm công bằng JSON_TABLE,
    lọc HIS theo cả tháng rồi khớp khóa ngày (nhanh hơn UNION ALL + BETWEEN từng ngày).
    Cần Oracle 12c+ (JSON_TABLE).
    """
    coded = [row for row in rows if str(row.get("code") or "").strip()]
    if not coded:
        raise ValueError("Không có dòng ngày nghỉ nào đã gắn mã bác sĩ để quét y lệnh.")
    year = int(coded[0]["year"])
    month = int(coded[0]["month"])
    last_day = calendar.monthrange(year, month)[1]
    month_start = year * 10000000000 + month * 100000000 + 1 * 1000000
    month_end = year * 10000000000 + month * 100000000 + last_day * 1000000 + 235959
    payload = json.dumps(
        [
            {
                "c": str(row["code"]).strip(),
                "y": int(row["year"]),
                "m": int(row["month"]),
                "d": int(row["day"]),
                "s": str(row.get("morning") or ""),
                "e": str(row.get("afternoon") or ""),
            }
            for row in coded
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    oracle_json = _oracle_clob_literal(payload)
    return f"""-- Quét y lệnh trùng ngày nghỉ (tháng {month}/{year})
-- Không cần INSERT bvxa_lichnghi: lịch nghỉ lấy từ file chấm công (JSON_TABLE).
-- Buổi sáng: 00:00–12:59 | Buổi chiều: từ 13:00 | Cả ngày (hai cột trống hoặc cả hai X): cả ngày.
-- Gợi ý index HIS: (REQUEST_LOGINNAME, INTRUCTION_TIME) hoặc (INTRUCTION_TIME).

WITH leave_days AS (
  SELECT
      jt.mabs,
      jt.nam,
      jt.thang,
      jt.ngay,
      jt.sang,
      jt.chieu,
      (jt.nam * 10000000000 + jt.thang * 100000000 + jt.ngay * 1000000)
        + CASE
            WHEN NVL(TRIM(jt.sang), ' ') = 'X' AND NVL(TRIM(jt.chieu), ' ') <> 'X' THEN 0
            WHEN NVL(TRIM(jt.chieu), ' ') = 'X' AND NVL(TRIM(jt.sang), ' ') <> 'X' THEN 130000
            ELSE 0
          END AS t_from,
      (jt.nam * 10000000000 + jt.thang * 100000000 + jt.ngay * 1000000)
        + CASE
            WHEN NVL(TRIM(jt.sang), ' ') = 'X' AND NVL(TRIM(jt.chieu), ' ') <> 'X' THEN 125959
            WHEN NVL(TRIM(jt.chieu), ' ') = 'X' AND NVL(TRIM(jt.sang), ' ') <> 'X' THEN 235959
            ELSE 235959
          END AS t_to
  FROM JSON_TABLE(
    {oracle_json},
    '$[*]' COLUMNS (
      mabs  VARCHAR2(64)  PATH '$.c',
      nam   NUMBER        PATH '$.y',
      thang NUMBER        PATH '$.m',
      ngay  NUMBER        PATH '$.d',
      sang  VARCHAR2(8)   PATH '$.s',
      chieu VARCHAR2(8)   PATH '$.e'
    )
  ) jt
),
per_order AS (
  SELECT
      ln.mabs,
      ln.nam,
      ln.thang,
      ln.ngay,
      ln.sang,
      ln.chieu,
      sr.SERVICE_REQ_CODE AS ma_y_lenh,
      RTRIM(
        XMLAGG(
          XMLELEMENT(e, ss.TDL_SERVICE_NAME || ', ')
          ORDER BY ss.TDL_SERVICE_NAME
        ).EXTRACT('//text()').GetClobVal(),
        ', '
      ) AS ten_y_lenh
  FROM leave_days ln
  JOIN HIS_SERVICE_REQ sr
    ON sr.REQUEST_LOGINNAME = ln.mabs
   AND sr.IS_DELETE = 0
   AND sr.INTRUCTION_TIME BETWEEN {month_start} AND {month_end}
   AND sr.INTRUCTION_TIME BETWEEN ln.t_from AND ln.t_to
  JOIN HIS_SERE_SERV ss
    ON ss.SERVICE_REQ_ID = sr.ID
   AND ss.PATIENT_TYPE_ID = 1
  GROUP BY
      ln.mabs,
      ln.nam,
      ln.thang,
      ln.ngay,
      ln.sang,
      ln.chieu,
      sr.SERVICE_REQ_CODE
),
orders_by_leave AS (
  SELECT
      po.mabs,
      po.nam,
      po.thang,
      po.ngay,
      po.sang,
      po.chieu,
      SUBSTR(
        XMLAGG(
          XMLELEMENT(e, CHR(10) || po.ma_y_lenh || ' | ' || po.ten_y_lenh)
          ORDER BY po.ma_y_lenh
        ).EXTRACT('//text()').GetClobVal(),
        2
      ) AS y_lenh
  FROM per_order po
  GROUP BY
      po.mabs,
      po.nam,
      po.thang,
      po.ngay,
      po.sang,
      po.chieu
)
SELECT
    ln.mabs AS ma_bac_si,
    ln.nam,
    ln.thang,
    ln.ngay,
    ln.sang,
    ln.chieu,
    obl.y_lenh
FROM leave_days ln
LEFT JOIN orders_by_leave obl
  ON obl.mabs = ln.mabs
 AND obl.nam = ln.nam
 AND obl.thang = ln.thang
 AND obl.ngay = ln.ngay
 AND NVL(obl.sang, ' ') = NVL(ln.sang, ' ')
 AND NVL(obl.chieu, ' ') = NVL(ln.chieu, ' ')
ORDER BY
    ln.mabs,
    ln.nam,
    ln.thang,
    ln.ngay;
"""


def build_leave_order_scan_from_file(source: Path) -> dict:
    """Đọc chấm công, gắn mã, trả sql_text quét y lệnh ngày nghỉ + số liệu."""
    from dailytool.doctors import apply_to_people

    source = source.expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Không thấy file: {source}")
    month, year, people = parse_attendance(source)
    codes = apply_to_people(people)
    rows = build_rows(month, year, people)
    sql_text = build_leave_order_scan_sql(rows)
    coded_rows = sum(1 for row in rows if str(row.get("code") or "").strip())
    return {
        "sql_text": sql_text,
        "month": month,
        "year": year,
        "doctors": len(people),
        "rows": len(rows),
        "coded_rows": coded_rows,
        "coded": codes["coded"],
        "uncoded": codes["uncoded"],
        "ambiguous": codes["ambiguous"],
    }


def export_leave_workbook(
    source: Path,
    output: Path | None = None,
    *,
    write_excel: bool = True,
) -> dict:
    """Đọc file chấm công và xuất ngày nghỉ. Trả về số liệu + sql_text.

    Khi write_excel=True: ghi .xlsx và .sql cạnh đó.
    Khi write_excel=False: chỉ trả sql_text, không ghi file.
    """
    from dailytool.doctors import apply_to_people

    source = source.expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Không thấy file: {source}")
    month, year, people = parse_attendance(source)
    codes = apply_to_people(people)
    rows = build_rows(month, year, people)
    sql_text = build_oracle_script(rows)

    output_path: Path | None = None
    sql_path: Path | None = None
    if write_excel:
        if output is None:
            output_path = source.parent / f"Ngày Nghỉ BS T{month}.xlsx"
        else:
            output_path = output.expanduser().resolve()
        write_workbook(output_path, rows)
        sql_path = output_path.with_suffix(".sql")
        write_oracle_script(sql_path, rows)

    morning = sum(1 for row in rows if row["morning"] == "X")
    afternoon = sum(1 for row in rows if row["afternoon"] == "X")
    full = sum(1 for row in rows if not row["morning"] and not row["afternoon"])
    return {
        "output": output_path,
        "sql": sql_path,
        "sql_text": sql_text,
        "month": month,
        "year": year,
        "doctors": len(people),
        "rows": len(rows),
        "full": full,
        "morning": morning,
        "afternoon": afternoon,
        "coded": codes["coded"],
        "uncoded": codes["uncoded"],
        "ambiguous": codes["ambiguous"],
        "codes_added": codes["added"],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Xuất ngày nghỉ bác sĩ từ file chấm công.")
    parser.add_argument("cham_cong", type=Path, help="File Excel chấm công tháng")
    parser.add_argument("-o", "--output", type=Path, help="File Excel kết quả")
    args = parser.parse_args(argv)

    try:
        result = export_leave_workbook(args.cham_cong, args.output)
    except (FileNotFoundError, ValueError) as error:
        print(error, file=sys.stderr)
        return 1
    print(f"Tháng {result['month']}/{result['year']}: {result['doctors']} bác sĩ, {result['rows']} dòng ngày nghỉ")
    print(f"  Cả ngày: {result['full']}    Nghỉ sáng: {result['morning']}    Nghỉ chiều: {result['afternoon']}")
    print(f"  Đã gắn mã: {result['coded']}    Chưa có mã: {len(result['uncoded'])}    Trùng mã: {len(result['ambiguous'])}")
    print(f"Đã ghi {result['output']}")
    print(f"Đã ghi {result['sql']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
