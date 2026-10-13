"""Danh mục mã bác sĩ cho tác vụ xuất ngày nghỉ.

Mã được ghép theo họ tên khi xuất file. Danh mục thêm, sửa, xóa được;
bác sĩ mới trong file chấm công được bổ sung với mã trống.
"""

from __future__ import annotations

import json
import os
import sys
import threading
import uuid
from pathlib import Path

from xuat_ngay_nghi import norm_name, parse_attendance

_LOCK = threading.Lock()

# Nhãn khoa trên sheet chấm công -> hậu tố mã (hoapt_kkb, nghiath_cc, ...).
# Chỉ dùng khi một người có nhiều mã, để chọn mã đúng khoa.
_SHEET_SUFFIX = {
    "KB": "kkb",
    "CC": "cc",
    "NHH": "nhh",
    "NTQ": "ntq",
    "NT-LMNT": "ntlm",
    "NTK": "ntk",
    "NTM": "ntm",
    "NNT": "nnt",
    "NGTQ": "ngtq",
    "CTCH": "ctch",
    "NGTN": "ngtn",
    "NGTK": "ngtk",
    "UB": "ub",
    "SAN": "san",
    "NHI": "nhi",
    "ICU": "icu",
    "GMHS": "gmhs",
    "LCK": "lck",
    "MẮT": "mat",
    "YHCT": "yhct",
    "CĐHA": "cdha",
    "NỘI SOI": "ns",
    "XN": "xn",
    "DƯỢC": "duoc",
}


def store_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("APPDATA", str(Path.home()))) / "DailyTool"
    else:
        base = Path(__file__).resolve().parent.parent / "data"
    base.mkdir(parents=True, exist_ok=True)
    return base / "ma_bac_si.json"


def _seed_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", "")) / "dailytool" / "ma_bac_si_seed.json"
    return Path(__file__).resolve().parent / "ma_bac_si_seed.json"


def _clean(value: str) -> str:
    return " ".join(str(value).replace("\n", " ").split())


def _blank(doctor: dict) -> dict:
    return {
        "id": str(doctor.get("id") or uuid.uuid4().hex[:12]),
        "name": _clean(doctor.get("name") or ""),
        "code": _clean(doctor.get("code") or ""),
    }


def _read_json(path: Path) -> list[dict]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    rows = raw["doctors"] if isinstance(raw, dict) else raw
    if not isinstance(rows, list):
        raise ValueError("Danh mục mã bác sĩ không đúng định dạng.")
    doctors = []
    for item in rows:
        if not isinstance(item, dict):
            continue
        doctor = _blank(item)
        if doctor["name"]:
            doctors.append(doctor)
    return doctors


def _load() -> list[dict]:
    path = store_path()
    if path.exists():
        return _read_json(path)
    seed = _seed_path()
    if seed.exists():
        doctors = _read_json(seed)
        _save(doctors)
        return doctors
    return []


def _save(doctors: list[dict]) -> None:
    path = store_path()
    payload = json.dumps({"doctors": doctors}, ensure_ascii=False, indent=2)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(payload + "\n", encoding="utf-8")
    temporary.replace(path)


def _same(left: dict, right_name: str, right_code: str) -> bool:
    return norm_name(left["name"]) == norm_name(right_name) and left["code"].casefold() == right_code.casefold()


def _require(name: str, code: str) -> tuple[str, str]:
    name = _clean(name)
    code = _clean(code)
    if not name:
        raise ValueError("Hãy nhập họ tên bác sĩ.")
    return name, code


def list_doctors() -> list[dict]:
    with _LOCK:
        doctors = _load()
    return sorted(doctors, key=lambda item: (0 if not item["code"] else 1, norm_name(item["name"])))


def summary() -> dict:
    doctors = list_doctors()
    with_code = sum(1 for item in doctors if item["code"])
    return {"total": len(doctors), "with_code": with_code, "without_code": len(doctors) - with_code}


def add_doctor(name: str, code: str) -> dict:
    name, code = _require(name, code)
    with _LOCK:
        doctors = _load()
        if any(_same(item, name, code) for item in doctors):
            raise ValueError("Bác sĩ này đã có mã đó.")
        doctor = {"id": uuid.uuid4().hex[:12], "name": name, "code": code}
        doctors.append(doctor)
        _save(doctors)
    return doctor


def update_doctor(doctor_id: str, name: str, code: str) -> dict:
    name, code = _require(name, code)
    with _LOCK:
        doctors = _load()
        current = next((item for item in doctors if item["id"] == doctor_id), None)
        if current is None:
            raise ValueError("Không thấy mã bác sĩ cần sửa.")
        if any(item["id"] != doctor_id and _same(item, name, code) for item in doctors):
            raise ValueError("Bác sĩ này đã có mã đó.")
        current["name"] = name
        current["code"] = code
        _save(doctors)
    return current


def delete_doctor(doctor_id: str) -> None:
    with _LOCK:
        doctors = _load()
        kept = [item for item in doctors if item["id"] != doctor_id]
        if len(kept) == len(doctors):
            raise ValueError("Không thấy mã bác sĩ cần xóa.")
        _save(kept)


def sheet_label(sheet: str) -> str:
    labels = []
    for part in str(sheet).split("."):
        part = part.strip()
        if not part or part.isdigit():
            continue
        labels.append(part)
    return " ".join(labels).upper()


def _code_suffix(code: str) -> str:
    for separator in ("_", "-"):
        if separator in code:
            return code.rsplit(separator, 1)[-1].casefold()
    return ""


def resolve_code(doctors: list[dict], name: str, sheet: str = "") -> tuple[str, str]:
    """Trả về (mã, trạng thái). Trạng thái là ok, missing hoặc ambiguous."""
    matches = [
        item for item in doctors if norm_name(item["name"]) == norm_name(name) and item["code"]
    ]
    if not matches:
        return "", "missing"
    unique = {item["code"] for item in matches}
    if len(unique) == 1:
        return matches[0]["code"], "ok"
    suffix = _SHEET_SUFFIX.get(sheet_label(sheet), "")
    if suffix:
        narrowed = [item for item in matches if _code_suffix(item["code"]) == suffix]
        narrowed_codes = {item["code"] for item in narrowed}
        if len(narrowed_codes) == 1:
            return narrowed[0]["code"], "ok"
    return "", "ambiguous"


def _ensure_people(doctors: list[dict], people: list[dict]) -> int:
    known = {norm_name(item["name"]) for item in doctors}
    added = 0
    for person in people:
        name = _clean(person.get("name") or "")
        if not name:
            continue
        key = norm_name(name)
        if key in known:
            continue
        doctors.append({"id": uuid.uuid4().hex[:12], "name": name, "code": ""})
        known.add(key)
        added += 1
    return added


def apply_to_people(people: list[dict]) -> dict:
    """Thêm bác sĩ mới vào danh mục và gán mã cho từng người."""
    with _LOCK:
        doctors = _load()
        added = _ensure_people(doctors, people)
        if added:
            _save(doctors)
        coded_names: list[str] = []
        uncoded: list[str] = []
        ambiguous: list[str] = []
        for person in people:
            code, status = resolve_code(doctors, person["name"], person.get("sheet", ""))
            person["code"] = code
            if status == "ok":
                coded_names.append(person["name"])
            elif status == "ambiguous":
                ambiguous.append(person["name"])
            else:
                uncoded.append(person["name"])
    return {
        "added": added,
        "coded": len(coded_names),
        "uncoded": uncoded,
        "ambiguous": ambiguous,
    }


def import_attendance(path: Path) -> dict:
    """Đọc bác sĩ trong file chấm công và thêm người chưa có trong danh mục."""
    _month, _year, people = parse_attendance(path)
    with _LOCK:
        doctors = _load()
        added_names = []
        known = {norm_name(item["name"]) for item in doctors}
        for person in people:
            name = _clean(person.get("name") or "")
            key = norm_name(name)
            if not name or key in known:
                continue
            doctors.append({"id": uuid.uuid4().hex[:12], "name": name, "code": ""})
            known.add(key)
            added_names.append(name)
        if added_names:
            _save(doctors)
    return {"added": added_names, "doctors": len(people)}
