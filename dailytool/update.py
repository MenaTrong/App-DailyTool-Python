"""Kiểm tra và cài bản DailyTool mới từ GitHub Releases.

Người phát hành gắn file DailyTool.exe vào một Release. Ứng dụng đã đóng gói
tải file đó, thoát, rồi thay đúng file đang chạy và mở lại. Danh mục mã bác sĩ
nằm ngoài file exe nên không bị ghi đè.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from dailytool import GITHUB_REPO, __version__

ASSET_NAME = "DailyTool.exe"
_API = "https://api.github.com"


class UpdateError(Exception):
    """Không kiểm tra hoặc không cài được bản mới."""


@dataclass(frozen=True)
class Release:
    version: str
    notes: str
    page: str
    download_url: str
    api_url: str
    size: int


def version_key(value: str) -> tuple[int, ...]:
    core = value.strip().lstrip("vV").split("-", 1)[0].split("+", 1)[0]
    numbers: list[int] = []
    for part in core.split("."):
        digits = "".join(char for char in part if char.isdigit())
        numbers.append(int(digits) if digits else 0)
    return tuple(numbers) or (0,)


def is_newer(remote: str, current: str | None = None) -> bool:
    left = version_key(remote)
    right = version_key(current if current is not None else __version__)
    width = max(len(left), len(right))
    return left + (0,) * (width - len(left)) > right + (0,) * (width - len(right))


def installed_executable() -> Path | None:
    """File exe đang chạy. None khi mở bằng Python, không phải bản đã đóng gói."""
    if not getattr(sys, "frozen", False):
        return None
    return Path(sys.executable)


def github_token() -> str:
    env = os.environ.get("DAILYTOOL_GITHUB_TOKEN", "").strip()
    if env:
        return env
    path = Path(os.environ.get("APPDATA", str(Path.home()))) / "DailyTool" / "github.token"
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    return ""


def _update_dir() -> Path:
    path = Path(tempfile.gettempdir()) / "DailyTool-update"
    path.mkdir(parents=True, exist_ok=True)
    return path


def failure_log() -> Path:
    return _update_dir() / "update.log"


def consume_update_failure() -> str:
    path = failure_log()
    if not path.exists():
        return ""
    text = path.read_text(encoding="utf-8", errors="replace")
    path.unlink(missing_ok=True)
    if "FAILED" in text:
        return "Không thay được file DailyTool.exe. Hãy tải bản mới trên trang phát hành GitHub và chép đè file cũ khi ứng dụng đã đóng."
    return ""


def _request(url: str, accept: str) -> Request:
    headers = {"User-Agent": "DailyTool", "Accept": accept}
    token = github_token()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return Request(url, headers=headers)


def release_from_payload(payload: dict) -> Release | None:
    assets = payload.get("assets") or []
    asset = next(
        (item for item in assets if str(item.get("name", "")).casefold() == ASSET_NAME.casefold()),
        None,
    )
    if asset is None or not asset.get("browser_download_url"):
        return None
    tag = str(payload.get("tag_name") or payload.get("name") or "")
    return Release(
        version=tag.lstrip("vV") or tag,
        notes=str(payload.get("body") or "").strip(),
        page=str(payload.get("html_url") or ""),
        download_url=str(asset["browser_download_url"]),
        api_url=str(asset.get("url") or ""),
        size=int(asset.get("size") or 0),
    )


def fetch_latest() -> Release | None:
    """Bản mới hơn bản đang chạy, hoặc None nếu đang là bản mới nhất."""
    repo = GITHUB_REPO.strip().strip("/")
    if not repo or repo.count("/") != 1:
        raise UpdateError(
            "Chưa khai báo kho GitHub. Điền GITHUB_REPO trong dailytool/__init__.py, dạng chu-so-huu/ten-kho, rồi phát hành lại."
        )
    url = f"{_API}/repos/{repo}/releases/latest"
    try:
        with urlopen(_request(url, "application/vnd.github+json"), timeout=20) as response:
            payload = json.load(response)
    except HTTPError as error:
        if error.code == 404:
            raise UpdateError(
                "Không thấy bản phát hành. Hãy tạo Release có file DailyTool.exe. Nếu kho riêng tư, đặt token vào %APPDATA%\\DailyTool\\github.token."
            ) from error
        raise UpdateError(f"GitHub trả về lỗi {error.code}.") from error
    except (URLError, TimeoutError, json.JSONDecodeError) as error:
        raise UpdateError("Không kết nối được GitHub. Kiểm tra mạng rồi thử lại.") from error
    if not isinstance(payload, dict):
        raise UpdateError("GitHub trả về dữ liệu không đọc được.")
    release = release_from_payload(payload)
    if release is None:
        raise UpdateError("Bản phát hành mới nhất không kèm file DailyTool.exe.")
    if not is_newer(release.version):
        return None
    return release


def download(release: Release, on_progress=None) -> Path:
    """Tải DailyTool.exe vào thư mục tạm. on_progress(đã tải, tổng) nếu có tổng."""
    if github_token() and release.api_url:
        url = release.api_url
    else:
        url = release.download_url
    destination = _update_dir() / f"DailyTool-{release.version}.exe"
    partial = destination.with_suffix(".part")
    try:
        with urlopen(_request(url, "application/octet-stream"), timeout=180) as response:
            total = int(response.headers.get("Content-Length") or release.size or 0)
            written = 0
            with partial.open("wb") as handle:
                while True:
                    chunk = response.read(256 * 1024)
                    if not chunk:
                        break
                    handle.write(chunk)
                    written += len(chunk)
                    if on_progress is not None:
                        on_progress(written, total)
        blob = partial.read_bytes()[:2]
        if blob != b"MZ":
            raise UpdateError("File tải về không phải ứng dụng Windows.")
        if release.size and partial.stat().st_size != release.size:
            raise UpdateError("File tải về bị thiếu. Hãy thử lại.")
        partial.replace(destination)
    except UpdateError:
        partial.unlink(missing_ok=True)
        raise
    except (URLError, TimeoutError, OSError) as error:
        partial.unlink(missing_ok=True)
        raise UpdateError("Không tải được bản mới. Kiểm tra mạng rồi thử lại.") from error
    return destination


def install_downloaded(source: Path, target: Path) -> None:
    """Thoát ứng dụng sau khi gọi hàm này. Script ngoài sẽ chép đè file exe."""
    if sys.platform != "win32":
        raise UpdateError("Tự thay file chỉ chạy trên Windows.")
    script = _update_dir() / f"apply-{os.getpid()}.ps1"
    log = failure_log()
    script.write_bytes(b"\xef\xbb\xbf" + _swap_script(source, target, os.getpid(), log).encode("utf-8"))
    command = [
        "cmd",
        "/c",
        "start",
        "",
        "/min",
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        str(script),
    ]
    subprocess.Popen(
        command,
        close_fds=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _swap_script(source: Path, target: Path, process_id: int, log: Path) -> str:
    return f"""$ErrorActionPreference = 'Stop'
$target = {_ps_quote(str(target))}
$source = {_ps_quote(str(source))}
$log = {_ps_quote(str(log))}
$processId = {int(process_id)}
for ($i = 0; $i -lt 60; $i++) {{
    if (-not (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {{ break }}
    Start-Sleep -Seconds 1
}}
$copied = $false
for ($i = 0; $i -lt 30; $i++) {{
    try {{
        Copy-Item -LiteralPath $source -Destination $target -Force
        $copied = $true
        break
    }} catch {{
        Start-Sleep -Seconds 1
    }}
}}
if (-not $copied) {{
    Set-Content -LiteralPath $log -Value 'FAILED' -Encoding utf8
    exit 1
}}
Start-Process -FilePath $target
Remove-Item -LiteralPath $source -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
"""


def short_notes(notes: str, limit: int = 400) -> str:
    text = " ".join(notes.split())
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "…"
