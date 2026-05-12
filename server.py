from __future__ import annotations

from email import policy
from email.parser import BytesParser
import json
import mimetypes
import os
import platform
import subprocess
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote

from formatter import format_docx, format_options_from_mapping


def app_root() -> Path:
    if getattr(sys, "_MEIPASS", None):
        return Path(sys._MEIPASS)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_public_dir(root: Path) -> Path:
    candidates = [
        root / "public",
        Path(sys.executable).resolve().parent / "public" if getattr(sys, "frozen", False) else None,
        Path.cwd() / "public",
        Path(__file__).resolve().parent / "public",
    ]
    for candidate in candidates:
        if candidate and (candidate / "index.html").exists():
            return candidate
    return root / "public"


ROOT = app_root()
PUBLIC = resolve_public_dir(ROOT)
if getattr(sys, "frozen", False):
    STORAGE = Path.home() / "Documents" / "公文格式助手输出"
else:
    STORAGE = ROOT / "storage"
UPLOADS = STORAGE / "uploads"
OUTPUTS = STORAGE / "outputs"
FONT_CACHE: list[str] | None = None


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path.startswith("/download/"):
            self.serve_download()
            return

        path = self.path.split("?", 1)[0]
        if path == "/api/fonts":
            self.respond_json({"fonts": system_font_names()})
            return

        if path == "/":
            path = "/index.html"
        self.serve_static(PUBLIC / path.lstrip("/"))

    def do_POST(self) -> None:
        if self.path != "/api/format":
            self.send_error(404)
            return

        try:
            upload = self.parse_upload()
            if upload is None:
                self.respond_json({"error": "请先选择一个文稿文件"}, status=400)
                return

            UPLOADS.mkdir(parents=True, exist_ok=True)
            safe_name, file_bytes, fields = upload
            upload_path = UPLOADS / safe_name
            with upload_path.open("wb") as f:
                f.write(file_bytes)

            options = format_options_from_mapping(fields)
            result = format_docx(upload_path, safe_name, OUTPUTS, options)
            self.respond_json(
                {
                    "downloadUrl": f"/download/{result.output_path.name}",
                    "fileName": result.output_path.name,
                    "report": [
                        {"index": item.index, "text": item.text, "role": item.role}
                        for item in result.report
                    ],
                    "warnings": [
                        {"level": item.level, "message": item.message}
                        for item in result.warnings
                    ],
                }
            )
        except Exception as exc:
            self.respond_json({"error": str(exc)}, status=500)

    def parse_upload(self) -> tuple[str, bytes, dict[str, str]] | None:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        if not content_type.startswith("multipart/form-data") or content_length <= 0:
            return None

        body = self.rfile.read(content_length)
        message_bytes = (
            f"Content-Type: {content_type}\r\n"
            "MIME-Version: 1.0\r\n\r\n"
        ).encode("utf-8") + body
        message = BytesParser(policy=policy.default).parsebytes(message_bytes)

        fields: dict[str, str] = {}
        upload: tuple[str, bytes] | None = None
        for part in message.iter_parts():
            field_name = part.get_param("name", header="content-disposition")
            filename = part.get_filename()
            if not field_name:
                continue
            payload = part.get_payload(decode=True) or b""
            if field_name == "file" and filename:
                if payload:
                    upload = (Path(filename).name, payload)
                continue
            if filename:
                continue
            charset = part.get_content_charset() or "utf-8"
            fields[field_name] = payload.decode(charset, errors="replace").strip()
        if upload is None:
            return None
        return upload[0], upload[1], fields

    def do_HEAD(self) -> None:
        if self.path.startswith("/download/"):
            self.serve_download(head_only=True)
            return
        self.send_error(404)

    def serve_static(self, path: Path) -> None:
        if not is_public_file(path):
            self.send_error(404)
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.end_headers()
        self.wfile.write(path.read_bytes())

    def serve_download(self, head_only: bool = False) -> None:
        name = Path(unquote(self.path.removeprefix("/download/"))).name
        path = OUTPUTS / name
        if not path.exists() or not path.is_file():
            self.send_error(404)
            return
        encoded_name = quote(name)
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.send_header(
            "Content-Disposition",
            f"attachment; filename=formatted.docx; filename*=UTF-8''{encoded_name}",
        )
        self.send_header("Content-Length", str(path.stat().st_size))
        self.end_headers()
        if not head_only:
            self.wfile.write(path.read_bytes())

    def respond_json(self, data: dict, status: int = 200) -> None:
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"公文格式助手已启动：http://127.0.0.1:{port}")
    server.serve_forever()


def is_public_file(path: Path) -> bool:
    try:
        resolved_path = path.resolve()
        resolved_public = PUBLIC.resolve()
    except OSError:
        return False
    return resolved_path.is_file() and (
        resolved_path == resolved_public or resolved_public in resolved_path.parents
    )


def system_font_names() -> list[str]:
    global FONT_CACHE
    if FONT_CACHE is not None:
        return FONT_CACHE

    defaults = {
        "Times New Roman",
        "Arial",
        "Calibri",
        "宋体",
        "仿宋",
        "黑体",
        "楷体",
        "微软雅黑",
        "方正公文小标宋",
        "方正公文仿宋",
        "方正公文黑体",
        "方正公文楷体",
    }
    names = set(defaults)
    system = platform.system()

    if system == "Darwin":
        names.update(macos_font_names())
    elif system == "Windows":
        names.update(windows_font_names())
    else:
        names.update(linux_font_names())

    FONT_CACHE = sorted(name for name in names if name)
    return FONT_CACHE


def macos_font_names() -> set[str]:
    names: set[str] = set()
    try:
        result = subprocess.run(
            ["system_profiler", "SPFontsDataType", "-json"],
            capture_output=True,
            text=True,
            timeout=10,
            check=True,
        )
        payload = json.loads(result.stdout)
        collect_font_names(payload, names)
    except Exception:
        pass

    scan_font_dirs(
        names,
        [
            Path("/System/Library/Fonts"),
            Path("/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
        ],
    )
    return names


def windows_font_names() -> set[str]:
    names: set[str] = set()
    try:
        import winreg

        registry_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ]
        for root, path in registry_paths:
            try:
                with winreg.OpenKey(root, path) as key:
                    for index in range(winreg.QueryInfoKey(key)[1]):
                        raw_name = winreg.EnumValue(key, index)[0]
                        names.add(clean_font_display_name(raw_name))
            except OSError:
                continue
    except Exception:
        pass

    windir = os.environ.get("WINDIR")
    if windir:
        scan_font_dirs(names, [Path(windir) / "Fonts"])
    return names


def linux_font_names() -> set[str]:
    names: set[str] = set()
    try:
        result = subprocess.run(
            ["fc-list", ":", "family"],
            capture_output=True,
            text=True,
            timeout=5,
            check=True,
        )
        for line in result.stdout.splitlines():
            for name in line.split(","):
                names.add(name.strip())
    except Exception:
        pass
    scan_font_dirs(names, [Path("/usr/share/fonts"), Path.home() / ".fonts"])
    return names


def collect_font_names(value: object, names: set[str]) -> None:
    if isinstance(value, dict):
        for key, item in value.items():
            if key in {"_name", "family", "familyName", "name"} and isinstance(item, str):
                names.add(clean_font_display_name(item))
            collect_font_names(item, names)
    elif isinstance(value, list):
        for item in value:
            collect_font_names(item, names)


def scan_font_dirs(names: set[str], directories: list[Path]) -> None:
    for directory in directories:
        if not directory.exists():
            continue
        for path in directory.rglob("*"):
            if path.suffix.lower() in {".ttf", ".ttc", ".otf"}:
                names.add(clean_font_display_name(path.stem))


def clean_font_display_name(name: str) -> str:
    name = name.strip()
    name = name.replace(" Regular", "")
    name = name.replace(" Bold", "")
    name = name.replace(" Italic", "")
    name = name.replace(" Oblique", "")
    name = name.replace(" (TrueType)", "")
    name = name.replace(" (OpenType)", "")
    return name.strip()


if __name__ == "__main__":
    main()
