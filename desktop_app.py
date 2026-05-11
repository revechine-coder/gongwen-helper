from __future__ import annotations

import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import ThreadingHTTPServer
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from formatter import FormatResult, SUPPORTED_INPUT_EXTENSIONS, format_docx
from server import Handler


APP_DIR = Path(__file__).resolve().parent
if getattr(sys, "frozen", False):
    OUTPUT_DIR = Path.home() / "Documents" / "公文格式助手输出"
else:
    OUTPUT_DIR = APP_DIR / "storage" / "outputs"


def ui_font(size: int, weight: str = "normal") -> tuple[str, int, str]:
    if sys.platform == "darwin":
        family = "PingFang SC"
    elif sys.platform.startswith("win"):
        family = "Microsoft YaHei UI"
    else:
        family = "Noto Sans CJK SC"
    return (family, size, weight)


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("公文格式助手")
        self.root.geometry("920x700")
        self.root.minsize(760, 560)
        self.root.configure(bg="#eef2f7")

        self.selected_file = tk.StringVar()
        self.summary_text = tk.StringVar(value="支持 .docx、.doc、.rtf、.txt、.html、.odt")
        self.status_text = tk.StringVar(value="请选择一个文稿文件。")

        self._build_ui()

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg="#eef2f7", padx=24, pady=24)
        shell.pack(fill="both", expand=True)

        tk.Label(shell, text="公文格式助手", bg="#eef2f7", fg="#111827", font=ui_font(22, "bold")).pack(anchor="w")
        tk.Label(
            shell,
            text="本地桌面版。选择文稿后，软件会按公文规则处理并生成新的 Word 文件。",
            bg="#eef2f7",
            fg="#4b5563",
            font=ui_font(11),
        ).pack(anchor="w", pady=(8, 0))

        top_card = tk.Frame(shell, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        top_card.pack(fill="x", pady=(18, 14))
        tk.Label(top_card, text="选择文稿", bg="#ffffff", fg="#111827", font=ui_font(12, "bold")).pack(anchor="w")

        file_row = tk.Frame(top_card, bg="#ffffff")
        file_row.pack(fill="x", pady=(14, 10))
        file_row.grid_columnconfigure(0, weight=1)

        self.file_entry = tk.Entry(
            file_row,
            textvariable=self.selected_file,
            font=ui_font(10),
            relief="solid",
            bd=1,
            bg="#ffffff",
            fg="#111827",
        )
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=8)

        tk.Button(
            file_row,
            text="浏览文件",
            command=self.choose_file,
            font=ui_font(10, "bold"),
            bg="#0f766e",
            fg="#ffffff",
            activebackground="#115e59",
            activeforeground="#ffffff",
            relief="flat",
            padx=14,
            pady=8,
        ).grid(row=0, column=1)

        tk.Button(
            top_card,
            text="开始处理",
            command=self.start_formatting,
            font=ui_font(10, "bold"),
            bg="#1d4ed8",
            fg="#ffffff",
            activebackground="#1e40af",
            activeforeground="#ffffff",
            relief="flat",
            padx=14,
            pady=8,
        ).pack(anchor="w", pady=(6, 0))

        tk.Label(top_card, textvariable=self.summary_text, bg="#ffffff", fg="#0f766e", font=ui_font(10)).pack(anchor="w", pady=(14, 0))
        tk.Label(top_card, textvariable=self.status_text, bg="#ffffff", fg="#9a3412", font=ui_font(10), wraplength=800, justify="left").pack(anchor="w", pady=(8, 0))

        result_card = tk.Frame(shell, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        result_card.pack(fill="both", expand=True)
        tk.Label(result_card, text="处理结果", bg="#ffffff", fg="#111827", font=ui_font(12, "bold")).pack(anchor="w")

        actions = tk.Frame(result_card, bg="#ffffff")
        actions.pack(fill="x", pady=(12, 10))
        tk.Button(
            actions,
            text="打开输出目录",
            command=self.open_output_dir,
            font=ui_font(10),
            bg="#e5e7eb",
            fg="#111827",
            activebackground="#d1d5db",
            relief="flat",
            padx=12,
            pady=7,
        ).pack(side="left")

        sections = tk.Frame(result_card, bg="#ffffff")
        sections.pack(fill="both", expand=True)
        sections.grid_columnconfigure(0, weight=1)
        sections.grid_columnconfigure(1, weight=1)
        sections.grid_rowconfigure(1, weight=1)

        tk.Label(sections, text="格式识别", bg="#ffffff", fg="#111827", font=ui_font(11, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))
        tk.Label(sections, text="编号校验", bg="#ffffff", fg="#111827", font=ui_font(11, "bold")).grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.report_box = tk.Text(sections, wrap="word", font=ui_font(10), bg="#ffffff", fg="#111827", relief="solid", bd=1)
        self.report_box.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(8, 0))
        self.report_box.configure(state="disabled")

        self.warning_box = tk.Text(sections, wrap="word", font=ui_font(10), bg="#ffffff", fg="#7c2d12", relief="solid", bd=1)
        self.warning_box.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(8, 0))
        self.warning_box.configure(state="disabled")

    def choose_file(self) -> None:
        patterns = " ".join(f"*{extension}" for extension in sorted(SUPPORTED_INPUT_EXTENSIONS))
        path = filedialog.askopenfilename(
            title="选择文稿文件",
            filetypes=[("支持的文稿", patterns), ("所有文件", "*.*")],
        )
        if path:
            self.selected_file.set(path)
            self.status_text.set("已选择文件，准备处理。")

    def start_formatting(self) -> None:
        file_path = Path(self.selected_file.get().strip())
        if not file_path:
            self.status_text.set("请先选择一个文稿文件。")
            return
        if file_path.suffix.lower() not in SUPPORTED_INPUT_EXTENSIONS:
            supported = "、".join(sorted(SUPPORTED_INPUT_EXTENSIONS))
            self.status_text.set(f"当前支持这些格式：{supported}")
            return

        self.summary_text.set(f"处理中：{file_path.name}")
        self.status_text.set("正在处理，请稍候...")
        threading.Thread(target=self._format_in_background, args=(file_path,), daemon=True).start()

    def _format_in_background(self, file_path: Path) -> None:
        try:
            result = format_docx(file_path, file_path.name, OUTPUT_DIR)
            self.root.after(0, lambda: self.show_result(result))
        except Exception as exc:
            self.root.after(0, lambda: self.show_error(str(exc)))

    def show_result(self, result: FormatResult) -> None:
        self.summary_text.set(f"已生成：{result.output_path.name}")
        self.status_text.set("处理完成。")
        report_text = "\n".join(f"第 {item.index} 段  {item.role}  {item.text}" for item in result.report) or "没有识别到内容。"
        warning_text = "\n".join(f"{item.level}：{item.message}" for item in result.warnings) or "未发现一级、二级标题编号跳号。"
        self._write_box(self.report_box, report_text)
        self._write_box(self.warning_box, warning_text)
        messagebox.showinfo("处理完成", f"已生成文件：\n{result.output_path}")

    def show_error(self, message: str) -> None:
        self.summary_text.set("处理失败")
        self.status_text.set(message)
        messagebox.showerror("处理失败", message)

    def _write_box(self, box: tk.Text, text: str) -> None:
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.insert("1.0", text)
        box.configure(state="disabled")

    def open_output_dir(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(str(OUTPUT_DIR))
            return
        subprocess.run(["open", str(OUTPUT_DIR)], check=False)


def main() -> None:
    if getattr(sys, "frozen", False):
        launch_web_app()
        return

    root = tk.Tk()
    DesktopApp(root)
    root.mainloop()


def launch_web_app() -> None:
    port = find_available_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"

    threading.Thread(target=server.serve_forever, daemon=True).start()
    time.sleep(0.3)
    open_url(url)

    try:
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        server.shutdown()


def find_available_port(start: int = 8780, stop: int = 8899) -> int:
    for port in range(start, stop + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError("没有找到可用的本地端口。")


def open_url(url: str) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", url], check=False)
        return
    webbrowser.open(url)


if __name__ == "__main__":
    main()
