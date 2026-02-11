#!/usr/bin/env python3
"""EML to PDF batch converter."""

from __future__ import annotations

import email
import html
import os
import re
import sys
from email import policy
from pathlib import Path
from typing import Callable

try:
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import (
        HRFlowable,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
    )
except ImportError:
    sys.exit("reportlab が必要です: pip install reportlab")

pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))

# ── Styles ────────────────────────────────────

STYLE_SUBJECT = ParagraphStyle(
    "Subject",
    fontName="HeiseiKakuGo-W5",
    fontSize=14,
    leading=20,
    spaceAfter=4 * mm,
    textColor=HexColor("#111111"),
)
STYLE_HEADER = ParagraphStyle(
    "Header",
    fontName="HeiseiKakuGo-W5",
    fontSize=9,
    leading=14,
    textColor=HexColor("#333333"),
)
STYLE_BODY = ParagraphStyle(
    "Body",
    fontName="HeiseiMin-W3",
    fontSize=10,
    leading=16,
    textColor=HexColor("#1a1a1a"),
    spaceAfter=1 * mm,
)

# ── EML parsing ──────────────────────────────

_STRIP_HTML = re.compile(r"<[^>]+>")


def _extract_body(msg: email.message.Message) -> str:
    """Extract plain-text body from an email message."""
    if not msg.is_multipart():
        content = msg.get_content()
        return content if isinstance(content, str) else ""

    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            content = part.get_content()
            if isinstance(content, str):
                return content

    # Fallback: strip HTML tags
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            content = part.get_content()
            if isinstance(content, str):
                return _STRIP_HTML.sub("", content)

    return ""


def parse_eml(filepath: Path) -> dict[str, str]:
    """Parse .eml file into subject, from, to, cc, date, body."""
    with open(filepath, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    return {
        "subject": str(msg.get("Subject", "")) or "(件名なし)",
        "from": str(msg.get("From", "")),
        "to": str(msg.get("To", "")),
        "cc": str(msg.get("Cc", "")),
        "date": str(msg.get("Date", "")),
        "body": _extract_body(msg).strip(),
    }


# ── PDF generation ───────────────────────────


def _escape(text: str) -> str:
    """Escape text for reportlab Paragraph markup."""
    return html.escape(text, quote=False).replace("  ", "&nbsp; ")


def _build_story(data: dict[str, str]) -> list:
    """Build reportlab flowable list from email data."""
    story: list = [Paragraph(_escape(data["subject"]), STYLE_SUBJECT)]

    headers = [("差出人", data["from"]), ("宛先", data["to"])]
    if data["cc"]:
        headers.append(("CC", data["cc"]))
    headers.append(("日時", data["date"]))

    for label, value in headers:
        markup = (
            f'<font color="#888888">{_escape(label)}:</font>'
            f"&nbsp;&nbsp;{_escape(value)}"
        )
        story.append(Paragraph(markup, STYLE_HEADER))

    story.append(Spacer(1, 5 * mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=HexColor("#cccccc")))
    story.append(Spacer(1, 5 * mm))

    for line in data["body"].split("\n"):
        if line.strip():
            story.append(Paragraph(_escape(line), STYLE_BODY))
        else:
            story.append(Spacer(1, 3 * mm))

    return story


def eml_to_pdf(eml_path: Path, pdf_path: Path) -> None:
    """Convert a single .eml file to PDF."""
    data = parse_eml(eml_path)
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        title=data["subject"],
        author=data["from"],
    )
    doc.build(_build_story(data))


# ── Batch conversion ─────────────────────────

ProgressCallback = Callable[[int, int, str, bool, str | None], None]


def batch_convert(
    input_dir: str,
    output_dir: str,
    on_progress: ProgressCallback | None = None,
) -> tuple[int, int, str]:
    """Convert all .eml files in *input_dir*, writing PDFs to *output_dir*."""
    src = Path(input_dir)
    dst = Path(output_dir)
    dst.mkdir(parents=True, exist_ok=True)

    eml_files = sorted(src.glob("*.eml"))
    if not eml_files:
        return 0, 0, ".eml ファイルが見つかりません"

    total = len(eml_files)
    ok = ng = 0

    for i, eml in enumerate(eml_files, 1):
        try:
            eml_to_pdf(eml, dst / f"{eml.stem}.pdf")
            ok += 1
            if on_progress:
                on_progress(i, total, eml.name, True, None)
        except (OSError, ValueError) as e:
            ng += 1
            if on_progress:
                on_progress(i, total, eml.name, False, str(e))

    summary = f"完了: {ok}/{total} 件成功"
    if ng:
        summary += f", {ng} 件エラー"
    return ok, ng, summary


# ── GUI ──────────────────────────────────────


def _run_gui() -> None:
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class App:
        def __init__(self, root: tk.Tk) -> None:
            self.root = root
            root.title("EML → PDF 一括変換")
            root.geometry("620x500")
            root.minsize(500, 400)

            self.input_dir = tk.StringVar()
            self.output_dir = tk.StringVar()
            self._build(root)

        def _build(self, root: tk.Tk) -> None:
            main = ttk.Frame(root, padding=20)
            main.pack(fill=tk.BOTH, expand=True)

            ttk.Label(
                main, text="EML → PDF 一括変換", font=("Helvetica", 16, "bold")
            ).pack(pady=(0, 15))

            for label, var, cmd in [
                (" 入力フォルダ (.eml) ", self.input_dir, self._pick_input),
                (" 出力フォルダ (PDF) ", self.output_dir, self._pick_output),
            ]:
                frame = ttk.LabelFrame(main, text=label, padding=8)
                frame.pack(fill=tk.X, pady=(0, 8))
                row = ttk.Frame(frame)
                row.pack(fill=tk.X)
                ttk.Entry(row, textvariable=var).pack(
                    side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8)
                )
                ttk.Button(row, text="選択...", command=cmd).pack(side=tk.RIGHT)

            self.btn = ttk.Button(main, text="  変換開始  ", command=self._start)
            self.btn.pack(pady=12)

            self.pvar = tk.DoubleVar()
            ttk.Progressbar(main, variable=self.pvar, maximum=100).pack(
                fill=tk.X, pady=(0, 4)
            )

            self.status = tk.StringVar(value="フォルダを選択して「変換開始」を押してください")
            ttk.Label(main, textvariable=self.status, foreground="#555").pack()

            log_frame = ttk.LabelFrame(main, text=" ログ ", padding=4)
            log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
            self.log = tk.Text(log_frame, height=8, font=("Menlo", 9), wrap=tk.WORD)
            sb = ttk.Scrollbar(log_frame, command=self.log.yview)
            self.log.configure(yscrollcommand=sb.set)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            self.log.pack(fill=tk.BOTH, expand=True)

        # ── helpers ──

        def _after(self, fn: Callable) -> None:
            self.root.after(0, fn)

        def _append_log(self, msg: str) -> None:
            self.log.insert(tk.END, msg + "\n")
            self.log.see(tk.END)

        def _pick_input(self) -> None:
            if p := filedialog.askdirectory(title="入力フォルダを選択"):
                self.input_dir.set(p)
                if not self.output_dir.get():
                    self.output_dir.set(os.path.join(p, "pdf_output"))

        def _pick_output(self) -> None:
            if p := filedialog.askdirectory(title="出力フォルダを選択"):
                self.output_dir.set(p)

        def _start(self) -> None:
            src, dst = self.input_dir.get(), self.output_dir.get()
            if not src or not dst:
                messagebox.showwarning("警告", "フォルダを選択してください")
                return
            self.btn.configure(state="disabled")
            self.log.delete("1.0", tk.END)
            self.pvar.set(0)
            threading.Thread(target=self._convert, args=(src, dst), daemon=True).start()

        def _convert(self, src: str, dst: str) -> None:
            def on_progress(i: int, total: int, name: str, ok: bool, err: str | None) -> None:
                mark = "✓" if ok else "✗"
                detail = f"  ({err})" if err else ""
                self._after(lambda: self.pvar.set(i / total * 100))
                self._after(lambda: self.status.set(f"{i} / {total} 件処理中..."))
                self._after(lambda m=f"{mark} {name}{detail}": self._append_log(m))

            try:
                ok, _, summary = batch_convert(src, dst, on_progress)
                self._after(lambda: self.status.set(summary))
                self._after(lambda: self._append_log(f"\n{summary}\n出力先: {dst}"))
                if ok:
                    self._after(lambda: messagebox.showinfo("完了", summary))
            except Exception as e:
                self._after(lambda: messagebox.showerror("エラー", str(e)))
            finally:
                self._after(lambda: self.btn.configure(state="normal"))

    root = tk.Tk()
    App(root)
    root.mainloop()


# ── CLI ──────────────────────────────────────


def _run_cli() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="EML to PDF batch converter")
    parser.add_argument("input_dir", help="入力フォルダ")
    parser.add_argument("-o", "--output", help="出力フォルダ")
    args = parser.parse_args()

    output = args.output or os.path.join(args.input_dir, "pdf_output")
    print(f"入力: {args.input_dir}\n出力: {output}\n")

    def on_progress(i: int, total: int, name: str, ok: bool, err: str | None) -> None:
        mark = "✓" if ok else "✗"
        detail = f"  ({err})" if err else ""
        print(f"  [{i}/{total}] {mark} {name}{detail}")

    _, _, summary = batch_convert(args.input_dir, output, on_progress)
    print(f"\n{summary}")


# ── Entry point ──────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        _run_cli()
    else:
        _run_gui()
