#!/usr/bin/env python3
"""
EML to PDF Batch Converter
.emlファイルをPDFに一括変換するツール

使い方:
  GUI: python eml2pdf.py
  CLI: python eml2pdf.py <入力フォルダ> [-o 出力フォルダ]
"""

import os
import sys
import email
from email import policy
from email.header import decode_header as _decode_header
from pathlib import Path
import traceback


def check_dependencies():
    try:
        import reportlab
        return True
    except ImportError:
        print("=" * 50)
        print("reportlab が必要です。以下のコマンドでインストールしてください：")
        print()
        print("  pip install reportlab")
        print()
        print("または:")
        print("  pip3 install reportlab")
        print("=" * 50)
        return False


if not check_dependencies():
    sys.exit(1)

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.colors import HexColor

# Register Japanese fonts
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))


# ──────────────────────────────────────────────
# EML Parsing
# ──────────────────────────────────────────────

def decode_mime_header(value):
    """Decode MIME-encoded header value (e.g., =?UTF-8?B?...?=)."""
    if not value:
        return ""
    parts = []
    for decoded_bytes, charset in _decode_header(str(value)):
        if isinstance(decoded_bytes, bytes):
            charset = charset or 'utf-8'
            for enc in [charset, 'utf-8', 'shift_jis', 'iso-2022-jp', 'euc-jp', 'latin-1']:
                try:
                    parts.append(decoded_bytes.decode(enc))
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            else:
                parts.append(decoded_bytes.decode('latin-1'))
        else:
            parts.append(str(decoded_bytes))
    return ' '.join(parts)


def parse_eml(filepath):
    """Parse an .eml file and return structured email data."""
    with open(filepath, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    subject = str(msg.get('Subject', '')) or '(件名なし)'
    from_addr = str(msg.get('From', ''))
    to_addr = str(msg.get('To', ''))
    cc_addr = str(msg.get('Cc', ''))
    date_str = str(msg.get('Date', ''))

    # Extract plain text body
    body = ''
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                try:
                    payload = part.get_content()
                    if isinstance(payload, str):
                        body = payload
                        break
                    elif isinstance(payload, bytes):
                        charset = part.get_content_charset() or 'utf-8'
                        body = payload.decode(charset, errors='replace')
                        break
                except Exception:
                    continue
        # Fallback: try text/html if no text/plain found
        if not body:
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    try:
                        payload = part.get_content()
                        if isinstance(payload, str):
                            # Simple HTML tag removal
                            import re
                            body = re.sub(r'<[^>]+>', '', payload)
                            body = body.strip()
                            break
                    except Exception:
                        continue
    else:
        try:
            payload = msg.get_content()
            if isinstance(payload, str):
                body = payload
            elif isinstance(payload, bytes):
                charset = msg.get_content_charset() or 'utf-8'
                body = payload.decode(charset, errors='replace')
        except Exception:
            body = '(本文を読み取れませんでした)'

    return {
        'subject': subject,
        'from': from_addr,
        'to': to_addr,
        'cc': cc_addr,
        'date': date_str,
        'body': body.strip(),
    }


# ──────────────────────────────────────────────
# PDF Generation
# ──────────────────────────────────────────────

def escape_para(text):
    """Escape text for reportlab Paragraph (XML markup)."""
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace('  ', '&nbsp; ')
    return text


def eml_to_pdf(eml_path, pdf_path):
    """Convert a single .eml file to a PDF file."""
    data = parse_eml(eml_path)

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        title=data['subject'],
        author=data['from'],
    )

    # Styles
    style_subject = ParagraphStyle(
        'Subject',
        fontName='HeiseiKakuGo-W5',
        fontSize=14,
        leading=20,
        spaceAfter=4 * mm,
        textColor=HexColor('#111111'),
    )
    style_header = ParagraphStyle(
        'Header',
        fontName='HeiseiKakuGo-W5',
        fontSize=9,
        leading=14,
        textColor=HexColor('#333333'),
    )
    style_body = ParagraphStyle(
        'Body',
        fontName='HeiseiMin-W3',
        fontSize=10,
        leading=16,
        textColor=HexColor('#1a1a1a'),
        spaceAfter=1 * mm,
    )

    story = []

    # Subject
    story.append(Paragraph(escape_para(data['subject']), style_subject))

    # Headers
    header_items = [
        ('差出人', data['from']),
        ('宛先', data['to']),
    ]
    if data['cc']:
        header_items.append(('CC', data['cc']))
    header_items.append(('日時', data['date']))

    for label, value in header_items:
        line = (
            f'<font color="#888888">{escape_para(label)}:</font>'
            f'&nbsp;&nbsp;{escape_para(value)}'
        )
        story.append(Paragraph(line, style_header))

    # Separator
    story.append(Spacer(1, 5 * mm))
    story.append(HRFlowable(
        width="100%", thickness=0.5, color=HexColor('#cccccc')
    ))
    story.append(Spacer(1, 5 * mm))

    # Body
    for line in data['body'].split('\n'):
        if line.strip() == '':
            story.append(Spacer(1, 3 * mm))
        else:
            story.append(Paragraph(escape_para(line), style_body))

    doc.build(story)


# ──────────────────────────────────────────────
# Batch Conversion
# ──────────────────────────────────────────────

def batch_convert(input_dir, output_dir, progress_callback=None):
    """Convert all .eml files in input_dir to PDF in output_dir."""
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    eml_files = sorted(input_path.glob('*.eml'))

    if not eml_files:
        return 0, 0, ".eml ファイルが見つかりません"

    total = len(eml_files)
    success = 0
    errors = 0

    for i, eml_file in enumerate(eml_files):
        try:
            pdf_name = eml_file.stem + '.pdf'
            pdf_file = output_path / pdf_name
            eml_to_pdf(eml_file, pdf_file)
            success += 1
            if progress_callback:
                progress_callback(i + 1, total, eml_file.name, True, None)
        except Exception as e:
            errors += 1
            if progress_callback:
                progress_callback(i + 1, total, eml_file.name, False, str(e))

    summary = f"完了: {success}/{total} 件成功"
    if errors:
        summary += f", {errors} 件エラー"

    return success, errors, summary


# ──────────────────────────────────────────────
# GUI (tkinter)
# ──────────────────────────────────────────────

def run_gui():
    """Launch the GUI application."""
    import tkinter as tk
    from tkinter import filedialog, ttk, messagebox
    import threading

    class App:
        def __init__(self, root):
            self.root = root
            self.root.title("EML → PDF 一括変換")
            self.root.geometry("620x500")
            self.root.minsize(500, 400)

            self.input_dir = tk.StringVar()
            self.output_dir = tk.StringVar()
            self._build_ui()

        def _build_ui(self):
            main = ttk.Frame(self.root, padding=20)
            main.pack(fill=tk.BOTH, expand=True)

            # Title
            ttk.Label(
                main, text="EML → PDF 一括変換",
                font=('Helvetica', 16, 'bold')
            ).pack(pady=(0, 15))

            # Input folder
            in_frame = ttk.LabelFrame(main, text=" 入力フォルダ (.eml) ", padding=8)
            in_frame.pack(fill=tk.X, pady=(0, 8))
            in_row = ttk.Frame(in_frame)
            in_row.pack(fill=tk.X)
            ttk.Entry(in_row, textvariable=self.input_dir).pack(
                side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
            ttk.Button(in_row, text="選択...", command=self._pick_input).pack(side=tk.RIGHT)

            # Output folder
            out_frame = ttk.LabelFrame(main, text=" 出力フォルダ (PDF) ", padding=8)
            out_frame.pack(fill=tk.X, pady=(0, 8))
            out_row = ttk.Frame(out_frame)
            out_row.pack(fill=tk.X)
            ttk.Entry(out_row, textvariable=self.output_dir).pack(
                side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
            ttk.Button(out_row, text="選択...", command=self._pick_output).pack(side=tk.RIGHT)

            # Convert button
            self.btn = ttk.Button(main, text="  変換開始  ", command=self._on_convert)
            self.btn.pack(pady=12)

            # Progress
            self.pvar = tk.DoubleVar()
            ttk.Progressbar(main, variable=self.pvar, maximum=100).pack(
                fill=tk.X, pady=(0, 4))

            self.status = tk.StringVar(
                value="フォルダを選択して「変換開始」を押してください")
            ttk.Label(main, textvariable=self.status, foreground='#555').pack()

            # Log
            log_frame = ttk.LabelFrame(main, text=" ログ ", padding=4)
            log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

            self.log = tk.Text(log_frame, height=8, font=('Menlo', 9), wrap=tk.WORD)
            sb = ttk.Scrollbar(log_frame, command=self.log.yview)
            self.log.configure(yscrollcommand=sb.set)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            self.log.pack(fill=tk.BOTH, expand=True)

        def _log(self, msg):
            self.log.insert(tk.END, msg + '\n')
            self.log.see(tk.END)

        def _pick_input(self):
            p = filedialog.askdirectory(title="入力フォルダを選択")
            if p:
                self.input_dir.set(p)
                if not self.output_dir.get():
                    self.output_dir.set(os.path.join(p, 'pdf_output'))

        def _pick_output(self):
            p = filedialog.askdirectory(title="出力フォルダを選択")
            if p:
                self.output_dir.set(p)

        def _on_convert(self):
            ind = self.input_dir.get()
            outd = self.output_dir.get()
            if not ind:
                messagebox.showwarning("警告", "入力フォルダを選択してください")
                return
            if not outd:
                messagebox.showwarning("警告", "出力フォルダを選択してください")
                return
            self.btn.configure(state='disabled')
            self.log.delete('1.0', tk.END)
            self.pvar.set(0)
            threading.Thread(
                target=self._convert, args=(ind, outd), daemon=True
            ).start()

        def _convert(self, input_dir, output_dir):
            def cb(current, total, name, ok, err):
                pct = current / total * 100
                self.root.after(0, lambda: self.pvar.set(pct))
                self.root.after(0, lambda: self.status.set(
                    f"{current} / {total} 件処理中..."))
                mark = "✓" if ok else "✗"
                detail = f"  ({err})" if err else ""
                msg = f"{mark} {name}{detail}"
                self.root.after(0, lambda m=msg: self._log(m))

            try:
                ok, ng, summary = batch_convert(input_dir, output_dir, cb)
                self.root.after(0, lambda: self.status.set(summary))
                self.root.after(0, lambda: self._log(f"\n{summary}"))
                self.root.after(0, lambda: self._log(f"出力先: {output_dir}"))
                if ok > 0:
                    self.root.after(0, lambda: messagebox.showinfo("完了", summary))
                elif ng == 0:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "情報", ".eml ファイルが見つかりません"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("エラー", str(e)))
                self.root.after(0, lambda: self._log(f"エラー: {e}"))
            finally:
                self.root.after(0, lambda: self.btn.configure(state='normal'))

    root = tk.Tk()
    App(root)
    root.mainloop()


# ──────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────

def run_cli():
    """Run the command-line version."""
    import argparse

    parser = argparse.ArgumentParser(
        description='EML to PDF Batch Converter - .emlファイルをPDFに一括変換',
    )
    parser.add_argument('input_dir', help='入力フォルダ（.emlファイルがあるフォルダ）')
    parser.add_argument(
        '-o', '--output',
        help='出力フォルダ（デフォルト: 入力フォルダ/pdf_output）')
    args = parser.parse_args()

    input_dir = args.input_dir
    output_dir = args.output or os.path.join(input_dir, 'pdf_output')

    print(f"入力: {input_dir}")
    print(f"出力: {output_dir}")
    print()

    def cb(current, total, name, ok, err):
        mark = "✓" if ok else "✗"
        detail = f"  ({err})" if err else ""
        print(f"  [{current}/{total}] {mark} {name}{detail}")

    ok, ng, summary = batch_convert(input_dir, output_dir, cb)
    print()
    print(summary)


# ──────────────────────────────────────────────
# Entry Point
# ──────────────────────────────────────────────

def main():
    if len(sys.argv) > 1:
        run_cli()
    else:
        try:
            run_gui()
        except ImportError:
            print("GUI モードには tkinter が必要です。")
            print("CLI モードで使用: python eml2pdf.py <入力フォルダ> [-o 出力フォルダ]")
            sys.exit(1)


if __name__ == '__main__':
    main()
