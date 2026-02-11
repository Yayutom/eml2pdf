# eml2pdf

.eml ファイル（メール）を PDF に一括変換する Python ツール。GUI・CLI 両対応。

## Features

- フォルダ単位で .eml ファイルを一括変換
- 件名・差出人・宛先・CC・日時をヘッダーとして PDF に記載
- 日本語メール完全対応（UTF-8 / ISO-2022-JP / Shift_JIS / EUC-JP）
- MIME エンコード（Base64 / Quoted-Printable）を自動デコード
- マルチパートメールから text/plain を自動抽出
- GUI（tkinter）と CLI の2モード搭載

## Quick Start

```bash
git clone https://github.com/Yayutom/eml2pdf.git
cd eml2pdf
bash run.sh
```

`run.sh` が初回のみ仮想環境の作成と依存ライブラリのインストールを自動で行います。

2回目以降は `EML→PDF変換.command` をダブルクリックするだけで起動します（macOS）。

## Usage

### GUI

```bash
bash run.sh
```

フォルダ選択画面が開くので、入力フォルダと出力フォルダを選んで「変換開始」を押してください。

### CLI

```bash
source venv/bin/activate
python3 eml2pdf.py /path/to/eml_folder -o /path/to/output
```

`-o` を省略すると入力フォルダ内に `pdf_output/` が作成されます。

## Requirements

- Python 3.9+
- reportlab（`run.sh` で自動インストール）

## License

MIT
