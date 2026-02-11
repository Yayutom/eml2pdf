# EML → PDF 一括変換ツール

.eml ファイル（メール）をPDFに一括変換するツールです。GUI / CLI 両対応。

## セットアップ（初回のみ）

```bash
bash run.sh
```

自動で仮想環境の作成と依存ライブラリのインストールが行われます。

## 使い方

### GUI（ダブルクリック）

`EML→PDF変換.command` をダブルクリックするとGUIが起動します。

### CLI

```bash
source venv/bin/activate
python3 eml2pdf.py /path/to/eml_folder -o /path/to/output
```

## 動作環境

- macOS / Linux / Windows
- Python 3.9+
- reportlab
