#!/bin/bash
# EML → PDF 一括変換 起動スクリプト
# 初回は自動でセットアップします

cd "$(dirname "$0")"

# 仮想環境がなければ作成
if [ ! -d "venv" ]; then
    echo ""
    echo "========================================="
    echo "  初回セットアップ中...（1分ほどかかります）"
    echo "========================================="
    echo ""
    python3 -m venv venv
    source venv/bin/activate
    pip install --quiet reportlab
    echo ""
    echo "  セットアップ完了！"
    echo ""
else
    source venv/bin/activate
fi

python3 eml2pdf.py
