# PDF Renamer Scan Tool

PDFファイルから特定範囲の数字を自動抽出してファイル名を変更するツールです。

## 機能概要

- PDFファイルの順次処理
- 指定範囲からの数字OCR抽出（8桁+999パターン対応）
- 高精度画像前処理（黒字抽出、ノイズ除去）
- 直感的なUI（左側PDFビューア、右側OCRイメージ表示）
- 動的OCR範囲再設定機能
- ファイル名編集・保存
- 処理ログ出力

## 必要環境

- Windows 10/11
- Tesseract OCR（自動検出）
- Python 3.8+ （開発時のみ）

## フォルダ構成

```
06_PDF_Scan/
├── pdf_input/          # 入力PDFファイル
├── pdf_output/         # リネーム済みPDFファイル
├── log_output/         # 処理ログファイル
├── ocr_get_image/      # OCR抽出画像
├── config.txt          # 設定ファイル
├── pdf_renamer.py      # メインプログラム
└── requirements.txt    # 依存関係
```

## 使用方法

### 1. 実行ファイル版
```bash
PDF_Renamer_Scan.exe
```

### 2. Python版
```bash
pip install -r requirements.txt
python pdf_renamer.py
```

## 操作手順

1. **入力フォルダ選択**: PDFファイルが格納されたフォルダを選択
2. **OCR範囲設定**: 必要に応じて数字抽出範囲を再設定
3. **自動処理**: PDFが表示され、OCRで数字を自動抽出
4. **確認・編集**: 抽出された8桁数字を確認・編集
5. **保存**: 「名前を付けて保存」でリネーム済みPDFを出力
6. **次のファイル**: 「次へ」ボタンで次のPDFを処理

## UI仕様

- **起動時最大化**: exeファイル実行時に自動で最大化
- **左右1:1比率**: PDFビューア（左）とOCRエリア（右）
- **赤枠表示**: OCR抽出範囲を視覚的に表示
- **フォントサイズ**: 全UI要素16pt統一
- **入力制限**: テキストボックスは8桁数字のみ入力可能

## OCR精度向上機能

- **CLAHE処理**: コントラスト適応的ヒストグラム均等化
- **ノイズ除去**: メディアンフィルタ適用
- **適応的二値化**: 局所的な閾値処理
- **モルフォロジー処理**: 文字の補強・整形
- **高解像度化**: 1.8倍スケールアップ
- **文字制限**: 数字とハイフンのみ抽出

## 設定ファイル（config.txt）

```
pdf_input_folder=pdf_input
pdf_output_folder=pdf_output
log_output_folder=log_output
ocr_image_folder=ocr_get_image
ocr_x=600
ocr_y=250
ocr_width=300
ocr_height=50
digit_length=8
```

## ログ出力

- **日次ログ**: `log_output/YYYYMMDD.txt`
- **形式**: `[時刻] 元ファイル名 -> 新ファイル名.pdf`
- **OCR画像**: `ocr_get_image/元ファイル名_ocr.png`

## トラブルシューティング

### Tesseract OCRが見つからない場合
以下のパスにTesseractをインストール：
- `C:\Program Files\Tesseract-OCR\tesseract.exe`
- `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`

### OCR精度が低い場合
1. 「OCR範囲を設定」で抽出範囲を調整
2. PDF画像の解像度・品質を確認
3. 数字部分が明確に見える範囲を選択

### ファイルが保存されない場合
- 出力フォルダの書き込み権限を確認
- ファイル名が8桁数字であることを確認
- ディスク容量を確認

## 技術仕様

- **PDF処理**: PyMuPDF (fitz)
- **OCR**: Tesseract + pytesseract
- **画像処理**: OpenCV + NumPy
- **UI**: tkinter
- **ビルド**: PyInstaller

## 更新履歴

- v1.0: 初期リリース
  - 基本的なPDF処理・OCR機能
  - UI実装（最大化、赤枠表示）
  - 高精度画像前処理
  - 動的範囲再設定機能
