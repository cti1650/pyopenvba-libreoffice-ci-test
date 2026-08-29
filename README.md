# pyOpenVBA + LibreOffice CI Test

pyOpenVBAを使用してExcelファイルのVBAを操作し、GitHub ActionsのWindows RunnerでLibreOfficeを使ってVBAの動作確認を行う検証リポジトリです。

## 概要

このリポジトリは以下の検証を行います：

1. **pyOpenVBA**: PythonからExcelファイルのVBAマクロを読み書き
2. **LibreOffice**: CIでExcel/VBAファイルを処理（変換・マクロ実行）
3. **GitHub Actions**: Windows/Ubuntu/macOSでのクロスプラットフォームテスト

## 必要要件

- Python 3.10以上
- LibreOffice（VBA実行テスト用）

## セットアップ

```bash
# リポジトリをクローン
git clone https://github.com/YOUR_USERNAME/pyopenvba-libreoffice-ci-test.git
cd pyopenvba-libreoffice-ci-test

# 依存関係をインストール
pip install -r requirements.txt
```

## 使い方

### VBA付きExcelファイルの作成

```bash
python scripts/create_excel_with_vba.py
```

### VBAの抽出

```bash
python scripts/extract_vba.py output/test_workbook.xlsm ./extracted
```

### VBAの注入

```bash
python scripts/inject_vba.py output/test_workbook.xlsm vba/sample_module.bas
```

### テストの実行

```bash
python tests/test_vba.py
```

### LibreOfficeでの変換テスト

```bash
python scripts/run_vba_libreoffice.py
```

## プロジェクト構成

```
.
├── .github/
│   └── workflows/
│       └── vba-test.yml        # GitHub Actions ワークフロー
├── scripts/
│   ├── create_excel_with_vba.py  # VBA付きExcel作成
│   ├── extract_vba.py            # VBA抽出
│   ├── inject_vba.py             # VBA注入
│   └── run_vba_libreoffice.py    # LibreOffice連携
├── vba/
│   └── sample_module.bas         # サンプルVBAコード
├── tests/
│   └── test_vba.py               # テストスクリプト
├── requirements.txt
└── README.md
```

## GitHub Actions

CIワークフローでは以下を実行します：

1. **pyOpenVBAテスト**: 複数OS・PythonバージョンでのpyOpenVBAテスト
2. **LibreOffice Windowsテスト**: Windows RunnerでLibreOfficeをインストールしてVBAファイルを処理
3. **LibreOffice Ubuntuテスト**: Ubuntu RunnerでLibreOfficeをインストールしてVBAファイルを処理

## 参考リンク

- [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) - PythonでVBAマクロを読み書きするライブラリ
- [LibreOffice](https://www.libreoffice.org/) - オープンソースのOfficeスイート

## ライセンス

MIT License
