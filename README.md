# pyOpenVBA + LibreOffice CI Test

pyOpenVBAを使用してExcelファイルのVBAを操作し、GitHub ActionsのWindows RunnerでLibreOfficeを使ってVBAの動作確認を行う検証リポジトリです。

## 概要

このリポジトリは以下の検証を行います：

1. **pyOpenVBA**: PythonからExcelファイルのVBAマクロを読み書き
2. **LibreOffice UNO API**: CIでVBAマクロを実際に実行して動作確認
3. **Win32 API**: VBAからWindows APIを呼び出すコードの検証
4. **VBAクラスモジュール**: pyOpenVBAでのクラス生成と、LibreOfficeでの対応状況の実測
5. **GitHub Actions**: Windows/Ubuntu/macOSでのクロスプラットフォームテスト
6. **CIキャッシュ**: LibreOfficeインストールのキャッシュによる高速化
7. **Docker**: ホストにLibreOfficeを入れずにローカル検証

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

## Docker（ローカル検証）

ホストにLibreOfficeをインストールせずに、VBAの実行検証と変換処理を試せます。

```bash
docker compose build

docker compose run --rm class-test   # ワークブック生成 + VBAクラス対応状況の実測
docker compose run --rm build        # .xlsm の生成のみ
docker compose run --rm convert      # output/*.xlsm を .ods に変換（引数で形式変更可）
docker compose run --rm shell        # コンテナ内で対話的に確認
```

`scripts/` `vba/` はread-onlyでマウントされるため、ホスト側の編集は再ビルドなしで反映されます。結果は `output/` に書き出されます。

## プロジェクト構成

```
.
├── .github/
│   └── workflows/
│       └── vba-test.yml          # GitHub Actions ワークフロー
├── docker/
│   ├── Dockerfile                # LibreOffice + python3-uno
│   ├── run-class-test.sh         # 生成 → 実測 の一括実行
│   └── convert.sh                # xlsm → ods 変換
├── docker-compose.yml
├── scripts/
│   ├── create_excel_with_vba.py    # VBA付きExcel作成
│   ├── create_class_test_excel.py  # クラスモジュール検証用ワークブック生成
│   ├── extract_vba.py              # VBA抽出（.bas / .cls）
│   ├── inject_vba.py               # VBA注入（.bas / .cls、新規追加対応）
│   ├── run_vba_libreoffice.py      # LibreOffice連携
│   ├── run_class_test_libreoffice.py  # クラス対応状況プローブ（実行側）
│   └── uno_class_probe.py          # クラス対応状況プローブ（UNO側）
├── vba/
│   ├── sample_module.bas         # サンプルVBAコード
│   ├── win32api_module.bas       # Win32 API呼び出しVBAコード
│   └── class_test/               # クラスモジュール検証用VBA一式
├── tests/
│   └── test_vba.py               # テストスクリプト
├── requirements.txt
└── README.md
```

## GitHub Actions

CIワークフローでは以下を実行します：

1. **pyOpenVBAテスト**: 複数OS・PythonバージョンでのpyOpenVBAテスト
2. **LibreOffice Windowsテスト**: Windows RunnerでLibreOfficeをインストールしてVBAマクロを実行
3. **LibreOffice Ubuntuテスト**: Ubuntu RunnerでLibreOfficeをインストールしてVBAマクロを実行
4. **VBAクラスモジュールテスト**: Ubuntu/WindowsでVBAクラスの対応状況を実測

### LibreOfficeキャッシュ

CIの実行時間短縮のため、LibreOfficeのインストールをキャッシュしています：

- **Windows**: `actions/cache@v4` で `C:\Program Files\LibreOffice` をキャッシュ
- **Ubuntu**: `awalsh128/cache-apt-pkgs-action` でaptパッケージをキャッシュ

初回実行時にキャッシュが保存され、2回目以降はインストールをスキップします。

### VBAマクロ実行テスト

LibreOffice UNO APIを使用して、実際にVBAマクロを実行し動作確認を行います：

```
--- Testing: test_workbook.xlsm ---
VBA Modules found: ['Module1']
MACRO EXECUTED: AddNumbers(10,20) = 30
CELL WRITE TEST: PASS (wrote to A1)
MACRO TEST: PASS
```

実行されるテスト：
- `AddNumbers(10, 20)` 関数の実行と結果検証（期待値: 30）
- セルへの書き込みテスト

## Win32 API VBAモジュール

`vba/win32api_module.bas` には以下のWindows API呼び出しが含まれています：

| API関数 | DLL | 説明 |
|---------|-----|------|
| `GetComputerNameA` | kernel32 | コンピュータ名を取得 |
| `GetUserNameA` | advapi32 | ユーザー名を取得 |
| `GetTickCount` | kernel32 | システム起動からの経過時間(ms) |
| `GetSystemMetrics` | user32 | 画面解像度などのシステム情報 |
| `GetTempPathA` | kernel32 | 一時フォルダのパス |
| `Sleep` | kernel32 | 処理を指定ミリ秒停止 |

### 注意事項

- Win32 APIはWindows専用のため、LibreOffice（Linux/macOS）では動作しません
- 64bit/32bit両対応の`PtrSafe`宣言を使用しています
- CIではVBAコードの注入確認のみ行い、Win32 APIの実行はスキップされます

## VBAクラスモジュール

### pyOpenVBA側：何が作れるか

`VBAProject.add_module()` を使うと、標準モジュールに加えて**クラスモジュールを新規作成できます**。クラスは汎用CLSID（`{FCFB3D2A-...}`）だけで済むため、pyOpenVBAが `synthesize_class_header()` でヘッダを合成できるからです。

| 種別 | 新規作成 | コード編集 |
|------|:--------:|:----------:|
| 標準モジュール (`.bas`) | ✅ | ✅ |
| クラスモジュール (`.cls`) | ✅ | ✅ |
| ドキュメントモジュール (ThisWorkbook / Sheet1) | ❌ ホスト固有CLSIDが必要 | ✅ |
| UserForm（デザイナ） | ❌ 上記に加え `.frx` レイアウト非対応 | ✅ コードのみ |

pyOpenVBA自身のソースに明記されています：

```
"pyOpenVBA does not invent document or designer module headers."
```

VBEからエクスポートした `VERSION 1.0 CLASS` 形式の `.cls` はそのまま渡せます（`normalize_class_source()` がストリーム形式に変換し、`VB_PredeclaredId` などの属性は保持されます）。

### LibreOffice側：実測結果

`vba/class_test/` のVBAをLibreOfficeで実行した結果です。**3環境すべてで 10/16 PASS、失敗項目も完全に一致**しました。

| 環境 | LibreOffice | 結果 |
|------|-------------|------|
| Docker（Debian bookworm） | 7.4.7.2 | 10/16 |
| GitHub Actions（ubuntu-latest） | 24.2.7.2 | 10/16 |
| GitHub Actions（windows-latest） | choco `libreoffice-fresh` | 10/16 |

7.4 から 24.2 というメジャーバージョン差、およびLinux/Windowsの違いを越えて結果が変わらないため、以下はバージョン固有の不具合ではなく**VBA互換モードの構造的な制限**と判断できます。

LibreOfficeはpyOpenVBAが書いた `.cls` を**クラスモジュールとして正しく認識**します（生成されたBasicに `Option ClassModule` と `Rem Attribute VBA_ModuleType=VBAClassModule` が付く）。

| 機能 | 結果 | 備考 |
|------|:----:|------|
| `Set c = New MyClass` | ✅ | |
| `Dim c As New MyClass` | ✅ | |
| `Property Get` / `Let` | ✅ | |
| `Property Set`（オブジェクト保持） | ✅ | |
| Privateフィールドの状態保持 | ✅ | |
| `Class_Initialize` | ✅ | 初回使用前に1回だけ発火 |
| `Class_Terminate` | ✅ | `Set x = Nothing` で発火 |
| `Collection` へのインスタンス格納 | ✅ | |
| `TypeName()` | ❌ | クラス名ではなく `"Object"` を返す |
| `VB_PredeclaredId = True`（既定インスタンス） | ❌ | エラー423 `Property or method not found` |
| `Implements`（インスタンス化のみ） | ✅ | クラス自体は生成できる |
| `Implements`（インターフェース経由の呼び出し） | ❌ | エラー425 `Invalid use of an object` |
| `Public Event` / `RaiseEvent` / `WithEvents` | ❌ | エラー420 `Invalid object reference`（`New` すら失敗） |

**結論**: 通常のデータ保持クラス（メソッド＋プロパティ＋ライフサイクル）はLibreOfficeでそのまま動きます。インターフェース、カスタムイベント、既定インスタンスに依存した設計は移植できません。

### 実行方法

```bash
# Docker（ホストにLibreOffice不要）
docker compose run --rm class-test

# ローカルにLibreOfficeがある場合
python scripts/create_class_test_excel.py
python scripts/run_class_test_libreoffice.py
```

個別のVBA機能が失敗しても終了コードは0です（失敗も「結果」のため）。すべてPASSを要求する場合は `--strict` を付けます。

### 検証時に踏んだLibreOfficeの落とし穴

この検証で判明した、UNO経由でVBAを実行する際の2点です。どちらも `scripts/uno_class_probe.py` で対処しています。

1. **`MacroExecutionMode` は 9**
   `ALWAYS_EXECUTE_NO_WARN` は **9** です。よく使われる `4` は `USE_CONFIG_REJECT_CONFIRMATION` で、**エラーも出さずにマクロを一切実行しません**（関数は既定値を返すだけ）。

2. **`VBAProject` ライブラリはスクリプトプロバイダから直接実行できない**
   `getScript()` は成功するのに `invoke()` が既定値を返し、コードが走りません。中身に関係なく、自明な `Function VbEcho() As String` を `VBAProject` に直接入れても空でした。一方、同じ関数をドキュメントの `Standard` ライブラリに入れると正常に動きます。
   そのため、ドキュメントの `Standard` ライブラリに `VBAProject.<Module>.<Function>()` を呼ぶブリッジモジュールを注入して実行しています。

環境依存の落とし穴も2点あります。

3. **Ubuntu: aptキャッシュは `execute_install_scripts: true` が必須**
   `cache-apt-pkgs-action` はキャッシュ復元時にdpkgのpostinstを実行しないため、LibreOfficeがPATHには居るのに起動できない状態になります（`soffice --version` が無出力、起動時に `Unspecified Application Error`）。初回（キャッシュミス）は通り、2回目（キャッシュヒット）から壊れるため気付きにくい挙動です。CIでは復元後に `soffice --version` を検証して即失敗させています。

4. **Windows: `soffice.exe --version` が返らないことがある**
   バージョン取得は20秒で打ち切って続行します。UNO用のPythonは `C:\Program Files\LibreOffice\program\python.exe`（バンドル版）が使われます。

## CI出力例

CIログでは以下のような詳細な検証結果が表示されます：

```
=== VBA Verification (via pyOpenVBA) ===

--- Verifying: test_workbook.xlsm ---
  File size: 10240 bytes
  Modules: ['Module1']

  [Module1] (15 lines)
    Functions/Subs: 2
    Preview:
      ' Sample VBA Module for testing
      Option Explicit
      Public Function AddNumbers(a As Long, b As Long) As Long
      ...

=== LibreOffice VBA Macro Test ===

--- Testing: test_workbook.xlsm ---
VBA Modules found: ['Module1']
MACRO EXECUTED: AddNumbers(10,20) = 30
MACRO TEST: PASS
```

## 参考リンク

- [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) - PythonでVBAマクロを読み書きするライブラリ
- [LibreOffice](https://www.libreoffice.org/) - オープンソースのOfficeスイート

## ライセンス

MIT License
