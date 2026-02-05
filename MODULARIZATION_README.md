# 出退勤確認システム - モジュール化ドキュメント

## 概要

2200行以上あった`出退勤確認システム.py`を、保守性向上のためにモジュール化しました。
メインファイルは約1000行に縮小され、機能ごとに分離されたモジュールで構成されています。

## ファイル構成

### メインファイル
- **出退勤確認システム.py** (約1000行)
  - GUIメインクラス
  - 画面表示ロジック
  - イベントハンドラー

### バックアップ
- **出退勤確認システム_backup_before_modularization.py**
  - モジュール化前の元ファイル（2200行以上）

### モジュールディレクトリ (modules/)

#### 1. **constants.py**
- 定数定義
- JST（日本時間タイムゾーン）
- パスワードハッシュ
- ディレクトリパス
- ウィンドウサイズ設定

#### 2. **database_manager.py**
- データベース操作クラス (`DatabaseManager`)
- データベース初期化
- 講師情報の管理（登録・取得・更新）
- 打刻記録の保存・取得
- 日付別・月別のデータ集計

#### 3. **card_reader_manager.py**
- カードリーダー管理クラス (`CardReaderManager`)
- リーダーの初期化
- カード検出
- カードUID読み取り
- リーダーへの接続・切断

#### 4. **csv_exporter.py**
- CSV出力クラス (`CSVExporter`)
- 日次集計CSVエクスポート
- 講師別日次集計
- ファイル名の一意性管理
- oldフォルダへの自動バックアップ

#### 5. **monthly_exporter.py**
- 月次集計クラス (`MonthlyExporter`)
- 月次集計CSVエクスポート
- 授業用・会議用の統合集計
- 講師別月次レポート作成

#### 6. **utils.py**
- ユーティリティクラス群
- `ConfigManager`: 設定ファイル管理
- `SoundManager`: 音声フィードバック管理

#### 7. **__init__.py**
- モジュールパッケージ初期化
- 公開APIの定義

## モジュール化の利点

### 1. **保守性の向上**
- 各機能が独立したモジュールに分離
- 変更の影響範囲が明確
- バグの発見・修正が容易

### 2. **可読性の向上**
- 1ファイルあたりの行数が削減
- 機能ごとにファイルが分かれているため理解しやすい
- クラス・メソッドの責務が明確

### 3. **再利用性**
- 各モジュールを他のプロジェクトでも利用可能
- テストが容易

### 4. **拡張性**
- 新機能の追加が容易
- 既存コードへの影響を最小限に抑えられる

## 使用方法

### 基本的な実行方法（変更なし）

```bash
python 出退勤確認システム.py
```

### 必要なパッケージ

```bash
pip install pyscard
```

### ディレクトリ構造

```
KinTouch/
├── 出退勤確認システム.py          # メインファイル（モジュール化版）
├── 出退勤確認システム_backup_before_modularization.py  # バックアップ
├── modules/                        # モジュールディレクトリ
│   ├── __init__.py
│   ├── constants.py
│   ├── database_manager.py
│   ├── card_reader_manager.py
│   ├── csv_exporter.py
│   ├── monthly_exporter.py
│   └── utils.py
├── data/                           # データディレクトリ
│   └── attendance.db
├── daily/                          # 日次集計出力
├── monthly/                        # 月次集計出力
└── reader_config.json              # リーダー設定
```

## 各モジュールの依存関係

```
出退勤確認システム.py
  ├── modules.constants
  ├── modules.database_manager
  │     └── modules.constants (JST)
  ├── modules.card_reader_manager
  ├── modules.csv_exporter
  │     └── modules.database_manager
  ├── modules.monthly_exporter
  │     ├── modules.database_manager
  │     └── modules.csv_exporter
  └── modules.utils
        └── modules.constants (CONFIG_PATH)
```

## 変更履歴

### 2026-02-05: モジュール化完了
- 元ファイル（2200行以上）を機能別に7つのモジュールに分割
- メインファイルを約1000行に縮小
- 全機能を保持したまま、コード構造を改善
- バックアップファイルを作成

## 注意事項

1. **後方互換性**: 既存のデータベース、設定ファイル、出力ファイルとの互換性を維持しています
2. **機能の変更なし**: モジュール化によって機能は一切変更されていません
3. **実行方法の変更なし**: 従来通りの方法で実行できます
4. **バックアップ**: 元のファイルは`出退勤確認システム_backup_before_modularization.py`として保存されています

## 今後の改善案

1. **ユニットテストの追加**
   - 各モジュールのテストケース作成
   
2. **ロギング機能の強化**
   - エラーログの詳細化
   
3. **設定の外部化**
   - パスワードハッシュなどの設定をコンフィグファイルに移行
   
4. **エラーハンドリングの改善**
   - より詳細なエラーメッセージ
   
5. **GUI部分のモジュール化**
   - 必要に応じて画面ごとにモジュール分割

## トラブルシューティング

### モジュールが見つからないエラー
```
ModuleNotFoundError: No module named 'modules'
```
**解決方法**: `modules/`ディレクトリが`出退勤確認システム.py`と同じディレクトリにあることを確認してください。

### インポートエラー
```
ImportError: cannot import name 'DatabaseManager'
```
**解決方法**: `modules/__init__.py`が正しく配置されていることを確認してください。

## サポート

問題が発生した場合は、バックアップファイル（`出退勤確認システム_backup_before_modularization.py`）を使用して元の状態に戻すことができます。