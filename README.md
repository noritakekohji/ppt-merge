# PPT Merge Tool

ディレクトリ内の複数の PowerPoint ファイル(`.pptx`)を、指定した順番で 1 つにマージし、続けて PDF へ変換する GUI ツール。

## 概要

- フォルダ内の pptx を一覧表示し、上下移動で順番を決めてマージ
- マージ後に同名の PDF を自動生成(ON/OFF 可能)
- 各ファイル元のデザイン(テーマ・配色・フォント)を保持

## 動作要件

- Windows + **Windows PowerShell 5.1**(標準同梱)
- **Microsoft PowerPoint** がインストールされていること(COM automation を利用)

## セットアップ / 使い方

1. リポジトリを任意のフォルダに配置
2. `PptMerge.bat` をダブルクリックして起動
3. 「フォルダから一括追加」または「ファイルを個別追加」で pptx を追加(ウィンドウへのドラッグ&ドロップも可)
4. リストの上下移動・チェックで順番と対象を調整
5. 出力先フォルダとファイル名を指定し、「マージ実行」

## テスト

```powershell
Invoke-Pester tests/PptMergeCore.Tests.ps1
```

## バージョン

現在のバージョン: 0.1.0 — 変更履歴は [CHANGELOG.md](CHANGELOG.md) を参照。

## ライセンス

社内利用。
