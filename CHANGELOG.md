# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- 「名前順ソート」ボタンを追加(チェック状態を保持したままファイル名昇順に並べ替え)
- ファイルリストで複数選択しての上下移動・削除に対応
- 対象フォルダのパスを直接入力して一括追加できる入力欄と「パス追加」ボタンを追加(Enter キーでも追加可能)

### Fixed
- マージ時の `Slides.InsertFromFile` で `SlideEnd=0` を渡していたため「範囲外の整数 0」エラーで失敗していた問題を修正(引数を省略し全スライドを挿入)
- マージ結果のレイアウトが崩れる問題を修正。空の新規プレゼンに挿入する方式をやめ、先頭ファイルをベースとして開き、その元サイズ・マスター・デザインを基準にするよう変更

### Changed
- ファイルリストを `CheckedListBox` から複数選択対応の `ListView`(チェックボックス付き)に変更。ファイル名クリックは選択のみ、チェックはチェックボックスのクリックまたは行のダブルクリックで切り替え
- 出力先フォルダが存在しない場合、作成するか確認するようにした(直接パス入力に対応)
- 起動時、出力先フォルダが未設定ならデスクトップを初期値として設定

## [0.1.0] - 2026-06-11

### Added
- 初回リリース
- フォルダ一括 / 個別ファイル / ドラッグ&ドロップによる pptx 追加
- CheckedListBox + 上下移動ボタンによるマージ順序指定
- 各ファイル元デザインを保持したマージ
- マージ後の PDF 変換(ON/OFF 切替可能)
- 設定の記憶(出力先フォルダ・PDF 作成設定)
- 出力ファイル名の自動初期値(merged_yyyyMMdd)
