# Sheaside

Excelのシートを縦型一覧で管理するWindows向けVSTOアドインです。

- ワークシート・グラフシートの一覧と検索
- Ctrl / Shiftによる複数選択、ドラッグによる並び替え
- 追加・名前変更・複製・削除・表示状態・タブ色の管理
- 複数ブック・複数ウィンドウ、Officeテーマへの追従

## ダウンロードと対応状況

[Sheaside 1.0.0-rc.6](https://github.com/nyo-ro/Sheaside/releases/tag/v1.0.0-rc.6)が現在のプレリリースです。.NET Framework・VSTO Runtimeの検出不具合と、配布DLL内の開発PCパスを修正しました。rc.5の問題のある配布物は配布終了の対象です。正式v1.0ではありません。

デスクトップ版Excel・.NET Framework 4.8以上・VSTO Runtimeを使用します。開発検証はExcel 2013です。ユーザー実機報告により、.NET Framework 4.8.1・VSTO Runtime導入後のMSIインストール、Microsoft 365でのアドイン起動・主要機能・UIを確認済みです。旧版からの更新、修復・削除・再インストール、Office 32bit/64bit両方での検証は未完了です。Web版・Mac版Excelは対象外です。

## ドキュメント

- [インストール・更新・削除・トラブルシューティング](docs/deployment.md)
- [操作説明](docs/usage.md)
- [開発・ビルド・テスト・配布物作成](docs/development.md)
- [検証項目・既知の問題](docs/testing.md)
- [公開履歴の取り扱い](docs/publication.md)
- [過去の開発記録](docs/archive/)

MSIは未署名、VSTOマニフェストは開発証明書で署名されています。対象PCの信頼条件によって読み込みを拒否される場合があります。セキュリティ機能を無条件に無効化せず、導入手順を確認してください。
