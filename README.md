# Sheaside

Excelのシートを縦型一覧で管理するWindows向けVSTOアドインです。

- ワークシート・グラフシートの一覧と検索
- Ctrl / Shiftによる複数選択、ドラッグによる並び替え
- 追加・名前変更・複製・削除・表示状態・タブ色の管理
- 複数ブック・複数ウィンドウ、Officeテーマへの追従

## ダウンロードと対応状況

[GitHub Releases](https://github.com/nyo-ro/Sheaside/releases)でRC版を配布しています。公開済みrc.5には.NET Framework 4.8.1が入っていてもインストールを拒否する既知の問題があります。修正候補rc.6は準備中で、まだ公開していません。正式v1.0ではありません。

デスクトップ版Excel・.NET Framework 4.8以上・VSTO Runtimeを使用します。開発検証はExcel 2013、Microsoft 365でのrc.6導入確認は未実施です。Web版・Mac版Excelは対象外です。

## ドキュメント

- [インストール・更新・削除・トラブルシューティング](docs/deployment.md)
- [操作説明](docs/usage.md)
- [開発・ビルド・テスト・配布物作成](docs/development.md)
- [検証項目・既知の問題](docs/testing.md)
- [公開履歴の取り扱い](docs/publication.md)
- [過去の開発記録](docs/archive/)

MSIは未署名、VSTOマニフェストは開発証明書で署名されています。対象PCの信頼条件によって読み込みを拒否される場合があります。セキュリティ機能を無条件に無効化せず、導入手順を確認してください。
