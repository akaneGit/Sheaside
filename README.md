# Sheaside

Excelのシートを縦型一覧で管理するWindows向けVSTOアドインです。

- ワークシート・グラフシートの一覧と検索
- Ctrl / Shiftによる複数選択、ドラッグによる並び替え
- 追加・名前変更・複製・削除・表示状態・タブ色の管理
- 複数ブック・複数ウィンドウ、Officeテーマへの追従

## ダウンロードと対応状況

[Sheaside 1.0.0](https://github.com/nyo-ro/Sheaside/releases/tag/v1.0.0)が正式版です。rc.6の既存機能と、.NET/VSTO検出・配布DLLの個人パス対策を維持しています。

Windowsデスクトップ版Excel・.NET Framework 4.8以上・Visual Studio Tools for Office Runtimeが必要です。Excel 2013で自動検証しています。Microsoft 365 **64bit版**ではrc.6のインストール・起動・主要機能・UI・アンインストール・再インストールをユーザー実機確認済みです。新しい1.0.0 MSIのMicrosoft 365実インストールは未実施です。

**更新時はExcelをすべて終了し、既存Sheasideをアンインストールしてから新しいMSIを実行してください。** 旧版からの直接更新は保証せず、同一ユーザーの旧MSIを検出した場合はインストーラーが削除を案内して停止します。

32bit版Officeは利用禁止ではありませんが、今回の対応保証・実機検証の対象外です。その他の未検証Office環境、Web版・Mac版Excelは動作保証しません。

## ドキュメント

- [インストール・更新・削除・トラブルシューティング](docs/deployment.md)
- [操作説明](docs/usage.md)
- [開発・ビルド・テスト・配布物作成](docs/development.md)
- [検証項目・既知の問題](docs/testing.md)
- [公開履歴の取り扱い](docs/publication.md)
- [過去の開発記録](docs/archive/)

MSIは未署名、VSTOマニフェストは開発証明書で署名されています。対象PCの信頼条件によって読み込みを拒否される場合があります。セキュリティ機能を無条件に無効化せず、導入手順を確認してください。
