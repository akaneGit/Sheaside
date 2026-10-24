# Sheaside 1.0.0-rc.5 開発・検証記録

基点 `b0e16b3`（RC4と承認済みメタデータ整理）。開始時の作業ツリーはクリーン。WinForms/VSTO・既存シート管理・リボン構成を維持する。MSI再作成、GitHub Releases公開、Microsoft 365インストールは対象外。

## 段階別変更

| Phase | 変更・原因 | 検証 | コミット |
| --- | --- | --- | --- |
| 1 | 非表示の選択時にActivateWorksheetが案内文字列を返し、ShowStatusが通知していた。Hidden/VHはnullを返し、ローカル選択のみとする。エラー通知と明示的な再表示は維持 | 管理370、Excel込み534、Debug/Release警告0・エラー0 | eb8431e |
| 2 | 標準色だけFlowLayoutPanelにDock=Fillボタンを入れていた不整合をTableLayoutPanelへ統一。固定330px高を廃止し内容に応じたAutoSize。標準10色をExcel 2013風の固定RGBへ整理 | 管理379、100/125/150%相当で全色ボタンの非ゼロ寸法・領域内配置・スクロールなし・選択結果 | dc45279 |
| 3 | 黒枠の四角を左端4論理pxの縦バーへ変更。色なしは描画せず、低コントラスト時のみ細いグレーの補助を追加 | 管理391、黒/白/淡色/色なし×明暗/選択背景 | 768cbf3 |
| 4 | 通常アイコンは明背景で濃いグレー、暗背景で明るいグレー。選択・高コントラストはシステム色。状態を名前の右へ移動し、個別シート保護の盾を追加 | 管理397、Excel込み567、Debug/Release警告0・エラー0。Worksheet/Chart保護・解除、ブック構成保護との区別 | ad0d041 |
| 5 | DPIに応じた行高・アイコン間隔・共通配置計算。横スクロールを廃止し名前を省略表示。種類/状態/保護それぞれの日本語ツールチップ、名前全文の説明 | 管理427、180/260/400論理px×100/125/150%の領域・整列・ツールチップ、既存選択/ドラッグ回帰 | 2095638 |
| 6 | RC5版番号・最終回帰・ホストUI検証・README/チェックリスト | 最終結果は下記 | 本記録を含むコミット |

## 保護状態の判定

WorksheetはProtectContents、ProtectDrawingObjects、ProtectScenariosのいずれか。ChartはProtectContents、ProtectDrawingObjects、ProtectData、ProtectFormatting、ProtectSelectionのいずれかで盾を表示する。Workbook.ProtectStructureは参照しない。更新時に再取得し、保護状態だけの変更でも行を更新する。

型ごとのAPIは[MicrosoftのChart Interop参照](https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.chart?view=excel-pia)と[Chart.ProtectData](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectdata)を確認し、Excel 2013で保護・解除の取得を検証した。シートの保護/解除操作自体は今回追加していない。

## 最終検証

- Debug/Release：警告0・エラー0で成功。
- 管理テスト：428チェック成功。色なし選択も追加。
- 全Excel連携回帰：Excel 2013（15.0）で598チェック成功（管理428を含む）。既存の単一/複数操作・ドラッグ・複数ブック/ウィンドウ・色適用の回帰を含む。
- `tools/Test-ExcelUi.ps1`：Excel 2013（15.0）実VSTOホスト15項目成功。RC4のリボン・表示設定・検索・複数ウィンドウ13項目に加え、Hidden/VHそれぞれの実クリックを検証。同期メッセージが通知でブロックせず、ローカル選択1件・ExcelのVisibleProbeアクティブ維持を確認。
- パレットの実WinForms表示：通常環境ClientSize 326×372、縦横Scroll=False、最下部キャンセルを含め全内容が領域内。生成したフォーム全体・一覧プレビューを目視確認。
- UIプレビュー：`Sheaside/obj/SmokeTests/rc5-pane.png`、`rc5-palette.png`（生成物のため非コミット）。一覧プレビューは選択行と別のアクティブ行、色バー、長い名前、Hidden/VH、盾を含む。
- テストは専用の未保存Excelブックを使用。ホストUIテストで変更した表示設定は終了時に元ファイルへ戻す。既存ユーザーブックや既存Excelプロセスは変更・終了しない。

## 変更ファイル

- UI：`Sheaside/Controls/ColorPaletteDialog.cs`、`ModernButton.cs`、`ReorderableSheetListBox.cs`、`SheetListControl.cs`
- モデル：`Sheaside/Models/WorksheetItem.cs`
- Excel操作：`Sheaside/Services/ExcelSheet.cs`、`ExcelWorksheetService.cs`
- 版番号：`Sheaside/Properties/AssemblyInfo.cs`（InformationalVersion 1.0.0-rc.5、FileVersion 1.0.0.5）
- テスト：`tests/AppearanceTests.cs`、`ChartTests.cs`、`ReleaseTests.cs`、`tools/Test-ExcelUi.ps1`
- 文書：`README.md`、本記録、`docs/rc.5-checklist.md`

## 制約・確認待ち

- Microsoft 365は未検証。MSI作成・公開・インストールは今回実施しない。
- 125/150%はWinForms Scaleと描画/座標検証。実OS倍率切替、モニター間DPI変更、各Officeテーマの実画面での読みやすさは手動確認待ち。
- パレット全体は実WinFormsフォームで確認し、色適用はExcel COM回帰で検証。Excelホストの右クリックからパレットを開き色を選ぶ一連の実マウス操作は手動確認待ち。
- 標準10色は固定RGB。Microsoft 365の将来の標準色変更へ自動追従する仕様ではない。テーマ色・濃淡は従来どおりブックから取得する。
- 極端に小さな作業領域・極端に狭いパネルは手動確認対象。通常の180論理px以上で名前・右側アイコンの分離を検証済み。
- シート保護の表示は読み取りのみ。保護されたシートへの操作可否は既存の検証とExcel自身の拒否に従う。
- Excel 2013の既存削除拒否は対象保持・DisplayAlerts復元を検証する分岐を含む。全バージョンでの削除成功を意味しない。
- 実ドラッグ・スクリーンリーダー・パネル×同期などRC4からの手動確認項目は引き続き確認する。

## RC5追加修正：タイトルと上部レイアウト

タスクペインのタイトルを「シート一覧」へ変更。内部のSheasideラベル（28px）・ブック名ラベルとコンテナ（26px）を削除し、検索欄を最上段へ移動した。通常倍率で54pxを一覧の高さへ再配分。共通の外周余白は維持する。

リボン名・ボタン名・ブック/ウィンドウ管理・版番号は変更していない。変更は `Controls/SheetListControl.cs`、`Services/SheetPaneManager.cs`、`tests/ReleaseTests.cs`、`tools/Test-ExcelUi.ps1` と本記録。

- Debug/Release：両方成功、警告0・エラー0。
- 管理回帰：430チェック成功。タイトル用コントロールがなく、検索→操作ボタン→一覧が余分な領域なしに配置されること、残り高さを一覧が使うことを追加検証。
- Excel 2013実VSTOホスト：21項目成功。タイトル、内部ラベル撤去、同一ブック別ウィンドウ、別ブックの初期非表示・一覧切替・追加先の分離・元ブック復帰を確認。
- 実行ロジックを変更しないUI修正のため、今回は管理回帰と実ホストの関連検証を実施し、全COM操作回帰の再実行は省略。RC5本体の598チェック成功記録は上記参照。
