> 過去の開発記録です。現行手順は[インストール](../deployment.md)・[操作](../usage.md)・[検証](../testing.md)を参照してください。

# 1.0.0-rc.4 開発記録

基点7434a51（rc.3）。既存slnx Project Id/csproj SubType・CanonicalName差分を保持し、コミットしない。MSI再作成・公開・M365インストール対象外。

1. リボン統合: 実装済み。TabViewにSheasideグループ、Hidden/VeryHidden独立保存・反映。管理370チェック、両ビルド警告0・エラー0、Excel 2013連携込み524チェック成功。
2. 専用アイコン: 完了。16/32/256px透過PNGとGDI+描画原本、IPictureDispリボン適用、状態線画。管理373/Excel込み527、両ビルド警告0・エラー0。初回Excel保護解除でCOM例外、無変更で再実行成功（記録保持）。Phase 1: 1945c83。
3. パネルUI整理: 実装済み。旧フィルター/設定画面撤去、検索解除、タイトル/ブック名/一覧、新規ウィンドウ非表示。管理343、両ビルド警告0・エラー0。旧UIの不要テストを置換したため件数減。非表示Excel COMで保護解除・DisplayAlertsの不安定な例外を観測し、表示したテスト専用Excelで全497チェック成功。以後は実利用に近い表示条件で継続する。Phase 2: e999000。
4. 一覧アイコン: 実装済み。種類・状態を別アイコンにし、Visibleは状態アイコンなし。名称のみ描画、ツールチップとAccessibleDescriptionを維持。管理361、両ビルド警告0・エラー0、Excel 2013連携込み515チェック成功。Phase 3: c892260。
5. 複数/検索中ドラッグ: 完了。実タブ順とIDから表示行の挿入先を解決し、選択グループの相対順を保持。既存単一Moveと共通一括結果基盤を再利用し、直前再検証・途中中断・実順序再取得。管理368、Excel 2013連携込み530、両ビルド警告0・エラー0。Hidden/VH/Chart混在、構成保護、ウィンドウ切替拒否を追加検証。Phase 4: ff25521。
6. 総合検証・RC文書: 最終管理370チェック、Excel 2013連携込み532チェック成功。Debug/Release各警告0・エラー0。版番号1.0.0-rc.4、FileVersion 1.0.0.4。README・手動チェックリスト・ホストUI検証を更新。Phase 5: b44987b。

## コミット

| Phase | 内容 | コミット |
| --- | --- | --- |
| 1 | 表示タブ統合・表示設定保存 | 1945c83 |
| 2 | 専用ロゴ・線画アイコン | e999000 |
| 3 | パネル整理・新規ウィンドウ手動表示 | c892260 |
| 4 | 種類・状態アイコンと説明 | ff25521 |
| 5 | 複数・絞り込み中のドラッグ | b44987b |
| 6 | 総合テスト・RC文書 | 本文を含むコミット（git logで参照） |

## 最終検証と範囲

- `tools/Verify-Phase.ps1`: Release、Debug、管理テスト370、Excel 2013（15.0）連携込み532成功。連携込み件数には管理テストを含む。
- `tools/Test-ExcelUi.ps1`: 専用の未保存ブックを開き、実VSTOホストで13項目成功。接続、標準表示タブ内のグループ、新規ブック非表示、リボン開閉、追加/削除整列、旧UI撤去、Hidden/VH独立反映、検索と解除、同一ブックの新規ウィンドウ非表示・元ウィンドウの表示維持を確認。終了時に元の設定ファイルを復元する。
- Excel 2013のWinForms UI Automationは文字入力のValuePatternを公開しないため、検索は対象ハンドルに限定したWM_SETTEXT、ボタンはBM_CLICKで検証。OS全体へのキー・マウス送信は使わない。
- グラフ/通常シート、Hidden/VH混在の実COM移動・保護拒否・ウィンドウ切替拒否を検証。途中失敗の中断と結果分類は共通SheetBatchRunnerの例外注入テストで検証。
- 描画プレビューは `Sheaside/obj/SmokeTests/rc4-pane.png`（生成物、Git対象外）。線画は96/120/144 DPIの描画テストを実施。
- 自動承認サービスの利用枠エラーで最終記録作業が一時停止したが、再開後にテスト完了結果を回収し、最終ホスト検証を実施した。

## 変更ファイル

- リボン: `Sheaside/SheasideRibbon.xml`、`Sheaside/SheasideRibbon.cs`
- UI: `Sheaside/Controls/SheetListControl.cs`、`ReorderableSheetListBox.cs`、`ModernButton.cs`、新規`RibbonArtwork.cs`。`PreferencesDialog.cs`は設定画面廃止に伴い削除。
- モデル: `Sheaside/Models/WorksheetItem.cs`、`WorksheetDragEventArgs.cs`
- サービス: `Sheaside/Services/UserPreferences.cs`、`SheetPaneManager.cs`、`ExcelWorksheetService.cs`、`WorksheetMoveRules.cs`
- プロジェクト: `Sheaside/Sheaside.csproj`の描画クラス追加・廃止画面削除のみコミット。`Sheaside/Properties/AssemblyInfo.cs`のRC版番号更新。
- アイコン: `Sheaside/Assets/README.md`、`sheaside-16.png`、`sheaside-32.png`、`sheaside-256.png`
- テスト: `tests/AppearanceTests.cs`、`BatchTests.cs`、`ChartTests.cs`、`ReleaseTests.cs`、`ReorderTests.cs`、`SelectionTests.cs`、`SmokeTests.cs`
- ツール・文書: `tools/Test-Sheaside.ps1`、`tools/Test-ExcelUi.ps1`、`README.md`、本記録、`docs/rc.4-checklist.md`

## 制約・手動確認待ち

- Microsoft 365は未検証。MSIは再作成せず、GitHub Releases公開・Microsoft 365インストールも行っていない。
- パネルの×によるリボン同期、実マウスのOLEドラッグ操作、アイコンの実画面での見やすさ、実OSの各DPI・モニター間移動、スクリーンリーダー読み上げは手動確認待ち。自動テストではドラッグしきい値・イベント・ID選択と実COM移動を個別検証している。
- 絞り込み末尾は最後の表示シートの直後。除外された末尾シートより後への移動を意味しない。見えていない行は直接移動しない。
- 複数移動はExcelのトランザクションではない。失敗時の自動ロールバックはせず、成功済み変更が残る。エラー対象の実状態も一覧で確認する。
- Officeテーマ検出は既存の安全なフォールバックを維持。未知のテーマ値や取得不能時はシステム色へ戻す。
- Excel 2013の既存削除拒否は、対象保持と警告設定復元を検証する分岐がある。すべてのExcelバージョンで削除成功を保証するテストではない。
- 既存 `.slnx` Project Id、`.csproj` SubType/CanonicalNameのメタデータ差分は未コミットのまま保持。

操作確認は [rc.4-checklist.md](rc.4-checklist.md) を使用する。

## 既存メタデータ差分の整理（RC4完了後）

ユーザーから確認・修正・コミットの許可を受け、上記の保留差分を採用した。
`.slnx` のIdは `.csproj` のProjectGuidと一致、ReorderableSheetListBoxのSubTypeはComponent、HostItemのCanonicalNameは生成済みXMLのprimaryCookieとコンストラクターが使用する `AddIn` と一致している。追加のコード修正は不要と判断した。
`.vscode/settings.json` の `dotnet.preferCSharpExtension: true` も個人情報を含まない開発用設定として追加した。
採用状態でDebug/Releaseを再ビルドし、いずれも警告0・エラー0。実行コードの変更はなく、RC4で成功済みの回帰テストは再実行していない。
