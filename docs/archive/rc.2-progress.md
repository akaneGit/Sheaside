> 過去の開発記録です。現行手順は[インストール](../deployment.md)・[操作](../usage.md)・[検証](../testing.md)を参照してください。

# Sheaside 1.0.0-rc.2 開発・検証記録

基点はaaeeb7c（1.0.0-rc.1）。WinForms/VSTOと既存のサービス・ID選択管理・バッチ安全性を維持した。Microsoft 365実機確認前のRCであり、正式版ではない。

## Phase別結果

| Phase | 実装内容 | 管理コード / Excel込み | Debug / Release | コミット |
|---|---|---:|---|---|
| 1 | 指定メニュー順、横並び追加・削除、成功メッセージ撤廃 | 315 / 462 | 警告0・エラー0 | e33124f |
| 2 | ペイン内フィルターポップアップ、適用色・条件・解除、検索分離、VeryHidden連動 | 321 / 468 | 警告0・エラー0 | 120130d |
| 3 | テーマ10色×6段階、標準色、色なし、任意色。ThemeColor/TintAndShade保持、取得不可RGB代替 | 328 / 479 | 警告0・エラー0 | 6198201 |
| 4 | Officeテーマ検出、旧テーマ/配置設定とUIを削除、旧設定読み込み互換 | 334 / 485 | 警告0・エラー0 | 47e54a3 |
| 5 | ユーザー単位MSI・配布ZIP、前提条件、削除/更新登録、MSI検証・手順 | 334 / 485 | 警告0・エラー0 | ac9b12b |
| 6 | RC2バージョン、テーマ変更追従・パレット操作追加テスト、全回帰、配布物再生成、実機チェックリスト | 336 / 490 | 警告0・エラー0 | 本記録を含む最終コミット（git log参照） |

各Phaseで管理テストとExcel 2013（15.0）の実COM連携、Debug/Releaseビルド後にコミット。最終FileVersion=1.0.0.2、InformationalVersion=1.0.0-rc.2、AssemblyVersion=1.0.0.0（バインディング互換維持）。MSI内部バージョンは0.99.2、正式版1.0.0へ増加できる。

## MSI成果物・検証

`artifacts/Sheaside-1.0.0-rc.2.msi`と同名ZIPを生成。ZIPにはMSI・INSTALL.md・SHA256SUMS.txtを収録。GitHub公開は未実施。

WiX 4.0.6のMSIビルド・ICE検証成功。MSI読み取り専用検証14項目成功。展開した実ペイロード4ファイルのSHA256がRelease出力と一致。MSI生成はユーザー環境へのインストールを行わない。

実インストール・修復・削除・再インストール・次版への更新・Office 32/64ビット双方でのVSTOホスト起動は未実施。Microsoft 365チェックリストで確認する。MSIは未署名、VSTOは開発証明書署名であり、配布先の信頼確認・企業ポリシーでロードが拒否される可能性がある。一般公開用証明書・自動証明書信頼・全ユーザー導入・ARM版Office対応は今回の成果物に含まない。

## 設計・互換性・制約

- フィルター解除は状態/種類だけを解除し、検索文字は残す。検索を消せば全解除。フィルターで消えた行のIDは選択集合から取り除く。
- VeryHidden表示OFFの対象を一覧・検索・操作に含めない。初期状態はOFF。旧設定のtheme/dockLeftは無視し、次回の設定保存で出力しない。
- Officeテーマは公開された安定的な取得/通知APIを使用できないため、HKCUのOfficeバージョン別`Common/UI Theme`をベストエフォートで読む。Office 15の白/明るい灰色/濃い灰色、16のカラフル/白/濃い灰色/黒/システムに対応。未知値・読取拒否はWindows標準色。高コントラスト優先。既存の前面・操作可能・表示中に限る3秒更新で再取得し、変更時のみ配色を更新する。アカウント同期やポリシーで値が得られない場合、完全なOffice追従は保証しない。
- カラフルテーマは一覧の明るい背景・文字色を合わせる基本対応。Officeのリボン描画を複製するものではない。パレットのプレビューはHSL近似、適用後の一覧はExcelの実RGB。ブックテーマが取得できない場合は固定RGBの代替色。
- テーマ色と任意RGBは区別し、同じ見た目でもRGB選択はテーマリンクを解除する。WorksheetとChartを同じアダプターで操作する。
- パネルの初期位置は右。Excel標準のドラッグで移動し、設定変更で強制再配置しない。手動位置をアドイン独自に永続保存しない。
- Excel 2013の一部削除拒否（Hidden複製後、複数ウィンドウ操作後のChart削除）は既存の条件付きテストで対象保持・DisplayAlerts復元を検証。正常削除の代替成功とは扱わずMicrosoft 365確認待ち。
- セル編集・COM切断・保護・読み取り専用・対象変更を既存の検証/例外処理で拒否。部分成功はロールバックせず報告、残りは中断。Excelグループ化は使わない。
- 旧ClickOnceスクリプト・プロジェクト資産は無断削除せず保存。新配布経路はMSIに一本化。MSIでも必要なVSTOマニフェストは維持。

## 変更ファイル一覧（rc.1比）

- UI: `Sheaside/Controls/SheetListControl.cs`、`ColorPaletteDialog.cs`（新規）、`PreferencesDialog.cs`。
- モデル: `Sheaside/Models/TabColorChoice.cs`（新規）。
- サービス: `Sheaside/Services/ExcelWorksheetService.cs`、`ExcelSheetBatchService.cs`、`SheetPaneManager.cs`、`UserPreferences.cs`、`OfficeThemeReader.cs`（新規）。
- プロジェクト/版数: `Sheaside/Sheaside.csproj`（新規ソースのCompile登録だけ）、`Sheaside/Properties/AssemblyInfo.cs`。
- テスト: `tests/SmokeTests.cs`、`BatchTests.cs`、`AppearanceTests.cs`、`ReleaseTests.cs`、`SelectionTests.cs`、`tools/Test-Sheaside.ps1`、`tools/Verify-Phase.ps1`（新規）。
- 配布: `.gitignore`、`installer/Sheaside.wxs`（新規）、`tools/Build-Msi.ps1`（新規）、`tools/Test-Msi.ps1`（新規）。
- 文書: `README.md`、`docs/deployment.md`、`docs/rc.2-progress.md`（本書）、`docs/rc.2-checklist.md`（新規）。

開始時から存在した`.slnx`のProject Idと`.csproj`のSubType/CanonicalNameメタデータ差分は、保持したままコミットから除外した。

## 再現と残る実機確認

`tools/Verify-Phase.ps1`で両ビルド・管理テスト・Excel連携を再現できる。`tools/Build-Msi.ps1`でMSI/ZIPと検証を再生成する。開発ビルドはExcelの開発用登録を更新するので、MSI確認は別PC/ユーザーで行う。

[Microsoft 365チェックリスト](rc.2-checklist.md)はすべて未実施として作成した。[配布手順](../deployment.md)に従って導入し、記録欄に結果を残す。実機結果を踏まえて正式v1.0への移行を判断する。

技術資料: [VSTO MSI配布](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer)、[VSTO登録](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins)、[Tab.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.themecolor)、[Tab.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.tintandshade)。
