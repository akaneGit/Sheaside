> 過去の開発記録です。現行手順は[インストール](../deployment.md)・[操作](../usage.md)・[検証](../testing.md)を参照してください。

# Sheaside 1.0.0-rc.3 開発記録

基点: e025ea3（1.0.0-rc.2）。WinForms/VSTO、選択ID管理、シート操作サービス、安全性検証を維持し、UIを改善した。

## 各Phase

| Phase | 内容 | 管理 / Excel連携込み | Debug・Release | コミット |
|---|---|---:|---|---|
| 1 | 解除を外側だけに統一・適用中のみ表示、外クリック/再クリック/Esc/別ウィンドウ/ペイン非表示で閉じる | 341 / 495 | 警告0・エラー0 | 3674bb6 |
| 2 | フラット・控えめ角丸ボタン、独自線画アイコン、統一高さ・余白、ホバー/無効/選択/フォーカス表示 | 348 / 502 | 警告0・エラー0 | 577eaed |
| 3 | 設定の説明・余白・自動サイズ・最小幅、ペインと共通のOffice配色 | 363 / 517 | 警告0・エラー0 | 44cc4e5 |
| 4 | ネイティブExcelへのフォーカス移動時の閉鎖補強、総合回帰、ホストUI検証、RC文書・版数更新 | 366 / 520 | 警告0・エラー0 | 本書を含む最終コミット（git log参照） |

各Phaseのテスト・両ビルド成功後にコミット。RC2の490チェックから最終520チェックへ増加。さらに専用Excel 2013ホストUIテスト6項目に成功。

FileVersionは1.0.0.3、InformationalVersionは1.0.0-rc.3。AssemblyVersionは互換性のため1.0.0.0を維持。

## 実装上の要点

- 状態/種類の解除は外側のみ、検索文字列は維持。検索だけの場合はフィルターアイコンを強調せず解除ボタンも隠す。
- ポップアップ表示中だけIMessageFilterを登録し、終了・破棄時に解除する。クリックは消費しない。Excelのネイティブメッセージループも考慮し、ペインからフォーカスが離れた時にも閉じる。ExcelのWindowDeactivateとペインVisibleChangedでも閉じる。
- アイコンは16単位座標のオリジナル線画。追加/削除/検索/フィルター/設定をGDI+で拡大描画する。外部画像やフォントを追加せず、第三者素材のライセンス追加なし。
- 新しいModernButtonは既存Buttonを継承し、クリック・キーボード・DialogResultを維持。Officeテーマ/高コントラストに応じた描画。選択一覧の描画・ドラッグロジックは書き換えていない。
- ポップアップはコンパクトなフラット選択肢。選択はチェックと背景・文字色で表示し、狭いペインではスクロール可能。
- 設定は既存の2項目のみ。VeryHiddenの意味、Excelの［表示］→［新しいウィンドウを開く］と新規ブックの初期表示への適用、現在のペイン表示は変えないことを説明。
- 自動サイズの縮みすぎを描画プレビューで発見し、最小幅と全体収まりテストで修正。説明をAutoSizeで折り返し、125/150%相当でも全体が収まることを検証。

## 検証結果

- 管理コード366チェック成功。解除・クローズ・VeryHidden・混在種類・選択維持・設定保存/復元・テーマ・ホバー・線画描画・設定レイアウトなど。
- Excel 2013（15.0）の実COM連携込み520チェック成功。追加/削除/複製/名前変更/色/表示状態/並び替え/複数ブック・ウィンドウなど既存回帰を維持。
- Debug/Releaseビルドは両方警告0・エラー0。
- Excel 2013の専用の一時ブックでVSTO接続、実ペインの追加/削除ボタンの高さ一致、フィルターを開く、表示条件を適用、外側の解除でリセット、再クリックで閉じる、計6項目成功。テスト用Excelを終了しブックは保存していない。
- Excel 2013のUI AutomationではWinFormsがPaneとして公開され、InvokePattern/AccessibleNameによる取得ができなかったため、対象Excel内のWinFormsネイティブボタンに限定してBM_CLICKを送る。タイムアウト付き。コードはtools/Test-ExcelUi.ps1に保存。
- 線画は96/120/144 DPIのビットマップ描画を検証。設定は画面外でフォームを生成し100/125/150%相当のScaleを適用して、説明・全体・チェックボックス位置・配色を検証。これはOS表示倍率切替や異なるDPIのモニター間移動の実機確認の代替ではない。
- フィルターと黒テーマ設定の描画プレビューを目視確認。テスト時の画像はSheaside/obj/SmokeTests/rc3-filter.png、rc3-settings.png。Git対象外。

## 未確認・既知の制約

- Excel 2013ホスト上で実マウスによる範囲選択・ドラッグ・ペインリサイズ・セル外クリック・Escの全組合せ、設定画面の実OS倍率125/150%はユーザー確認待ち。管理コード・COMテストが通ったことと区別する。
- Microsoft 365は今回全項目未確認。Office 16のテーマ、32/64ビット、実DPI・マルチモニター、ネイティブのフォーカス/外クリックを確認する。
- Officeテーマ検出はrc.2と同じベストエフォートで、取得不可時Windows標準色、高コントラスト優先。設定画面は開いた時点のテーマを使用する。OS標準タイトルバーの色はWinFormsのクライアント配色とは別。
- Excel 2013の既知のHidden複製後/複数ウィンドウ後のグラフ削除拒否は、既存テストで対象保持・警告設定復元を確認する分岐を継続。Microsoft 365の正常削除成功と同一視しない。
- MSIの変更・生成、Microsoft 365へのインストール、GitHub Releases公開、自動更新は実施していない。既存のinstaller、Build-Msi、artifactsはRC2のまま。今回のRC3をそれらのRC2設定で配布しない。

## 変更ファイル一覧

- Sheaside/Controls/SheetListControl.cs
- Sheaside/Controls/ModernButton.cs（新規、共通配色と線画ボタン）
- Sheaside/Controls/PreferencesDialog.cs
- Sheaside/Services/SheetPaneManager.cs
- Sheaside/Properties/AssemblyInfo.cs
- Sheaside/Sheaside.csproj（ModernButton.csのCompile登録のみ）
- tests/ReleaseTests.cs
- tools/Test-Sheaside.ps1
- tools/Test-ExcelUi.ps1（新規）
- README.md
- docs/rc.3-progress.md（本書、新規）
- docs/rc.3-checklist.md（新規）

開始時からのSheaside.slnxのProject IdとSheaside.csprojのSubType/CanonicalName差分は保持し、コミットしていない。

## 再確認

```powershell
.\tools\Verify-Phase.ps1
# Debugの開発登録・初期表示ONが必要。新しいExcelを一時表示して閉じる。
.\tools\Test-ExcelUi.ps1
```

[Excel 2013操作確認とMicrosoft 365の確認待ち項目](rc.3-checklist.md)に沿って確認する。Excel 2013の操作性確認後、配布・Microsoft 365検証へ進む。

デザイン参考: [Microsoftのアイコン設計ガイド](https://learn.microsoft.com/en-us/windows/apps/develop/ui/controls/icons)。簡潔で意味が伝わるアイコンと拡大描画の方針を参考にし、外部素材を転載していない。
