> 過去の開発記録です。現行手順は[インストール](../deployment.md)・[操作](../usage.md)・[検証](../testing.md)を参照してください。

# RC5配布準備記録

- 開始: main、HEAD a3329050de5e3889456d12aff820dcb236b9571e、作業ツリークリーン。
- origin: https://github.com/nyo-ro/Sheaside.git。認証・push権限を確認。非公開設定を維持する。
- タグ/Release: v1.0.0-rc.5（開始時は同名なし）。プレリリースのみ。公開先URLは https://github.com/nyo-ro/Sheaside/releases/tag/v1.0.0-rc.5 。
- アプリケーションソースはa332905をそのまま使用。配布準備ではUI/機能を変更しない。
- AssemblyVersion/マニフェスト識別子1.0.0.0、FileVersion 1.0.0.5、InformationalVersion 1.0.0-rc.5。
- MSI: ProductVersion 0.99.5、ProductCode {834E3417-08C4-4078-A73B-6A05AF410909}。UpgradeCode {E2087C97-A6B5-4089-B8F5-86DB64EBB425}とComponent GUIDをrc.2から維持。MajorUpgrade・旧版削除・新版検出条件を検証したが、実更新は未検証。
- Release/Debugビルド警告0・エラー0。管理430、Excel 2013連携込み600チェック成功。
- MSI構造18チェック、抽出4ファイルとRelease出力のハッシュ一致。旧rc.2 MSI構造17チェック。
- `tools/Test-Distribution.ps1`：24チェック成功。ZIP展開は4ファイルの許可リスト、マニフェスト依存先のローカル参照・SHA-256・版番号・MSIL・リボンリソース・旧成果物保全を確認。
- `mage.exe -Verify`：Sheaside.vsto、Sheaside.dll.manifestとも有効な署名。MSIはNotSigned。証明書は変更せず、秘密鍵の同梱なし。
- アイコンは同梱アセンブリの描画コード、リボンXMLは埋め込みリソース。実行時に外部PNGは不要。
- 出力: artifacts/rc.5/Sheaside-1.0.0-rc.5.msi、同.zip。ZIP内はMSI、INSTALL.md、CHECKLIST.md、MSI-SHA256.txtのみ。外部SHA256SUMS.txtに両成果物のハッシュを記録。
- 既存artifacts直下のrc.2 MSI/ZIP・INSTALL.md・SHA256SUMS.txtは上書きしない。rc.2 MSI SHA-256: 5BD7B5BB93596FBFA9E8DE0D678CAF01269A7397C038B36C2C55EF37D3F8ABFD、ZIP: 7CAB5FA1B0AED9601D135D8CB783222542E16385052F251A096E27542863DC6B。
- 成果物は.gitignoreのartifacts/配下に置き、Gitへバイナリを追加しない。ハッシュと公開説明は[Release Notes](rc.5-release-notes.md)で管理。
- 開発環境のアドイン登録を保護するためMSIの実インストールは行わない。Microsoft 365・実更新・修復・アンインストールはユーザーによる別環境確認待ち。
- 中断要因: 自動承認サービスの利用枠エラーにより最終検証前に一時停止。再開後、PowerShell引数の既定値評価を修正し配布検証を完了。
