# 開発・検証・配布

## 開発・検証

Visual Studio 2026 Community、Office/SharePoint開発ツール、.NET Framework 4.8開発ツール、VSTO Runtime、デスクトップExcelを使用します。`Sheaside.slnx`をDebug / Any CPUでビルドし、F5で起動します。

別環境でPFXがない場合は、プロジェクトの［署名］で開発用証明書を指定するか、`tools/New-DevelopmentCertificate.ps1`を実行します。PFXはGit管理しません。

```powershell
# Debugビルド後
.\tools\Test-Sheaside.ps1
.\tools\Test-Sheaside.ps1 -Excel
# 切り分け用
.\tools\Test-Sheaside.ps1 -Excel -BatchOnly
# 専用Excelを一時表示してホスト内のリボン・検索・ウィンドウ別UIを確認
.\tools\Test-ExcelUi.ps1

```


## rc.6候補の作成

```powershell
.\tools\Verify-Phase.ps1
.\tools\Test-ExcelUi.ps1
.\tools\Build-Msi.ps1 -SkipBuild
.\tools\Test-MsiPrerequisites.ps1 -Path artifacts/rc.6/Sheaside-1.0.0-rc.6.msi
.\tools\Test-Distribution.ps1
```

WiX 4.0.6を使用します。出力先はartifacts/rc.6。既存ディレクトリがある場合は上書きを拒否します。公開済みrc.5は再生成・差し替えません。ビルド時の署名は既存の開発証明書を使用し、秘密鍵はコミットしません。

製品のAssemblyVersionは1.0.0.0を維持、FileVersionは1.0.0.6、InformationalVersionは1.0.0-rc.6。MSIは0.99.6、ProductCodeは新規、UpgradeCodeは維持します。

Release構成はデバッグ情報を生成しません。DLL内のPDB絶対パスによる開発PC情報の露出を防ぎます。Debug構成のシンボルは従来どおりです。配布検証ではDLL・マニフェストにローカルユーザーパスや秘密情報パターンがないことも確認します。既存rc.2/rc.5成果物がローカルにない場合、Test-Distribution.ps1の旧成果物保全チェックは実行できないため、取得してハッシュ確認するか、未検証として扱ってください。

MSI前提条件テストは実際のMSIでAppSearch・LaunchConditionsのみ実行します。OSの.NETレジストリを書き換えず、製品をインストールしません。境界値はセッションプロパティへ投入します。実インストール、更新、アンインストールとは別の検証です。

公開前に全履歴・タグ・配布物も監査してください。GitHub noreplyを使用し、履歴や公開資産の書き換えは承認後に行います。リモートoriginとprivate-backupの混同を避けます。
