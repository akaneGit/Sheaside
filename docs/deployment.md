# Sheaside 1.0.0-rc.2 インストール・配布

現在の配布形式はMSIです。自動更新は行わず、GitHub Releasesから手動で取得します。リポジトリ内の旧ClickOnce発行スクリプト・プロジェクト設定は削除せず保存していますが、今後の配布には使いません。`.vsto`と`.dll.manifest`はMSI方式でもVSTOが必要とするファイルなので引き続き同梱します。

## 利用環境と前提条件

- WindowsのデスクトップExcel。開発連携テストはExcel 2013、配布先Microsoft 365の実機確認は未実施です。
- [.NET Framework 4.8以上](https://learn.microsoft.com/en-us/dotnet/framework/install/on-windows-10)。MSIはRelease値528040以上を確認します。
- [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=105522)。CLR 4対応ランタイムをOSに合わせて導入してください。MSIは32/64ビット両方の登録場所を確認します。前提ソフトは自動ダウンロードしません。
- Excelを含むOfficeアプリケーションを終了してから導入・修復・更新・削除してください。

MSI自体はx86ですが、アドインはAny CPU、登録はHKCUの共有領域です。同じMSIで32/64ビットOfficeを想定しています。ARM版Office、全ユーザー一括導入は対象外です。両ビット数のMicrosoft 365実機確認が完了するまで互換性保証は行いません。

## インストール

1. GitHub Releasesの`Sheaside-1.0.0-rc.2.zip`を展開するか、MSIと本手順を取得します。公開操作は別途実施します。
2. `Get-FileHash .\Sheaside-1.0.0-rc.2.msi -Algorithm SHA256`で配布元のSHA256SUMS.txtと照合します。
3. 旧ClickOnce版がある場合は先にWindowsの「インストールされているアプリ」からアンインストールします。異なる配布形式の自動移行は行いません。
4. MSIをダブルクリックします。必要なら次のログ付きコマンドを使用します。ユーザー単位のため通常は管理者起動しません。

```powershell
msiexec /i "Sheaside-1.0.0-rc.2.msi" /L*v "Sheaside-install.log"
```

5. Excelを起動し、VSTOの信頼確認が出た場合は配布元を確認して判断します。リボンのSheasideと右側の一覧を確認してください。

配置先は`%LOCALAPPDATA%\Programs\Sheaside`、登録先は`HKCU\Software\Microsoft\Office\Excel\Addins\Sheaside`です。設定は`%LOCALAPPDATA%\Sheaside\settings.xml`に保持します。開発PCではVisual Studioビルドが同じ登録先を上書きするため、MSIの実機確認には別PC/別ユーザーを推奨します。

## 修復・削除・再インストール

同一版を修復する場合はWindowsのアプリ管理から修復、または次のコマンドを使用します。

```powershell
msiexec /fa "Sheaside-1.0.0-rc.2.msi" /L*v "Sheaside-repair.log"
msiexec /x "Sheaside-1.0.0-rc.2.msi" /L*v "Sheaside-uninstall.log"
```

アンインストールではアドイン本体とMSIが登録したExcelアドイン値を削除します。ユーザー設定とExcelブックは残します。削除後に同じMSIを再実行すれば再インストールできます。現時点では修復・削除・再インストールの実動作はMicrosoft 365確認待ちです。

## 手動更新

Excelを終了し、新バージョンのMSIをダウンロードして実行します。固定UpgradeCodeにより旧MSIを置き換え、設定を保持します。新しい版が入っている場合は古い版への上書きを拒否します。RC2のMSI内部バージョンは`0.99.2`（表示名`1.0.0-rc.2`）。正式版は`1.0.0`以上とし、将来の各リリースではMSIの数値バージョンを増加・ProductCodeを更新し、UpgradeCodeと既存Component GUIDを維持してください。同じリリース番号の異なるMSIを公開し直さないでください。

## 署名と制約

現在のMSIはAuthenticode未署名です。VSTOマニフェストはローカル開発証明書署名で、初回に発行元不明の確認が出る場合があります。企業ポリシーによってはロードを拒否されます。MSIは証明書を自動的に信頼ストアへ追加せず、セキュリティ設定も緩和しません。一般公開前には配布用コード署名証明書によるマニフェスト・MSI署名を整備してください。秘密鍵・PFXはGitや配布ZIPに含めません。

MSI生成・ICE検証・データベース検証の成功は、別PCのインストール成功や証明書の信頼成立を意味しません。Microsoft 365の実機確認が必要です。Excel不在の環境をMSIは自動検出しません。Officeとランタイムの導入後に実行してください。

## 開発者向けビルド

```powershell
.\tools\Verify-Phase.ps1
.\tools\Build-Msi.ps1
.\tools\Test-Msi.ps1 -Path .\artifacts\Sheaside-1.0.0-rc.2.msi
```

Visual StudioのVSTO開発ツール・署名用開発証明書に加えて.NET SDKと初回のNuGet接続が必要です。WiX 4.0.6を`Sheaside/obj/wix`に固定バージョンで導入します。`artifacts`にMSI、INSTALL.md、SHA256SUMS.txt、配布ZIPを生成します。成果物はGit対象外です。GitHub Releasesへのアップロードは自動では実施しません。

MSIにはSheaside.dll、Sheaside.dll.manifest、Sheaside.vsto、Microsoft.Office.Tools.Common.v4.0.Utilities.dllの4ファイルを含めます。Excel Interopは型埋め込みを使用し、開発用PIAやPDB、証明書秘密鍵は含めません。

## 設計の根拠

- [Microsoft: VSTOのWindows Installer配布・前提条件・信頼](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer)
- [Microsoft: VSTOアドインのレジストリ登録](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins) — Manifestは`file:///[INSTALLFOLDER]Sheaside.vsto|vstolocal`とし、ClickOnceキャッシュへ転送しません。
- [Microsoft: .NET Frameworkの検出](https://learn.microsoft.com/en-us/dotnet/framework/install/how-to-determine-which-versions-are-installed)
