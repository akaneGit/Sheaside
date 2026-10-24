# Sheaside 1.0.0-rc.5 インストール・配布

Excelのシートを縦一覧で管理するWindows向けVSTOアドインです。今回の公開はプレリリースです。Microsoft 365での実インストール・更新・動作は未検証です。

## ダウンロードと前提条件

公開先: https://github.com/akaneGit/Sheaside/releases/tag/v1.0.0-rc.5
リポジトリは非公開のため、閲覧権限のあるGitHubアカウントでログインしてください。

- WindowsのデスクトップExcel（導入予定: Microsoft 365）。Web版Excelは対象外。
- [.NET Framework 4.8以上](https://learn.microsoft.com/en-us/dotnet/framework/install/on-windows-10)。MSIはRelease値528040以上を確認します。
- [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=105522)のCLR 4対応版。OS/Officeに適したMicrosoft公式ランタイムを導入してください。MSIは前提ソフトを自動取得しません。
- MSIはx86・ユーザー単位、アドインはAny CPU。HKCUへの登録により32/64bit Officeを想定しますが、両ビット数のMicrosoft 365検証は未実施です。ARM版Office・全ユーザー一括導入は対象外です。

## インストール

1. 作業中のブックを保存し、Excelをすべて終了します。
2. Releaseから `Sheaside-1.0.0-rc.5.zip` を取得し、新しいフォルダーへ展開します。MSI単体でも導入可能ですが、ZIPにはこの手順とCHECKLIST.mdが入っています。
3. Release NotesのSHA-256とファイルを照合します。ZIP内のMSI-SHA256.txtでもMSIを照合できます。

```powershell
Get-FileHash .\Sheaside-1.0.0-rc.5.zip -Algorithm SHA256
Get-FileHash .\Sheaside-1.0.0-rc.5.msi -Algorithm SHA256
msiexec /i "Sheaside-1.0.0-rc.5.msi" /L*v "Sheaside-install.log"
```

4. MSIを実行します。通常は管理者として起動せず、Excelを使用するユーザーで実行してください。旧ClickOnce版は先にWindowsのアプリ管理から削除します。MSIとClickOnceの自動移行は行いません。
5. Excelを起動します。リボン［表示］→［Sheaside］→［シート一覧］でパネルを開きます。新しいExcelウィンドウではパネルは自動表示されません。パネルのタイトルは「シート一覧」です。
6. ［非表示も表示］［VeryHiddenも表示］は独立した切り替えです。初期値はOFF、既存設定がある場合は引き継ぎます。

配置先: `%LOCALAPPDATA%\Programs\Sheaside`
登録先: `HKCU\Software\Microsoft\Office\Excel\Addins\Sheaside`
ユーザー設定: `%LOCALAPPDATA%\Sheaside\settings.xml`

アドインが表示されない場合はExcelの［ファイル→オプション→アドイン］からCOMアドインと使用できないアイテムを確認し、読み込みエラーとインストールログを記録してください。Officeのセキュリティ設定を無条件に緩めないでください。

## 署名と信頼条件

MSIはAuthenticode未署名です。VSTOの2つのマニフェストは既存の開発証明書で署名されています。署名の暗号的な有効性と、導入先で発行元を信頼できることは別です。

- 証明書: `CN=Sheaside Local Development`
- 拇印（SHA-1識別子）: `C61660B9D07850DE0D156C9127355F4124B5093A`
- 有効期限: 2027-10-22（署名タイムスタンプの有無や導入先ポリシーによる扱いは別途確認が必要）

導入先では発行元不明の確認や信頼エラーが発生し、企業ポリシーによって読み込みが拒否される可能性があります。MSIは証明書を信頼ストアへ自動登録せず、Office/SmartScreen/ウイルス対策の設定も変更しません。配布元・ハッシュ・証明書を確認し、組織端末では管理者の承認を受けて信頼条件を整備してください。警告が出た場合の無条件な実行・保護無効化は案内しません。正式配布用コード署名証明書による署名は今後の課題です。秘密鍵・PFX・パスワードは配布しません。

## rc.2からの手動更新

Excelを終了してrc.5のMSIを実行します。同じユーザーにインストールされた旧MSIを検出し、置き換える構成です。

| 項目 | rc.2 | rc.5 |
| --- | --- | --- |
| MSI ProductVersion | 0.99.2 | 0.99.5 |
| ProductCode | {5E6A06B4-0794-49AD-8F9B-5084BBD8C2E2} | {834E3417-08C4-4078-A73B-6A05AF410909} |
| UpgradeCode | {E2087C97-A6B5-4089-B8F5-86DB64EBB425} | 同一 |

MajorUpgradeで旧版を検出し、新版導入済み環境へのダウングレードを拒否します。Component GUIDは維持しています。構造検証は実更新成功を意味しません。rc.2からの実更新・設定引継ぎは未検証です。旧版が別ユーザーや別配布方式で登録されている場合は自動移行を期待しないでください。同じrc.5の別バイナリでの上書き公開は行いません。

## アンインストール・再インストール

Excelを終了し、Windowsのアプリ管理からSheasideを削除するか、以下を実行します。

```powershell
msiexec /x "Sheaside-1.0.0-rc.5.msi" /L*v "Sheaside-uninstall.log"
# 同じ版の修復
msiexec /fa "Sheaside-1.0.0-rc.5.msi" /L*v "Sheaside-repair.log"
```

本体とMSIが登録したアドイン情報を削除し、ユーザー設定とExcelブックは残す構成です。削除後に同じMSIを実行して再インストールできます。実際の削除・修復・再インストールは未検証です。

## 配布物と既知の制約

MSIにはSheaside.dll、Sheaside.dll.manifest、Sheaside.vsto、Microsoft.Office.Tools.Common.v4.0.Utilities.dllの4ファイルを格納します。リボンXMLはアセンブリ内、アイコンは同梱コードから描画するため別PNGは実行時に不要です。Excel Interopは型埋め込みです。開発用PIA・PDB・秘密鍵・中間生成物は含めません。

`.vsto`と`.dll.manifest`はMSIでも必要なVSTOマニフェストです。登録には`|vstolocal`を使用し、ClickOnce配布や自動更新は使用しません。既存のClickOnce関連設定・スクリプトは削除せず保存しています。

開発PCはVisual StudioとMSIが同じアドイン登録先を使用するため、実インストール検証は別PC/別ユーザーを推奨します。今回は開発環境の登録を保護し、MSIの実インストールは行っていません。Microsoft 365・Officeビット数・実DPI・企業ポリシーによる差はCHECKLIST.mdで確認してください。

## 再生成（開発者向け）

`tools/Build-Msi.ps1`はWiX 4.0.6を再利用し、`artifacts/rc.5`へ出力します。既存出力フォルダーがある場合は中止し、既存成果物を上書きしません。配布バイナリはGit管理せずGitHub Releasesへ添付します。

参考: [MicrosoftのWindows Installer配布](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer)、[VSTOの信頼](https://learn.microsoft.com/en-us/visualstudio/vsto/granting-trust-to-office-solutions)、[Major Upgrade](https://learn.microsoft.com/en-us/windows/win32/msi/major-upgrades)。