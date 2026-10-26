# Sheaside 1.0.0 インストール・配布

正式版: https://github.com/nyo-ro/Sheaside/releases/tag/v1.0.0

## 必要環境と確認範囲

- Windowsデスクトップ版Excel、.NET Framework 4.8以上、Visual Studio Tools for Office Runtime（CLR 4対応版）。前提ソフトは自動取得しません。
- Microsoft 365 **64bit版**では、rc.6のMSIインストール・起動・主要機能・UI・アンインストール・再インストールをユーザー実機確認済みです。1.0.0の新しいMSIの実インストールは未実施です。
- 正式版の開発検証はExcel 2013です。.NETはRelease値528040以上を要件とし、4.8.1環境も含みます。
- 32bit版Officeは利用禁止ではありませんが未検証・保証対象外です。その他の未検証Office環境も保証しません。Web版・Mac版Excelには対応しません。
- MSIはx86・ユーザー単位、アドインはAny CPU。ARM版Office・全ユーザー一括導入は対象外です。

## インストール・推奨更新手順

1. ブックを保存し、Excelをすべて終了します。
2. 既存Sheasideがある場合は、Windowsのアプリ管理からアンインストールします。旧ClickOnce版も事前に削除してください。
3. Releaseから `Sheaside-1.0.0.zip` を取得して新しいフォルダーへ展開し、SHA-256を照合します。MSI単体でも導入できます。
4. Excelを使用するユーザーで `Sheaside-1.0.0.msi` を実行します。通常は管理者として起動しません。
5. Excelを起動し、［表示］→［Sheaside］→［シート一覧］で確認します。新規ウィンドウではパネルは自動表示されません。

```powershell
Get-FileHash .\Sheaside-1.0.0.zip -Algorithm SHA256
Get-FileHash .\Sheaside-1.0.0.msi -Algorithm SHA256
msiexec /i "Sheaside-1.0.0.msi" /L*v "Sheaside-install.log"
```

旧版からの直接更新は保証しません。旧MSIを同一ユーザーの同じUpgradeCodeとして検出した場合、正式版はファイル配置・旧版削除の前に停止し、アンインストールを案内します。ClickOnce・別ユーザー・手動登録・Visual Studioによる開発登録は同じ方法で検出できるとは限りません。開発PCではMSIとVSTO開発登録を併用しないでください。

配置先: `%LOCALAPPDATA%\Programs\Sheaside`
登録先: `HKCU\Software\Microsoft\Office\Excel\Addins\Sheaside`
設定: `%LOCALAPPDATA%\Sheaside\settings.xml`

## アンインストール・再インストール・修復

Excelをすべて終了し、Windowsのアプリ管理から削除します。同じバージョンのMSIでも削除できます。

```powershell
msiexec /x "Sheaside-1.0.0.msi" /L*v "Sheaside-uninstall.log"
msiexec /fa "Sheaside-1.0.0.msi" /L*v "Sheaside-repair.log"
```

本体とMSIが登録したアドイン情報を削除し、設定とExcelブックは残す構成です。削除後は新しいMSIを実行して導入します。rc.6のMicrosoft 365 64bit環境での削除・再インストールは確認済みですが、1.0.0の実インストール・削除・再インストール・修復は未実施です。修復はrc.6でも未確認です。

## 署名と信頼条件

MSIはAuthenticode未署名です。VSTOの2つのマニフェストは既存の開発証明書で署名されています。署名の暗号的な有効性と、導入先で発行元を信頼できることは別です。

- 証明書: `CN=Sheaside Local Development`
- 拇印（SHA-1識別子）: `C61660B9D07850DE0D156C9127355F4124B5093A`
- 有効期限: 2027-10-22（署名タイムスタンプの有無や導入先ポリシーによる扱いは別途確認が必要）

導入先では発行元不明の確認や信頼エラーが発生し、企業ポリシーによって読み込みが拒否される可能性があります。MSIは証明書を信頼ストアへ自動登録せず、Office/SmartScreen/ウイルス対策の設定も変更しません。配布元・ハッシュ・証明書を確認し、組織端末では管理者の承認を受けて信頼条件を整備してください。警告が出た場合の無条件な実行・保護無効化は案内しません。正式配布用コード署名証明書による署名は今後の課題です。秘密鍵・PFX・パスワードは配布しません。

## トラブルシューティング

rc.5の.NET/VSTO誤判定は、MSIのraw DWORDが`#533509`などの文字列になることを考慮していなかったためです。rc.6と1.0.0は型を考慮して両レジストリビューを検査します。

```powershell
Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release
```

528040以上なら.NETの要件を満たします。既に満たす場合はOSレジストリを書き換えず、ログのNET48_32/NET48_64/VSTO32/VSTO64を確認してください。ログを共有する際は個人名・ユーザーパスをマスクしてください。

アドインが出ない場合はExcelの［ファイル→オプション→アドイン］でCOMアドイン・使用できないアイテムを確認します。署名・信頼エラーでは、保護機能を無効化せず発行元と組織ポリシーを確認してください。

## 配布構成

MSI内はSheaside.dll、Sheaside.dll.manifest、Sheaside.vsto、Microsoft.Office.Tools.Common.v4.0.Utilities.dllの4ファイルです。PDB・秘密鍵・開発用中間生成物は配布しません。VSTOマニフェストはMSIでも必要で、`|vstolocal`登録を使います。ClickOnce配布・自動更新は行いません。

MSI ProductVersionは1.0.0、ProductCodeは`{493B22A8-C746-4D1C-9749-14108ED59252}`、UpgradeCodeは従来の`{E2087C97-A6B5-4089-B8F5-86DB64EBB425}`を維持します。1つの安定したComponent GUID・登録先を使用します。生成MSIの順序・条件検証は、実インストール後の共存や削除の全ケースを保証するものではありません。
