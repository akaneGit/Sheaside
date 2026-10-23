# RCのインストール・配布

Windows版Excel用VSTOアドインです。Excel Online、Mac、Office.jsには対応しません。

## 開発環境

Visual Studio 2026 Community、Office/SharePoint開発ツール、.NET Framework 4.8開発ツール、VSTO Runtime、デスクトップExcelを用意します。Debug / Any CPUでビルドしてF5起動します。プロジェクトのデバッグ設定でMicrosoft 365のExcelを起動していることを確認してください。自動テスト環境はExcel 2013です。

署名PFXがない場合はVisual Studioの［署名］で開発用証明書を作成するか、`tools/New-DevelopmentCertificate.ps1`を使用します。既存PFXは上書きしません。証明書を変更した場合はThumbprintの差分も確認します。

## ローカル発行

```powershell
.\tools\Publish-ReleaseCandidate.ps1
# 出力先を変更する場合
.\tools\Publish-ReleaseCandidate.ps1 -OutputDirectory C:\Sheaside-RC
```

既定出力先は`Sheaside/obj/ReleaseCandidate`です。`setup.exe`、`Sheaside.vsto`、`Application Files/Sheaside_1_0_0_0/`以下を一式で保持します。DLLだけではインストールできません。

ReleaseのPublishターゲットを使用し、`IsWebBootstrapper=false`を明示しています。.NET Framework 4.8とVSTO Runtimeを前提条件に登録しました。前提ソフトはベンダーサイトから取得するため、未導入端末では接続と導入権限が必要です。

本環境では警告・エラー0で発行し、Mageで`.vsto`とアプリケーションマニフェストの署名を検証済みです。setup.exeは開発用自己署名のため、Authenticodeでは信頼ルート未登録の`UnknownError`になります。署名の整合性と配布先での信頼は別です。証明書の信頼設定変更、実インストール、外部公開は実施していません。

## 検証端末への導入と正式配布

1. Microsoft 365検証端末のWindows、Excelビルド・32/64bitを記録し、前提ランタイムを確認します。
2. 配布用の信頼されたコード署名証明書を設定します。社内配布は組織の証明書配布ポリシーに従います。PFXや秘密鍵はパッケージへ含めません。
3. Visual Studioの［発行］で正式な配布先・更新元・発行バージョンを設定して再発行します。本スクリプトは固定バージョンのローカルRC検証用で、自動更新元は設定しません。
4. 発行フォルダー一式を検証端末へ渡し、Excelを終了してsetup.exeから導入します。前提条件導入済みなら.vstoからも導入できますが、前提条件を確認する経路はsetup.exeです。
5. Excelの［オプション］→［アドイン］→［COMアドイン］でSheaside、リボン、パネルを確認します。表示されない場合は［使用できないアイテム］も確認します。
6. [実機チェックリスト](v1.0-checklist.md)に沿って導入・更新・アンインストール・再導入も確認します。開発版と配布版は検証環境を分け、二重登録を避けます。
7. 正式版ではRC表記を外し、ClickOnce発行バージョンをRCより大きくして再発行します。

アンインストールはWindowsのインストール済みアプリから行います。個人設定`%LOCALAPPDATA%/Sheaside/settings.xml`は別に保持されます。初期化する場合はExcelをすべて終了し、このファイルをバックアップして移動してください。

配布方式はMicrosoftの[OfficeソリューションのClickOnce展開](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce)に基づきます。管理者向けMSIや企業の一括配布パッケージは本RCでは作成していません。
