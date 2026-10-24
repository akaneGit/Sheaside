# Sheaside 1.0.0-rc.5

**プレリリースです。Microsoft 365での実インストール・動作確認は未実施です。** 正式版v1.0への移行は実機確認後に判断します。

SheasideはExcelのワークシート・グラフシートを縦型一覧で管理するWindows向けVSTOアドインです。

## rc.5の改善

- Hidden／VeryHidden選択時の不要な通知を撤廃。アクティブシートや表示状態は変更しません。
- 標準10色の表示不具合を修正し、パレットを内容に応じたサイズへ変更。
- タブ色を細い縦バーへ変更。種類・非表示状態・個別シート保護のアイコンを整理。
- 長い名前の省略表示と日本語ツールチップに対応。
- パネルのタイトルを「シート一覧」に変更し、内部のタイトル・ブック名を削除。一覧の高さを拡大。
- 既存の単一/複数選択、一括操作、検索中のドラッグ、複数ブック・ウィンドウ対応を維持。

最新UI修正 `b44d01e` を含みます。リボンは［表示］タブの［Sheaside］グループです。

## 導入

1. WindowsのデスクトップExcel、.NET Framework 4.8以上、Visual Studio 2010 Tools for Office Runtime（CLR 4対応版）を用意します。
2. このReleaseの `Sheaside-1.0.0-rc.5.zip` を取得し、下記SHA-256と照合して展開します。MSI単体でも導入できます。
3. Excelをすべて終了し、`Sheaside-1.0.0-rc.5.msi` をExcelを使用するユーザーで実行します。
4. Excelを起動し、［表示 → Sheaside → シート一覧］でパネルを開きます。自動表示はしません。
5. ZIP内の `INSTALL.md` と `CHECKLIST.md` に沿って確認します。

配布先リポジトリは公開です。ZIP内のINSTALL.mdには公開前の「非公開・アクセス権が必要」という説明が残っていますが、現在は不要です。配布物の同一性を保つためMSI・ZIPは再作成していません。

## 更新・アンインストール

rc.2からの更新は、Excel終了後にrc.5 MSIを実行します。内部製品バージョンは `0.99.2 → 0.99.5`、ProductCodeを変更しUpgradeCodeを維持したMajor Upgradeです。**更新構造は検証済みですが、実更新は未検証です。** 旧ClickOnce版は先にアンインストールしてください。

削除はExcelを終了し、Windowsのアプリ管理からSheasideをアンインストールします。ログ付きコマンド例:

```powershell
msiexec /i "Sheaside-1.0.0-rc.5.msi" /L*v "Sheaside-install.log"
msiexec /x "Sheaside-1.0.0-rc.5.msi" /L*v "Sheaside-uninstall.log"
```

ユーザー設定とExcelブックは残す構成です。再インストール・修復・アンインストールの実動作も確認待ちです。

## 署名・既知の制約

- MSIはAuthenticode未署名です。
- VSTOマニフェストは既存の開発証明書 `CN=Sheaside Local Development` で署名されています。署名検証は成功しましたが、導入先での信頼成立を意味しません。
- 証明書拇印: `C61660B9D07850DE0D156C9127355F4124B5093A`。有効期限: 2027-10-22。
- 発行元不明の警告や企業ポリシーによる読み込み拒否が発生する可能性があります。Officeの保護やセキュリティ機能を無条件に無効化せず、配布元・管理者へ確認してください。証明書の自動信頼登録は行いません。
- 秘密鍵・PFX・認証情報は配布物に含みません。ClickOnce配布・自動更新は使用しません。
- MSIはx86、アドインはAny CPU、登録はユーザー単位です。32/64bit Officeを想定していますが、Microsoft 365の両環境での互換性は未検証です。ARM版Office・全ユーザー導入は対象外です。
- 実OSのDPI変更・複数モニター・Microsoft 365固有のリボン/テーマ動作は確認待ちです。

## 検証結果

- Debug/Releaseビルド成功、警告0・エラー0。
- 管理回帰430チェック、Excel 2013連携込み600チェック成功。
- rc.5タイトル修正時の実VSTOホストUI 21項目成功。
- MSI構造18チェック、MSIから抽出した4ファイルのRelease出力とのSHA-256一致。
- ZIP・マニフェスト参照/依存ハッシュ・バージョン・既存rc.2成果物の保全など24チェック成功。
- 2つのVSTOマニフェストはMageで署名検証成功。
- **MSIの実インストール・更新・削除、Microsoft 365での動作は未検証。構造検証だけで動作保証はしません。**

## SHA-256

```text
33942038BCA5B104A3AFF244D3EFE3C0037957720E2729FD94CD41B5C44D98F0  Sheaside-1.0.0-rc.5.msi
0BB09DDA87108529EC91F91732CA07A8F87ABD960B126F43731478CB7CF133CA  Sheaside-1.0.0-rc.5.zip
```

詳細は[導入手順](https://github.com/akaneGit/Sheaside/blob/v1.0.0-rc.5/docs/deployment.md)と[Microsoft 365確認リスト](https://github.com/akaneGit/Sheaside/blob/v1.0.0-rc.5/docs/rc.5-install-checklist.md)を参照してください。
