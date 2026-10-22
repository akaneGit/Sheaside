# Sheaside v0.1

Excelの右側にワークシート一覧を表示するVSTOアドインです。既存のVisual Studioソリューションと.NET Framework 4.8プロジェクトを使用しています。

## 実装した動作

- ブックを開くと、右側に「Sheaside」カスタムタスクペインを表示します。
- ワークシートをブック内の順序で縦に表示します。クリック、または選択してEnterで切り替えられます。
- ブック／ウィンドウの切り替え時と、Excel側でのシート切り替え時に一覧を更新します。
- リボンの［Sheaside］タブ → ［Sheaside］ボタンで、そのウィンドウのパネルを表示・非表示にします。パネルの×ボタンとも状態を同期します。
- ウィンドウのHwndごとにパネルを管理します。同じブックの［新しいウィンドウを開く］にも対応する構成です。新しいウィンドウでは初期表示し、表示状態の永続保存は行いません。
- Excelにウィンドウ終了後・シート名変更の専用イベントがないため、1秒ごとのWinFormsタイマーでも実際のウィンドウと表示中の一覧を照合します。閉じる操作のキャンセル時にはパネルを削除しません。
- 非表示／VeryHiddenシートには「（非表示）」を付けます。クリックしても再表示せず、パネル内で案内します。グラフシートは一覧の対象外です。
- シートの追加・削除・並び替え操作は実装していません。Excel側で行われた変更は一覧に反映します。

## ファイル構成と変更内容

| ファイル | 区分 | 役割 |
| --- | --- | --- |
| `Sheaside/ThisAddIn.cs` | 変更 | 起動・終了時の管理、Ribbon XMLの接続 |
| `Sheaside/Sheaside.csproj` | 変更 | ソース／XMLリソースの登録、ClickOnce署名設定 |
| `Sheaside/Controls/SheetListControl.cs` | 新規 | Excel Interopに依存しないWinForms一覧UI |
| `Sheaside/Models/WorksheetItem.cs` | 新規 | 一覧表示用データ、選択イベント引数 |
| `Sheaside/Services/ExcelWorksheetService.cs` | 新規 | 対象ウィンドウのブック取得、シート列挙・アクティブ化 |
| `Sheaside/Services/SheetPaneManager.cs` | 新規 | ウィンドウごとのパネル、Excelイベント、タイマー、終了処理 |
| `Sheaside/SheasideRibbon.cs` | 新規 | 表示切り替えとボタン状態同期のコールバック |
| `Sheaside/SheasideRibbon.xml` | 新規 | リボンのタブ・トグルボタン定義 |
| `Sheaside/Sheaside_TemporaryKey.pfx` | ローカル生成 | 開発用のClickOnce署名証明書（Git対象外） |
| `tools/New-DevelopmentCertificate.ps1` | 新規 | 開発用証明書の生成とプロジェクトの拇印更新 |
| `.gitignore` | 新規 | ビルド生成物、IDE設定、秘密鍵の除外 |
| `README.md` | 新規 | 実装内容と確認手順 |

既存の`*.Designer.cs`、`ThisAddIn.Designer.xml`、ソリューションファイルは変更していません。UIはUserControl内のコードで構築し、リボンにはRibbon XMLを使用しています。

## Visual Studioでの動作確認

1. Visual Studio 2026 Communityで`Sheaside.slnx`を開き、Sheasideをスタートアッププロジェクトにします。Office/SharePoint開発ツール、.NET Framework 4.8開発ツール、VSTO Runtime、Windows版Excelが必要です。
2. 構成を`Debug / Any CPU`にし、［ビルド］→［ソリューションのビルド］を実行します。
3. F5でExcelを起動します。ブックを開いていない開始画面では、空白のブックを作成してください。
4. ワークシートを2枚以上用意します。右側に「Sheaside」が表示され、ブック内のワークシート名と順番が一致することを確認します。
5. 一覧で別のシートをクリックし、Excelのアクティブシートが切り替わることを確認します。キーボードでは矢印キーで選んでEnterを押します。
6. ［Sheaside］タブの［Sheaside］ボタンでパネルを非表示・再表示します。パネルの×で閉じたときにボタンの押下状態も解除され、ボタンから再表示できることを確認します。
7. 別のブックを開き、ウィンドウを行き来して、それぞれのブックの一覧とアクティブシートが表示されることを確認します。
8. ［表示］→［新しいウィンドウを開く］で同じブックの2つ目のウィンドウを作成します。一方のパネルを非表示にしても他方が表示されたままで、一覧のクリックがそのウィンドウのシートを切り替えることを確認します。
9. ブックを変更して閉じる際、保存確認で［キャンセル］を選んでもパネルが残ることを確認します。その後、不要なウィンドウを閉じても残りのウィンドウで操作できること、最後のブックを閉じて新規ブックを作成しても一覧が表示されることを確認します。
10. 非表示シート、長いシート名、日本語のシート名、グラフシートのあるブックでも確認します。非表示シートをクリックしても表示状態は変わりません。セル編集中などに切り替えが失敗した場合は、編集を完了して再試行してください。

アドインが表示されない場合は、Excelの［ファイル］→［オプション］→［アドイン］でCOMアドインのSheasideと、使用できないアイテムを確認してください。F5が目的と異なるExcelを起動する場合は、プロジェクトのデバッグ設定で起動対象のExcel.exeを確認してください。

## 開発用署名

元のプロジェクトにClickOnce署名設定がなく、標準VSTOビルドが失敗したため、1年間有効なローカル開発用証明書を作成し、署名設定を追加しています。PFXはこの作業環境に生成済みです。配布には配布用の署名証明書を使用してください。

別の環境へソースのみを持っていく場合は、Visual Studioのプロジェクトプロパティ［署名］でテスト証明書を作成・選択するか、PFXが存在しない状態で次を実行してください。

```powershell
powershell.exe -NoProfile -File .\tools\New-DevelopmentCertificate.ps1
```

スクリプトは既存の証明書を上書きしません。証明書を生成すると、プロジェクト内の`ManifestCertificateThumbprint`も更新します。秘密鍵を含むPFXはGit管理から除外しています。通常のVSTOビルドでは、標準のOfficeToolsターゲットが開発用証明書の処理と現在のユーザーへのアドイン登録を行います。

## 検証記録

- Visual Studio 2026付属MSBuildで、ソリューションのDebug / Releaseビルドに成功。
- `Sheaside.dll`、`Sheaside.dll.manifest`、`Sheaside.vsto`を生成。
- Releaseビルドは警告0件・エラー0件。
- Excel Microsoft 365上でのパネル表示・リボン・クリック操作は未確認。作業環境のExcel登録先はOffice15であり、上記の手動確認が必要です。

コマンドラインでのDebugビルド例:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' `
    .\Sheaside.slnx /t:Build /p:Configuration=Debug '/p:Platform=Any CPU' /p:VisualStudioVersion=18.0
```

タスクペインはウィンドウを明示して作成し、終了時にはVSTO自身に破棄を任せます。このライフサイクルとRibbon XMLの接続は、Microsoftの[カスタムタスクペイン](https://learn.microsoft.com/en-us/visualstudio/vsto/custom-task-panes)と[Ribbon XML](https://learn.microsoft.com/en-us/visualstudio/vsto/ribbon-xml)の説明を参照しています。
