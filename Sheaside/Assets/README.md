# Sheaside artwork

参考: ユーザー提供の重なったシート＋波ロゴ。深緑と青緑、白の3色を用いた独自のGDI+ベクター描画。

- sheaside-16.png: 小サイズ用、背面シート1枚を省略。
- sheaside-32.png: リボン用。
- sheaside-256.png: 高解像度プレビュー。
- 描画原本: Controls/RibbonArtwork.cs。透過PNG。ランタイムは同じコードで32pxを生成しIPictureDispとしてOfficeへ渡す。
- 種類/表示状態の線画: Controls/ModernButton.cs。外部フォント・素材は未使用。

画像の加工ではなく参考意匠を小サイズに適したベクターで描き起こした。ユーザー提供資料以外の第三者素材を追加していない。
