# 1.0.0-rc.2 開発記録

基点: aaeeb7c (rc.1)。WinForms/VSTOを維持。開始時の未コミット差分はslnxのProject Id、csprojのSubTypeとCanonicalName。保持し、今回のコミットには含めない。

1. メニュー・操作UI: 完了。指定順序・区切り線、追加/削除ボタン、通常通知撤廃。管理315/Excel込み462チェック、Debug/Release警告0・エラー0。
2. フィルターポップアップ: 完了。検索と状態/種類を分離、ペイン内配置、条件表示と解除、VeryHidden設定連動。管理321/Excel込み468、両構成警告0・エラー0。Phase 1: e33124f。
3. Excelテーマ色パレット: 未着手。
4. Officeテーマ追従・設定整理: 未着手。
5. MSI: 未着手。WiXはPATHに存在せず、dotnetは利用可能。実際に作成・検査する。
6. 総合テスト・RC文書: 未着手。

各Phaseで管理コード・Excel 2013連携テスト、Debug/Releaseビルド後にコミットする。Microsoft 365とMSI実インストールは自動検査と区別して記録する。既存ClickOnce資産は削除せず、今後の配布経路から外す。
