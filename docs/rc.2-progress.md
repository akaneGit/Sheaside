# 1.0.0-rc.2 開発記録

基点: aaeeb7c (rc.1)。WinForms/VSTOを維持。開始時の未コミット差分はslnxのProject Id、csprojのSubTypeとCanonicalName。保持し、今回のコミットには含めない。

1. メニュー・操作UI: 完了。指定順序・区切り線、追加/削除ボタン、通常通知撤廃。管理315/Excel込み462チェック、Debug/Release警告0・エラー0。
2. フィルターポップアップ: 完了。検索と状態/種類を分離、ペイン内配置、条件表示と解除、VeryHidden設定連動。管理321/Excel込み468、両構成警告0・エラー0。Phase 1: e33124f。
3. Excelテーマ色パレット: 完了。テーマ10色×6段階、標準色、色なし、任意RGB。ThemeColor/TintAndShadeを保持し、取得不可時はRGBへフォールバック。管理328／Excel込み479チェック、Debug/Release警告・エラー0。Phase 2: 120130d。
4. Officeテーマ追従・設定整理: 完了。Office 15/16のユーザー設定を可能な範囲で検出、取得不可はWindows標準色。既存の条件付き更新で再検出。旧テーマ/配置設定は無視し、UI・保存項目を削除。管理334／Excel込み485、両ビルド警告・エラー0。Phase 3: 6198201。
5. MSI: 未着手。WiXはPATHに存在せず、dotnetは利用可能。実際に作成・検査する。
6. 総合テスト・RC文書: 未着手。

各Phaseで管理コード・Excel 2013連携テスト、Debug/Releaseビルド後にコミットする。Microsoft 365とMSI実インストールは自動検査と区別して記録する。既存ClickOnce資産は削除せず、今後の配布経路から外す。
