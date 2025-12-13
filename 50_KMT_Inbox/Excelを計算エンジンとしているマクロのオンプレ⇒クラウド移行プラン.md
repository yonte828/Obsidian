## 📊 現状分析：技術的課題の整理

現行システムの**クリティカルな依存関係**：

1. **Excel数式エンジンへの依存** - 質量データ更新後、Excel の数式が自動再計算し、総質量・接地圧・重心位置等を算出
2. **[[VBAとは|VBA]] による [[COMとは|COM]] 制御** - GDMS [[APIとは|API]] 連携、Outlook メール送信
3. **[[オンプレミスとは|オンプレミス]]環境** - 社内サーバー (10.2.3.154) への [[HTTPとは|HTTP]] 通信

---

## 🔍 F/S案1: Power Automate + Office Scripts

### ✅ 実現可能性：**中～高（条件付き）**

#### 実装アーキテクチャ

```
[[Power Automateとは|Power Automate]] Desktop (オンプレ)
  ↓ [[データゲートウェイとは|データゲートウェイ]]経由
[[Power Automateとは|Power Automate]] Cloud
  ↓
[[Office Scriptsとは|Office Scripts]] (Excel Online)
  ↓
[[SharePointとは|SharePoint]]/[[OneDriveとは|OneDrive]] に保存
```

### 📋 技術的実現性の詳細評価

| 要件                   | [[Office Scriptsとは|Office Scripts]] での実現 | 制約・注意点                                                        |
| -------------------- | ----------------------- | ------------------------------------------------------------- |
| **Excel 数式の再計算** | ✅ 可能                    | `workbook.getActiveWorksheet().calculate()` で強制再計算可能          |
| **セル値の読み書き**         | ✅ 可能                    | Range [[APIとは|API]] で完全対応                                           |
| **外部 [[APIとは|API]] 呼び出し**  | ✅ 可能                    | [[fetch() APIとは|fetch() API]] で[[HTTPとは|HTTP]]リクエスト可能（[[CORSとは|CORS]]制約あり） |
| **複雑な条件分岐**          | ✅ 可能                    | [[TypeScriptとは|TypeScript]] ベースで実装                                         |
| **ファイル操作**           | ⚠️ 部分的                  | SharePoint/OneDrive 上のファイルのみ                                  |
| **メール送信**            | ✅ 可能                    | [[Power Automateとは|Power Automate]] 側で Outlook/Gmail コネクタ使用                    |
| **実行時間**             | ⚠️ 制約あり                 | [[Office Scriptsとは|Office Scripts]]: 最大5分/実行                                   |

### ⚠️ 重大な制約事項

1. **[[オンプレミスとは|オンプレミス]] [[APIとは|API]] へのアクセス**
    
    - [[Office Scriptsとは|Office Scripts]] は[[クラウド]]サンドボックスで実行されるため、社内サーバー (`10.2.3.154`) への直接通信は**不可能**
    - **解決策**: [[Power Automateとは|Power Automate]] の[[オンプレミスとは|オンプレミス]] [[データゲートウェイとは|データゲートウェイ]]経由で [[APIとは|API]] 呼び出し
2. **実行時間制限**
    
    - [[Office Scriptsとは|Office Scripts]] は**最大 5 分**で強制終了
    - 大量データ処理や複雑な計算がある場合、タイムアウトリスクあり
3. **[[VBAとは|VBA]] マクロの互換性**
    
    - 既存の [[VBAとは|VBA]] コードは**そのままでは動作しない**
    - [[TypeScriptとは|TypeScript]] で完全リライト必要

### 🏗️ 推奨実装パターン

```typescript
Copy// Office Scripts のサンプル実装イメージ
async function main(workbook: ExcelScript.Workbook) {
  // 1. データ取得（更新前）
  const sheet = workbook.getWorksheet("機種シート");
  const beforeData = captureBeforeData(sheet);
  
  // 2. Power Automate からの API データを受け取る
  // (Power Automate 側で GDMS API を呼び出し、結果を Scripts に渡す)
  
  // 3. Excel に値を書き込み
  updateMassData(sheet, apiResponse);
  
  // 4. 再計算を強制実行
  workbook.getActiveWorksheet().calculate();
  
  // 5. データ取得（更新後）
  const afterData = captureAfterData(sheet);
  
  // 6. 差異比較
  const diff = compareData(beforeData, afterData);
  
  // 7. 結果を Power Automate に返却
  return { success: true, differences: diff };
}
```

### 💰 コスト試算（月額換算）

- **[[Power Automateとは|Power Automate]] Premium**: 約 ¥4,000/ユーザー
- **[[Office 365とは|Office 365]] E3/E5**: Excel Online + [[Office Scriptsとは|Office Scripts]] 含む
- **[[オンプレミスとは|オンプレミス]]ゲートウェイ**: 無料（[[Windows Serverとは|Windows Server]] 必要）

**推定総コスト**: 月額 ¥5,000～10,000（既存 [[Office 365とは|Office 365]] 契約次第）

---

## 🚀 F/S案2: Python + xlwings（ハイブリッドクラウド）【推奨】

### ✅ 実現可能性：**高**

#### アーキテクチャ

```
[[Azure Functionsとは|Azure Functions]] / [[AWS Lambdaとは|AWS Lambda]] ([[Pythonとは|Python]])
  ↓ [[VPNとは|VPN]]/[[ExpressRouteとは|ExpressRoute]]
オンプレ Excel サーバー (Windows)
  - [[xlwingsとは|xlwings]] + Excel Desktop
  - GDMS [[APIとは|API]] アクセス
```

### 📋 技術的優位性

|項目|評価|詳細|
|---|---|---|
|**Excel 計算エンジン活用**|✅✅|[[xlwingsとは|xlwings]] で既存の Excel 数式をそのまま利用可能|
|**[[VBAとは|VBA]] ロジック移植**|✅✅|[[Pythonとは|Python]] でロジックを再現、計算部分は Excel に委譲|
|**[[APIとは|API]] 連携**|✅✅|[[requestsライブラリとは|requests]] ライブラリで柔軟に実装|
|**拡張性**|✅✅|[[pandasとは|pandas]], [[Numpyとは|numpy]] 等の高度なデータ処理が可能|
|**保守性**|✅✅|[[Gitとは|Git]] 管理、テスト自動化が容易|

### 🔧 実装例

```python
Copyimport xlwings as xw
import requests
from datetime import datetime

def update_unified_calc_sheet(file_path: str):
    # 1. Excel を開く（計算エンジンとして活用）
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets['機種シート']
    
    # 2. 更新前データ取得
    before_data = {
        '総質量': sheet.range('X10').value,
        '接地圧': sheet.range('Y10').value,
        # ... その他の計算結果
    }
    
    # 3. GDMS [[APIとは|API]] から最新データ取得
    pn_list = extract_part_numbers(sheet)
    api_data = fetch_gdms_data(pn_list)
    
    # 4. Excel に値を書き込み（数式はそのまま）
    for row, mass in enumerate(api_data, start=5):
        sheet.range(f'Q{row}').value = mass['質量']
    
    # 5. Excel の再計算を実行
    wb.app.calculate()
    
    # 6. 更新後データ取得
    after_data = {
        '総質量': sheet.range('X10').value,
        '接地圧': sheet.range('Y10').value,
    }
    
    # 7. 差異比較とレポート作成
    diff = compare_results(before_data, after_data)
    
    # 8. 保存とクローズ
    wb.save()
    wb.close()
    app.quit()
    
    return diff

def fetch_gdms_data(pn_list):
    """GDMS API から質量データ取得"""
    response = requests.post(
        'http://10.2.3.154:37001/api/mass',
        json={'part_numbers': pn_list},
        timeout=30
    )
    return response.json()
```

### 🏗️ デプロイ構成案

**オプション A: [[オンプレミスとは|オンプレミス]]専用サーバー**

- [[Windows Serverとは|Windows Server]] 上で [[Pythonとは|Python]] + Excel Desktop を常駐
- Azure DevOps で[[CI_CD|CI/CD]] パイプライン構築
- タスクスケジューラーで定期実行

**オプション B: ハイブリッド[[クラウド]]**

- [[Azure VMとは|Azure VM]] (Windows) に [[xlwingsとは|xlwings]] 環境構築
- [[オンプレミスとは|オンプレミス]] GDMS へは [[VPNとは|VPN]]/[[ExpressRouteとは|ExpressRoute]] 経由
- [[Azure Functionsとは|Azure Functions]] でトリガー管理

### 💰 コスト試算

- **[[Azure VMとは|Azure VM]] (Standard B2s)**: 約 ¥4,000/月
- **Excel ライセンス**: [[Office 365とは|Office 365]] 既存契約流用
- **[[Pythonとは|Python]] 環境**: 無料（オープンソース）

---

## 🎯 F/S案3: フルクラウド移行（計算ロジック移植）

### ✅ 実現可能性：**中（工数大）**

#### アーキテクチャ

```
[[Azure Functionsとは|Azure Functions]] ([[Pythonとは|Python]]/[[C#とは|C#]])
  ↓
[[pandasとは|pandas]] + [[Numpyとは|numpy]] で計算ロジック実装
  ↓
[[Azure SQL Databaseとは|Azure SQL Database]] / [[Azure Cosmos DBとは|Cosmos DB]]
```

### 📋 アプローチ

Excel の数式を**完全に解析**し、[[Pythonとは|Python]]/[[C#とは|C#]] でロジックを再実装：

```python
Copydef calculate_total_mass(parts_df):
    """総質量計算（Excel 数式を Python で再現）"""
    return parts_df['質量'].sum()

def calculate_ground_pressure(total_mass, contact_area):
    """接地圧計算"""
    return total_mass / contact_area

def calculate_center_of_gravity(parts_df):
    """重心位置計算"""
    total_mass = parts_df['質量'].sum()
    cog_x = (parts_df['質量'] * parts_df['X座標']).sum() / total_mass
    cog_y = (parts_df['質量'] * parts_df['Y座標']).sum() / total_mass
    return (cog_x, cog_y)
```

### ⚠️ 課題

1. **Excel 数式の完全解析が必要** - 複雑な IF/[[VLOOKUP関数とは|VLOOKUP]]/[[配列数式]]等のロジック抽出に工数
2. **検証コスト** - 既存 Excel との計算結果一致を確認する膨大なテストケース
3. **保守性** - Excel 側の計算ロジック変更時、[[Pythonとは|Python]] 側も同期修正必要

### 💰 コスト試算

- **開発工数**: 3～6人月（計算ロジック解析・実装・テスト）
- **[[Azure Functionsとは|Azure Functions]]**: 約 ¥1,000/月（従量課金）
- **ランニング**: 月額 ¥2,000～5,000

---

## 📊 総合比較表

|項目|[[Power Automateとは|Power Automate]] + [[Office Scriptsとは|Office Scripts]]|[[Pythonとは|Python]] + [[xlwingsとは|xlwings]]|フル[[クラウド]]移行|
|---|---|---|---|
|**実現可能性**|⭐⭐⭐ 中～高|⭐⭐⭐⭐⭐ 高|⭐⭐⭐ 中|
|**Excel 数式活用**|✅ 可能|✅✅ 完全対応|❌ 要移植|
|**[[オンプレミスとは|オンプレミス]] [[APIとは|API]] 接続**|⚠️ ゲートウェイ必要|✅ 直接可能|✅ [[VPNとは|VPN]]経由可能|
|**開発工数**|2～3人月|1～2人月|4～6人月|
|**初期コスト**|¥50万～100万|¥30万～70万|¥200万～400万|
|**月額ランニング**|¥5,000～10,000|¥4,000～8,000|¥2,000～5,000|
|**保守性**|⭐⭐⭐|⭐⭐⭐⭐⭐|⭐⭐⭐⭐|
|**拡張性**|⭐⭐|⭐⭐⭐⭐|⭐⭐⭐⭐⭐|
|**リスク**|タイムアウト、[[CORSとは|CORS]]制約|Windows依存|計算ロジック検証|

---

## 🏆 推奨アプローチ

### **第一推奨: [[Pythonとは|Python]] + [[xlwingsとは|xlwings]]（段階的移行）**

#### フェーズ1: 最小限の移行（3ヶ月）

1. [[VBAとは|VBA]] の処理フロー制御部分を [[Pythonとは|Python]] に移植
2. Excel の計算エンジンは**そのまま活用**（[[xlwingsとは|xlwings]] 経由）
3. GDMS [[APIとは|API]] 連携、メール送信を [[Pythonとは|Python]] で実装
4. [[オンプレミスとは|オンプレミス]] [[Windows Serverとは|Windows Server]] でスケジュール実行

#### フェーズ2: [[クラウド]]化（+3ヶ月）

1. [[Azure VMとは|Azure VM]] (Windows) に環境構築
2. [[VPNとは|VPN]]/[[ExpressRouteとは|ExpressRoute]] で[[オンプレミスとは|オンプレミス]]接続
3. Azure DevOps で [[CI_CD|CI/CD]] 構築

#### フェーズ3: 完全[[クラウド]]化（オプション、+6ヶ月）

1. 計算ロジックを段階的に [[Pythonとは|Python]] に移植
2. Excel 依存を削減
3. フルマネージドサービス化

### **第二推奨: [[Power Automateとは|Power Automate]]（小規模・迅速移行向け）**

以下の条件を満たす場合に検討：

- 処理時間が 5 分以内に収まる
- [[Office 365とは|Office 365]] E3/E5 の契約がある
- 開発リソースが限られている
- [[オンプレミスとは|オンプレミス]][[データゲートウェイとは|データゲートウェイ]]の設置が可能

---

## 📝 次のステップ

### 即座に実施すべきこと

1. **Excel 計算式の棚卸し**
    
    - どのシートのどのセルが重要な計算を行っているか文書化
    - 数式の複雑度を評価（[[VLOOKUP関数とは|VLOOKUP]]、[[配列数式]]、[[ユーザー定義関数]]等）
2. **パフォーマンステスト**
    
    - 現行 [[VBAとは|VBA]] の実行時間を計測
    - [[Office Scriptsとは|Office Scripts]] の 5分制約に収まるか検証
3. **ネットワーク要件確認**
    
    - [[クラウド]]から[[オンプレミスとは|オンプレミス]] GDMS への接続方式（[[VPNとは|VPN]]/[[ExpressRouteとは|ExpressRoute]]/[[データゲートウェイとは|データゲートウェイ]]）
    - セキュリティポリシーの確認
4. **[[POC]]（概念実証）の実施**
    
    - 小規模な機種データで [[Pythonとは|Python]] + [[xlwingsとは|xlwings]] を試験実装
    - [[Power Automateとは|Power Automate]] + [[Office Scriptsとは|Office Scripts]] でプロトタイプ作成
    - 両者を比較評価

### 詳細な F/S が必要な場合

以下の情報があれば、より精緻な提案が可能です：

- 処理対象ファイルの数とサイズ
- GDMS [[APIとは|API]] の詳細仕様（レスポンス時間、データ量）
- 現行の実行頻度とタイミング
- 既存の IT インフラ（Azure 契約の有無、[[VPNとは|VPN]] 環境等）

何か追加で確認したい点や、特定の案について深掘りしたい内容はありますか？
#VBA #Excelマクロ #Python移行 #GDMS連携 #自動化 #質量管理 #システム連携