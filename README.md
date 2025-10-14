# Excel VBA 業務自動化プロジェクト

![VBA歴](https://img.shields.io/badge/VBA歴-15年-blue.svg)
![運用実績](https://img.shields.io/badge/運用実績-8年-green.svg)
![言語](https://img.shields.io/badge/言語-VBA-yellow.svg)
![プラットフォーム](https://img.shields.io/badge/プラットフォーム-Excel-brightgreen.svg)
![ライセンス](https://img.shields.io/badge/ライセンス-Private-red.svg)

## 概要

製造業における発注業務と検査証管理を自動化するExcel VBAプロジェクト。基幹システム（ERP）との連携により、手作業を大幅に削減し、業務効率を劇的に改善しました。

### ビジネスインパクト

- **工数削減:** 月間100時間以上の削減
- **運用実績:** 8年間の安定稼働
- **横展開:** 一部から開始し、段階的に製造課→品質保証課→総務課→技術課へ展開
- **累計削減効果:** 約9,600時間(8年間)

---

## プロジェクト構成

### 1. 基幹システム連携.xlsm
**用途**: 基幹システム（ERPSystem）への発注データ自動入力

**主な機能**:
- Excelシートから発注情報を読み取り
- 基幹システム画面への自動入力（SendKeys使用）
- マウス操作の自動化
- 単位変換辞書による柔軟な入力

**技術的特徴**:
- クラス設計による責務の分離
- Windows API活用（ウィンドウ操作、マウス制御）
- 段階的リファクタリングの履歴保存

**モジュール構成**:
```
├── OrderData.cls          # 発注データ管理クラス
├── ERPSystemOperator.cls  # ERP操作クラス
├── Module3.bas            # 最新版メイン処理
├── Module2.bas            # 中間リファクタリング版
└── Module1.bas            # 初期実装版
```

---

### 2. 検査証_一般.xlsm
**用途**: 検査証データの管理・印刷システム

**主な機能**:
- 外部Excelファイルからの過去データインポート
- 日付範囲によるフィルタリング
- データの並び替え・集計
- 数量の分割・統合処理
- 検査証の自動印刷（最大32件、8件/ページ）

**技術的特徴**:
- イベントドリブン設計（Worksheet_Change）
- 外部ブック管理クラス
- 配列処理による高速化
- エラーハンドリングの徹底

**モジュール構成**:
```
├── WorkbookManager.cls    # 外部ブック管理
├── import.bas             # データインポート処理
├── Adjust.bas             # 数量分割処理
├── printing.bas           # 印刷処理
├── Module1.bas            # 数量統合処理
└── Sheet3.cls             # 日付検証イベント
```

---

## 開発の経緯

**背景:**
Excel VBAは2009年から15年以上使用していますが、このツールは2017年頃に開発を開始しました。

### Phase 1: 初期実装(2017年)
- 一部の業務から小規模にスタート
- 手続き型で一気に実装
- とにかく動くものを作る
- 現場でのフィードバック収集

**技術的特徴**:
- 全処理を1つのSubプロシージャに記述
- セル座標のハードコーディング
- 基本的なエラーハンドリング

**コード例** (Module1.bas):
```vba
Sub CoreSystemOperation()
    ' 事前確認
    Dim check As Variant
    check = MsgBox("基幹システムはメイン画面になっていますか?", vbYesNo)

    ' データ読み込み
    Dim Requester As String
    Requester = UCase(Cells(11 + Resize, 6))

    ' 基幹システム操作
    AppActivate "ERPSystem", True
    .SendKeys "" & Requester & "", True
    ' ... 以下、直列的に処理が続く
End Sub
```

---

### Phase 2: 機能拡張と改善(2018-2020年)
- 関数分割でコードを整理
- 辞書による単位変換
- エラーハンドリング追加
- 対象業務を徐々に拡大

**技術的改善**:
- 関数による処理の分割
- Dictionary使用による柔軟な変換
- ウィンドウ情報の構造化

**コード例** (Module2.bas):
```vba
Function GetInputData(Resize As Long) As Dictionary
    Dim data As New Dictionary
    With Cells
        data.Add "Requester", UCase(.Item(11 + Resize, 6))
        data.Add "Order_quantity", .Item(47 + Resize, 1)
        ' ... データをDictionaryで管理
    End With
    Set GetInputData = data
End Function

Function GetUnitNumber(Unit As String) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "kg", 3
    dict.Add "t", 4
    ' ... 単位辞書
    If dict.Exists(Unit) Then
        GetUnitNumber = dict(Unit)
    End If
End Function
```

---

### Phase 3: リファクタリング(2021-2023年)
- クラス設計の導入
- 責務の分離
- テストしやすい構造へ
- コード品質の継続的改善

**技術的進化**:
- クラスモジュールによるOOP
- 単一責任の原則（SRP）
- 依存性の注入パターン

**クラス設計** (Module3.bas + OrderData.cls + ERPSystemOperator.cls):

```vba
' === メイン処理 ===
Sub CoreSystemOperation()
    ' データ読み込み
    Dim data As New OrderData
    data.LoadFromSheet Resize

    ' ERP操作
    Dim erp As New ERPSystemOperator
    erp.Initialize
    erp.SendOrderData data
    erp.SendNote data.Note
End Sub

' === OrderData.cls (データ管理) ===
Public Requester As String
Public OrderQuantity As Long
Public Unit As String

Public Sub LoadFromSheet(resizeValue As Long)
    Requester = UCase(Cells(11 + resizeValue, 6))
    OrderQuantity = Cells(47 + resizeValue, 1)
    Unit = Cells(47 + resizeValue, 4)
End Sub

Public Function GetUnitNumber() As Variant
    ' 単位辞書変換
End Function

' === ERPSystemOperator.cls (ERP操作) ===
Private shell As Object

Public Sub Initialize()
    AppActivate "ERPSystem", True
    Set shell = CreateObject("Wscript.Shell")
    ' ウィンドウ情報取得
End Sub

Public Sub SendOrderData(data As OrderData)
    shell.SendKeys data.Requester, True
    shell.SendKeys data.OrderQuantity, True
    ' ... 入力処理
End Sub
```

**責務の分離**:
| クラス | 責務 |
|--------|------|
| OrderData | データの保持と変換 |
| ERPSystemOperator | ERP画面操作 |
| CoreSystemOperation | オーケストレーション |

---

### Phase 4: 横展開(2023年)
- 製造課で安定稼働を確認
- 品質保証課・総務課・技術課へ展開
- 8年の運用実績を証明

**展開戦略**:
1. 製造課での実証（2017-2020年）
2. 成功事例の共有
3. 他部署からの要望対応
4. カスタマイズと展開（2021-2023年）

---

## 技術スタック

### 言語・プラットフォーム
- **VBA (Visual Basic for Applications)**
- **Excel 2013/2016/2019/365**
- **Windows 10/11**

### 使用技術

#### 1. Windows API
```vba
' ウィンドウ操作
Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Declare PtrSafe Function GetWindowPlacement Lib "user32" _
    (ByVal hwnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As Long

' マウス操作
Declare PtrSafe Sub SetCursorPos Lib "user32" _
    (ByVal x As Long, ByVal y As Long)
Declare PtrSafe Sub mouse_event Lib "user32" _
    (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
     ByVal cButtons As Long, ByVal dwExtraInfo As Long)
```

#### 2. COM オブジェクト
```vba
' WScript.Shell (キー送信、アプリ切替)
Set shell = CreateObject("Wscript.Shell")
shell.SendKeys "{F6}", True
shell.AppActivate "ERPSystem"

' Scripting.Dictionary (辞書)
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "kg", 3
```

#### 3. Excel VBA 機能
- **イベントハンドリング**: `Worksheet_Change`, `Workbook_Open`
- **AutoFilter**: 高速フィルタリング
- **SpecialCells**: 可視セルの取得
- **配列処理**: メモリ内での高速データ操作
- **PrintOut**: 自動印刷

---

## プロジェクトが示すスキル

### 1. 長期運用能力: 8年間の安定稼働と継続的改善
- 8年間の運用を通じた保守性の証明
- エラーハンドリングの徹底
- 現場フィードバックの継続的反映

### 2. 段階的リファクタリング
**手続き型 → 関数分割 → クラス設計**
- 動作を維持しながら品質向上
- レガシーコードの保存（履歴として）
- 新旧比較によるスキル成長の可視化

### 3. クラス設計とOOP
- 単一責任の原則（SRP）
- 依存性の注入パターン
- カプセル化とプロパティ管理

### 4. Windows API活用
- ウィンドウ操作の自動化
- マウス座標の動的計算
- 画面解像度対応

### 5. パフォーマンス最適化
- 配列処理による高速化
- AutoFilterの効率的使用
- 外部ブック管理の最適化

### 6. 業務分析と要件定義
- 現場の課題を技術で解決
- 測定可能な成果（工数削減）
- 部署横断での展開能力

---

## セキュリティ対応

### 実施済みサニタイゼーション
1. **システム名のマスク**
   - 実システム名 → `ERPSystem` に置換
   - 19箇所を自動置換
   - バックアップファイルを保管

2. **外部ファイル参照の匿名化**
   - 実ファイル名 → ダミー名に変更
   - シート名の匿名化

3. **機密情報の確認**
   - パスワード・認証情報: なし
   - 個人情報: 検出されず
   - データベース接続情報: なし

### 分析ツール
- **olevba**: VBAコード抽出と分析
- **カスタムPythonスクリプト**: バイナリレベルの置換

詳細: [サニタイゼーション報告書.md](サニタイゼーション報告書.md)

---

## ファイル構成

```
excel/
├── README.md                          # 本ファイル
├── .gitignore                         # Git除外設定
├── 基幹システム連携.xlsm               # 発注自動化ツール
├── 検査証_一般.xlsm                    # 検査証管理システム
├── 分析レポート.md                     # 詳細分析レポート
├── サニタイゼーション報告書.md         # セキュリティ対応報告
└── 基幹システム連携.xlsm.backup       # バックアップ（非公開）
```

---

## 使用方法

### 基幹システム連携.xlsm

1. **事前準備**
   - ERPシステムをメイン画面で起動
   - Excelシートに発注データを入力

2. **実行**
   - ボタンをクリック
   - 確認ダイアログで「はい」

3. **自動処理**
   - データ読み込み
   - ERP画面への自動入力
   - 完了メッセージ表示

### 検査証_一般.xlsm

1. **日付範囲入力**
   - D3またはE3セルに日付を入力
   - 自動的にデータインポート開始

2. **数量調整**
   - 分割: Q列のセルを選択 → 分割数入力
   - 統合: R列のセルを選択 → 自動統合

3. **印刷**
   - 印刷ボタンをクリック
   - 最大32件まで自動ページネーション

---

## トラブルシューティング

### Q1: ERPシステムが見つからない
**A**: `AppActivate "ERPSystem"` の名前を実際のアプリケーション名に変更してください。

### Q2: マウス座標がずれる
**A**: 画面解像度が異なる場合、`refWidth`, `refHeight` の値を調整してください。

### Q3: データインポートが失敗する
**A**: 外部ファイルのパスとシート名を確認してください（[import.bas](import.bas):203行目）。

---

## 今後の改善予定

### Phase 5: 現代化（計画中）
1. **Python移行検討**
   - pywinauto による GUI自動化
   - pandas によるデータ処理
   - openpyxl による Excel操作

2. **CI/CD導入**
   - Git管理の強化
   - 自動テストの導入
   - バージョン管理の改善

3. **ログ機能追加**
   - 実行履歴の記録
   - エラーログの保存
   - パフォーマンスモニタリング

---

## ライセンス

このプロジェクトは私的利用のため、非公開です。
コードの一部または全部の商用利用・再配布は禁止されています。

---

## 謝辞

本プロジェクトは製造現場の業務効率化要請に応え、8年間にわたり改善を続けてきました。以下のスキルを証明します:

- **15年のVBA経験:** 2009年から継続的に使用
- **8年の運用実績:** 段階的な拡大と継続的改善
- **スモールスタート戦略:** 一部から始め、成功を確認しながら横展開
- **累計9,600時間の削減:** 測定可能な巨大なインパクト

現場の課題を技術で解決し、継続的に改善する姿勢を示すプロジェクトです。

---

## 連絡先

このプロジェクトに関するご質問は、ポートフォリオサイトからお問い合わせください。

---

**作成日**: 2025-10-13
**最終更新**: 2025-10-13
