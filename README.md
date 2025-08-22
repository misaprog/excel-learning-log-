# excel-learning-log-

# Excel 学習ログ - Week 1

## 📌 課題内容
- 複数シートの整理（Q1〜Q4、売上サマリー）
- タブの色を設定（赤＝四半期売上、青＝サマリー）
- 3Dセル参照で四半期合計を算出
- Consolidateツールで在庫を集計

---

## 📊 成果画面

### 売上サマリーシート

<img width="665" height="912" alt="スクリーンショット 2025-08-22 152557" src="https://github.com/user-attachments/assets/fcfa6323-04b5-44a3-8018-879344339237" />







### 在庫統合（Consolidate結果）

<img width="1499" height="904" alt="スクリーンショット 2025-08-22 154055" src="https://github.com/user-attachments/assets/145cba62-d960-4429-bb39-e881514aca4b" />


---

## ✨ 学んだこと
- 3Dセル参照でシートをまたいだ集計が簡単にできる  
- タブの色を工夫するとシート構造が直感的になる  
- Consolidateツールは異なるデータ構造でも合算可能

---

## 第2回 練習課題：ヘルプデスクチケットの統合サマリー

この練習課題では、前回より少し難しいステップを含む Excel の操作を行います。ステップ1〜6は基本的な作業で、ステップ7〜9ではビデオで学んだ「Consolidate（統合）ツール」を使用して、異なる順序やカテゴリのデータをまとめます。

---

### ステップ 1

ワークブックをダウンロードして Excel で開きます。

### ステップ 2

シート名を変更します：

* Sheet1 → 4月
* Sheet2 → 5月
* Sheet3 → 6月
* Sheet4 → レポート
  
  <img width="458" height="834" alt="スクリーンショット 2025-08-22 160418" src="https://github.com/user-attachments/assets/6a138367-a9ce-4dc1-b511-da082f22dd9c" />


### ステップ 3

6月シートのコピーを作成し、ワークブックの最初に移動して名前を「サマリー」に変更します。

<img width="509" height="828" alt="スクリーンショット 2025-08-22 161512" src="https://github.com/user-attachments/assets/87d6754a-ca45-4404-a3c6-0852cdf10caa" />


### ステップ 4

タブの色を変更して視覚的な区別をつけます：

* 月次シート → 赤
* サマリーシート → 青
  
  <img width="384" height="86" alt="スクリーンショット 2025-08-22 161736" src="https://github.com/user-attachments/assets/3c3c925d-5177-489e-83a0-d98e888321e3" />


### ステップ 5

3Dセル参照を使用して、4月〜6月のヘルプデスク時間をサマリーシートに合計します。

<img width="586" height="875" alt="スクリーンショット 2025-08-22 162839" src="https://github.com/user-attachments/assets/6e923b18-8829-4ec1-a80f-5f837386d308" />



### ステップ 6

レポートシートでリンク式を使用して、4月〜6月の合計チケット数を取得します。

<img width="925" height="874" alt="スクリーンショット 2025-08-22 172931" src="https://github.com/user-attachments/assets/d9555cf6-a291-41fc-af72-cb6f344a761d" />

---

### ステップ 7（少し難しい）

Consolidate ツールを使用して、以下のサマリーを作成します：

* 対象：5月第4週、6月第1週、6月第2週
* 集計：優先度ごとのチケット件数
* 並べ替え：優先度順
  
  <img width="455" height="544" alt="スクリーンショット 2025-08-22 215935" src="https://github.com/user-attachments/assets/9e8d730f-f1c4-4a01-a3c8-107250479c4b" />


### ステップ 8

Consolidate ツールで以下を集計：

* 平均オープン日数（平均値を使用）
* 対象：5月第4週、6月第1週、6月第2週
* 小数点以下2桁に書式変更（四捨五入は使用しない）
* 並べ替え：優先度順
  
  <img width="455" height="544" alt="スクリーンショット 2025-08-22 215935" src="https://github.com/user-attachments/assets/dba5f5bd-73b8-40dc-ab60-4bafce11c35d" />


> ⚠️ 他の参照は必ず削除してから次のステップへ進む

### ステップ 9

Consolidate ツールで以下を集計：

* 満足度評価ごとのチケット数
* 使用関数：COUNT
* 並べ替え：満足度順
  
  <img width="455" height="544" alt="スクリーンショット 2025-08-22 215935" src="https://github.com/user-attachments/assets/62772430-af06-4e92-8380-02fd08c966d0" />

### ステップ７．９はcountif関数で、合計を出してから統合する！

---

🎉 **お疲れさまでした！**
第2回練習課題は完了です。ダウンロードしたワークブックと自分の解答を比較して、正しく集計できているか確認しましょう。

---



