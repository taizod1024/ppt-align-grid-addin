# ppt-align-grid-addin

PowerPointの図形をグリッドに揃えるアドイン

## 使い方

1. ppt-align-grid.ppamをPowerPointアドインとして読み込みます。
2. [表示]タブの右側に[図形]セクションが追加されます。

![ppt-align-grid](img/ppt-ailgn-grid.png)

### グリッド線に揃える

- 図形を選択していない状態で[グリッド線に揃える]を押すとスライドマスター以外のすべての図形をグリッド線に整列します。
- 図形を選択している状態で[グリッド線に揃える]を押すと選択している図形をグリッド線に整列します。

### 片側接続のコネクタ

- 片側だけ図形に接続しているコネクタを探します。

## .ppam編集方法

1. ./src/ppt-align-grid.pptxを編集
2. 編集結果はbasとしてエクスポート
3. PowerPointで新規ファイルを作成し、basをインポートして./ppt-align-grid.ppamとして保存
4. 一時的に./ppt-align-grid.ppam.zipとしてから./src/ppt-align-grid/配下をコピー
5. ./ppt-align-grid.ppamに名前を戻す