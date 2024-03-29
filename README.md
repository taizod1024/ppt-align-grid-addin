# ppt-align-grid-addin

PowerPoint の図形をグリッド線に揃えるアドイン

## 使い方

1. ppt-align-grid.ppam を PowerPoint アドインとして読み込みます。
2. [表示]タブの右側に[図形]セクションが追加されます。

![ppt-align-grid](images/ppt-ailgn-grid.png)

### グリッド線に揃える

- 図形を選択していない状態で[グリッド線に揃える]を押すとスライドマスター以外のすべての図形をグリッド線に整列します。
- 図形を選択している状態で[グリッド線に揃える]を押すと選択している図形をグリッド線に整列します。

### 片側接続のコネクタ

- 片側だけ図形に接続しているコネクタを探します。

### リンク切れの URL

- すべてのスライドの図形のテキストおよび画像からリンク切れしている URL を探して選択します。
- HTTP ステータス 200 以外の場合をリンク切れと判定しています。

## .ppam 編集方法

1. ./src/ppt-align-grid.pptm を編集
2. 編集結果は bas としてエクスポート
3. PowerPoint で新規ファイルを作成し、bas をインポートして./ppt-align-grid.ppam として保存
4. 一時的に./ppt-align-grid.ppam.zip としてから./src/ppt-align-grid.ppam/配下をコピー
5. ./ppt-align-grid.ppam に名前を戻す
