# J2E Excel Add-in

JSONテキストをExcelの値、Entity、spill配列に変換するExcelカスタム関数アドインです。

公開サイト: <https://koizumihiroo.github.io/j2e-excel-addin/>

## 利用上の注意

このアドインは現時点の状態のまま提供しており、動作、継続的な提供、特定目的への適合性を含め、明示または黙示を問わず一切保証しません。公開後でも、セキュリティ上または運用上の理由により、予告なく公開を一時停止または終了する場合があります。

## 配布ファイル

- [Excel用manifest](https://koizumihiroo.github.io/j2e-excel-addin/manifest.xml)
- [GitHub Pagesの公開ページ](https://koizumihiroo.github.io/j2e-excel-addin/)
- [関数ランタイムHTML](https://koizumihiroo.github.io/j2e-excel-addin/functions.html)
- [関数JavaScript](https://koizumihiroo.github.io/j2e-excel-addin/functions.js)
- [関数メタデータ](https://koizumihiroo.github.io/j2e-excel-addin/functions.json)

## このアドインを試す

このアドインは現在Microsoft Marketplaceには掲載されていません。試す場合は、利用者自身が以下の手順で[manifest.xml](https://koizumihiroo.github.io/j2e-excel-addin/manifest.xml)をダウンロードし、自分が使用するExcelに追加します。

### 必要条件

- Excel for the web、またはCustom Functions Runtime 1.4以降に対応するMicrosoft 365版Excel for Windowsを使用すること。
- 個人用Microsoftアカウントまたは組織のMicrosoft 365アカウントで、Excelにサインインしていること。
- GitHub PagesとMicrosoftがホストするOffice.jsへ接続できるインターネット接続があること。

ネットワーク接続は初回の関数実行時だけでなく、アドインの読み込み、数式の再計算、ブックを再度開くときにも必要になる可能性があります。Excelのキャッシュ動作は環境によって異なるため、オフラインでの動作は保証されません。

### Excel for the web / Excel for Windows（Microsoft 365デスクトップアプリ）

1. [manifest.xml](https://koizumihiroo.github.io/j2e-excel-addin/manifest.xml) を自分の端末へダウンロードします。
2. Excelで、自分が使用するブックを開きます。
3. **ホーム > アドイン** を選びます。
4. メニュー右下の **+ その他のアドイン** を選びます。
5. **Office アドイン** 画面で、**ストア** ではなく **個人用アドイン** を選びます。
6. 右上の **個人用アドインの管理** を選びます。
7. **アドインのアップロード** を選び、自分でダウンロードした `manifest.xml` を選択して **アップロード** します。

利用者の組織ポリシーによっては、独自アドインの追加に管理者の許可が必要です。追加後は、Excelのセルで `=J2E.JSON(A1)` などのカスタム関数を使用できます。

**個人用アドイン** に **アドインのアップロード** が表示されない場合は、組織ポリシーで独自アドインのアップロードが無効になっています。組織のMicrosoft 365管理者へ許可を依頼してください。

Excelのバージョン、更新チャネル、アカウントの種類、組織ポリシーによって画面構成やボタン名が異なる場合があります。その場合は、**アドイン**、**個人用アドイン**、**個人用アドインの管理**、**アドインのアップロード** に相当する項目を探してください。

## アドインを削除する

アドインを追加前の状態へ戻す場合は、Excelの **マイ アドイン** からJ2Eを削除します。

1. Excel for the webでは **ホーム > アドイン > その他の設定**、Excel for Windowsでは **挿入 > アドイン > マイ アドイン** を開きます。
2. 一覧から **JSON to Excel Entity** または **J2E** を選び、**削除** を実行します。
3. ダウンロードした `manifest.xml` が不要であれば、端末から削除します。

ダウンロード済みの `manifest.xml` を端末から削除するだけでは、Excelに追加済みのアドインは削除されません。アドインを削除しても、既存のセルに入力した数式や計算結果は自動的には削除されません。

## 関数

関数例を続けて試す場合は、次のタブ区切りデータを空の`Sheet1`の`A1`へ一度だけ貼り付けます。`A1`にJSON object、`B1`にJSON array、`C1`に空のJSON objectが入ります。

```text
{"商品":"りんご","単価":120,"在庫":25}	[[1,2],[3,4]]	{}
```

コードブロックの1行全体をコピーし、`Sheet1!A1`を選択して **Ctrl+V** で貼り付けてください。右クリックで貼り付ける場合は、**貼り付けのオプション > 値** を選びます。3つの値がA1の1セルに入った場合は、A1を選択して **データ > 区切り位置 > 区切り記号付き > タブ** を選びます。

`OBJECT_KEYS`、`OBJECT_VALUES`、`ARRAY` は複数セルへ結果をスピルします。各関数の数式セルの下と右側を空けてください。

### `J2E.JSON(jsonText)`

任意のJSON文字列をExcelの値へ変換します。文字列、数値、真偽値は通常のExcel値として返し、オブジェクトまたは配列はExcel Entityとして返します。

**既知の不具合:** Excel JavaScript APIには[空セル値（`EmptyCellValue`）](https://learn.microsoft.com/en-us/javascript/api/excel/excel.emptycellvalue?view=excel-js-preview)が定義されており、Excelカスタム関数も[同じJSONスキーマ](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-data-types-concepts)を使用します。それでも、現在のExcelアプリでは `{"key":null}` のようなJSONを `J2E.JSON` で評価すると `#VALUE!` エラーになります。

**例**

`A1`のJSON objectを変換する数式を`D1`に入力します。

```excel
=J2E.JSON(A1)
```

`D1`は`商品`、`単価`、`在庫`をプロパティに持つEntityになります。

### `J2E.OBJECT(jsonText)`

JSONオブジェクトだけをExcel Entityへ変換します。最上位が配列、文字列、数値、真偽値、`null` のJSONは受け付けません。

`A1`のobjectをExcel Entityへ変換する数式を`E1`に入力します。

```excel
=J2E.OBJECT(A1)
```

`E1`は`商品`、`単価`、`在庫`をプロパティに持つEntityになります。

### `J2E.OBJECT_KEYS(input)`

JSONオブジェクトの文字列、または `J2E.JSON` / `J2E.OBJECT` が返したEntityを受け取り、キーを1列にスピルします。空のオブジェクトはスピルできません。

`A1`のobjectのキーを取得する数式を`F1`に入力します。

```excel
=J2E.OBJECT_KEYS(A1)
```

`F1:F3`に`商品`、`単価`、`在庫`が順に表示されます。

### `J2E.OBJECT_VALUES(input)`

JSONオブジェクトの文字列、または `J2E.JSON` / `J2E.OBJECT` が返したEntityを受け取り、値を1列にスピルします。空のオブジェクトはスピルできません。

`A1`のobjectの値を取得する数式を`G1`に入力します。

```excel
=J2E.OBJECT_VALUES(A1)
```

`G1:G3`に`りんご`、`120`、`25`が順に表示されます。

### `J2E.IS_EMPTY_OBJECT(input)`

JSON文字列またはExcel Entityが空のオブジェクトかどうかを判定し、`TRUE` または `FALSE` を返します。オブジェクト以外のJSONは `FALSE` になります。

空のobjectが入った`C1`を判定する数式を`H1`に入力します。

```excel
=J2E.IS_EMPTY_OBJECT(C1)
```

`H1`は`TRUE`になります。Entityを判定するには、`H2`に`=J2E.IS_EMPTY_OBJECT(D1)`を入力します。結果は`FALSE`です。

### `J2E.ARRAY(jsonText)`

JSON配列をExcelのスピル配列へ変換します。1次元配列は1列に、入れ子の配列は行列としてスピルします。空配列と、行ごとの列数が異なる配列は受け付けません。

`B1`のJSON arrayをスピルする数式を`I1`に入力します。

```excel
=J2E.ARRAY(B1)
```

`I1:J2`に1行目`1` / `2`、2行目`3` / `4`が表示されます。

## プライバシー

J2Eの関数コードは、受け取ったJSONテキストをExcel内で処理します。関数実装からJSONの内容をJ2Eのサーバーへ送信する処理はありません。Excelはこのサイトから関数ファイルを読み込み、Office.jsはMicrosoftがホストするURLから読み込みます。

## サポート

不具合や質問は[GitHub Issues](https://github.com/koizumihiroo/j2e-excel-addin/issues)へ報告してください。
個別のサポート対応は行っておらず、問い合わせへの回答や修正対応は保証しません。

## 公開リポジトリ管理者向け: デプロイ

GitHub Actionsは、`main`へのpush時に配布ファイルを検証し、GitHub Pagesへデプロイします。Pull Requestでは検証だけを実行します。

検証環境は次のファイルで固定しています。

- Node.js: [`.node-version`](.node-version) の `24.21.0`
- pnpm: [`package.json`](package.json) の `12.8.1`

検証スクリプトはNode.js標準ライブラリだけを使用し、指定したNode.js以外の実行を拒否します。ローカルでは次のコマンドで同じ検証とPages成果物の作成を行えます。

```bash
corepack enable
pnpm install --frozen-lockfile
pnpm run verify
pnpm run site
```

`pnpm run site` は配布物を検証してから、`_site/` に公開対象の静的ファイルを作成します。`pnpm run verify` は検証だけを実行します。`pnpm run site` は既存の `_site/` も再生成できますが、生成マーカーのない旧出力は配布対象の6ファイルだけで構成されている場合に限って引き継ぎます。想定外のファイル、ディレクトリ、シンボリックリンク、または不正なマーカーがある場合は、内容を保護するため更新を拒否します。

初回公開時は、次のGitHub CLIコマンドでPagesビルド方式をGitHub Actionsへ設定してから`main`へpushします。

```bash
REPOSITORY="$(gh repo view --json nameWithOwner --jq .nameWithOwner)"
if PAGES_RESPONSE="$(gh api --include "repos/${REPOSITORY}/pages" 2>&1)"; then
  gh api --method PUT "repos/${REPOSITORY}/pages" -f build_type=workflow
elif printf '%s\n' "${PAGES_RESPONSE}" | grep -Eq '(^HTTP/[0-9.]+ 404|HTTP 404)'; then
  gh api --method POST "repos/${REPOSITORY}/pages" -f build_type=workflow
else
  printf '%s\n' "${PAGES_RESPONSE}" >&2
  exit 1
fi
```

このコマンドはPagesリソースを取得し、存在する場合は`PUT`、存在せず404の場合だけ`POST`を実行します。認証・権限・通信など404以外の失敗は中断するため、設定失敗を「未設定」と誤判定しません。

設定後、公開対象を`main`へpushするとデプロイが開始されます。

```bash
git add .
git commit -m "Prepare GitHub Pages deployment"
git push -u origin main
```
