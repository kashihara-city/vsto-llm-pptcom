# vsto-llm-pptcom

## 概要

PowerPoint で開いているプレゼンテーションを、LLM が作ったり直したりするアドインです。
主要依存ライブラリのバージョンは以下の通りです。npm への依存はありません。
- PDF.js 6.4.299
- anydoc 0.2.4
- Newtonsoft.Json 13.0.3

## 動作環境

- Windows 11
- PowerPoint 2019 以降
- Microsoft Edge WebView2 ランタイム
- OpenAI 互換の API を持つ接続先（チャット用モデル。任意で画像生成用モデル。）

## 主な機能

- 「こんなパワポ作って」などの自然言語でパワーポイントを作成できます。
- Word・Excel・PowerPoint・PDF・テキストのファイルを添付し、「この資料を5枚にまとめて」のような依頼も可能です。
- 資料の図や写真を、そのままスライドに貼ることもできます。
- LLM は VBA に近い短いスクリプト（pptcom スクリプト）を書きます。アドインがそれを検査してから PowerPoint で実行します。

![橿原市第5期地域福祉推進計画（PDF）を添付して「この資料を5枚にまとめて」と依頼し、続けて1枚目の文字の高さを揃えるよう頼んだところ。右が作業ウィンドウで、オンプレミスで動かしているモデルを使っている](docs/image.png)

## 必要な設定

アドインとしての動作のために、所定の場所にファイルを配置する必要があります。  
このため、インストーラ（MSI）でのインストールが必要です。管理者権限は不要です。

設定は、インストール後に、作業ウィンドウの上の歯車から行います。設定項目は以下のとおりです。

- （必須）接続先URLと、使うモデル
- （必要な場合）API キー
- （既定でオフ）途中経過を保存する: オンにすると、依頼ごとの記録を `%LOCALAPPDATA%\KashiharaCity\pptcom\` に残します
- （既定でオフ）外部からの操作を許可する: テスト用です。オンにすると、ほかのプログラムからこのアドインに依頼を送れます

## インストールされるものと書き込む場所

このアドインは、ログオンしているユーザーの範囲に次のものを置きます。コンピューター全体の場所（Program Files、HKEY_LOCAL_MACHINE）には書きません。

- **Windows 資格情報マネージャー**に、接続先の API キー
- **レジストリ**（`HKEY_CURRENT_USER`）に、PowerPoint へのアドインの登録、信頼の設定、作業ウィンドウの設定
- **`%LOCALAPPDATA%\KashiharaCity\pptcom\`** に、アドインのファイル（`app\`）、作業中の一時ファイル（`temp\`）、「途中経過を保存する」がオンのときの記録

インストーラー（MSI）が置いたもの（`app\` とレジストリの登録）はアンインストールで消えます。アドインが使いながら書いたもの（API キー、作業ウィンドウの設定、`temp\` と記録）は残るので、不要なら手で消してください。

### Windows 資格情報マネージャー

| 名前 | 内容 | 書くとき |
|---|---|---|
| `KashiharaCity/pptcom:api_key`（汎用資格情報） | 接続先の API キー。作業ウィンドウには表示しない | 作業ウィンドウで API キーを設定したとき（アドイン） |

### レジストリ

アドインの設定は以下のレジストリに保存します。

| キー（`HKEY_CURRENT_USER\Software\…`） | 値 | 内容 | 書くもの |
|---|---|---|---|
| `KashiharaCity\pptcom` | `endpoint`、`chat_model`、`image_model`（文字列）、`save_progress`、`external_control`（0/1） | 作業ウィンドウの設定（接続先、モデル、2 つのスイッチ） | アドイン（設定を変えたとき） |
| `KashiharaCity\pptcom\Installer` | `Files`、`WebView2LoaderX86`、`WebView2LoaderX64` | インストーラーが自分の部品の有無を判定するための目印 | MSI |

次のキーは、Office と VSTO が場所を決めているものです。

| キー（`HKEY_CURRENT_USER\Software\…`） | 値 | 内容 | 書くもの |
|---|---|---|---|
| `Microsoft\Office\PowerPoint\Addins\pptcom` | `Description`、`FriendlyName`、`LoadBehavior`（3）、`Manifest`（`file:///…/pptcom.vsto\|vstolocal`） | PowerPoint の起動時にアドインを読み込む登録。PowerPoint はこの場所でアドインを探す | MSI |
| `Microsoft\VSTO\Security\Inclusion\{9E192468-0F34-4BCE-9263-052B3358849B}` | `Url`（上の .vsto）、`PublicKey`（.vsto に署名した証明書の公開鍵） | VSTO の信頼済み一覧。この .vsto をこの鍵の署名で確認なしに読み込ませる（証明書をストアに入れる必要はない） | MSI |
| `Microsoft\VSTO\SolutionMetadata` | 読み込んだ .vsto の URL ごとの情報 | VSTO ランタイムが自分で記録する | VSTO |
| `Microsoft\Office\16.0\PowerPoint\AddInLoadTimes` | `pptcom` | アドインの読み込みにかかった時間 | Office |


### `%LOCALAPPDATA%\KashiharaCity\pptcom\`

`app\`（MSI が置く。アンインストールで消える）:

| ファイル | 内容 |
|---|---|
| `pptcom.vsto`、`pptcom.dll.manifest` | VSTO がアドインを読み込むための署名付きの設定ファイル |
| `pptcom.dll` | アドイン本体（Newtonsoft.Json と WebView2 SDK、作業ウィンドウの画面、資料の読み取りに使う anydoc と PDF.js を含む） |
| `Microsoft.Office.Tools.Common.v4.0.Utilities.dll` | VSTO の部品 |
| `runtimes\win-x86\native\WebView2Loader.dll`、`runtimes\win-x64\native\WebView2Loader.dll` | WebView2 の読み込み用（PowerPoint のビット数に合うほうを使う） |
| `THIRD-PARTY-NOTICES.md` | 同梱している第三者のコンポーネントとライセンス |

`temp\`（アドインが書く。PowerPoint を閉じているときなら、いつ消してもかまわない）:

| ファイル | 内容 |
|---|---|
| `WebView2\` | 作業ウィンドウ（WebView2）の閲覧データ。作業ウィンドウを開くと作られる |
| `work\` | 作業中の一時ファイル（スライドの描画、生成した画像、添付した資料の画像と PDF のページの画像、デッキの既定値を測るためのコピー）。描画と測定用のコピーは使い終わると消し、生成した画像と資料の画像は会話が終わる（新しい会話、別のプレゼンテーション、PowerPoint の終了）と消す |

記録（アドインが書く。「途中経過を保存する」がオンのときだけ）:

| ファイル | 内容 |
|---|---|
| `session.log` | 依頼、モデルの推論と応答、ツールの実行結果 |
| `com_usage.jsonl` | 実行した PowerPoint の操作の記録 |
| `generated-pptx\addin_<日時>\` | 依頼ごとのデッキのコピー、スライドの画像、記録、添付した資料から読み取った文字 |

## pptcom.dllの構造

```
PowerPoint
└─ pptcom.dll
   ├─ C# のコード
   │   ├─ ThisAddIn             エージェントのループ、会話（添付資料も会話に属す）
   │   ├─ ComProxy              pptcom（資料の画像も d1_img3 などの変数として貼れる）
   │   ├─ Documents             添付資料の保持、docread、画像の変換
   │   └─ AiTaskPaneControl     作業ウィンドウ（WebView2）の土台。DLL 内のファイルを返す
   ├─ 結合したライブラリ         Newtonsoft.Json、WebView2 SDK（Core、WinForms）
   └─ 埋め込みのファイル（https://pptcom.invalid/ で作業ウィンドウに渡す）
       ├─ 画面          index.html / taskpane.js / taskpane.css
       ├─ docs.js       ドロップと添付の受付、PDF のページの描画
       ├─ docworker.js  Word・Excel・PowerPoint → 文字（anydoc）
       ├─ pdfdoc.js     PDF → 文字と画像（PDF.js）
       └─ vendor/       anydoc（WebAssembly）、PDF.js
```

## ビルド

次のものが必要です。

- Visual Studio 2022 以降（「Office/SharePoint 開発」ワークロード）
- Rust 1.99.0 と wasm32-unknown-unknown ターゲット（`rustup toolchain install 1.99.0 --profile minimal --target wasm32-unknown-unknown`）

```
powershell -ExecutionPolicy Bypass -File scripts\build\build.ps1
```

作業ウィンドウで使う PDF.js と anydoc は、公式の配布元から取ったものを `deps/` に置き、版とハッシュを `deps/web-sources.lock.json` で固定しています。npm は使いません。

- PDF.js: Mozilla が GitHub のリリースで配布している `pdfjs-<版>-dist.zip` から、使うファイルだけを取り出します。
- anydoc: タグのコミットのソース（`deps/anydoc.tar.gz`）から、anydoc 自身のリリースと同じ手順（cargo → wasm-bindgen → wasm-opt）で WebAssembly に組みます。wasm-bindgen と wasm-opt は公式のビルド済みを、lock のハッシュで確かめてから使います。最初のビルドは数分かかり、Rust のライブラリ（crates.io、Cargo.lock で固定）と 2 つの道具を取るためにネットワークが要ります。組んだ結果は `pptcom/obj/webvendor/` に残し、lock が変わるまで組み直しません。

版を上げるときは `scripts/deps/fetch-web.py` の版を書き換えて実行します。公開から 7 日たっていない版は取りません。

`python -X utf8 scripts/test/no-secrets.py` は、Git が追跡しているファイルに、私有 IP アドレス、社内のホスト名、API キーらしき文字列、開発機のパス、鍵や証明書の中身が混ざっていないかを調べます（CI でも毎回流します）。`--history` を付けると、すべてのコミットのファイルとコミットのメッセージを調べます。パターンでは見分けられない接続先のホスト名やモデル名は、Git の管理外の `notes/secrets-denylist.txt`（1 行に 1 語）に書いておくと、それも照合します。

## リポジトリの構成

- `pptcom/` アドイン本体（C#、VSTO）
- `pptcom/ComProxy/` pptcom スクリプトの解析と実行、使ってよいメンバーの検査
- `pptcom/UI/` 作業ウィンドウ（`web/` が画面の HTML・CSS・JavaScript。DLL に埋め込み、WebView2 に渡す）
- `pptcom/Documents/` 添付した資料の保持、docread、資料の画像
- `deps/` 作業ウィンドウで使う PDF.js と anydoc の、公式の配布元から取ったもの（`web-sources.lock.json` に版・取得元・ハッシュ）
- `scripts/deps/` `deps/` を更新するスクリプト
- `scripts/build/` MSI のビルド（WiX）、anydoc に含まれる Rust ライブラリのライセンス一覧の生成
- `scripts/test/` テストと実験用の道具（構文テスト、PowerPoint を外から動かすランナー、アドインを外から操作するスクリプトなど）
- `licenses/` 同梱している第三者のコンポーネントとライセンス（`THIRD-PARTY-NOTICES.md`。MSI にも入れます）
- `LICENSE` このリポジトリのライセンス（MIT）
- `SECURITY.md` 脆弱性の報告のしかた
- `docs/` README の画像
