# わが家の家計簿（ホーム画面アプリ）

## 👉 [アプリを開く](https://24tu.github.io/kakeibo-app/)

https://24tu.github.io/kakeibo-app/

スマホで開いたら「ホーム画面に追加」すると、アプリのように全画面で使えます。

- iPhone：Safariの共有ボタン（□に↑）→「ホーム画面に追加」
- Android：Chromeの右上「︙」→「ホーム画面に追加」

---

スマホのホーム画面に追加して使う、個人用の家計簿アプリです。
データはこのリポジトリには入っておらず、持ち主のGoogleスプレッドシートに保存されます（合言葉つき）。

- `index.html` … アプリ本体
- `sw.js` / `manifest.webmanifest` … ホーム画面アプリにするための設定
- `apps_script/Code.gs` … スプレッドシートに貼り付ける保存係のプログラム
