# BAMBITCH宇都宮 ご来店予告掲示板 アーカイバ

[BAMBITCH宇都宮 ご来店予告掲示板](https://bambitch-utsunomiya.com/cgi-bin-utnmy/imgboard.cgi)(KENT-WEB `imgboard.cgi`)を **毎日 GitHub Actions で自動取得** し、投稿と画像をこのリポジトリに蓄積します。Instagram アーカイバ(`instagram-archiver`)と同じ「Actions を毎日回して日々の情報を貯める」方式です。

## 何を保存するか

| 保存先 | 内容 |
|--------|------|
| `posts.json` | 全投稿のインデックス(No. をキー / 日時・タイトル・本文・画像ファイル名) |
| `data/YYYY-MM-DD.md` | 投稿日ごとの読みやすい Markdown ログ |
| `img-box/imgYYYYMMDDHHMMSS.jpg` | 添付画像(掲示板と同じファイル名) |

投稿 No. で重複排除するので、毎日実行しても **新規分だけが追記** されます(冪等)。

## 自動実行

`.github/workflows/bbs-archiver.yml`

- **スケジュール**: 毎日 22:10 UTC(= 翌 07:10 JST)
- 取得 → 新規分をコミット&プッシュ → スナップショットを artifact 保存
- 認証情報は不要(公開掲示板のため Secrets 設定なし)

### 手動実行 / 過去分の取得

GitHub の **Actions → BBS Archiver → Run workflow** から実行できます。
`max_pages` を大きく(例: `18`)すると過去ページまで遡って取得します。

## ローカル実行

```bash
pip install -r requirements.txt
python bambitch_bbs.py           # 既定は直近3ページ
BBS_MAX_PAGES=18 python bambitch_bbs.py   # 全ページ遡り
```

## 環境変数

| 変数 | 既定 | 説明 |
|------|------|------|
| `BBS_MAX_PAGES` | `3` | 取得ページ数(1ページ=約15〜18投稿) |
| `BBS_BASE_URL` | BAMBITCH宇都宮の CGI URL | 対象掲示板の URL |

## 注意

- 公開掲示板の情報を個人的にアーカイブする目的の構成です。ページ取得の間に 1.5 秒のウェイトを入れ、サーバ負荷に配慮しています。
- 文字コードは Shift_JIS を UTF-8 に変換して保存します。
