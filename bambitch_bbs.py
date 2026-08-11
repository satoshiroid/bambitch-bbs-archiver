#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BAMBITCH宇都宮 ご来店予告掲示板(KENT-WEB imgboard.cgi) アーカイバ。

公開掲示板の投稿(No./日時/タイトル/投稿者/本文)と添付画像を取得し、
- 画像: img-box/ に保存
- 投稿インデックス: posts.json に No. をキーとして蓄積(重複は追記しない)
- 日次ログ: data/YYYY-MM-DD.md に投稿日ごとにまとめて出力

認証は不要。GitHub Actions から毎日実行し、差分をコミットして履歴を蓄積する。

環境変数:
  BBS_MAX_PAGES  取得するページ数(既定: 3 ページ = 直近の投稿を十分カバー)
  BBS_BASE_URL   掲示板 CGI の URL(既定は BAMBITCH宇都宮)
"""
from __future__ import annotations

import html as html_lib
import json
import os
import re
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

import requests

JST = timezone(timedelta(hours=9))

BASE_URL = os.environ.get(
    "BBS_BASE_URL",
    "https://bambitch-utsunomiya.com/cgi-bin-utnmy/imgboard.cgi",
)
# CGI と同じディレクトリ(相対 ./img-box/ の解決に使う)
ROOT_URL = BASE_URL.rsplit("/", 1)[0] + "/"

MAX_PAGES = int(os.environ.get("BBS_MAX_PAGES", "3"))

REPO_DIR = Path(__file__).resolve().parent
IMG_DIR = REPO_DIR / "img-box"
DATA_DIR = REPO_DIR / "data"
POSTS_JSON = REPO_DIR / "posts.json"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15"
)

session = requests.Session()
session.headers.update({"User-Agent": USER_AGENT})


def log(msg: str) -> None:
    print(f"[{datetime.now(JST):%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def fetch_page(page: int) -> str:
    """指定ページの HTML を Shift_JIS からデコードして返す。"""
    if page <= 1:
        url = BASE_URL
    else:
        url = f"{BASE_URL}?amode=&p1=&p2=&bbsaction=page_change&page={page}"
    resp = session.get(url, timeout=30)
    resp.raise_for_status()
    # KENT-WEB imgboard は Shift_JIS
    resp.encoding = "shift_jis"
    return resp.text


def strip_tags(fragment: str) -> str:
    """<BR> を改行に、その他タグを除去してテキスト化する。"""
    text = re.sub(r"(?i)<br\s*/?>", "\n", fragment)
    text = re.sub(r"(?is)<[^>]+>", "", text)
    text = html_lib.unescape(text)
    # 行末の空白を整理
    lines = [ln.rstrip() for ln in text.splitlines()]
    return "\n".join(lines).strip()


# 各記事は削除用チェックボックス <INPUT ... NAME="rmid...S{timestamp}"> で始まる
ARTICLE_SPLIT = re.compile(r'(?i)<INPUT\s+TYPE="?CHECKBOX"?\s+NAME="rmid')

RE_NO = re.compile(r"No\.(\d+)")
RE_DATE = re.compile(r"\[(\d{4}/\d{2}/\d{2}),(\d{2}:\d{2}:\d{2})\]")
RE_IMG = re.compile(r'(?i)<IMG\s+SRC="(\.?/?img-box/[^"]+)"')
RE_IMG_TITLE = re.compile(r'(?i)画像タイトル：<A[^>]*>([^<]+)')
RE_TITLE = re.compile(r'(?i)<FONT\s+SIZE="?\+1"?[^>]*>\s*<B>(.*?)</B>', re.S)
RE_BODY = re.compile(r"(?is)<BLOCKQUOTE[^>]*>(.*?)</BLOCKQUOTE>")
# 投稿者名: 日付/No. の直前に <FONT COLOR=...><B>名前</B></FONT> が並ぶことが多い。
# タイトル FONT(+1) の後、最初の [日付] までの間の <B>...</B> を投稿者候補とする。


def parse_articles(page_html: str) -> list[dict]:
    """1 ページ分の HTML から記事レコードのリストを返す。"""
    chunks = ARTICLE_SPLIT.split(page_html)
    # split の先頭はヘッダ部なので捨てる
    articles: list[dict] = []
    for chunk in chunks[1:]:
        m_no = RE_NO.search(chunk)
        m_date = RE_DATE.search(chunk)
        if not m_no or not m_date:
            continue
        no = int(m_no.group(1))
        date_str = f"{m_date.group(1)} {m_date.group(2)}"  # 2026/08/11 04:25:06
        try:
            dt = datetime.strptime(date_str, "%Y/%m/%d %H:%M:%S").replace(tzinfo=JST)
        except ValueError:
            continue

        img_url = None
        img_filename = None
        m_img = RE_IMG.search(chunk)
        if m_img:
            rel = m_img.group(1).lstrip("./")  # img-box/imgXXXX.jpg
            img_url = ROOT_URL + rel
            img_filename = rel.split("/")[-1]

        m_imgtitle = RE_IMG_TITLE.search(chunk)
        img_title = html_lib.unescape(m_imgtitle.group(1).strip()) if m_imgtitle else None

        m_title = RE_TITLE.search(chunk)
        title = strip_tags(m_title.group(1)) if m_title else ""

        m_body = RE_BODY.search(chunk)
        body = strip_tags(m_body.group(1)) if m_body else ""

        articles.append(
            {
                "no": no,
                "datetime": dt.isoformat(),
                "date": dt.strftime("%Y-%m-%d"),
                "title": title,
                "body": body,
                "image_filename": img_filename,
                "image_url": img_url,
                "image_title": img_title,
            }
        )
    return articles


def download_image(img_url: str, filename: str) -> bool:
    """画像を img-box/ に保存。既存ならスキップ。成功時 True。"""
    dest = IMG_DIR / filename
    if dest.exists() and dest.stat().st_size > 0:
        return False
    try:
        resp = session.get(img_url, timeout=60)
        resp.raise_for_status()
        IMG_DIR.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(resp.content)
        log(f"  画像取得: {filename} ({len(resp.content)//1024} KB)")
        return True
    except requests.RequestException as e:
        log(f"  画像取得失敗: {filename} ({e})")
        return False


def load_index() -> dict:
    if POSTS_JSON.exists():
        try:
            return json.loads(POSTS_JSON.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            log("posts.json が壊れています。空で再作成します。")
    return {}


def save_index(index: dict) -> None:
    # No. の降順(新しい順)で保存
    ordered = dict(sorted(index.items(), key=lambda kv: int(kv[0]), reverse=True))
    POSTS_JSON.write_text(
        json.dumps(ordered, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def write_daily_markdown(new_posts: list[dict]) -> None:
    """新規投稿を投稿日ごとの Markdown に追記する。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    by_date: dict[str, list[dict]] = {}
    for p in new_posts:
        by_date.setdefault(p["date"], []).append(p)

    for date, posts in by_date.items():
        md_path = DATA_DIR / f"{date}.md"
        existing = md_path.read_text(encoding="utf-8") if md_path.exists() else f"# {date} の投稿\n"
        already = set(re.findall(r"No\.(\d+)", existing))
        blocks = [existing.rstrip() + "\n"]
        for p in sorted(posts, key=lambda x: x["no"]):
            if str(p["no"]) in already:
                continue
            t = datetime.fromisoformat(p["datetime"]).strftime("%H:%M:%S")
            blocks.append(f"\n## No.{p['no']}  {t}")
            if p["title"]:
                blocks.append(f"**{p['title']}**\n")
            if p["image_filename"]:
                blocks.append(f"![{p['image_filename']}](../img-box/{p['image_filename']})\n")
            if p["body"]:
                blocks.append(p["body"] + "\n")
        md_path.write_text("\n".join(blocks).rstrip() + "\n", encoding="utf-8")


def main() -> int:
    log(f"取得開始: {BASE_URL} (最大 {MAX_PAGES} ページ)")
    index = load_index()
    seen_nos = set(index.keys())

    all_articles: list[dict] = []
    for page in range(1, MAX_PAGES + 1):
        try:
            page_html = fetch_page(page)
        except requests.RequestException as e:
            log(f"ページ {page} 取得失敗: {e}")
            break
        arts = parse_articles(page_html)
        log(f"ページ {page}: {len(arts)} 件の記事を検出")
        if not arts:
            break
        all_articles.extend(arts)
        time.sleep(1.5)  # サーバ負荷に配慮

    # No. で重複排除(ページ跨ぎ)
    unique: dict[int, dict] = {}
    for a in all_articles:
        unique[a["no"]] = a

    new_posts: list[dict] = []
    for no, art in unique.items():
        if str(no) in seen_nos:
            continue
        if art["image_url"] and art["image_filename"]:
            download_image(art["image_url"], art["image_filename"])
        index[str(no)] = art
        new_posts.append(art)

    if new_posts:
        save_index(index)
        write_daily_markdown(new_posts)
        log(f"新規投稿 {len(new_posts)} 件を追加しました。")
        for p in sorted(new_posts, key=lambda x: x["no"]):
            log(f"  + No.{p['no']} [{p['date']}] {p['title'][:30]}")
    else:
        log("新規投稿はありませんでした。")

    log(f"完了。総投稿数(累計): {len(index)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
