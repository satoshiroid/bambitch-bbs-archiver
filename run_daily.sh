#!/bin/bash
# BBS アーカイバ 毎日実行ラッパー(launchd から呼ばれる)。
# 取得 → bbs_log.xlsx / img-box 更新 → GitHub へ push。ログは logs/ に残す。
set -uo pipefail

REPO_DIR="$HOME/bambitch-bbs-archiver"
cd "$REPO_DIR" || exit 1

mkdir -p logs
LOG="logs/run_$(date +%Y%m%d).log"

{
  echo "==================== $(date '+%Y-%m-%d %H:%M:%S') 実行開始 ===================="

  # 最新化(リモートに変更があれば取り込む)
  git pull --rebase --autostash origin main 2>&1 || echo "git pull スキップ(オフライン等)"

  # 取得(直近3ページ)
  BBS_MAX_PAGES="${BBS_MAX_PAGES:-3}" ./.venv/bin/python bambitch_bbs.py
  STATUS=$?

  # 差分があればコミット & push
  git add -A
  if git diff --cached --quiet; then
    echo "差分なし: コミットなし"
  else
    git -c user.name="satoshiroid" -c user.email="satoshi.yasui1123@gmail.com" \
        commit -m "chore: BBS取得 $(date '+%Y-%m-%d %H:%M')" 2>&1
    if git push origin main 2>&1; then
      echo "push 成功"
    else
      echo "push 失敗(次回再試行。ローカルには保存済み)"
    fi
  fi

  echo "==================== $(date '+%Y-%m-%d %H:%M:%S') 実行終了 (exit=$STATUS) ===================="
  echo ""
} >> "$LOG" 2>&1
