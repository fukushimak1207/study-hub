# -*- coding: utf-8 -*-
"""
学びハブへ PC からデータを送る道具。

使い方:
  python hub.py setup                    … 最初に1回。URL と合言葉を聞いて設定ファイルに保存する
  python hub.py check                    … つながるか確かめる(件数だけ表示)
  python hub.py seminars <勉強会.json>   … 勉強会を追記する(同じ会は空欄だけ埋める)
  python hub.py papers <latest.json> [IssueのURL] … 論文候補を追記する(論文ウォッチャーの形式)

設定ファイル: ~/.config/study-hub/config.json(GitHub には上げない場所)
環境変数 STUDYHUB_URL / STUDYHUB_TOKEN があればそちらを優先する(GitHub Actions 用)。
"""

import os
import sys
import json
import re
import getpass
import subprocess
import urllib.request
import urllib.parse

CONF = os.path.join(os.path.expanduser("~"), ".config", "study-hub", "config.json")


def load_conf():
    url = os.environ.get("STUDYHUB_URL", "").strip()
    token = os.environ.get("STUDYHUB_TOKEN", "").strip()
    if url and token:
        return url, token
    if not os.path.exists(CONF):
        sys.exit("[エラー] 設定がありません。先に `python hub.py setup` を実行してください。")
    with open(CONF, encoding="utf-8") as f:
        c = json.load(f)
    return c["url"].strip(), c["token"].strip()


def call(payload=None):
    url, token = load_conf()
    if payload is None:
        req = urllib.request.Request(url + "?" + urllib.parse.urlencode({"token": token}))
    else:
        payload = dict(payload, token=token)
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        # text/plain にしておくと Apps Script がそのまま受け取れる
        req = urllib.request.Request(url, data=data,
                                     headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=60) as r:
        res = json.loads(r.read().decode("utf-8"))
    if not res.get("ok"):
        sys.exit("[エラー] " + res.get("error", "不明なエラー"))
    return res


def read_clipboard():
    """クリップボードの中身を読んで、そのあと空にする(合言葉を画面に出さないため)"""
    ps = ["powershell", "-NoProfile", "-Command"]
    out = subprocess.run(ps + ["[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-Clipboard -Raw"],
                         capture_output=True, text=True, encoding="utf-8").stdout
    subprocess.run(ps + ["Set-Clipboard -Value ' '"], capture_output=True)
    return (out or "").strip()


def cmd_setup():
    url = input("ウェブアプリの URL(https://script.google.com/macros/s/.../exec): ").strip()
    if not url.startswith("https://script.google.com/"):
        sys.exit("[エラー] URL は https://script.google.com/ で始まるものを貼ってください")
    if os.name == "nt":
        input("合言葉をコピーしてから Enter を押してください(画面には出ません)")
        token = read_clipboard()
        if not re.fullmatch(r"[0-9a-f]{32}", token):
            sys.exit("[エラー] クリップボードの中身が合言葉の形(英数字32文字)ではありません。"
                     "合言葉だけをコピーし直して、もう一度実行してください")
        print("合言葉を読み取りました(クリップボードは空にしました)")
    else:
        token = getpass.getpass("合言葉(入力しても画面には出ません): ").strip()
    os.makedirs(os.path.dirname(CONF), exist_ok=True)
    with open(CONF, "w", encoding="utf-8") as f:
        json.dump({"url": url, "token": token}, f)
    print("保存しました: " + CONF)
    cmd_check()


def cmd_check():
    res = call()
    print("つながりました。勉強会 %d 件 / 論文 %d 件" % (len(res["seminars"]), len(res["papers"])))


def cmd_seminars(path):
    with open(path, encoding="utf-8") as f:
        items = json.load(f)
    res = call({"action": "addSeminars", "items": items})
    print("勉強会: 新規 %d 件 / 空欄を補完 %d 件" % (res["added"], res["filled"]))


def authors_short(a):
    # 論文ウォッチャーの authors は [["姓", "イニシャル"], ...] の形
    au = a.get("authors") or []
    if isinstance(au, str):
        return au
    if not au:
        return ""
    first = " ".join(au[0]) if isinstance(au[0], (list, tuple)) else str(au[0])
    return first + (" et al." if len(au) > 1 else "")


def cmd_papers(path, issue_url=""):
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    items = [{
        "pmid": a["pmid"], "title": a.get("title", ""), "journal": a.get("journal", ""),
        "year": a.get("year", ""), "authors": authors_short(a),
        "theme": a.get("_theme", ""), "summary_ja": a.get("summary_ja", ""),
    } for a in d["items"]]
    res = call({"action": "addPapers", "week": d["date"], "issue_url": issue_url, "items": items})
    print("論文: 新規 %d 件(%s 週)" % (res["added"], d["date"]))


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    mode = sys.argv[1]
    if mode == "setup":
        cmd_setup()
    elif mode == "check":
        cmd_check()
    elif mode == "seminars" and len(sys.argv) >= 3:
        cmd_seminars(sys.argv[2])
    elif mode == "papers" and len(sys.argv) >= 3:
        cmd_papers(sys.argv[2], sys.argv[3] if len(sys.argv) >= 4 else "")
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
