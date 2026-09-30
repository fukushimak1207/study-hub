# -*- coding: utf-8 -*-
"""
スマホに URL と合言葉を渡すための QR コードを PC の画面に出す。

使い方: python qr_setup.py
  - hub.py setup で保存した設定(~/.config/study-hub/config.json)を使う
  - スマホのカメラで読むと学びハブが開き、設定が自動で入る
  - QR の画像は一時ファイルに作り、Enter を押すと消す
"""

import os
import sys
import json
import base64
import tempfile

import qrcode

APP = "https://fukushimak1207.github.io/study-hub/"
CONF = os.path.join(os.path.expanduser("~"), ".config", "study-hub", "config.json")


def main():
    if not os.path.exists(CONF):
        sys.exit("[エラー] 先に hub.py setup を実行してください")
    with open(CONF, encoding="utf-8") as f:
        c = json.load(f)
    raw = json.dumps({"url": c["url"], "token": c["token"]}).encode("utf-8")
    code = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    img = qrcode.make(APP + "#setup=" + code, box_size=8, border=3)

    path = os.path.join(tempfile.gettempdir(), "study-hub-setup-qr.png")
    img.save(path)
    os.startfile(path)
    input("スマホのカメラで QR コードを読んでください。終わったら Enter(画像を消します)")
    try:
        os.remove(path)
        print("QR コードの画像を消しました")
    except OSError:
        print("画像を消せませんでした。表示中のアプリを閉じてから消してください: " + path)


if __name__ == "__main__":
    main()
