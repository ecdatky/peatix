#!/usr/bin/env python3
"""Peatix「RIK SQUARE! 2026」の参加者リストCSVをダウンロードする。
使い方: python3 fetch_peatix_csv.py <出力先ディレクトリ>
環境変数 PEATIX_USER / PEATIX_PASSWORD が必要。"""
import os, sys
from playwright.sync_api import sync_playwright

EVENT = "5227975"
out_dir = sys.argv[1]
os.makedirs(out_dir, exist_ok=True)
with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/opt/pw-browsers/chromium")
    pg = b.new_page()
    pg.goto("https://peatix.com/signin", wait_until="domcontentloaded")
    pg.wait_for_timeout(3000)
    pg.fill("input[name=username]", os.environ["PEATIX_USER"])
    pg.click('button:has-text("Next")'); pg.wait_for_timeout(3000)
    pg.fill("input[type=password]", os.environ["PEATIX_PASSWORD"])
    pg.keyboard.press("Enter"); pg.wait_for_timeout(6000)
    pg.goto(f"https://peatix.com/event/{EVENT}/list_sales", wait_until="domcontentloaded")
    pg.wait_for_timeout(4000)
    with pg.expect_download(timeout=60000) as d:
        pg.click("text=参加者リスト(CSV)")
    path = os.path.join(out_dir, "participants.csv")
    d.value.save_as(path)
    print(path)
