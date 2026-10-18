#!/usr/bin/env python3
"""【参加者リスト】RIK SQUARE2026.xlsx を作り直す(列構成や書式を変えたときに使う)。
使い方: python3 build_master_template.py <fetch_peatix_csv.py の出力ディレクトリ> <出力xlsx>
シート「参加申込リスト」の3行目からが Excel テーブル「参加申込」。B1 は最終更新日時。
以降は Power Automate から Office スクリプト update_master_list.ts を呼び、「チケット番号」をキーに
新しい人は行追加、既存の人は Peatix 由来の列(ステータス・懇親会)だけ更新し、B1 を書き換える。
「備考欄」〜「情報・商談メモ」の手入力列は自動更新では触らない。"""
import json, os, sys
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

# 列の並び(ご参加数・ステータスは集計用で非表示)
cols = ["チケット番号", "担当営業所", "会社名", "申込ご担当者", "部署・役職", "懇親会", "メールアドレス", "電話番号", "ヨミ",
        "郵便番号", "都道府県", "住所", "業種", "業種（その他）", "備考欄", "新規", "ユーザー", "取引先種類",
        "情報・商談メモ ※複数申込の場合は話をした方の名前も記載", "申込日時", "ご参加数", "ステータス"]
MANUAL = {"備考欄", "新規", "ユーザー", "取引先種類", "情報・商談メモ ※複数申込の場合は話をした方の名前も記載"}
HIDDEN = {"ご参加数", "ステータス"}
WIDTH = {"チケット番号": 12, "担当営業所": 11, "会社名": 30, "申込ご担当者": 14, "部署・役職": 20, "都道府県": 10,
         "メールアドレス": 30, "電話番号": 15, "ヨミ": 16, "郵便番号": 10, "住所": 36, "ご参加数": 9, "懇親会": 11, "業種": 22,
         "業種（その他）": 20, "ステータス": 11, "備考欄": 24, "新規": 9, "ユーザー": 9, "取引先種類": 12,
         "情報・商談メモ ※複数申込の場合は話をした方の名前も記載": 40,
         "申込日時": 17}
L = {h: chr(65 + i) for i, h in enumerate(cols)}  # 列名 → 列記号(22列なので1文字で足りる)
FIRST, LAST = 4, 2001  # データ行の範囲(条件付き書式・集計は行追加で伸びても届くよう広めに取る)

rows = json.load(open(os.path.join(sys.argv[1], "master_rows.json"), encoding="utf-8"))
n = max(len(rows), 1)  # テーブルは最低1行必要
end = FIRST - 1 + n
now = datetime.now(ZoneInfo("Asia/Tokyo"))


def font(sz=11, b=False, color=None):
    return Font(name="游ゴシック", size=sz, bold=b, color=color)


wb = Workbook()
ws = wb.active
ws.title = "参加申込リスト"
ws["A1"], ws["B1"] = "最終更新日時", f"{now.year}年{now.month}月{now.day}日（{'月火水木金土日'[now.weekday()]}）　{now:%H:%M}"
ws["A1"].font = ws["B1"].font = font()
ws["A2"] = "RIK SQUARE!2026 お申込状況"
ws["A2"].font = font(20, True)
# 集計(キャンセルと社内=ユニマットリックの申込は数えない)
st, kon, co = (f"{L[h]}{FIRST}:{L[h]}{LAST}" for h in ("ステータス", "懇親会", "会社名"))
rik = f'{co},"<>*ユニマットリック*"'
red = "FFFF0000"
for addr, v, f, al in (
        ("E2", "申込者数\n※社員除く", font(16, True, red), "right"),
        ("F2", f'=COUNTIFS({st},"<>",{st},"<>*キャンセル*",{rik})', font(36, True, red), "left"),
        ("G2", "懇親会", font(16, True, red), "right"),
        ("H2", f'=COUNTIFS({kon},"参加する",{st},"<>*キャンセル*",{rik})', font(24, True, red), "left"),
        ("J2", "キャンセル", font(11, True), None),
        ("K2", f'=COUNTIFS({st},"*キャンセル*",{rik})', font(11, True), None)):
    ws[addr] = v
    ws[addr].font = f
    ws[addr].alignment = Alignment(horizontal=al, vertical="center", wrap_text=addr == "E2")
ws.row_dimensions[2].height = 69.75

thin = Side(style="thin", color="BFBFBF")
grid = Border(left=thin, right=thin, top=thin, bottom=thin)
manual_fill = PatternFill("solid", fgColor="FFF2CC")  # 手入力列の見出しは黄色で区別
for j, h in enumerate(cols, 1):
    c = ws.cell(3, j, h)
    c.font = font(11, h in MANUAL, "FF000000" if h in MANUAL else None)
    c.alignment = Alignment(vertical="center", wrap_text=True)
    c.border = grid
    if h in MANUAL:
        c.fill = manual_fill
    ws.column_dimensions[L[h]].width = WIDTH.get(h, 9)
    ws.column_dimensions[L[h]].hidden = h in HIDDEN
ws.row_dimensions[3].height = 32.1
for i in range(FIRST, end + 1):
    r = rows[i - FIRST] if rows else {}
    ws.row_dimensions[i].height = 18.75
    for j, h in enumerate(cols, 1):
        c = ws.cell(i, j, r.get(h, "") or None)
        c.font = font(11, h == "会社名")
        c.alignment = Alignment(vertical="center")
        c.number_format = "@"  # 日時・番号を文字列のまま保つ
        c.border = grid
ws.freeze_panes = f"D{FIRST}"

# 条件付き書式(上ほど優先)
# 薄い罫線(Power Automate で追加される行にも付くよう条件付き書式でも引く) / 社内(ユニマットリック)の行は全体グレー
# 会社名: 2行以上ある会社はオレンジ、「個人」は黄色 / メールアドレス: 2行以上あれば赤 / 懇親会: 「参加する」は緑
area = f"A{FIRST}:{L[cols[-1]]}{LAST}"
c, m, k = L["会社名"], L["メールアドレス"], L["懇親会"]
ws.conditional_formatting.add(area, FormulaRule(formula=[f'${L["チケット番号"]}{FIRST}<>""'], border=grid))
ws.conditional_formatting.add(area, FormulaRule(
    formula=[f'ISNUMBER(SEARCH("ユニマットリック",${c}{FIRST}))'], fill=PatternFill("solid", bgColor="D9D9D9"), stopIfTrue=True))
ws.conditional_formatting.add(f"{c}{FIRST}:{c}{LAST}", FormulaRule(
    formula=[f'AND(${c}{FIRST}<>"",${c}{FIRST}<>"個人",COUNTIF(${c}${FIRST}:${c}${LAST},${c}{FIRST})>1)'],
    fill=PatternFill("solid", bgColor="FCE4D6")))
ws.conditional_formatting.add(f"{c}{FIRST}:{c}{LAST}", FormulaRule(
    formula=[f'${c}{FIRST}="個人"'], fill=PatternFill("solid", bgColor="FFF2CC")))
ws.conditional_formatting.add(f"{m}{FIRST}:{m}{LAST}", FormulaRule(
    formula=[f'AND(${m}{FIRST}<>"",COUNTIF(${m}${FIRST}:${m}${LAST},${m}{FIRST})>1)'],
    fill=PatternFill("solid", bgColor="FFC7CE")))
ws.conditional_formatting.add(f"{k}{FIRST}:{k}{LAST}", FormulaRule(
    formula=[f'${k}{FIRST}="参加する"'], fill=PatternFill("solid", bgColor="E2EFDA")))

t = Table(displayName="参加申込", ref=f"A3:{L[cols[-1]]}{end}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False)  # 交互の色付けはしない
ws.add_table(t)
wb.save(sys.argv[2])
print(sys.argv[2], f"{len(rows)} rows")
