#!/usr/bin/env python3
"""Peatix「RIK SQUARE! 2026」の参加者リストをダウンロードし、確認用シート付きxlsxに整形する。
使い方: python3 fetch_peatix_csv.py <出力先ディレクトリ>
環境変数 PEATIX_USER / PEATIX_PASSWORD が必要。
TEAMS_UPLOAD_URL(Power Automateの HTTP トリガーURL)があれば、xlsxと参加者数をPOSTして Teams に投稿する。"""
import base64, csv, io, json, os, re, sys, unicodedata, urllib.request, zipfile
from datetime import datetime
from zoneinfo import ZoneInfo
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

# Peatix標準はUTF-16LE・タブ区切り。
text = open(path, encoding="utf-16").read()
rows = [[c.strip() for c in r] for r in csv.reader(io.StringIO(text, newline=""), delimiter="\t")]
header, data = rows[0], [r for r in rows[1:] if any(r)]

now = datetime.now(ZoneInfo("Asia/Tokyo"))


def fmt_zip(v):
    d = re.sub(r"\D", "", unicodedata.normalize("NFKC", v))
    return f"{d[:3]}-{d[3:]}" if len(d) == 7 else v


# 市外局番が3桁の主な地域(それ以外の固定電話は4桁として扱う簡易判定。ハイフン済みの値は触らない)
AREA3 = {"011", "015", "017", "018", "019", "022", "023", "024", "025", "026", "027", "028", "029", "042", "043", "044",
         "045", "046", "047", "048", "049", "052", "053", "055", "058", "059", "072", "073", "075", "076", "077", "078",
         "079", "082", "083", "084", "086", "087", "088", "089", "092", "093", "095", "096", "097", "098", "099"}


def fmt_tel(v):
    if "-" in v or "－" in v or "ー" in v:
        return v
    d = re.sub(r"\D", "", unicodedata.normalize("NFKC", v))
    if len(d) == 11 and d[:3] in ("070", "080", "090", "050", "020"):
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    if len(d) == 10 and d.startswith("0120"):
        return f"{d[:4]}-{d[4:7]}-{d[7:]}"
    if len(d) == 10 and d[:2] in ("03", "06"):
        return f"{d[:2]}-{d[2:6]}-{d[6:]}"
    if len(d) == 10 and d[0] == "0":
        return f"{d[:3]}-{d[3:6]}-{d[6:]}" if d[:3] in AREA3 else f"{d[:4]}-{d[4:6]}-{d[6:]}"
    return v


def col(n):
    return chr(65 + n)


def sheet_xml(rows, header_style=None, widths=None, freeze=False, style_fn=None):
    """最小限のシートXML(インライン文字列・UTF-8)。header_style: 1行目に当てるスタイル番号。"""
    out = ['<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    if freeze:
        out.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    if widths:
        out.append("<cols>" + "".join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths)) + "</cols>")
    out.append("<sheetData>")
    for r, row in enumerate(rows, 1):
        out.append("<row>")
        for c, v in enumerate(row):
            sid = header_style if r == 1 and header_style else (style_fn(r, c, v) if style_fn else 0)
            st = f' s="{sid}"' if sid else ""
            if v == "":
                out.append(f"<c{st}/>")  # r属性を省いて順番で位置を決めるため、空セルも1つ書く
            elif isinstance(v, int):
                out.append(f"<c{st}><v>{v}</v></c>")
            else:
                t = v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                out.append(f'<c{st} t="inlineStr"><is><t>{t}</t></is></c>')
        out.append("</row>")
    out.append("</sheetData></worksheet>")
    return "".join(out)


def width(vals):
    return min(max(max((len(v) * (2 if ord(v[0]) > 255 else 1) for v in vals if v), default=0) + 2, 10), 50)


# 確認用シート: チケット番号以降の列だけ残し、郵便番号をハイフン入りに、見出しに色を付ける
start = header.index("チケット番号")
ci_zip = header.index("郵便番号") - start
ci_tel = header.index("電話番号") - start
ci_party = len(header) - 1 - start  # 懇親会の列(最終列)
assert header[-1].startswith("懇親会"), header[-1]
check = [header[start:]]
for r in data:
    row = r[start:]
    row[ci_zip] = fmt_zip(row[ci_zip])
    row[ci_tel] = fmt_tel(row[ci_tel])
    check.append(row)
widths = [width([r[i] for r in check]) for i in range(len(check[0]))]

# 元データ(CSVそのまま)シート
int_cols = {header.index(h) for h in ("注文番号", "会員 ID", "枚数")}
raw = [header] + [[int(v) if i in int_cols and v.isdigit() else v for i, v in enumerate(r)] for r in data]

# メール添付をbase64で受け渡す際、AIが文字列を転記しても壊れにくいよう、openpyxl等を使わず最小構成で書く
X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
NS = "http://schemas.openxmlformats.org/"
parts = {
    "[Content_Types].xml": X + f'<Types xmlns="{NS}package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    "_rels/.rels": X + f'<Relationships xmlns="{NS}package/2006/relationships"><Relationship Id="rId1" Type="{NS}officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": X + f'<workbook xmlns="{NS}spreadsheetml/2006/main" xmlns:r="{NS}officeDocument/2006/relationships"><sheets><sheet name="確認用" sheetId="1" r:id="rId1"/><sheet name="participants{now:%Y%m%d}" sheetId="2" r:id="rId2"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": X + f'<Relationships xmlns="{NS}package/2006/relationships"><Relationship Id="rId1" Type="{NS}officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{NS}officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{NS}officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    "xl/styles.xml": X + f'<styleSheet xmlns="{NS}spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color rgb="FF000000"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    "xl/worksheets/sheet1.xml": X + sheet_xml(check, header_style=1, widths=widths, freeze=True, style_fn=lambda r, c, v: 2 if c == ci_party and v == "参加する" else 0),
    "xl/worksheets/sheet2.xml": X + sheet_xml(raw),
}
out = os.path.join(out_dir, f"participants_{now:%Y%m%d_%H%M}.xlsx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for name, xml in parts.items():
        z.writestr(name, xml.encode("utf-8"))
print(out)

# 参加者数(キャンセルを除く)と、前日末(昨日23:59:59 JST)時点からの増減。申込日で判定するため前日分の保存は不要。
status_i, date_i = header.index("ステータス"), header.index("申込日")
active = [r for r in data if "キャンセル" not in r[status_i] and "cancel" not in r[status_i].lower()]
today = now.strftime("%Y-%m-%d")
added = sum(1 for r in active if r[date_i][:10] >= today)
print(f"COUNT total={len(active)} delta={added}")

# Power Automate(HTTP受信 → SharePoint保存 → Teamsカード投稿)へ送信
url = os.environ.get("TEAMS_UPLOAD_URL")
if url:
    body = json.dumps({"filename": os.path.basename(out), "contentBase64": base64.b64encode(open(out, "rb").read()).decode(),
                       "total": len(active), "delta": added}).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=90) as r:
        print(f"POST status={r.status}")
