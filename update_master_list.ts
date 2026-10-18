// Office スクリプト(Excel の「自動化」タブで作成・保存し、Power Automate の「スクリプトの実行」から呼ぶ)
// 【参加者リスト】RIK SQUARE2026.xlsx のテーブル「参加申込」を Peatix の最新データで更新する。
//   rowsJson: fetch_peatix_csv.py が送る参加者の配列(JSON文字列)。キーは表の見出し名と同じ。
//   updated : 最終更新日時の表示文字列。B1 に書き込む。
// 「チケット番号」が表にない人は末尾に追加し、ある人は「ステータス」「懇親会」だけ更新する。
// それ以外の列(備考欄・商談メモなどの手入力列)には触らない。
function main(workbook: ExcelScript.Workbook, rowsJson: string, updated: string) {
  const table = workbook.getTable("参加申込");
  const headers = table.getHeaderRowRange().getValues()[0].map(String);
  const key = headers.indexOf("チケット番号");
  const body = table.getRangeBetweenHeaderAndTotal();
  const values = body.getValues();

  const pos = new Map<string, number>();
  values.forEach((r, i) => { if (String(r[key]) !== "") pos.set(String(r[key]), i); });

  const rows = JSON.parse(rowsJson) as { [h: string]: string }[];
  const add: string[][] = [];
  let changed = 0;
  for (const r of rows) {
    const i = pos.get(r["チケット番号"]);
    if (i === undefined) {
      add.push(headers.map(h => r[h] ?? ""));
      continue;
    }
    for (const h of ["ステータス", "懇親会"]) {
      const j = headers.indexOf(h);
      if (j >= 0 && r[h] !== undefined && String(values[i][j]) !== r[h]) {
        body.getCell(i, j).setValue(r[h]);
        changed++;
      }
    }
  }
  if (add.length > 0) {
    // 表が空行1行だけの場合はその行を使う
    if (values.length === 1 && values[0].every(v => String(v) === "")) {
      body.setValues([add.shift() as string[]]);
    }
    if (add.length > 0) table.addRows(-1, add);
  }
  table.getWorksheet().getRange("B1").setValue(updated);
  console.log(`added=${add.length} changed=${changed}`);
}
