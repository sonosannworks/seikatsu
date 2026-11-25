/**
 * サポートナビ：「通いの場」タブのデータを返す関数（追加分）
 *
 * 【使い方】
 * 1. このファイルの中身を、スプレッドシートの Apps Script に「新しいファイル」として貼り付けます
 *    （既存の doGet などが入ったファイルはそのまま）。
 * 2. 既存の doGet の中（result を組み立てている箇所）に、次の3行を追加します。
 *    既存の処理はそのままです。
 *
 *      if (action === 'all' || action === 'kayoinoba') {
 *        result.kayoinoba = getKayoinoba_();
 *      }
 *
 *    ※ SPREADSHEET_ID は既存のコードで定義している定数をそのまま使います。
 *
 * 3. 「デプロイ」→「デプロイを管理」→ 既存のデプロイを編集 →
 *    バージョン「新バージョン」で更新します（URLは変わりません）。
 *
 * 見出し行（1行目）の列名で読み取るので、列の並び替えや列の追加をしても動きます。
 * 名称が空の行は返しません。
 */
function getKayoinoba_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('通いの場');
  if (!sheet) return [];

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const header = values[0].map(function (h) { return String(h).trim(); });
  const col = function (name) { return header.indexOf(name); };
  const idx = {
    no: col('No'),
    school: col('校区'),
    area: col('地区'),
    name: col('名称'),
    kind: col('区分'),
    place: col('場所'),
    url: col('詳細URL'),
    contact: col('問い合わせ先'),
  };
  const get = function (row, key) {
    return idx[key] === -1 ? '' : String(row[idx[key]] || '').trim();
  };

  return values.slice(1)
    .filter(function (row) { return get(row, 'name') !== ''; })
    .map(function (row) {
      return {
        no: get(row, 'no'),
        school: get(row, 'school'),
        area: get(row, 'area'),
        name: get(row, 'name'),
        kind: get(row, 'kind'),
        place: get(row, 'place'),
        url: get(row, 'url'),
        contact: get(row, 'contact'),
      };
    });
}
