// Генератор книги uchet_prodazh.xlsx по PRD. Запуск: node build.js
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const fs = require('fs');

const OUT = 'uchet_prodazh.xlsx';
const FONT = 'Arial';
const CUR = '#,##0 "₸"';
const DATE = 'dd.mm.yyyy';

const wb = new ExcelJS.Workbook();
wb.calcProperties.fullCalcOnLoad = true;

const base = { name: FONT, size: 10 };
const inputFont = { ...base, color: { argb: 'FF0000FF' } };
const inputFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const headFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
const headFont = { ...base, bold: true, color: { argb: 'FFFFFFFF' } };
const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const border = { top: thin, bottom: thin, left: thin, right: thin };

function header(ws, row, titles, widths) {
  titles.forEach((t, i) => {
    const c = ws.getCell(row, i + 1);
    c.value = t; c.font = headFont; c.fill = headFill; c.border = border;
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    if (widths) ws.getColumn(i + 1).width = widths[i];
  });
  ws.getRow(row).height = 30;
}
function input(c, fmt) { c.font = inputFont; c.fill = inputFill; c.border = border; if (fmt) c.numFmt = fmt; }
function calc(c, fmt) { c.font = base; c.border = border; if (fmt) c.numFmt = fmt; }

const wsInfo = wb.addWorksheet('Инструкция');
const wsS = wb.addWorksheet('Продажи');
const wsT = wb.addWorksheet('Товары');
const wsP = wb.addWorksheet('Приход');
const wsC = wb.addWorksheet('Клиенты');
const wsR = wb.addWorksheet('Отчет');
const wsI = wb.addWorksheet('Счет');

// ---------- Инструкция ----------
wsInfo.getColumn(1).width = 110;
[
  ['Учёт продаж: как работать с файлом', true],
  ['', false],
  ['Жёлтые ячейки с синим текстом вводит пользователь. Остальные ячейки содержат формулы — их не нужно менять.', false],
  ['', false],
  ['1. Товары — заполните справочник: артикул, название, цена продажи, закупочная цена, начальный и минимальный остаток.', false],
  ['2. Клиенты — заполните название, телефон, email и реквизиты клиента.', false],
  ['3. Приход — вносите поступления товара (дата, артикул из списка, количество, поставщик). Остатки обновятся сами.', false],
  ['4. Продажи — вносите: дата, № счёта, клиент и артикул (из выпадающих списков), количество. Остальное считается автоматически.', false],
  ['5. Отчет — выберите год в ячейке B2: помесячная выручка, прибыль, маржа, график, рейтинги товаров и клиентов.', false],
  ['6. Счет — введите номер счёта в B3, реквизиты поставщика в B7 (один раз). Лист готов к печати (A4) или сохранению в PDF.', false],
  ['', false],
  ['Допущения и ограничения', true],
  ['• Валюта — тенге (₸); меняется форматом ячеек.', false],
  ['• Цена берётся из справочника «Товары». При смене цены пересчитываются и прошлые продажи. Чтобы зафиксировать цену, перезапишите её вручную в строке продаж.', false],
  ['• Один счёт = несколько строк журнала с одинаковым № счёта, максимум 20 позиций на счёт.', false],
  ['• Лимиты формул: 999 строк продаж, 999 строк прихода, 199 товаров, 199 клиентов. В отчёте показываются первые 30 товаров и клиентов.', false],
  ['• Если данных больше — протяните формулы вниз и расширьте диапазоны.', false],
  ['• Демонстрационные данные нужно удалить из жёлтых ячеек и внести реальные.', false],
].forEach(([t, b], i) => {
  const c = wsInfo.getCell(i + 1, 1);
  c.value = t; c.font = { ...base, bold: b, size: b ? 12 : 10 };
  c.alignment = { wrapText: true, vertical: 'top' };
});

// ---------- Товары ----------
header(wsT, 1, ['Артикул', 'Название', 'Цена продажи', 'Закупочная цена', 'Начальный остаток', 'Мин. остаток', 'Приход', 'Продано', 'Остаток', 'Статус'],
  [12, 30, 14, 14, 12, 12, 10, 10, 10, 12]);
const goods = [
  ['A-001', 'Товар 1', 15000, 9000, 50, 10],
  ['B-002', 'Товар 2', 8500, 5000, 30, 10],
  ['C-003', 'Товар 3', 25000, 15000, 20, 5],
  ['D-004', 'Товар 4', 4000, 2500, 100, 20],
];
for (let r = 2; r <= 200; r++) {
  const g = goods[r - 2] || [];
  for (let c = 1; c <= 6; c++) {
    const cell = wsT.getCell(r, c);
    cell.value = g[c - 1] ?? null;
    input(cell, c === 3 || c === 4 ? CUR : null);
  }
  const f = {
    7: `IF(A${r}="","",SUMIFS(Приход!$C$2:$C$1000,Приход!$B$2:$B$1000,A${r}))`,
    8: `IF(A${r}="","",SUMIFS(Продажи!$E$2:$E$1000,Продажи!$D$2:$D$1000,A${r}))`,
    9: `IF(A${r}="","",E${r}+G${r}-H${r})`,
    10: `IF(A${r}="","",IF(I${r}<=F${r},"Заказать","ОК"))`,
  };
  for (const c in f) { const cell = wsT.getCell(r, +c); cell.value = { formula: f[c] }; calc(cell); }
}
wsT.views = [{ state: 'frozen', ySplit: 1 }];
wsT.addConditionalFormatting({
  ref: 'J2:J200',
  rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"Заказать"'], priority: 1,
    style: { font: { bold: true, color: { argb: 'FFC00000' } }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }],
});

// ---------- Клиенты ----------
header(wsC, 1, ['Клиент', 'Телефон', 'Email', 'Реквизиты / адрес', 'Сумма покупок', 'Последняя покупка'], [28, 18, 26, 45, 16, 16]);
const clients = [
  ['ТОО «Альфа»', '+7 701 111 22 33', 'info@alfa.kz', 'БИН 123456789012, г. Астана, пр. Республики 1'],
  ['ИП Сергеев', '+7 702 222 33 44', 'sergeev@mail.kz', 'ИИН 800101300123, г. Астана, ул. Кенесары 10'],
  ['ТОО «Бета»', '+7 707 333 44 55', 'office@beta.kz', 'БИН 210987654321, г. Алматы, ул. Абая 5'],
];
for (let r = 2; r <= 200; r++) {
  const cl = clients[r - 2] || [];
  for (let c = 1; c <= 4; c++) { const cell = wsC.getCell(r, c); cell.value = cl[c - 1] ?? null; input(cell); }
  wsC.getCell(r, 5).value = { formula: `IF(A${r}="","",SUMIFS(Продажи!$H$2:$H$1000,Продажи!$C$2:$C$1000,A${r}))` };
  calc(wsC.getCell(r, 5), CUR);
  wsC.getCell(r, 6).value = { formula: `IF(A${r}="","",IF(COUNTIF(Продажи!$C$2:$C$1000,A${r})=0,"",SUMPRODUCT(MAX((Продажи!$C$2:$C$1000=A${r})*Продажи!$A$2:$A$1000))))` };
  calc(wsC.getCell(r, 6), DATE);
}
wsC.views = [{ state: 'frozen', ySplit: 1 }];

// ---------- Приход ----------
header(wsP, 1, ['Дата', 'Артикул', 'Количество', 'Поставщик / комментарий'], [14, 14, 12, 40]);
const arrivals = [
  [new Date(2026, 0, 10), 'A-001', 20, 'Поставщик 1'],
  [new Date(2026, 1, 5), 'C-003', 10, 'Поставщик 2'],
  [new Date(2026, 2, 12), 'D-004', 50, 'Поставщик 1'],
];
for (let r = 2; r <= 1000; r++) {
  const a = arrivals[r - 2] || [];
  for (let c = 1; c <= 4; c++) { const cell = wsP.getCell(r, c); cell.value = a[c - 1] ?? null; input(cell, c === 1 ? DATE : null); }
  wsP.getCell(r, 2).dataValidation = { type: 'list', allowBlank: true, formulae: ['Товары!$A$2:$A$200'] };
}
wsP.views = [{ state: 'frozen', ySplit: 1 }];

// ---------- Продажи ----------
header(wsS, 1, ['Дата', '№ счёта', 'Клиент', 'Артикул', 'Кол-во', 'Название', 'Цена', 'Сумма', 'Себестоимость', 'Прибыль', 'Месяц', 'Год', 'Ключ строки'],
  [13, 10, 24, 12, 9, 28, 13, 14, 14, 14, 8, 8, 14]);
const sales = [
  [new Date(2026, 0, 15), 1001, 'ТОО «Альфа»', 'A-001', 10],
  [new Date(2026, 0, 15), 1001, 'ТОО «Альфа»', 'B-002', 5],
  [new Date(2026, 1, 3), 1002, 'ИП Сергеев', 'A-001', 10],
  [new Date(2026, 1, 20), 1003, 'ТОО «Бета»', 'C-003', 8],
  [new Date(2026, 2, 14), 1004, 'ТОО «Альфа»', 'C-003', 6],
  [new Date(2026, 3, 2), 1005, 'ИП Сергеев', 'D-004', 40],
  [new Date(2026, 4, 18), 1006, 'ТОО «Бета»', 'B-002', 12],
  [new Date(2026, 5, 9), 1007, 'ТОО «Альфа»', 'D-004', 25],
];
const lk = (col, r) => `INDEX(Товары!$${col}$2:$${col}$200,MATCH(D${r},Товары!$A$2:$A$200,0))`;
for (let r = 2; r <= 1000; r++) {
  const s = sales[r - 2] || [];
  for (let c = 1; c <= 5; c++) { const cell = wsS.getCell(r, c); cell.value = s[c - 1] ?? null; input(cell, c === 1 ? DATE : null); }
  wsS.getCell(r, 3).dataValidation = { type: 'list', allowBlank: true, formulae: ['Клиенты!$A$2:$A$200'] };
  wsS.getCell(r, 4).dataValidation = { type: 'list', allowBlank: true, formulae: ['Товары!$A$2:$A$200'] };
  const f = {
    6: [`IF(D${r}="","",IFERROR(${lk('B', r)},"нет в справочнике"))`],
    7: [`IF(D${r}="","",IFERROR(${lk('C', r)},0))`, CUR],
    8: [`IF(D${r}="","",E${r}*G${r})`, CUR],
    9: [`IF(D${r}="","",E${r}*IFERROR(${lk('D', r)},0))`, CUR],
    10: [`IF(D${r}="","",H${r}-I${r})`, CUR],
    11: [`IF(A${r}="","",MONTH(A${r}))`],
    12: [`IF(A${r}="","",YEAR(A${r}))`],
    13: [`IF(OR(B${r}="",D${r}=""),"",B${r}&"|"&COUNTIF($B$2:B${r},B${r}))`],
  };
  for (const c in f) { const cell = wsS.getCell(r, +c); cell.value = { formula: f[c][0] }; calc(cell, f[c][1]); }
}
wsS.views = [{ state: 'frozen', ySplit: 1 }];
wsS.autoFilter = 'A1:M1000';

// ---------- Отчет ----------
wsR.getColumn(1).width = 24;
[16, 16, 14, 12, 4].forEach((w, i) => (wsR.getColumn(i + 2).width = w));
wsR.getCell('A1').value = 'Отчёт по продажам'; wsR.getCell('A1').font = { ...base, bold: true, size: 14 };
wsR.getCell('A2').value = 'Год'; wsR.getCell('A2').font = { ...base, bold: true };
wsR.getCell('B2').value = 2026; input(wsR.getCell('B2'), '0');
header(wsR, 4, ['Месяц', 'Выручка', 'Прибыль', 'Позиций', 'Маржа']);
const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
months.forEach((m, i) => {
  const r = 5 + i, n = i + 1;
  wsR.getCell(r, 1).value = m; calc(wsR.getCell(r, 1));
  wsR.getCell(r, 2).value = { formula: `SUMIFS(Продажи!$H$2:$H$1000,Продажи!$L$2:$L$1000,$B$2,Продажи!$K$2:$K$1000,${n})` }; calc(wsR.getCell(r, 2), CUR);
  wsR.getCell(r, 3).value = { formula: `SUMIFS(Продажи!$J$2:$J$1000,Продажи!$L$2:$L$1000,$B$2,Продажи!$K$2:$K$1000,${n})` }; calc(wsR.getCell(r, 3), CUR);
  wsR.getCell(r, 4).value = { formula: `COUNTIFS(Продажи!$L$2:$L$1000,$B$2,Продажи!$K$2:$K$1000,${n})` }; calc(wsR.getCell(r, 4), '0');
  wsR.getCell(r, 5).value = { formula: `IF(B${r}=0,0,C${r}/B${r})` }; calc(wsR.getCell(r, 5), '0.0%');
});
wsR.getCell('A17').value = 'Итого';
wsR.getCell('B17').value = { formula: 'SUM(B5:B16)' };
wsR.getCell('C17').value = { formula: 'SUM(C5:C16)' };
wsR.getCell('D17').value = { formula: 'SUM(D5:D16)' };
wsR.getCell('E17').value = { formula: 'IF(B17=0,0,C17/B17)' };
['A', 'B', 'C', 'D', 'E'].forEach((c) => { const cell = wsR.getCell(c + '17'); calc(cell); cell.font = { ...base, bold: true }; });
wsR.getCell('B17').numFmt = CUR; wsR.getCell('C17').numFmt = CUR; wsR.getCell('D17').numFmt = '0'; wsR.getCell('E17').numFmt = '0.0%';

wsR.getCell('A19').value = 'По товарам (за выбранный год)'; wsR.getCell('A19').font = { ...base, bold: true, size: 12 };
header(wsR, 20, ['Артикул', 'Название', 'Продано, шт', 'Выручка', 'Прибыль', 'Место']);
wsR.getColumn(2).width = 22; wsR.getColumn(5).width = 16; wsR.getColumn(6).width = 10;
for (let i = 0; i < 30; i++) {
  const r = 21 + i, t = 2 + i;
  const f = [
    `IF(Товары!A${t}="","",Товары!A${t})`,
    `IF(A${r}="","",Товары!B${t})`,
    `IF(A${r}="","",SUMIFS(Продажи!$E$2:$E$1000,Продажи!$D$2:$D$1000,A${r},Продажи!$L$2:$L$1000,$B$2))`,
    `IF(A${r}="","",SUMIFS(Продажи!$H$2:$H$1000,Продажи!$D$2:$D$1000,A${r},Продажи!$L$2:$L$1000,$B$2))`,
    `IF(A${r}="","",SUMIFS(Продажи!$J$2:$J$1000,Продажи!$D$2:$D$1000,A${r},Продажи!$L$2:$L$1000,$B$2))`,
    `IF(A${r}="","",RANK(D${r},$D$21:$D$50))`,
  ];
  f.forEach((x, j) => { const cell = wsR.getCell(r, j + 1); cell.value = { formula: x }; calc(cell, j === 3 || j === 4 ? CUR : j === 2 ? '0' : null); });
}
wsR.getCell('A53').value = 'По клиентам (за выбранный год)'; wsR.getCell('A53').font = { ...base, bold: true, size: 12 };
header(wsR, 54, ['Клиент', 'Позиций', 'Выручка', 'Место']);
for (let i = 0; i < 30; i++) {
  const r = 55 + i, t = 2 + i;
  const f = [
    `IF(Клиенты!A${t}="","",Клиенты!A${t})`,
    `IF(A${r}="","",COUNTIFS(Продажи!$C$2:$C$1000,A${r},Продажи!$L$2:$L$1000,$B$2))`,
    `IF(A${r}="","",SUMIFS(Продажи!$H$2:$H$1000,Продажи!$C$2:$C$1000,A${r},Продажи!$L$2:$L$1000,$B$2))`,
    `IF(A${r}="","",RANK(C${r},$C$55:$C$84))`,
  ];
  f.forEach((x, j) => { const cell = wsR.getCell(r, j + 1); cell.value = { formula: x }; calc(cell, j === 2 ? CUR : j === 1 ? '0' : null); });
}

// ---------- Счет ----------
wsI.getColumn(1).width = 8; wsI.getColumn(2).width = 44;
wsI.getColumn(3).width = 12; wsI.getColumn(4).width = 16; wsI.getColumn(5).width = 18;
wsI.getCell('A1').value = 'СЧЁТ НА ОПЛАТУ'; wsI.getCell('A1').font = { ...base, bold: true, size: 16 };
const lbl = (r, t) => { const c = wsI.getCell(r, 1); c.value = t; c.font = { ...base, bold: true }; };
lbl(3, '№ счёта'); lbl(4, 'Дата'); lbl(5, 'Клиент'); lbl(6, 'Реквизиты'); lbl(7, 'Поставщик');
wsI.getColumn(1).width = 12;
wsI.getCell('B3').value = 1001; input(wsI.getCell('B3'), '0');
wsI.getCell('B4').value = { formula: 'IFERROR(INDEX(Продажи!$A$2:$A$1000,MATCH(B3&"|1",Продажи!$M$2:$M$1000,0)),"")' }; calc(wsI.getCell('B4'), DATE);
wsI.getCell('B4').alignment = { horizontal: 'left' };
wsI.getCell('B5').value = { formula: 'IF(B3="","",IFERROR(INDEX(Продажи!$C$2:$C$1000,MATCH(B3&"|1",Продажи!$M$2:$M$1000,0)),"счёт не найден"))' }; calc(wsI.getCell('B5'));
wsI.getCell('B6').value = { formula: 'IFERROR(INDEX(Клиенты!$D$2:$D$200,MATCH(B5,Клиенты!$A$2:$A$200,0)),"")' }; calc(wsI.getCell('B6'));
wsI.getCell('B7').value = 'ТОО «Моя компания», БИН 000000000000, г. Астана, ул. Примерная 1, р/с KZ00 0000 0000 0000 0000';
input(wsI.getCell('B7'));
[5, 6, 7].forEach((r) => { wsI.mergeCells(`B${r}:E${r}`); wsI.getCell(`B${r}`).alignment = { wrapText: true, vertical: 'top' }; });
wsI.getRow(6).height = 30; wsI.getRow(7).height = 42;
header(wsI, 9, ['№', 'Наименование', 'Кол-во', 'Цена', 'Сумма']);
const mt = 'Продажи!$M$2:$M$1000';
for (let i = 1; i <= 20; i++) {
  const r = 9 + i;
  const f = [
    `IF(ISNUMBER(MATCH($B$3&"|${i}",${mt},0)),${i},"")`,
    `IF(A${r}="","",INDEX(Продажи!$F$2:$F$1000,MATCH($B$3&"|"&A${r},${mt},0)))`,
    `IF(A${r}="","",INDEX(Продажи!$E$2:$E$1000,MATCH($B$3&"|"&A${r},${mt},0)))`,
    `IF(A${r}="","",INDEX(Продажи!$G$2:$G$1000,MATCH($B$3&"|"&A${r},${mt},0)))`,
    `IF(A${r}="","",INDEX(Продажи!$H$2:$H$1000,MATCH($B$3&"|"&A${r},${mt},0)))`,
  ];
  f.forEach((x, j) => { const cell = wsI.getCell(r, j + 1); cell.value = { formula: x }; calc(cell, j === 3 || j === 4 ? CUR : null); });
}
wsI.getCell('D30').value = 'Итого:'; wsI.getCell('D30').font = { ...base, bold: true }; wsI.getCell('D30').alignment = { horizontal: 'right' };
wsI.getCell('E30').value = { formula: 'SUM(E10:E29)' }; calc(wsI.getCell('E30'), CUR); wsI.getCell('E30').font = { ...base, bold: true };
wsI.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1, printArea: 'A1:E31' };

// ---------- сохранение + график ----------
async function addChart(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const sheetPath = 'xl/worksheets/sheet6.xml';
  let sheet = await zip.file(sheetPath).async('string');
  if (!sheet.includes('xmlns:r=')) sheet = sheet.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  sheet = sheet.replace('</worksheet>', '<drawing r:id="rIdChart1"/></worksheet>');
  zip.file(sheetPath, sheet);

  const relsPath = 'xl/worksheets/_rels/sheet6.xml.rels';
  const rel = '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>';
  const ex = zip.file(relsPath);
  zip.file(relsPath, ex
    ? (await ex.async('string')).replace('</Relationships>', rel + '</Relationships>')
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`);

  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Выручка по месяцам"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`);

  zip.file('xl/charts/chart1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/>
<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="ru-RU" sz="1200" b="1"/><a:t>Выручка по месяцам</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>
<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Отчет!$B$4</c:f></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="1F4E78"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>
<c:cat><c:strRef><c:f>Отчет!$A$5:$A$16</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Отчет!$B$5:$B$16</c:f></c:numRef></c:val></c:ser>
<c:gapWidth val="80"/><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>
<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111111"/><c:crosses val="autoZero"/></c:valAx>
</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`);

  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

(async () => {
  await wb.xlsx.writeFile(OUT);
  await addChart(OUT);
  console.log('OK', OUT);
})();
