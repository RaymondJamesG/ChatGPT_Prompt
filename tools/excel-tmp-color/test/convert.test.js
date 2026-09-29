const path = require('path');
const test = require('node:test');
const assert = require('node:assert/strict');

const deps = process.env.EXCEL_TOOL_DEPS || '/tmp/excel-tool-deps/node_modules';
globalThis.ExcelJS = require(path.join(deps, 'exceljs'));
globalThis.JSZip = require(path.join(deps, 'jszip'));
const TmpExcel = require('../convert.js');

function find(result, sheet, address) {
  return result.cells.find((cell) => cell.sheet === sheet && cell.address === address);
}

async function build(fill) {
  const workbook = new globalThis.ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('文案');
  fill(sheet, workbook);
  return workbook.xlsx.writeBuffer();
}

test('applyTint matches Excel luminance adjustment', () => {
  assert.equal(TmpExcel.applyTint('FF0000', 0.4), 'FF6666');
  assert.equal(TmpExcel.applyTint('FF0000', -0.5), '800000');
  assert.equal(TmpExcel.applyTint('FF0000', 0), 'FF0000');
  assert.equal(TmpExcel.applyTint('FF0000', 1), 'FFFFFF');
  assert.equal(TmpExcel.applyTint('FF0000', -1), '000000');
  assert.equal(TmpExcel.applyTint('C0504D', -0.25), '953735');
  assert.equal(TmpExcel.applyTint('4F81BD', -0.5), '254061');
  assert.equal(TmpExcel.applyTint('000000', 0.5), '808080');
});

test('applyTmpColor wraps plain text and only updates existing color tags', () => {
  assert.equal(TmpExcel.applyTmpColor('危险警告', 'FF0000'), '<color=#FF0000>危险警告</color>');
  assert.equal(
    TmpExcel.applyTmpColor('<color=#00FF00>治疗完成</color>', '0000FF'),
    '<color=#0000FF>治疗完成</color>'
  );
  assert.equal(
    TmpExcel.applyTmpColor('<color=red>暴击</color>', '0000FF'),
    '<color=#0000FF>暴击</color>'
  );
  assert.equal(
    TmpExcel.applyTmpColor('<color="#00FF00">治疗完成</color>', '112233'),
    '<color=#112233>治疗完成</color>'
  );
  assert.equal(
    TmpExcel.applyTmpColor("获得<color='#00FF00'>金币</color>奖励", 'FF0000'),
    '<color=#FF0000>获得</color><color=#FF0000>金币</color><color=#FF0000>奖励</color>'
  );
  assert.equal(
    TmpExcel.applyTmpColor('<b>暴击</b>', 'FF8800'),
    '<color=#FF8800><b>暴击</b></color>'
  );
  assert.equal(TmpExcel.applyTmpColor('<color=#00FF00>未闭合', 'FF0000'), '<color=#FF0000>未闭合</color>');
  assert.equal(TmpExcel.applyTmpColor('', 'FF0000'), '');
});

test('parseThemeXml uses spreadsheet theme indexes, not XML order', () => {
  const xml = `
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>`;
  const theme = TmpExcel.parseThemeXml(xml);
  assert.equal(theme[0], 'FFFFFF');
  assert.equal(theme[1], '000000');
  assert.equal(theme[4], '4F81BD');
  assert.equal(theme[5], 'C0504D');
  assert.equal(theme[10], '0000FF');
});

test('convert workbook colors into TextMeshPro tags', async () => {
  const buffer = await build((sheet) => {
    sheet.getCell('A1').value = '危险警告';
    sheet.getCell('A1').font = { color: { argb: 'FFFF0000' } };

    sheet.getCell('A2').value = '普通说明';
    sheet.getCell('A2').font = { color: { argb: 'FF000000' } };

    sheet.getCell('A3').value = '没有颜色';

    sheet.getCell('A4').value = '<color=#00FF00>治疗完成</color>';
    sheet.getCell('A4').font = { color: { argb: 'FF0000FF' } };

    sheet.getCell('A5').value = {
      richText: [
        { text: '获得' },
        { text: '稀有道具', font: { color: { argb: 'FFFF0000' } } },
        { text: '！', font: { color: { argb: 'FF00AAFF' } } }
      ]
    };

    sheet.getCell('A6').value = '主题强调';
    sheet.getCell('A6').font = { color: { theme: 4 } };

    sheet.getCell('A7').value = '主题加深';
    sheet.getCell('A7').font = { color: { theme: 5, tint: -0.25 } };

    sheet.getCell('A8').value = '索引红';
    sheet.getCell('A8').font = { color: { indexed: 10 } };

    sheet.getCell('A9').value = '<color=#112233>保持</color>';
    sheet.getCell('A9').font = { color: { argb: 'FF112233' } };

    sheet.getCell('A10').value = '<color=#ff0000>生命值</color>';
    sheet.getCell('A10').font = { color: { argb: 'FFFF0000' } };

    sheet.getCell('A11').value = '白字';
    sheet.getCell('A11').font = { color: { argb: 'FFFFFFFF' } };

    sheet.getCell('A12').value = '<b>暴击</b>';
    sheet.getCell('A12').font = { color: { argb: 'FFFF8800' } };

    sheet.getCell('A13').value = "获得<color='#00FF00'>金币</color>奖励";
    sheet.getCell('A13').font = { color: { argb: 'FFFF0000' } };

    sheet.getCell('A14').value = 1280;
    sheet.getCell('A14').font = { color: { argb: 'FF800080' } };

    sheet.getCell('A15').value = '只有底色';
    sheet.getCell('A15').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }
    };

    sheet.getCell('A16').value = new Date(Date.UTC(2024, 0, 2));
    sheet.getCell('A16').font = { color: { argb: 'FFFF0000' } };

    sheet.getCell('A17').value = { formula: '1+2', result: 3 };
    sheet.getCell('A17').font = { color: { argb: 'FF0000FF' } };

    sheet.getCell('A18').value = { text: '查看详情', hyperlink: 'https://example.com' };
    sheet.getCell('A18').font = { color: { theme: 10 } };

    sheet.getCell('A19').value = {
      richText: [
        { text: '甲', font: { color: { argb: 'FFFF0000' } } },
        { text: '乙', font: { color: { argb: 'FFFF0000' } } }
      ]
    };

    sheet.getCell('A20').value = {
      richText: [
        { text: '<color=#00FF00>甲</color>', font: { color: { argb: 'FFFF0000' } } },
        { text: '乙', font: { color: { argb: 'FFFF0000' } } }
      ]
    };

    sheet.getCell('A21').value = '半透明红';
    sheet.getCell('A21').font = { color: { argb: '80FF0000' } };

    sheet.getCell('A22').value = '透明字节当不透明';
    sheet.getCell('A22').font = { color: { argb: '00FF0000' } };

    sheet.getCell('A23').value = '默认主题色';
    sheet.getCell('A23').font = { color: { theme: 1 } };

    sheet.getCell('A24').value = '文本色变浅';
    sheet.getCell('A24').font = { color: { theme: 1, tint: 0.5 } };

    sheet.getCell('A25').value = {
      richText: [
        { text: '甲' },
        { text: '乙', font: { color: { argb: 'FF0000FF' } } },
        { text: '丙', font: { bold: true } }
      ]
    };
    sheet.getCell('A25').font = { color: { argb: 'FFFF0000' } };

    sheet.getRow(26).font = { color: { argb: 'FF008800' } };
    sheet.getCell('A26').value = '整行变绿';
  });

  const other = new globalThis.ExcelJS.Workbook();
  await other.xlsx.load(buffer);
  const sheet2 = other.addWorksheet('备注');
  sheet2.getCell('B2').value = '第二张表';
  sheet2.getCell('B2').font = { color: { argb: 'FF123456' } };
  const finalBuffer = await other.xlsx.writeBuffer();

  const result = await TmpExcel.convertArrayBuffer(finalBuffer);
  assert.equal(find(result, '文案', 'A1').tagged, '<color=#FF0000>危险警告</color>');
  assert.equal(find(result, '文案', 'A2'), undefined);
  assert.equal(find(result, '文案', 'A3'), undefined);
  assert.equal(find(result, '文案', 'A4').tagged, '<color=#0000FF>治疗完成</color>');
  assert.equal(
    find(result, '文案', 'A5').tagged,
    '获得<color=#FF0000>稀有道具</color><color=#00AAFF>！</color>'
  );
  assert.equal(find(result, '文案', 'A6').tagged, '<color=#4F81BD>主题强调</color>');
  assert.equal(find(result, '文案', 'A7').tagged, '<color=#953735>主题加深</color>');
  assert.equal(find(result, '文案', 'A8').tagged, '<color=#FF0000>索引红</color>');
  assert.equal(find(result, '文案', 'A9'), undefined);
  assert.equal(find(result, '文案', 'A10'), undefined);
  assert.equal(find(result, '文案', 'A11').tagged, '<color=#FFFFFF>白字</color>');
  assert.equal(find(result, '文案', 'A12').tagged, '<color=#FF8800><b>暴击</b></color>');
  assert.equal(
    find(result, '文案', 'A13').tagged,
    '<color=#FF0000>获得</color><color=#FF0000>金币</color><color=#FF0000>奖励</color>'
  );
  assert.equal(find(result, '文案', 'A14').tagged, '<color=#800080>1280</color>');
  assert.equal(find(result, '文案', 'A15'), undefined);
  assert.equal(find(result, '文案', 'A16').tagged, '<color=#FF0000>2024-01-02</color>');
  assert.equal(find(result, '文案', 'A17').tagged, '<color=#0000FF>3</color>');
  assert.equal(find(result, '文案', 'A17').formula, true);
  assert.equal(find(result, '文案', 'A18').tagged, '<color=#0000FF>查看详情</color>');
  assert.equal(find(result, '文案', 'A19').tagged, '<color=#FF0000>甲乙</color>');
  assert.equal(
    find(result, '文案', 'A20').tagged,
    '<color=#FF0000>甲</color><color=#FF0000>乙</color>'
  );
  assert.equal(find(result, '文案', 'A21').tagged, '<color=#FF000080>半透明红</color>');
  assert.equal(find(result, '文案', 'A22').tagged, '<color=#FF0000>透明字节当不透明</color>');
  assert.equal(find(result, '文案', 'A23'), undefined);
  assert.equal(find(result, '文案', 'A24').tagged, '<color=#808080>文本色变浅</color>');
  assert.equal(
    find(result, '文案', 'A25').tagged,
    '<color=#FF0000>甲</color><color=#0000FF>乙</color><color=#FF0000>丙</color>'
  );
  assert.equal(find(result, '文案', 'A26').tagged, '<color=#008800>整行变绿</color>');
  assert.equal(find(result, '备注', 'B2').tagged, '<color=#123456>第二张表</color>');
  assert.ok(result.scanned > result.cells.length);

  const noAlpha = await TmpExcel.convertArrayBuffer(finalBuffer, { keepAlpha: false });
  assert.equal(find(noAlpha, '文案', 'A21').tagged, '<color=#FF0000>半透明红</color>');

  const withBlack = await TmpExcel.convertArrayBuffer(finalBuffer, { includeBlack: true });
  assert.equal(find(withBlack, '文案', 'A2').tagged, '<color=#000000>普通说明</color>');

  const checkBook = new globalThis.ExcelJS.Workbook();
  await checkBook.xlsx.load(finalBuffer);
  TmpExcel.applyChanges(checkBook, result.cells);
  assert.equal(checkBook.getWorksheet('文案').getCell('A1').value, '<color=#FF0000>危险警告</color>');
  assert.equal(checkBook.getWorksheet('文案').getCell('A3').value, '没有颜色');
  assert.equal(checkBook.getWorksheet('备注').getCell('B2').value, '<color=#123456>第二张表</color>');
});

test('old xls files are rejected with a clear message', async () => {
  const ole = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0, 0, 0, 0]);
  await assert.rejects(
    () => TmpExcel.convertArrayBuffer(ole),
    /另存为 \.xlsx/
  );
});
