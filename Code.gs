const CONFIG = {
  TITLE: 'Controle Familiar - Avós',
  SHEET_ENTRIES: 'Lançamentos',
  PEOPLE: ['Valciria', 'Deco', 'Vivi', 'Nega', 'Valdi'],
  SAVINGS_GOAL: 3000,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Controle Familiar')
    .addItem('Configurar planilha', 'setupSpreadsheet')
    .addItem('Atualizar resumo', 'refreshDashboard')
    .addToUi();
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('pt_BR');
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_ENTRIES);
  setupEntriesSheet_(sheet);
  refreshDashboard();

  const oldDashboard = ss.getSheetByName('Resumo');
  if (oldDashboard && ss.getSheets().length > 1) ss.deleteSheet(oldDashboard);

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('Planilha configurada. Lançamentos e resumo ficam juntos na mesma aba.');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_ENTRIES) return;
  if (e.range.getRow() < 13) return;
  refreshDashboard();
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ENTRIES);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 13
    ? sheet.getRange(13, 1, lastRow - 12, 4).getValues()
    : [];

  const contributions = Object.fromEntries(CONFIG.PEOPLE.map(name => [name, 0]));
  let totalContributed = 0;
  let totalSpent = 0;

  rows.forEach(([person, date, amount]) => {
    const value = Number(amount) || 0;
    const personName = String(person || '').trim();
    if (!value) return;

    if (value > 0 && Object.prototype.hasOwnProperty.call(contributions, personName)) {
      contributions[personName] += value;
      totalContributed += value;
    } else if (value < 0) {
      totalSpent += Math.abs(value);
    }
  });

  const fairShare = CONFIG.PEOPLE.length ? totalSpent / CONFIG.PEOPLE.length : 0;
  const reserve = totalContributed - totalSpent;
  const progress = CONFIG.SAVINGS_GOAL > 0
    ? Math.max(0, Math.min(reserve / CONFIG.SAVINGS_GOAL, 1))
    : 0;

  sheet.getRange('B3').setValue(totalContributed);
  sheet.getRange('B4').setValue(totalSpent);
  sheet.getRange('B5').setValue(reserve);
  sheet.getRange('B6').setValue(CONFIG.SAVINGS_GOAL);
  sheet.getRange('B7').setValue(progress);

  const output = CONFIG.PEOPLE.map(name => [
    name,
    contributions[name],
    fairShare,
    contributions[name] - fairShare,
  ]);
  sheet.getRange(3, 6, CONFIG.PEOPLE.length, 4).setValues(output);
  sheet.getRange(3, 7, CONFIG.PEOPLE.length, 3).setNumberFormat('R$ #,##0.00');

  const balanceRange = sheet.getRange(3, 9, CONFIG.PEOPLE.length, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setBackground('#d9ead3').setFontColor('#274e13')
      .setRanges([balanceRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setBackground('#f4cccc').setFontColor('#990000')
      .setRanges([balanceRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0).setBackground('#eeeeee').setFontColor('#666666')
      .setRanges([balanceRange]).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  sheet.getRange('B7').setNumberFormat('0%');
  sheet.getRange('A8:D8').merge();
  sheet.getRange('A8').setFormula('=SPARKLINE(B7,{"charttype","bar";"max",1})');
}

function setupEntriesSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(12);
  sheet.setTabColor('#444444');

  sheet.getRange('A1:I1').merge().setValue(CONFIG.TITLE);
  sheet.getRange('A1:I1')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#222222').setFontColor('#ffffff');

  sheet.getRange('A3:A7').setValues([
    ['Total contribuído'],
    ['Total gasto'],
    ['Reserva atual'],
    ['Meta da reserva'],
    ['Progresso'],
  ]).setFontWeight('bold');
  sheet.getRange('B3:B6').setNumberFormat('R$ #,##0.00');

  sheet.getRange('F2:I2').setValues([['Pessoa', 'Contribuiu', 'Cota dos gastos', 'Saldo']]);
  sheet.getRange('F2:I2')
    .setFontWeight('bold').setBackground('#eeeeee').setHorizontalAlignment('center');

  sheet.getRange('A10:I10').merge().setValue('Lançamentos');
  sheet.getRange('A10:I10').setFontWeight('bold').setFontSize(13);

  sheet.getRange('A12:D12').setValues([['Pessoa', 'Data', 'Valor', 'Observação']]);
  sheet.getRange('A12:D12')
    .setFontWeight('bold').setBackground('#222222').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PEOPLE, true).setAllowInvalid(false).build();
  sheet.getRange('A13:A1000').setDataValidation(validation);
  sheet.getRange('B13:B1000').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('C13:C1000').setNumberFormat('R$ #,##0.00');
  sheet.getRange('A13:D1000').setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 95);
  sheet.setColumnWidth(3, 105);
  sheet.setColumnWidth(4, 240);
  sheet.setColumnWidth(5, 24);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 115);
  sheet.setColumnWidth(8, 125);
  sheet.setColumnWidth(9, 110);
  sheet.setRowHeights(1, 1000, 34);

  sheet.getRange('K1').setValue('Como usar');
  sheet.getRange('K2').setValue('Contribuição: valor positivo + nome da pessoa.');
  sheet.getRange('K3').setValue('Gasto com os avós: valor negativo. O nome pode ficar em branco.');
  sheet.hideColumns(11);

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange('A12:D1000').createFilter();
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
