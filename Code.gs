const CONFIG = {
  TITLE: 'Controle Familiar - Avós',
  SHEET_DASHBOARD: 'Resumo',
  SHEET_ENTRIES: 'Lançamentos',
  PEOPLE: ['Val Síria', 'Deco', 'Vivi', 'Nega', 'Valdi'],
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

  const entries = getOrCreateSheet_(ss, CONFIG.SHEET_ENTRIES);
  const dashboard = getOrCreateSheet_(ss, CONFIG.SHEET_DASHBOARD);

  setupEntriesSheet_(entries);
  setupDashboardSheet_(dashboard);
  refreshDashboard();

  ss.setActiveSheet(entries);
  SpreadsheetApp.getUi().alert('Planilha configurada. Use a aba “Lançamentos” para registrar contribuições e gastos.');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_ENTRIES) return;
  if (e.range.getRow() < 2) return;
  refreshDashboard();
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const entries = ss.getSheetByName(CONFIG.SHEET_ENTRIES);
  const dashboard = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (!entries || !dashboard) return;

  const lastRow = entries.getLastRow();
  const rows = lastRow >= 2
    ? entries.getRange(2, 1, lastRow - 1, 4).getValues()
    : [];

  const contributions = Object.fromEntries(CONFIG.PEOPLE.map(name => [name, 0]));
  let totalContributed = 0;
  let totalSpent = 0;

  rows.forEach(([person, date, amount, note]) => {
    const value = Number(amount) || 0;
    const personName = String(person || '').trim();
    const noteText = String(note || '').trim().toLowerCase();

    if (!value) return;

    // Regra simples para manter só 4 campos na entrada:
    // valores positivos = contribuição de uma pessoa;
    // valores negativos = gasto comum com os avós.
    if (value > 0 && contributions.hasOwnProperty(personName)) {
      contributions[personName] += value;
      totalContributed += value;
    } else if (value < 0) {
      totalSpent += Math.abs(value);
    }
  });

  const fairShare = CONFIG.PEOPLE.length ? totalSpent / CONFIG.PEOPLE.length : 0;
  const balances = CONFIG.PEOPLE.map(name => contributions[name] - fairShare);
  const reserve = totalContributed - totalSpent;
  const goal = CONFIG.SAVINGS_GOAL;
  const progress = goal > 0 ? Math.max(0, Math.min(reserve / goal, 1)) : 0;

  dashboard.getRange('B3').setValue(totalContributed);
  dashboard.getRange('B4').setValue(totalSpent);
  dashboard.getRange('B5').setValue(reserve);
  dashboard.getRange('B6').setValue(goal);
  dashboard.getRange('B7').setValue(progress);

  const output = CONFIG.PEOPLE.map((name, index) => [
    name,
    contributions[name],
    fairShare,
    balances[index],
  ]);

  dashboard.getRange(11, 1, CONFIG.PEOPLE.length, 4).setValues(output);
  dashboard.getRange(11, 2, CONFIG.PEOPLE.length, 3).setNumberFormat('R$ #,##0.00');

  const balanceRange = dashboard.getRange(11, 4, CONFIG.PEOPLE.length, 1);
  balanceRange.setConditionalFormatRules([]);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#d9ead3')
      .setFontColor('#274e13')
      .setRanges([balanceRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#f4cccc')
      .setFontColor('#990000')
      .setRanges([balanceRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#eeeeee')
      .setFontColor('#666666')
      .setRanges([balanceRange])
      .build(),
  ];
  balanceRange.setConditionalFormatRules(rules);

  dashboard.getRange('B7').setNumberFormat('0%');
  dashboard.getRange('A8').setFormula('=SPARKLINE(B7,{"charttype","bar";"max",1})');
}

function setupEntriesSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.setTabColor('#444444');

  sheet.getRange('A1:D1').setValues([['Pessoa', 'Data', 'Valor', 'Observação']]);
  sheet.getRange('A1:D1')
    .setFontWeight('bold')
    .setBackground('#222222')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PEOPLE, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A2:A1000').setDataValidation(validation);

  sheet.getRange('B2:B1000').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('C2:C1000').setNumberFormat('R$ #,##0.00');
  sheet.getRange('A2:D1000').setVerticalAlignment('middle');

  // Mobile-first: colunas enxutas, legíveis e sem excesso de informação.
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 95);
  sheet.setColumnWidth(3, 105);
  sheet.setColumnWidth(4, 240);
  sheet.setRowHeights(1, 1000, 34);

  sheet.getRange('F1').setValue('Como usar');
  sheet.getRange('F2').setValue('Contribuição: valor positivo + nome da pessoa.');
  sheet.getRange('F3').setValue('Gasto com os avós: valor negativo. O nome pode ficar em branco.');
  sheet.getRange('F4').setValue('Ex.: remédio de R$ 200 → -200 em Valor e “Remédio” em Observação.');
  sheet.hideColumns(6);

  sheet.getRange('A1:D1000').createFilter();
}

function setupDashboardSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(2);
  sheet.setTabColor('#777777');

  sheet.getRange('A1:D1').merge().setValue(CONFIG.TITLE);
  sheet.getRange('A1:D1')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#222222')
    .setFontColor('#ffffff');

  sheet.getRange('A3:A7').setValues([
    ['Total contribuído'],
    ['Total gasto'],
    ['Reserva atual'],
    ['Meta da reserva'],
    ['Progresso'],
  ]).setFontWeight('bold');

  sheet.getRange('B3:B6').setNumberFormat('R$ #,##0.00');
  sheet.getRange('A8:D8').merge();

  sheet.getRange('A10:D10').setValues([['Pessoa', 'Contribuiu', 'Cota dos gastos', 'Saldo']]);
  sheet.getRange('A10:D10')
    .setFontWeight('bold')
    .setBackground('#eeeeee')
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 125);
  sheet.setColumnWidth(4, 115);
  sheet.setRowHeights(1, 30, 34);

  sheet.getRange('A3:D16').setVerticalAlignment('middle');
  sheet.getRange('B3:B7').setHorizontalAlignment('right');
  sheet.getRange('A11:D15').setHorizontalAlignment('center');

  sheet.getRange('A17:D18').merge();
  sheet.getRange('A17').setValue(
    'Saldo = contribuição da pessoa − 1/5 dos gastos realizados.\n' +
    'Positivo: a pessoa está com valor a haver. Negativo: contribuiu menos que a parcela atual dos gastos.'
  );
  sheet.getRange('A17').setWrap(true).setFontColor('#666666').setFontSize(9);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
