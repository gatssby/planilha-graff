const CONFIG = {
  TITLE: 'Controle Familiar - Avós',
  SHEET_ENTRIES: 'Lançamentos',
  PEOPLE: ['Valciria', 'Deco', 'Vivi', 'Nega', 'Valdi'],
  SAVINGS_GOAL: 3000,
  HEADER_ROW: 12,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Controle Familiar')
    .addItem('Configurar / atualizar planilha', 'setupSpreadsheet')
    .addItem('Atualizar resumo', 'refreshDashboard')
    .addToUi();
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('pt_BR');
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_ENTRIES);
  const existingEntries = readExistingEntries_(sheet);

  setupEntriesSheet_(sheet);

  if (existingEntries.length) {
    sheet.getRange(CONFIG.HEADER_ROW + 1, 1, existingEntries.length, 4)
      .setValues(existingEntries);
  }

  refreshDashboard();

  const oldDashboard = ss.getSheetByName('Resumo');
  if (oldDashboard && oldDashboard.getSheetId() !== sheet.getSheetId() && ss.getSheets().length > 1) {
    ss.deleteSheet(oldDashboard);
  }

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    'Planilha atualizada. O resumo e os lançamentos ficam juntos na mesma aba.'
  );
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_ENTRIES) return;
  if (e.range.getRow() <= CONFIG.HEADER_ROW) return;
  refreshDashboard();
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ENTRIES);
  if (!sheet) return;

  const firstDataRow = CONFIG.HEADER_ROW + 1;
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= firstDataRow
    ? sheet.getRange(firstDataRow, 1, lastRow - CONFIG.HEADER_ROW, 4).getValues()
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

  // Resumo compacto.
  sheet.getRange('B2').setValue(reserve).setNumberFormat('R$ #,##0.00');
  sheet.getRange('D2').setValue(CONFIG.SAVINGS_GOAL).setNumberFormat('R$ #,##0.00');
  sheet.getRange('A3:D3').merge();
  sheet.getRange('A3').setFormula('=SPARKLINE(MAX(0,MIN(B2/D2,1)),{"charttype","bar";"max",1})');

  // Estilo Splitwise: ninguém "deve para" outra pessoa.
  // Se o saldo está negativo, a pessoa deve contribuir esse valor para o caixa/reserva.
  // Se está positivo, o valor continua sendo crédito individual dela dentro da reserva.
  const statuses = CONFIG.PEOPLE.map(name => {
    const balance = contributions[name] - fairShare;
    if (balance < -0.005) return [name, 'Deve contribuir', Math.abs(balance)];
    if (balance > 0.005) return [name, 'A haver', balance];
    return [name, 'Em dia', 0];
  });

  sheet.getRange(6, 1, CONFIG.PEOPLE.length, 3).setValues(statuses);
  sheet.getRange(6, 3, CONFIG.PEOPLE.length, 1).setNumberFormat('R$ #,##0.00');

  const backgrounds = [];
  const fontColors = [];
  statuses.forEach(([, status]) => {
    if (status === 'Deve contribuir') {
      backgrounds.push(['#ffffff', '#fce8e6', '#fce8e6']);
      fontColors.push(['#222222', '#b3261e', '#b3261e']);
    } else if (status === 'A haver') {
      backgrounds.push(['#ffffff', '#e6f4ea', '#e6f4ea']);
      fontColors.push(['#222222', '#137333', '#137333']);
    } else {
      backgrounds.push(['#ffffff', '#f1f3f4', '#f1f3f4']);
      fontColors.push(['#222222', '#5f6368', '#5f6368']);
    }
  });
  sheet.getRange(6, 1, CONFIG.PEOPLE.length, 3)
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors);
}

function setupEntriesSheet_(sheet) {
  // Permite reaplicar o layout mesmo sobre versões anteriores da planilha.
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.showColumns(1, sheet.getMaxColumns());
  sheet.setConditionalFormatRules([]);
  sheet.clear();

  sheet.setFrozenRows(0);
  sheet.setTabColor('#444444');

  // Cabeçalho pequeno para ocupar pouco espaço no celular.
  sheet.getRange('A1:D1').merge().setValue(CONFIG.TITLE);
  sheet.getRange('A1:D1')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBackground('#222222')
    .setFontColor('#ffffff');
  sheet.setRowHeight(1, 28);

  sheet.getRange('A2:D2').setValues([['Reserva', '', 'Meta', '']]);
  sheet.getRange('A2').setFontWeight('bold');
  sheet.getRange('C2').setFontWeight('bold');
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 16);

  sheet.getRange('A5:C5').setValues([['Pessoa', 'Situação', 'Valor']]);
  sheet.getRange('A5:C5')
    .setFontWeight('bold')
    .setFontColor('#5f6368')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(5, 24);

  sheet.getRange('A11:D11').merge().setValue('Lançamentos');
  sheet.getRange('A11:D11')
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor('#5f6368');
  sheet.setRowHeight(11, 24);

  sheet.getRange(CONFIG.HEADER_ROW, 1, 1, 4)
    .setValues([['Pessoa', 'Data', 'Valor', 'Observação']])
    .setFontWeight('bold')
    .setBackground('#222222')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PEOPLE, true)
    .setAllowInvalid(false)
    .build();

  const firstDataRow = CONFIG.HEADER_ROW + 1;
  sheet.getRange(firstDataRow, 1, 988, 1).setDataValidation(validation);
  sheet.getRange(firstDataRow, 2, 988, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(firstDataRow, 3, 988, 1).setNumberFormat('R$ #,##0.00');
  sheet.getRange(firstDataRow, 1, 988, 4).setVerticalAlignment('middle');

  // Larguras pensadas para uso no celular.
  sheet.setColumnWidth(1, 105);
  sheet.setColumnWidth(2, 92);
  sheet.setColumnWidth(3, 105);
  sheet.setColumnWidth(4, 220);

  sheet.setRowHeights(6, CONFIG.PEOPLE.length, 28);
  sheet.setRowHeights(firstDataRow, 988, 32);

  // Instruções ficam escondidas para não poluir a interface.
  sheet.getRange('F1').setValue('Como usar');
  sheet.getRange('F2').setValue('Contribuição: valor positivo + nome da pessoa.');
  sheet.getRange('F3').setValue('Gasto com os avós: valor negativo; Pessoa pode ficar em branco.');
  sheet.getRange('F4').setValue('“Deve contribuir” significa repor esse valor no caixa/reserva familiar.');
  sheet.hideColumns(6);

  sheet.getRange(CONFIG.HEADER_ROW, 1, 989, 4).createFilter();
}

function readExistingEntries_(sheet) {
  const lastRow = sheet.getLastRow();
  if (!lastRow) return [];

  const searchRows = Math.min(lastRow, 30);
  const values = sheet.getRange(1, 1, searchRows, 4).getValues();
  let headerRow = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(value => String(value || '').trim());
    if (
      row[0] === 'Pessoa' &&
      row[1] === 'Data' &&
      row[2] === 'Valor' &&
      row[3] === 'Observação'
    ) {
      headerRow = i + 1;
      break;
    }
  }

  if (!headerRow || lastRow <= headerRow) return [];

  return sheet
    .getRange(headerRow + 1, 1, lastRow - headerRow, 4)
    .getValues()
    .filter(row => row.some(value => value !== '' && value !== null));
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
