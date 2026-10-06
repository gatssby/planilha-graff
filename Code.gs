const CONFIG = {
  TITLE: 'Controle Familiar - Avós',
  SHEET_ENTRIES: 'Lançamentos',
  PEOPLE: ['Valciria', 'Deco', 'Vivi', 'Nega', 'Valdi'],
  HEADER_ROW: 9,
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
    'Planilha atualizada. Agora todo lançamento representa um valor pago em favor dos avós.'
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

  const paidByPerson = Object.fromEntries(CONFIG.PEOPLE.map(name => [name, 0]));
  let totalPaid = 0;

  rows.forEach(([person, date, amount]) => {
    const value = Number(amount) || 0;
    const personName = String(person || '').trim();

    if (value <= 0) return;
    if (!Object.prototype.hasOwnProperty.call(paidByPerson, personName)) return;

    paidByPerson[personName] += value;
    totalPaid += value;
  });

  const fairShare = CONFIG.PEOPLE.length ? totalPaid / CONFIG.PEOPLE.length : 0;
  const balances = CONFIG.PEOPLE.map(name => paidByPerson[name] - fairShare);
  const maxAbsBalance = Math.max(...balances.map(Math.abs), 0);

  const statuses = CONFIG.PEOPLE.map((name, index) => {
    const balance = balances[index];

    if (balance < -0.005) {
      return [
        name,
        'Deve contribuir',
        Math.abs(balance),
        makeBalanceBar_(balance, maxAbsBalance),
      ];
    }

    if (balance > 0.005) {
      return [
        name,
        'A haver',
        balance,
        makeBalanceBar_(balance, maxAbsBalance),
      ];
    }

    return [name, 'Em dia', 0, makeBalanceBar_(0, maxAbsBalance)];
  });

  sheet.getRange(3, 1, CONFIG.PEOPLE.length, 4).setValues(statuses);
  sheet.getRange(3, 3, CONFIG.PEOPLE.length, 1).setNumberFormat('R$ #,##0.00');

  const backgrounds = [];
  const fontColorsABC = [];
  const barColors = [];

  statuses.forEach(([, status]) => {
    if (status === 'Deve contribuir') {
      backgrounds.push(['#ffffff', '#fce8e6', '#fce8e6', '#ffffff']);
      fontColorsABC.push(['#222222', '#b3261e', '#b3261e']);
      barColors.push(['#b3261e']);
    } else if (status === 'A haver') {
      backgrounds.push(['#ffffff', '#e6f4ea', '#e6f4ea', '#ffffff']);
      fontColorsABC.push(['#222222', '#137333', '#137333']);
      barColors.push(['#137333']);
    } else {
      backgrounds.push(['#ffffff', '#f1f3f4', '#f1f3f4', '#ffffff']);
      fontColorsABC.push(['#222222', '#5f6368', '#5f6368']);
      barColors.push(['#9aa0a6']);
    }
  });

  sheet.getRange(3, 1, CONFIG.PEOPLE.length, 4).setBackgrounds(backgrounds);
  sheet.getRange(3, 1, CONFIG.PEOPLE.length, 3).setFontColors(fontColorsABC);
  sheet.getRange(3, 4, CONFIG.PEOPLE.length, 1)
    .setFontColors(barColors)
    .setHorizontalAlignment('center')
    .setFontFamily('Roboto Mono')
    .setFontSize(9);
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

  // Cabeçalho compacto para ocupar pouco espaço no celular.
  sheet.getRange('A1:D1').merge().setValue(CONFIG.TITLE);
  sheet.getRange('A1:D1')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBackground('#222222')
    .setFontColor('#ffffff');
  sheet.setRowHeight(1, 28);

  sheet.getRange('A2:D2').setValues([['Pessoa', 'Situação', 'Valor', 'Balanço']]);
  sheet.getRange('A2:D2')
    .setFontWeight('bold')
    .setFontColor('#5f6368')
    .setHorizontalAlignment('left');
  sheet.getRange('D2').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 24);

  sheet.setRowHeights(3, CONFIG.PEOPLE.length, 28);

  sheet.getRange('A8:D8').merge().setValue('Lançamentos');
  sheet.getRange('A8:D8')
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor('#5f6368');
  sheet.setRowHeight(8, 24);

  sheet.getRange(CONFIG.HEADER_ROW, 1, 1, 4)
    .setValues([['Pessoa', 'Data', 'Valor', 'Observação']])
    .setFontWeight('bold')
    .setBackground('#222222')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const personValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PEOPLE, true)
    .setAllowInvalid(false)
    .build();

  const amountValidation = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .setHelpText('Informe somente valores positivos. Cada lançamento representa um valor pago em favor dos avós.')
    .build();

  const firstDataRow = CONFIG.HEADER_ROW + 1;
  const dataRows = 991;
  sheet.getRange(firstDataRow, 1, dataRows, 1).setDataValidation(personValidation);
  sheet.getRange(firstDataRow, 2, dataRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(firstDataRow, 3, dataRows, 1)
    .setNumberFormat('R$ #,##0.00')
    .setDataValidation(amountValidation);
  sheet.getRange(firstDataRow, 1, dataRows, 4).setVerticalAlignment('middle');

  // Larguras pensadas para uso no celular.
  sheet.setColumnWidth(1, 105);
  sheet.setColumnWidth(2, 112);
  sheet.setColumnWidth(3, 105);
  sheet.setColumnWidth(4, 220);
  sheet.setRowHeights(firstDataRow, dataRows, 32);

  // Instruções escondidas para não poluir a interface.
  sheet.getRange('F1').setValue('Como usar');
  sheet.getRange('F2').setValue('Todo lançamento é um valor pago em favor dos avós.');
  sheet.getRange('F3').setValue('Selecione quem pagou e informe o valor sempre positivo.');
  sheet.getRange('F4').setValue('“Deve contribuir” = está abaixo da média paga pelos cinco. “A haver” = pagou acima da média.');
  sheet.hideColumns(6);

  sheet.getRange(CONFIG.HEADER_ROW, 1, dataRows + 1, 4).createFilter();
}

function makeBalanceBar_(balance, maxAbsBalance) {
  const units = 4;
  const empty = '·';
  const filled = '■';
  const axis = '│';

  if (maxAbsBalance < 0.005 || Math.abs(balance) < 0.005) {
    return empty.repeat(units) + axis + empty.repeat(units);
  }

  const fillCount = Math.max(
    1,
    Math.min(units, Math.round((Math.abs(balance) / maxAbsBalance) * units))
  );

  if (balance < 0) {
    return (
      empty.repeat(units - fillCount) +
      filled.repeat(fillCount) +
      axis +
      empty.repeat(units)
    );
  }

  return (
    empty.repeat(units) +
    axis +
    filled.repeat(fillCount) +
    empty.repeat(units - fillCount)
  );
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
