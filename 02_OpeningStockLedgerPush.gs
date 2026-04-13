function updateOpeningStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const activeSheet = ss.getActiveSheet();

  const sourceName = 'PRICE LIST MASTER';
  const ledgerName = 'OPENING STOCK LEDGER';
  const inOutName = 'IN-OUT ENTRY';

  if (!activeSheet || activeSheet.getName() !== sourceName) {
    ui.alert('Please open the PRICE LIST MASTER tab, then click "Update Opening Stock".');
    return;
  }

  const sourceSheet = ss.getSheetByName(sourceName);
  const ledgerSheet = ss.getSheetByName(ledgerName);
  const inOutSheet = ss.getSheetByName(inOutName);

  if (!sourceSheet || !ledgerSheet || !inOutSheet) {
    ui.alert('Required sheet missing. Check PRICE LIST MASTER / OPENING STOCK LEDGER / IN-OUT ENTRY.');
    return;
  }

  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastCol = sourceSheet.getLastColumn();

  if (sourceLastRow < 2) {
    ui.alert('No data found in PRICE LIST MASTER.');
    return;
  }

  const sourceData = sourceSheet.getRange(1, 1, sourceLastRow, sourceLastCol).getValues();
  const sourceHeaders = sourceData[0].map(h => String(h).trim());

  const ledgerHeaders = ledgerSheet
    .getRange(1, 1, 1, ledgerSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());

  const inOutHeaders = inOutSheet
    .getRange(1, 1, 1, inOutSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());

  const sourceMap = osCreateHeaderMap_(sourceHeaders);
  const ledgerMap = osCreateHeaderMap_(ledgerHeaders);
  const inOutMap = osCreateHeaderMap_(inOutHeaders);

  const sourceQtyHeader = 'OPENING STOCK ENTRY';
  const sourceRequiredHeaders = [
    'ITEMCODE',
    'MATERIAL NAME',
    'COLOR',
    'BIN CARD NUMBER',
    'SUPPLIER NAME',
    sourceQtyHeader
  ];

  const inOutRequiredHeaders = [
    'LEDGER',
    'ITEMCODE',
    'MATERIAL NAME',
    'COLOR',
    'BIN CARD NUMBER',
    'BIN CARD LOCATION',
    'SUPPLIER NAME',
    'RECEIVED QTY',
    'ISSUED QTY',
    'RELATED TO CUSTOMER NAME',
    'RELATED TO CUSTOMER PO NO',
    'RELATED TO STYLE NAME'
  ];

  const missingInSource = sourceRequiredHeaders.filter(h => sourceMap[h] === undefined);
  if (missingInSource.length > 0) {
    ui.alert('Missing required header(s) in PRICE LIST MASTER:\n' + missingInSource.join(', '));
    return;
  }

  const missingInLedger = [
    'ITEMCODE',
    'MATERIAL NAME',
    'COLOR',
    'BIN CARD NUMBER',
    'SUPPLIER NAME'
  ].filter(h => ledgerMap[h] === undefined);
  if (missingInLedger.length > 0) {
    ui.alert('Missing required header(s) in OPENING STOCK LEDGER:\n' + missingInLedger.join(', '));
    return;
  }

  const ledgerQtyHeader = osFindFirstExistingHeader_(
    ledgerMap,
    ['OPENING STOCK ENTRY', 'OPENING STOCK QTY', 'RECEIVED QTY']
  );
  if (!ledgerQtyHeader) {
    ui.alert('OPENING STOCK LEDGER must contain one of these headers:\nOPENING STOCK ENTRY, OPENING STOCK QTY, RECEIVED QTY');
    return;
  }

  const missingInInOut = inOutRequiredHeaders.filter(h => inOutMap[h] === undefined);
  if (missingInInOut.length > 0) {
    ui.alert('Missing required header(s) in IN-OUT ENTRY:\n' + missingInInOut.join(', '));
    return;
  }

  const existingLedgerItemCodes = osGetExistingItemCodes_(ledgerSheet, ledgerMap);
  const currentBatchItemCodes = new Set();

  const rowsToLedger = [];
  const rowsToInOut = [];
  const sourceRowNumbersToClear = [];

  const now = new Date();

  for (let r = 1; r < sourceData.length; r++) {
    const row = sourceData[r];

    const qtyValue = row[sourceMap[sourceQtyHeader]];
    const itemCode = String(row[sourceMap['ITEMCODE']]).trim();

    const isQualified =
      itemCode !== '' &&
      String(row[sourceMap['MATERIAL NAME']]).trim() !== '' &&
      String(row[sourceMap['COLOR']]).trim() !== '' &&
      String(row[sourceMap['BIN CARD NUMBER']]).trim() !== '' &&
      String(row[sourceMap['SUPPLIER NAME']]).trim() !== '' &&
      String(qtyValue).trim() !== '';

    if (!isQualified) continue;

    if (existingLedgerItemCodes.has(itemCode) || currentBatchItemCodes.has(itemCode)) {
      continue;
    }

    const ledgerRow = new Array(ledgerHeaders.length).fill('');
    ledgerHeaders.forEach((header, idx) => {
      if (header === 'DATE') {
        ledgerRow[idx] = now;
      } else if (header === ledgerQtyHeader) {
        ledgerRow[idx] = qtyValue;
      } else if (sourceMap[header] !== undefined) {
        ledgerRow[idx] = row[sourceMap[header]];
      }
    });
    rowsToLedger.push(ledgerRow);

    const inOutRow = new Array(inOutHeaders.length).fill('');
    inOutHeaders.forEach((header, idx) => {
      if (header === 'DATE') {
        inOutRow[idx] = now;
      } else if (header === 'LEDGER') {
        inOutRow[idx] = 'OPENING STOCK';
      } else if (header === 'RECEIVED QTY') {
        inOutRow[idx] = qtyValue;
      } else if (header === 'ISSUED QTY') {
        inOutRow[idx] = '';
      } else if (sourceMap[header] !== undefined) {
        inOutRow[idx] = row[sourceMap[header]];
      }
    });
    rowsToInOut.push(inOutRow);

    sourceRowNumbersToClear.push(r + 1);
    currentBatchItemCodes.add(itemCode);
  }

  if (rowsToLedger.length === 0) {
    ui.alert('No qualifying new rows found to update.');
    return;
  }

  osAppendRows_(ledgerSheet, rowsToLedger, ledgerHeaders.length);
  osAppendRows_(inOutSheet, rowsToInOut, inOutHeaders.length);

  SpreadsheetApp.flush();

  osClearAndMarkProcessedOpeningStock_(
    sourceSheet,
    sourceMap[sourceQtyHeader] + 1,
    sourceRowNumbersToClear
  );

  SpreadsheetApp.flush();

  ss.toast(
    rowsToLedger.length + ' row(s) updated successfully.',
    'Opening Stock Update',
    5
  );
}

function osCreateHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const cleanHeader = String(header).trim();
    if (cleanHeader !== '' && map[cleanHeader] === undefined) {
      map[cleanHeader] = index;
    }
  });
  return map;
}

function osFindFirstExistingHeader_(headerMap, headerOptions) {
  for (let i = 0; i < headerOptions.length; i++) {
    if (headerMap[headerOptions[i]] !== undefined) {
      return headerOptions[i];
    }
  }
  return '';
}

function osGetExistingItemCodes_(sheet, headerMap) {
  const set = new Set();
  const itemCodeCol = headerMap['ITEMCODE'];
  if (itemCodeCol === undefined) return set;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;

  const values = sheet.getRange(2, itemCodeCol + 1, lastRow - 1, 1).getDisplayValues();
  values.forEach(row => {
    const code = String(row[0]).trim();
    if (code !== '') set.add(code);
  });

  return set;
}

function osAppendRows_(sheet, rows, width) {
  const startRow = osGetNextAppendRow_(sheet, width);
  sheet.getRange(startRow, 1, rows.length, width).setValues(rows);
}

function osGetNextAppendRow_(sheet, width) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  const data = sheet.getRange(1, 1, lastRow, width).getDisplayValues();
  for (let r = data.length - 1; r >= 0; r--) {
    const hasAnyValue = data[r].some(cell => String(cell).trim() !== '');
    if (hasAnyValue) {
      return r + 2;
    }
  }
  return 2;
}

function osClearAndMarkProcessedOpeningStock_(sheet, colIndex, rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return;

  const contiguousRanges = osBuildContiguousRanges_(rowNumbers);

  contiguousRanges.forEach(rangeObj => {
    const range = sheet.getRange(rangeObj.startRow, colIndex, rangeObj.numRows, 1);
    range.clearContent();
    range.setBackground('#ff0000');
  });
}

function osBuildContiguousRanges_(rowNumbers) {
  const sorted = [...rowNumbers].sort((a, b) => a - b);
  const ranges = [];

  let startRow = sorted[0];
  let prevRow = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const currentRow = sorted[i];

    if (currentRow === prevRow + 1) {
      prevRow = currentRow;
      continue;
    }

    ranges.push({
      startRow: startRow,
      numRows: prevRow - startRow + 1
    });

    startRow = currentRow;
    prevRow = currentRow;
  }

  ranges.push({
    startRow: startRow,
    numRows: prevRow - startRow + 1
  });

  return ranges;
}
