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

  const sourceRequiredHeaders = [
    'Butler Item Code',
    'Description',
    'Color',
    'BIN CARD NUMBER',
    'SUPPLIER NAME',
    'OPENING STOCK ENTRY'
  ];

  const missingInSource = sourceRequiredHeaders.filter(h => sourceMap[h] === undefined);
  if (missingInSource.length > 0) {
    ui.alert('Missing required header(s) in PRICE LIST MASTER:\n' + missingInSource.join(', '));
    return;
  }

  const ledgerRequiredHeaders = [
    'ITEMCODE',
    'MATERIAL NAME',
    'COLOR',
    'BIN CARD NUMBER',
    'SUPPLIER NAME',
    'OPENING STOCK ENTRY'
  ];

  const missingInLedger = ledgerRequiredHeaders.filter(h => ledgerMap[h] === undefined);
  if (missingInLedger.length > 0) {
    ui.alert('Missing required header(s) in OPENING STOCK LEDGER:\n' + missingInLedger.join(', '));
    return;
  }

  const inOutRequiredHeaders = [
    'DATE',
    'LEDGER',
    'ITEMCODE',
    'MATERIAL NAME',
    'COLOR',
    'BIN CARD NUMBER',
    'SUPPLIER NAME',
    'RECEIVED QTY'
  ];

  const missingInInOut = inOutRequiredHeaders.filter(h => inOutMap[h] === undefined);
  if (missingInInOut.length > 0) {
    ui.alert('Missing required header(s) in IN-OUT ENTRY:\n' + missingInInOut.join(', '));
    return;
  }

  const existingLedgerItemCodes = osGetExistingItemCodes_(ledgerSheet, ledgerMap);
  const currentBatchItemCodes = new Set();

  const rowsToLedger = [];
  const rowsToInOut = [];
  const sourceRowNumbersToMark = [];

  const now = new Date();

  for (let r = 1; r < sourceData.length; r++) {
    const row = sourceData[r];

    const itemCode = String(row[sourceMap['Butler Item Code']]).trim();
    const materialName = String(row[sourceMap['Description']]).trim();
    const color = String(row[sourceMap['Color']]).trim();
    const binCardNumber = String(row[sourceMap['BIN CARD NUMBER']]).trim();
    const supplierName = String(row[sourceMap['SUPPLIER NAME']]).trim();
    const openingStockQty = row[sourceMap['OPENING STOCK ENTRY']];

    const isQualified =
      itemCode !== '' &&
      materialName !== '' &&
      color !== '' &&
      binCardNumber !== '' &&
      supplierName !== '' &&
      String(openingStockQty).trim() !== '' &&
      String(openingStockQty).trim() !== 'OPENING STOCK UPDATED';

    if (!isQualified) continue;

    if (existingLedgerItemCodes.has(itemCode) || currentBatchItemCodes.has(itemCode)) {
      continue;
    }

    const ledgerRow = new Array(ledgerHeaders.length).fill('');
    ledgerHeaders.forEach((header, idx) => {
      if (header === 'DATE') {
        ledgerRow[idx] = now;
      } else if (header === 'ITEMCODE') {
        ledgerRow[idx] = itemCode;
      } else if (header === 'MATERIAL NAME') {
        ledgerRow[idx] = materialName;
      } else if (header === 'COLOR') {
        ledgerRow[idx] = color;
      } else if (header === 'BIN CARD NUMBER') {
        ledgerRow[idx] = binCardNumber;
      } else if (header === 'SUPPLIER NAME') {
        ledgerRow[idx] = supplierName;
      } else if (header === 'OPENING STOCK ENTRY') {
        ledgerRow[idx] = openingStockQty;
      }
    });
    rowsToLedger.push(ledgerRow);

    const inOutRow = new Array(inOutHeaders.length).fill('');
    inOutHeaders.forEach((header, idx) => {
      if (header === 'DATE') {
        inOutRow[idx] = now;
      } else if (header === 'LEDGER') {
        inOutRow[idx] = 'OPENING STOCK';
      } else if (header === 'ITEMCODE') {
        inOutRow[idx] = itemCode;
      } else if (header === 'MATERIAL NAME') {
        inOutRow[idx] = materialName;
      } else if (header === 'COLOR') {
        inOutRow[idx] = color;
      } else if (header === 'BIN CARD NUMBER') {
        inOutRow[idx] = binCardNumber;
      } else if (header === 'SUPPLIER NAME') {
        inOutRow[idx] = supplierName;
      } else if (header === 'RECEIVED QTY') {
        inOutRow[idx] = openingStockQty;
      } else if (header === 'ISSUED QTY') {
        inOutRow[idx] = '';
      }
    });
    rowsToInOut.push(inOutRow);

    sourceRowNumbersToMark.push(r + 1);
    currentBatchItemCodes.add(itemCode);
  }

  if (rowsToLedger.length === 0) {
    ui.alert('No qualifying new rows found to update.');
    return;
  }

  osAppendRows_(ledgerSheet, rowsToLedger, ledgerHeaders.length);
  osAppendRowsSkippingFormulaColumns_(inOutSheet, rowsToInOut, inOutHeaders.length);

  SpreadsheetApp.flush();

  osMarkProcessedOpeningStock_(
    sourceSheet,
    sourceMap['OPENING STOCK ENTRY'] + 1,
    sourceRowNumbersToMark
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

function osAppendRowsSkippingFormulaColumns_(sheet, rows, width) {
  const startRow = osGetNextAppendRow_(sheet, width);
  if (!rows || rows.length === 0) return;

  const formulaFlags = osGetFormulaColumnFlags_(sheet, width);

  let blockStart = null;
  for (let c = 0; c < width; c++) {
    const isWritable = !formulaFlags[c];

    if (isWritable && blockStart === null) {
      blockStart = c;
    }

    const isBlockEnd = blockStart !== null && (!isWritable || c === width - 1);
    if (isBlockEnd) {
      const blockEnd = isWritable && c === width - 1 ? c : c - 1;
      const blockWidth = blockEnd - blockStart + 1;

      const values = rows.map(row => row.slice(blockStart, blockEnd + 1));
      sheet.getRange(startRow, blockStart + 1, rows.length, blockWidth).setValues(values);

      blockStart = null;
    }
  }
}

function osGetFormulaColumnFlags_(sheet, width) {
  if (sheet.getMaxRows() < 2 || width < 1) {
    return new Array(width).fill(false);
  }

  const formulas = sheet.getRange(2, 1, 1, width).getFormulas()[0];
  return formulas.map(f => String(f).trim() !== '');
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

function osMarkProcessedOpeningStock_(sheet, colIndex, rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return;

  const contiguousRanges = osBuildContiguousRanges_(rowNumbers);

  contiguousRanges.forEach(rangeObj => {
    const range = sheet.getRange(rangeObj.startRow, colIndex, rangeObj.numRows, 1);
    range.setValue('OPENING STOCK UPDATED');
    range.setBackground('#ff0000');

    const protection = range.protect().setDescription('Protected: OPENING STOCK UPDATED');
    const editors = protection.getEditors();
    if (editors.length > 0) {
      protection.removeEditors(editors);
    }
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  });
}

function osBuildContiguousRanges_(rowNumbers) {
  const ranges = [];
  if (!rowNumbers || rowNumbers.length === 0) return ranges;

  const sorted = [...rowNumbers].sort((a, b) => a - b);

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
