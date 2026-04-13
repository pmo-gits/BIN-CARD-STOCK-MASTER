function updateOpeningStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== "PRICE LIST MASTER") {
    ui.alert("This function can only be run from the PRICE LIST MASTER tab.");
    return;
  }

  const priceSheet = ss.getSheetByName("PRICE LIST MASTER");
  const openingLedgerSheet = ss.getSheetByName("OPENING STOCK LEDGER");

  if (!priceSheet || !openingLedgerSheet) {
    ui.alert("Required sheet not found. Please check sheet names.");
    return;
  }

  const priceData = priceSheet.getDataRange().getValues();
  const ledgerData = openingLedgerSheet.getDataRange().getValues();

  if (priceData.length < 2) {
    ui.alert("PRICE LIST MASTER has no data.");
    return;
  }

  if (ledgerData.length < 1) {
    ui.alert("OPENING STOCK LEDGER header row is missing.");
    return;
  }

  const priceHeaders = priceData[0];
  const ledgerHeaders = ledgerData[0];

  const priceHeaderMap = createHeaderMap_(priceHeaders);
  const ledgerHeaderMap = createHeaderMap_(ledgerHeaders);

  const priceItemCodeCol = findHeaderIndex_(priceHeaderMap, [
    "ITEMCODE",
    "BUTLER ITEM CODE",
    "BUTLER ITEMCODE"
  ]);

  const priceBinCardCol = findHeaderIndex_(priceHeaderMap, [
    "BIN CARD NUMBER"
  ]);

  const priceOpeningStockCol = findHeaderIndex_(priceHeaderMap, [
    "OPENING STOCK ENTRY"
  ]);

  const ledgerItemCodeCol = findHeaderIndex_(ledgerHeaderMap, [
    "ITEMCODE"
  ]);

  const ledgerBinCardCol = findHeaderIndex_(ledgerHeaderMap, [
    "BIN CARD NUMBER"
  ]);

  const ledgerOpeningStockCol = findHeaderIndex_(ledgerHeaderMap, [
    "OPENING STOCK ENTRY"
  ]);

  const missingHeaders = [];
  if (priceItemCodeCol === -1) missingHeaders.push("PRICE LIST MASTER -> ITEMCODE / BUTLER ITEM CODE");
  if (priceBinCardCol === -1) missingHeaders.push("PRICE LIST MASTER -> BIN CARD NUMBER");
  if (priceOpeningStockCol === -1) missingHeaders.push("PRICE LIST MASTER -> OPENING STOCK ENTRY");
  if (ledgerItemCodeCol === -1) missingHeaders.push("OPENING STOCK LEDGER -> ITEMCODE");
  if (ledgerBinCardCol === -1) missingHeaders.push("OPENING STOCK LEDGER -> BIN CARD NUMBER");
  if (ledgerOpeningStockCol === -1) missingHeaders.push("OPENING STOCK LEDGER -> OPENING STOCK ENTRY");

  if (missingHeaders.length > 0) {
    ui.alert("Missing required header(s):\n\n" + missingHeaders.join("\n"));
    return;
  }

  const existingItemCodes = new Set();
  for (let i = 1; i < ledgerData.length; i++) {
    const row = ledgerData[i];
    const itemCode = cleanValue_(row[ledgerItemCodeCol]);
    if (itemCode !== "") {
      existingItemCodes.add(itemCode);
    }
  }

  const rowsToAppend = [];
  const skippedDuplicates = [];
  const seenNewItemCodes = new Set();
  const rowsToClear = [];

  for (let i = 1; i < priceData.length; i++) {
    const row = priceData[i];

    const openingStock = row[priceOpeningStockCol];
    const itemCode = cleanValue_(row[priceItemCodeCol]);
    const binCardNumber = cleanValue_(row[priceBinCardCol]);

    if (isBlank_(openingStock)) continue;
    if (itemCode === "") continue;

    if (existingItemCodes.has(itemCode)) {
      skippedDuplicates.push(itemCode);
      continue;
    }

    if (seenNewItemCodes.has(itemCode)) {
      skippedDuplicates.push(itemCode);
      continue;
    }

    const newLedgerRow = new Array(ledgerHeaders.length).fill("");
    newLedgerRow[ledgerItemCodeCol] = itemCode;
    newLedgerRow[ledgerBinCardCol] = binCardNumber;
    newLedgerRow[ledgerOpeningStockCol] = openingStock;

    rowsToAppend.push(newLedgerRow);
    rowsToClear.push(i + 1); // actual sheet row number
    seenNewItemCodes.add(itemCode);
  }

  if (rowsToAppend.length > 0) {
    const startRow = openingLedgerSheet.getLastRow() + 1;
    openingLedgerSheet
      .getRange(startRow, 1, rowsToAppend.length, ledgerHeaders.length)
      .setValues(rowsToAppend);

    for (let j = 0; j < rowsToClear.length; j++) {
      const targetCell = priceSheet.getRange(rowsToClear[j], priceOpeningStockCol + 1);
      targetCell.clearContent();
      targetCell.setBackground("#ff0000");
      targetCell.setValue("OPENING STOCK UPDATED");
    }
  }

  let message = "Opening stock update completed.\n\n";
  message += "New rows added: " + rowsToAppend.length + "\n";
  message += "Duplicate itemcodes skipped: " + [...new Set(skippedDuplicates)].length + "\n";
  message += "Opening stock cleared for posted rows: " + rowsToClear.length;

  if (skippedDuplicates.length > 0) {
    message += "\n\nSkipped ITEMCODE(s):\n" + [...new Set(skippedDuplicates)].join(", ");
  }

  ui.alert(message);
}

function createHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const normalized = normalizeHeader_(header);
    if (normalized) {
      map[normalized] = index;
    }
  });
  return map;
}

function findHeaderIndex_(headerMap, possibleNames) {
  for (let i = 0; i < possibleNames.length; i++) {
    const normalized = normalizeHeader_(possibleNames[i]);
    if (normalized in headerMap) {
      return headerMap[normalized];
    }
  }
  return -1;
}

function normalizeHeader_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function cleanValue_(value) {
  return String(value == null ? "" : value).trim();
}

function isBlank_(value) {
  return value === null || value === "" || String(value).trim() === "";
}
