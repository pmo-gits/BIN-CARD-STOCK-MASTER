function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Update Script")
    .addItem("Update Opening Stock", "updateOpeningStock")
    .addSeparator()
    .addItem("Update Received Ledger", "updateReceivedLedger")
    .addSeparator()
    .addItem("Update Issued Ledger", "updateIssuedLedger")
    .addToUi();
}
