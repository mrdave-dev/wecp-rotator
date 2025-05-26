function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu")
    .addItem("Show Alert", "showAlert")
    .addToUi();
}

function showAlert() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Hello from TypeScript!");
}

