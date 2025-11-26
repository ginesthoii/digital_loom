function colorDMC() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7); 
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const r = values[i][2]; // column C
    const g = values[i][3]; // column D
    const b = values[i][4]; // column E

    if (!isNaN(r) && !isNaN(g) && !isNaN(b)) {
      const hex = rgbToHex(r, g, b);
      sheet.getRange(i + 2, 7).setBackground(hex); // column G
    }
  }
}

function rgbToHex(r, g, b) {
  return "#" +
    toHex(r) +
    toHex(g) +
    toHex(b);
}

function toHex(n) {
  n = parseInt(n, 10);
  if (isNaN(n)) return "00";
  return ("0" + n.toString(16)).slice(-2);
}
