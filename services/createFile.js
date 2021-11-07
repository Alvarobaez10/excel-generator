module.exports = function createFile(sheet, jsonDatos, keys, style) {
  let row = 1;
  for (let i = 0; i < keys.length; i++) {
    sheet
      .cell(row, i + 1)
      .string(keys[i].toUpperCase().replace(/_/g, " "))
      .style(style);
  }

  for (let item of jsonDatos) {
    row++;
    for (let j = 0; j < keys.length; j++) {
      let valor = item[keys[j]] === null ? "" : item[keys[j]];
      sheet.cell(row, j + 1).string(valor);
    }
  }
  return sheet;
};
