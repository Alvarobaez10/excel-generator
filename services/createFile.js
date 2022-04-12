module.exports = function createFile(
  sheet,
  jsonDatos,
  keys,
  style,
  header,
  styleHeader
) {
  let row = 1;
  if (header) {
    sheet
      .cell(row, 1, row, keys.length, true)
      .string(header.toUpperCase())
      .style(styleHeader);
    row++;
  }

  for (let i = 0; i < keys.length; i++) {
    sheet
      .cell(row, i + 1)
      .string(keys[i].toUpperCase().replace(/_/g, " "))
      .style(styleHeader);
  }

  for (let item of jsonDatos) {
    row++;
    for (let j = 0; j < keys.length; j++) {
      let valor = item[keys[j]] === null ? "" : item[keys[j]];
      if (typeof valor === "number") {
        sheet.cell(row, j + 1).number(valor);
      } else if (typeof valor === "object" && valor instanceof Date) {
        let dia = valor.getDate();
        let mes = valor.getMonth() + 1;
        let anio = valor.getFullYear();
        sheet.cell(row, j + 1).string(`${dia}/${mes}/${anio}`);
      } else {
        sheet.cell(row, j + 1).string(valor.toString());
      }
    }
  }
  return sheet;
};
