var express = require("express");
var router = express.Router();
var xl = require("excel4node");
const createFile = require("../services/createFile");
const encode = require("nodejs-base64-encode");

router.post("/plantillaInventario", function (req, res, next) {
  let jsonDatos = JSON.parse(encode.decode(req.body.datos, "base64"));
  let titulo = encode.decode(req.body.titulo, "base64");
  let datos = jsonDatos.data;

  let opcionesDropdown = [jsonDatos.opciones[0].dato];

  let dropdown = {
    type: "list",
    allowBlank: false,
    prompt: "Seleccione",
    error: "Opción no válida",
    showDropDown: true,
    formulas: opcionesDropdown,
  };

  var wb = new xl.Workbook({
    workbookView: {
      activeTab: 2, // Specifies an unsignedInt that contains the index to the active sheet in this book view.
    },
    author: "HyG consultores.", // Name for use in features such as comments
  });
  var options = {
    sheetProtection: {
      deleteColumns: false,
      deleteRows: false,
      formatCells: false,
      formatColumns: false,
      insertColumns: true,
      insertHyperlinks: false,
      insertRows: false,
    },
  };

  var style = wb.createStyle({
    font: {
      color: "#000000",
      size: 12,
    },
  });

  let keys = Object.keys(datos[0]);

  var ws;
  var ws2;
  if (titulo === "plantilla") {
    ws = wb.addWorksheet("Plantilla");
    ws2 = wb.addWorksheet("Ejemplo", options);

    dropdown["sqref"] = "A2:A100";
    ws.addDataValidation(dropdown);

    for (let i = 0; i < keys.length; i++) {
      ws.cell(1, i + 1)
        .string(keys[i].toUpperCase().replace(/_/g, " "))
        .style(style);
    }
  } else {
    ws2 = wb.addWorksheet("Detalle inventario");
    dropdown["sqref"] = `A2:A${datos.length + 20}`;
    ws2.addDataValidation(dropdown);
  }

  ws2 = createFile(ws2, datos, keys, style);

  wb.write("Plantilla_inventario.xlsx", res);
});

router.post("/generarReporte", function (req, res, next) {
  let origen = req.body.origen;
  let jsonDatos = JSON.parse(encode.decode(req.body.data, "base64"));

  var wb = new xl.Workbook({
    author: "HyG consultores.", // Name for use in features such as comments
  });
  var style = wb.createStyle({
    font: {
      color: "#000000",
      size: 12,
    },
  });

  let keys = Object.keys(jsonDatos[0]);

  var ws = wb.addWorksheet("reporte");

  ws = createFile(ws, jsonDatos, keys, style);

  wb.write(`Reporte_${origen}.xlsx`, res);
});

module.exports = router;
