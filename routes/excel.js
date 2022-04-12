var express = require("express");
var router = express.Router();
var xl = require("excel4node");
const createFile = require("../services/createFile");
const encode = require("nodejs-base64-encode");

router.post("/plantillaInventario", function (req, res, next) {
  let jsonDatos = JSON.parse(encode.decode(req.body.datos, "base64")); // jsonDatos es un array de objetos con la estructura para la generación de la plantilla
  let titulo = encode.decode(req.body.titulo, "base64"); //titulo de la plantilla
  let styleHeader = {}; // estilo de la cabecera - opcional

  let datos = jsonDatos.data;

  var wb = new xl.Workbook({
    workbookView: {
      activeTab: 2, // Specifies an unsignedInt that contains the index to the active sheet in this book view.
    },
    author: "HyG consultores.", // Name for use in features such as comments
  });

  if (req.body.styleHeader) {
    styleHeader = wb.createStyle(
      JSON.parse(encode.decode(req.body.data, "base64"))
    );
  }

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

    for (let i = 0; i < jsonDatos.opciones.length; i++) {
      let opcionesDropdown = [jsonDatos.opciones[i].dato];
      let dropdown = {
        type: "list",
        allowBlank: false,
        prompt: "Seleccione",
        error: "Opción no válida",
        showDropDown: true,
        formulas: opcionesDropdown,
      };

      dropdown["sqref"] = jsonDatos.opciones[i].rango;
      ws.addDataValidation(dropdown);
    }

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

  ws2 = createFile(ws2, datos, keys, style, "", styleHeader);

  wb.write("Plantilla_inventario.xlsx", res);
});

router.post("/generarReporte", function (req, res, next) {
  let origen = req.body.origen;
  let header = req.body.title;
  let jsonDatos = JSON.parse(encode.decode(req.body.data, "base64"));
  let styleHeader = {};
  var wb = new xl.Workbook({
    author: "HyG consultores.", // Name for use in features such as comments
  });
  let style = wb.createStyle({
    font: {
      color: "#000000",
      size: 12,
    },
  });

  if (req.body.styleHeader) {
    styleHeader = wb.createStyle(
      JSON.parse(encode.decode(req.body.styleHeader, "base64"))
    );
  }

  let keys = Object.keys(jsonDatos[0]);

  var ws = wb.addWorksheet("reporte");

  ws = createFile(ws, jsonDatos, keys, style, header, styleHeader);

  wb.write(`Reporte_${origen}.xlsx`, res);
});

router.post("/test", function (req, res, next) {
  res.send("Ok");
});

module.exports = router;
