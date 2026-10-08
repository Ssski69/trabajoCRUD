// ============================================================
// >>> AGREGAR CAMPO NUEVO (lugar 1 de 5)
// Agrega el nombre de la columna al final de la lista.
// Estos nombres se escriben como encabezados en la fila 1.
// ============================================================
var ENCABEZADOS = ["ID", "Fecha y Hora", "Nombre", "Edad", "Cargo"];

function obtenerHoja() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(ENCABEZADOS);
  }
  return hoja;
}

function responder(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// CONSULTAR: devuelve todos los registros
// ------------------------------------------------------------
function doGet() {
  var datos = obtenerHoja().getDataRange().getValues();
  datos.shift(); // quita los encabezados

  var registros = datos.map(function (fila) {
    return {
      id: fila[0],
      fecha: fila[1],
      nombre: fila[2],
      edad: fila[3],
      cargo: fila[4]
      // >>> AGREGAR CAMPO NUEVO (lugar 2 de 5)
      // Agrega una coma después de "cargo" y escribe el nuevo campo
      // con el siguiente número de fila. Ejemplo:
      // cargo: fila[4],
      // telefono: fila[5]
    };
  });

  return responder(registros);
}

// ------------------------------------------------------------
// GUARDAR, EDITAR o ELIMINAR
// ------------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    var p = e.parameter;
    var hoja = obtenerHoja();

    // ---------- GUARDAR ----------
    if (p.accion === "guardar") {
      // >>> AGREGAR CAMPO NUEVO (lugar 3 de 5)
      // Agrega el campo al final, en el mismo orden que ENCABEZADOS.
      // Ejemplo: [Date.now(), new Date(), p.nombre, p.edad, p.cargo, p.telefono]
      hoja.appendRow([Date.now(), new Date(), p.nombre, p.edad, p.cargo]);
      return responder({ ok: true });
    }

    var fila = buscarFila(hoja, p.id);
    if (fila === -1) {
      return responder({ ok: false, mensaje: "No se encontró el registro" });
    }

    // ---------- EDITAR ----------
    if (p.accion === "editar") {
      // >>> AGREGAR CAMPO NUEVO (lugar 4 de 5)
      // a) Agrega el campo dentro de setValues.
      // b) Cambia el ÚLTIMO número del getRange por la cantidad de
      //    campos que hay dentro de setValues.
      //
      // Ahora son 3 campos (nombre, edad, cargo)  ->  getRange(fila, 3, 1, 3)
      // Con 4 campos (agregando telefono)          ->  getRange(fila, 3, 1, 4)
      hoja.getRange(fila, 3, 1, 3).setValues([[p.nombre, p.edad, p.cargo]]);
      return responder({ ok: true });
    }

    // ---------- ELIMINAR ----------
    if (p.accion === "eliminar") {
      hoja.deleteRow(fila);
      return responder({ ok: true });
    }

    return responder({ ok: false, mensaje: "Acción no válida" });

  } catch (error) {
    return responder({ ok: false, mensaje: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// Busca el número de fila que tiene ese ID (-1 si no existe)
// ------------------------------------------------------------
function buscarFila(hoja, id) {
  var ids = hoja.getRange(1, 1, hoja.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      return i + 1; // +1 porque las filas empiezan en 1
    }
  }
  return -1;
}

// ============================================================
// >>> AGREGAR CAMPO NUEVO (lugar 5 de 5)
// Escribe el encabezado en la fila 1 de la hoja de cálculo,
// en la columna que sigue (Cargo está en E, el siguiente iría en F).
//
// DESPUÉS DE CAMBIAR ESTE ARCHIVO:
// Implementar > Administrar implementaciones > Editar (lápiz)
// > Versión: "Nueva versión" > Implementar
// ============================================================