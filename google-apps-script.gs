// ============================================================================
// GOOGLE APPS SCRIPT - STORE POS & CONCILIACIÓN DE INVENTARIOS
// Copiar este código en la ventana de Google Apps Script (Extensiones -> Apps Script)
// ============================================================================

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (e && e.parameter && e.parameter.obtener === "ventas_TEC") {
    var hoja = ss.getSheetByName("Ventas_TEC");
    if (!hoja) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var datos = hoja.getDataRange().getValues();
    var arr = [];
    for (var j = 1; j < datos.length; j++) {
      if (datos[j][0]) arr.push({ fecha: Utilities.formatDate(new Date(datos[j][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), producto: datos[j][2], cantidad: datos[j][3], total: datos[j][5], vendedor: datos[j][6] || '', tipoPago: datos[j][7] || '' });
    }
    return ContentService.createTextOutput(JSON.stringify(arr)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.obtener === "ventas_MILENIO") {
    var hoja = ss.getSheetByName("Ventas_MILENIO");
    if (!hoja) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var datos = hoja.getDataRange().getValues();
    var arr = [];
    for (var j = 1; j < datos.length; j++) {
      if (datos[j][0]) arr.push({ fecha: Utilities.formatDate(new Date(datos[j][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), producto: datos[j][2], cantidad: datos[j][3], total: datos[j][5], vendedor: datos[j][6] || '', tipoPago: datos[j][7] || '' });
    }
    return ContentService.createTextOutput(JSON.stringify(arr)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.obtener === "devoluciones_TEC") {
    var hoja = ss.getSheetByName("Devoluciones_TEC");
    if (!hoja) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var datos = hoja.getDataRange().getValues();
    var arr = [];
    for (var j = 1; j < datos.length; j++) {
      if (datos[j][0]) arr.push({ fecha: Utilities.formatDate(new Date(datos[j][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), producto: datos[j][2], cantidad: datos[j][3], total: datos[j][4], motivo: datos[j][5] || '', vendedor: datos[j][6] || '' });
    }
    return ContentService.createTextOutput(JSON.stringify(arr)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.obtener === "devoluciones_MILENIO") {
    var hoja = ss.getSheetByName("Devoluciones_MILENIO");
    if (!hoja) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var datos = hoja.getDataRange().getValues();
    var arr = [];
    for (var j = 1; j < datos.length; j++) {
      if (datos[j][0]) arr.push({ fecha: Utilities.formatDate(new Date(datos[j][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), producto: datos[j][2], cantidad: datos[j][3], total: datos[j][4], motivo: datos[j][5] || '', vendedor: datos[j][6] || '' });
    }
    return ContentService.createTextOutput(JSON.stringify(arr)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.obtener === "historial_inventario") {
    var hoja = ss.getSheetByName("Historial_Inventario");
    if (!hoja) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var datos = hoja.getDataRange().getValues();
    var arr = [];
    for (var j = 1; j < datos.length; j++) {
      if (datos[j][0]) arr.push({ fecha: Utilities.formatDate(new Date(datos[j][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), id: datos[j][1], descripcion: datos[j][2], accion: datos[j][3], stockAnterior: datos[j][4], stockNuevo: datos[j][5], precioAnterior: datos[j][6], precioNuevo: datos[j][7], vendedor: datos[j][8] || '' });
    }
    return ContentService.createTextOutput(JSON.stringify(arr)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.obtener === "archivos") {
    var hojas = ss.getSheets();
    var archivos = [];
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      if ((nombre.indexOf("Ventas_") === 0 || nombre.indexOf("Devoluciones_") === 0) && nombre.indexOf("_20") > -1) {
        archivos.push(nombre);
      }
    }
    return ContentService.createTextOutput(JSON.stringify(archivos)).setMimeType(ContentService.MimeType.JSON);
  }
  
  var hojaProd = ss.getSheetByName("Productos");
  if (!hojaProd) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
  var datos = hojaProd.getDataRange().getValues();
  var productos = [];
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0]) {
      productos.push({ 
        id: String(datos[i][0]).trim(), 
        descripcion: datos[i][1], 
        stock: Number(datos[i][2]), 
        precio: Number(datos[i][3]), 
        stockMinimo: Number(datos[i][4]) || 5, 
        categoria: datos[i][5] || 'TEC',
        fisico: datos[i][6] !== undefined ? Number(datos[i][6]) : null
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(productos)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var peticion = JSON.parse(e.postData.contents);
  var accion = peticion.accion;
  var fechaActual = new Date();
  
  if (accion === "cobrar") {
    var carrito = peticion.datos;
    var vendedor = peticion.vendedor || '';
    var tipoPago = peticion.tipoPago || '';
    var hojaProd = ss.getSheetByName("Productos");
    var datosProd = hojaProd.getDataRange().getValues();
    
    carrito.forEach(function(item) {
      var cat = item.categoria || 'TEC';
      var hojaVentas = ss.getSheetByName("Ventas_" + cat);
      if (!hojaVentas) return;
      for (var k = 0; k < item.cantidad; k++) {
        hojaVentas.appendRow([fechaActual, item.id, item.descripcion, 1, item.precio, item.precio, vendedor, tipoPago]);
      }
      for (var i = 1; i < datosProd.length; i++) {
        if (String(datosProd[i][0]).trim() === String(item.id).trim()) {
          var stockActual = Number(datosProd[i][2]);
          hojaProd.getRange(i + 1, 3).setValue(Math.max(0, stockActual - Number(item.cantidad)));
          break;
        }
      }
    });
    
    if (peticion.devoluciones && peticion.devoluciones.length > 0) {
      peticion.devoluciones.forEach(function(item) {
        var cat = item.categoria || 'TEC';
        var hojaDev = ss.getSheetByName("Devoluciones_" + cat);
        if (!hojaDev) return;
        for (var k = 0; k < item.cantidad; k++) {
          hojaDev.appendRow([fechaActual, item.id, item.descripcion, 1, item.precio, item.motivo || 'Cambio', vendedor]);
        }
        for (var i = 1; i < datosProd.length; i++) {
          if (String(datosProd[i][0]).trim() === String(item.id).trim()) {
            var stockActual = Number(datosProd[i][2]);
            hojaProd.getRange(i + 1, 3).setValue(stockActual + Number(item.cantidad));
            break;
          }
        }
      });
    }
    return ContentService.createTextOutput(JSON.stringify({status: "OK"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (accion === "devolver") {
    var carrito = peticion.datos;
    var vendedor = peticion.vendedor || '';
    var motivo = peticion.motivo || 'Cambio';
    var hojaProd = ss.getSheetByName("Productos");
    var datosProd = hojaProd.getDataRange().getValues();
    
    carrito.forEach(function(item) {
      var cat = item.categoria || 'TEC';
      var hojaDev = ss.getSheetByName("Devoluciones_" + cat);
      if (!hojaDev) return;
      for (var k = 0; k < item.cantidad; k++) {
        hojaDev.appendRow([fechaActual, item.id, item.descripcion, 1, item.precio, motivo, vendedor]);
      }
      for (var i = 1; i < datosProd.length; i++) {
        if (String(datosProd[i][0]).trim() === String(item.id).trim()) {
          var stockActual = Number(datosProd[i][2]);
          hojaProd.getRange(i + 1, 3).setValue(stockActual + Number(item.cantidad));
          break;
        }
      }
    });
    return ContentService.createTextOutput(JSON.stringify({status: "OK"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (accion === "guardar_producto") {
    var prod = peticion.datos;
    var registradoPor = peticion.registradoPor || '';
    var hojaProd = ss.getSheetByName("Productos");
    var hojaHist = ss.getSheetByName("Historial_Inventario");
    var datosProd = hojaProd.getDataRange().getValues();
    var encontrado = false;
    
    for (var i = 1; i < datosProd.length; i++) {
      if (String(datosProd[i][0]).trim() === String(prod.id).trim()) {
        var stockAnt = Number(datosProd[i][2]);
        var precioAnt = Number(datosProd[i][3]);
        hojaProd.getRange(i + 1, 2).setValue(prod.descripcion);
        hojaProd.getRange(i + 1, 3).setValue(Number(prod.stock));
        hojaProd.getRange(i + 1, 4).setValue(Number(prod.precio));
        hojaProd.getRange(i + 1, 6).setValue(prod.categoria || 'TEC');
        if (hojaHist) {
          var accionHist = 'Actualización';
          if (stockAnt !== Number(prod.stock)) accionHist = 'Suma stock';
          if (precioAnt !== Number(prod.precio)) accionHist = stockAnt !== Number(prod.stock) ? 'Cambio precio y stock' : 'Cambio precio';
          hojaHist.appendRow([fechaActual, prod.id, prod.descripcion, accionHist, stockAnt, Number(prod.stock), precioAnt, Number(prod.precio), registradoPor]);
        }
        encontrado = true;
        break;
      }
    }
    
    if (!encontrado) {
      hojaProd.appendRow([prod.id, prod.descripcion, Number(prod.stock), Number(prod.precio), 5, prod.categoria || 'TEC']);
      if (hojaHist) hojaHist.appendRow([fechaActual, prod.id, prod.descripcion, 'Alta', 0, Number(prod.stock), 0, Number(prod.precio), registradoPor]);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "OK"})).setMimeType(ContentService.MimeType.JSON);
  }

  // GUARDAR BORRADOR TEMPORAL DE CONTEO EN HOJA
  if (accion === "guardar_conteo_temp") {
    var datosConteo = peticion.datos;
    var hojaProd = ss.getSheetByName("Productos");
    if (hojaProd) {
      var datosProd = hojaProd.getDataRange().getValues();
      for (var i = 1; i < datosProd.length; i++) {
        if (String(datosProd[i][0]).trim() === String(datosConteo.id).trim()) {
          hojaProd.getRange(i + 1, 7).setValue(Number(datosConteo.fisico));
          break;
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "OK"})).setMimeType(ContentService.MimeType.JSON);
  }

  // APLICAR CONCILIACIÓN FINAL DE INVENTARIOS
  if (accion === "conciliar_inventario") {
    var ajustes = peticion.datos;
    var registradoPor = peticion.registradoPor || '';
    var motivo = peticion.motivo || 'Ajuste Conciliación';
    var hojaProd = ss.getSheetByName("Productos");
    var hojaHist = ss.getSheetByName("Historial_Inventario");
    var datosProd = hojaProd.getDataRange().getValues();

    ajustes.forEach(function(item) {
      for (var i = 1; i < datosProd.length; i++) {
        if (String(datosProd[i][0]).trim() === String(item.id).trim()) {
          var stockAnt = Number(datosProd[i][2]);
          var precioAnt = Number(datosProd[i][3]);
          var stockNvo = Number(item.stockNuevo);
          
          // Reemplazar stock real
          hojaProd.getRange(i + 1, 3).setValue(stockNvo);
          // Limpiar columna Fisico
          hojaProd.getRange(i + 1, 7).setValue(stockNvo);

          // Registrar en historial auditoria
          if (hojaHist) {
            var itemMotivo = item.motivo || motivo || 'Ajuste Conciliación';
            var difStr = (stockNvo - stockAnt) > 0 ? ('+' + (stockNvo - stockAnt)) : String(stockNvo - stockAnt);
            var accionStr = 'Conciliación (' + itemMotivo + ': ' + difStr + ' pzas)';
            hojaHist.appendRow([fechaActual, item.id, item.descripcion, accionStr, stockAnt, stockNvo, precioAnt, precioAnt, registradoPor]);
          }
          break;
        }
      }
    });

    return ContentService.createTextOutput(JSON.stringify({status: "OK"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (accion === "descargar_actual") {
    var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
    return ContentService.createTextOutput(JSON.stringify({url: url})).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("Acción no válida").setMimeType(ContentService.MimeType.JSON);
}

function doOptions() {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}
