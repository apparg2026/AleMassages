function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Churrería POS Inteligente')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, minimal-ui');
}

// --- BASE DE DATOS INICIAL ---
function iniciarSistema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName("Usuarios")) {
    var us = ss.insertSheet("Usuarios"); us.appendRow(["ID", "Usuario", "Password", "Rol"]);
    us.appendRow([Utilities.getUuid(), "admin", "1234", "Administrador"]);
  }
  
  if (!ss.getSheetByName("CierresDiarios")) {
    ss.insertSheet("CierresDiarios").appendRow(["ID", "Fecha", "Efectivo", "Tarjeta", "QR", "Total_Ingresos", "Mes", "Año", "Usuario"]);
  }
  
  if (!ss.getSheetByName("Movimientos")) {
    ss.insertSheet("Movimientos").appendRow(["ID", "Fecha", "Tipo", "Empleado_Detalle", "Monto", "Medio_Pago", "Observaciones", "Usuario"]);
  }
  
  if (!ss.getSheetByName("Empleados")) {
    var emps = ss.insertSheet("Empleados"); emps.appendRow(["ID", "Nombre"]);
    emps.appendRow([Utilities.getUuid(), "Cintia"]);
    emps.appendRow([Utilities.getUuid(), "Tamara"]);
    emps.appendRow([Utilities.getUuid(), "Johana"]);
    emps.appendRow([Utilities.getUuid(), "Ruben"]);
  }
}

function verificarLogin(usuario, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetUsr = ss.getSheetByName("Usuarios");
    if (!sheetUsr) { iniciarSistema(); sheetUsr = ss.getSheetByName("Usuarios"); }
    var data = sheetUsr.getDataRange().getValues().slice(1);
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(usuario).trim() && String(data[i][2]).trim() === String(password).trim()) {
        return { success: true, rol: String(data[i][3]), usuario: String(data[i][1]) };
      }
    }
    return { success: false, error: "Credenciales incorrectas" };
  } catch(e) { return { success: false, error: "Error de servidor" }; }
}

function obtenerDatos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var getSafely = function(name) { var sh = ss.getSheetByName(name); return sh ? sh.getDataRange().getValues().slice(1) : []; };

  var empleados = getSafely("Empleados").map(function(r){ return {id: String(r[0]), nombre: String(r[1])}; }).filter(function(e){ return e.nombre; });
  var usuarios = getSafely("Usuarios").map(function(r){ return {id: String(r[0]), usuario: String(r[1]), rol: String(r[3])}; }).filter(function(u){ return u.usuario; });
  
  return { empleados: empleados, usuarios: usuarios };
}

// --- REGISTROS DE CAJA Y GASTOS ---
function registrarCierreDiario(efectivo, tarjeta, qr, usr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("CierresDiarios") || ss.insertSheet("CierresDiarios");
  
  var e = parseFloat(efectivo) || 0; var t = parseFloat(tarjeta) || 0; var q = parseFloat(qr) || 0; var total = e + t + q;
  if(total <= 0) return "El total no puede ser 0.";

  var fechaObj = new Date();
  sh.appendRow([Utilities.getUuid(), fechaObj.toISOString(), e, t, q, total, fechaObj.getMonth()+1, fechaObj.getFullYear(), usr]);
  return "¡Cierre de caja guardado con éxito! 💰";
}

function registrarMovimiento(tipo, persona, monto, medio, obs, usr) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Movimientos") || ss.insertSheet("Movimientos");
  sh.appendRow([Utilities.getUuid(), new Date().toISOString(), tipo, persona, monto, medio, obs, usr]); 
  return "Registro guardado correctamente 📥"; 
}

// --- HISTORIAL DE OPERACIONES ---
function obtenerHistorialMovimientos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Movimientos");
    if(!sh) return JSON.stringify([]);
    var data = sh.getDataRange().getValues();
    if(data.length <= 1) return JSON.stringify([]);
    
    var rows = data.slice(1);
    var limit = Math.min(rows.length, 30);
    var recent = rows.slice(rows.length - limit).reverse();
    
    var tz = Session.getScriptTimeZone();
    var res = recent.map(function(r) {
      var d = new Date(r[1]);
      var fStr = isNaN(d.getTime()) ? r[1] : Utilities.formatDate(d, tz, "dd/MM HH:mm");
      return { fecha: fStr, tipo: String(r[2]), persona: String(r[3]), monto: Number(r[4]), medio: String(r[5]), obs: String(r[6]) };
    });
    
    return JSON.stringify(res);
  } catch(e) { return JSON.stringify([]); }
}

// --- HISTORIAL Y EDICIÓN DE CIERRES DIARIOS ---
function obtenerHistorialCierres() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("CierresDiarios");
    if(!sh) return JSON.stringify([]);
    var data = sh.getDataRange().getValues();
    if(data.length <= 1) return JSON.stringify([]);
    
    var rows = data.slice(1);
    var limit = Math.min(rows.length, 30);
    var recent = rows.slice(rows.length - limit).reverse();
    
    var tz = Session.getScriptTimeZone();
    var res = recent.map(function(r) {
      var d = new Date(r[1]);
      var fStr = isNaN(d.getTime()) ? r[1] : Utilities.formatDate(d, tz, "dd/MM/yyyy");
      return { id: String(r[0]), fecha: fStr, efectivo: Number(r[2]), tarjeta: Number(r[3]), qr: Number(r[4]), total: Number(r[5]), usr: String(r[8]) };
    });
    return JSON.stringify(res);
  } catch(e) { return JSON.stringify([]); }
}

function editarCierreDiario(id, efec, tarj, qr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CierresDiarios");
    if(!sheet) return "Error: Hoja no encontrada";
    var data = sheet.getDataRange().getValues();
    for(var i = 1; i < data.length; i++) {
      if(String(data[i][0]) === String(id)) {
        var e = parseFloat(efec) || 0;
        var t = parseFloat(tarj) || 0;
        var q = parseFloat(qr) || 0;
        var tot = e + t + q;
        
        sheet.getRange(i+1, 3).setValue(e);
        sheet.getRange(i+1, 4).setValue(t);
        sheet.getRange(i+1, 5).setValue(q);
        sheet.getRange(i+1, 6).setValue(tot);
        return "Cierre actualizado correctamente ✔️";
      }
    }
    return "No se encontró el registro.";
  } catch(e) { return "Error: " + e.message; }
}

// --- CRUD EMPLEADOS Y USUARIOS ---
function guardarEmpleado(id, nombre) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empleados");
  if (id && sheet) { var data = sheet.getDataRange().getValues(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i+1, 2).setValue(nombre); return "Personal actualizado ✔️"; } }
  } else { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empleados").appendRow([Utilities.getUuid(), nombre]); return "Personal agregado 👤"; }
}
function eliminarEmpleado(id) { 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empleados"); var data = sheet.getDataRange().getValues(); 
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return "Personal eliminado 🗑️"; } } 
}
function guardarUsuario(id, user, pass, rol) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (id && sheet) { var data = sheet.getDataRange().getValues(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i+1, 2).setValue(user); if(pass) sheet.getRange(i+1, 3).setValue(pass); sheet.getRange(i+1, 4).setValue(rol); return "Usuario actualizado ✏️"; } }
  } else { sheet.appendRow([Utilities.getUuid(), user, pass, rol]); return "Usuario creado 👤"; }
}
function eliminarUsuario(id) { 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios"); var data = sheet.getDataRange().getValues(); 
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { if (String(data[i][1]) === "admin") return "Error"; sheet.deleteRow(i+1); return "Usuario eliminado 🗑️"; } } 
}

// --- ESTADÍSTICAS E IA FINANCIERA (BLINDADO) ---
function obtenerEstadisticas(climaObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();
    var getRows = function(name) { var sh = ss.getSheetByName(name); return sh ? sh.getDataRange().getValues().slice(1) : []; };

    var cierres = getRows("CierresDiarios");
    var movimientos = getRows("Movimientos");
    
    var hoyObj = new Date(); var mes = hoyObj.getMonth() + 1; var anio = hoyObj.getFullYear();
    var diaStr = Utilities.formatDate(hoyObj, tz, "dd/MM/yyyy");
    
    var vHoy = 0, eHoy = 0, tHoy = 0, qHoy = 0;
    var ventasMes = 0, gastosHoy = 0;
    var eMes = 0, tMes = 0, qMes = 0;
    var ventasDia = {};
    var diasSemana = {"Domingo":0, "Lunes":0, "Martes":0, "Miércoles":0, "Jueves":0, "Viernes":0, "Sábado":0};
    var nombresDias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

    cierres.forEach(function(c) {
      if(!c[1] || String(c[1]).trim() === "") return;
      var f = new Date(c[1]); if(isNaN(f.getTime())) return;
      var fStr = Utilities.formatDate(f, tz, "dd/MM/yyyy");
      
      var efec = parseFloat(c[2]) || 0; var tarj = parseFloat(c[3]) || 0; var qr = parseFloat(c[4]) || 0; var tot = parseFloat(c[5]) || 0;

      if (fStr === diaStr) { vHoy += tot; eHoy += efec; tHoy += tarj; qHoy += qr; }
      if ((f.getMonth() + 1) === mes && f.getFullYear() === anio) {
        ventasMes += tot; eMes += efec; tMes += tarj; qMes += qr;
        var d = f.getDate(); ventasDia[d] = (ventasDia[d] || 0) + tot;
        diasSemana[nombresDias[f.getDay()]] += tot;
      }
    });

    movimientos.forEach(function(m) {
      if(!m[1] || String(m[1]).trim() === "") return;
      var f = new Date(m[1]); if(isNaN(f.getTime())) return;
      
      // Los adelantos o gastos descuentan del Efectivo SI SE PAGARON EN EFECTIVO.
      // Las compras "A Cuenta" no tocan el cajón, son deudas anotadas.
      if(Utilities.formatDate(f, tz, "dd/MM/yyyy") === diaStr) {
        if((String(m[2]) === "Gasto" || String(m[2]) === "Adelanto") && String(m[5]) === "Efectivo") {
          gastosHoy += (parseFloat(m[4]) || 0);
        }
      }
    });

    var hDias = [], hMontos = [];
    for (var dia in ventasDia) { hDias.push(String(dia) + "/" + mes); hMontos.push(Number(ventasDia[dia])); }
    if(hDias.length === 0) { hDias = ["Hoy"]; hMontos = [Number(vHoy)]; }

    var dMax = "Sin registros", dMaxVal = 0, dMin = "Sin registros", dMinVal = Infinity;
    for(var ds in diasSemana) {
      if(diasSemana[ds] > dMaxVal) { dMaxVal = diasSemana[ds]; dMax = ds; }
      if(diasSemana[ds] < dMinVal && diasSemana[ds] > 0) { dMinVal = diasSemana[ds]; dMin = ds; }
    }

    var diasPasados = hoyObj.getDate() || 1;
    var proyeccionMes = (ventasMes / diasPasados) * 30;
    var metaObj = Math.max(500000, Math.ceil(proyeccionMes / 500000) * 500000); 
    var pctMeta = Math.min(100, Math.round((ventasMes / metaObj) * 100));

    var digPct = ventasMes > 0 ? Math.round(((tMes + qMes) / ventasMes) * 100) : 0;
    var efecPct = ventasMes > 0 ? Math.round((eMes / ventasMes) * 100) : 0;

    var climaTxt = "Clima estable los próximos días. Flujo de ventas normal.";
    if(climaObj && climaObj.arrLluvias) climaTxt = "Hay pronóstico de lluvia a corto plazo. La venta de panificados aumenta con la humedad. Preparate.";
    else if(climaObj && climaObj.arrCalor) climaTxt = "Se viene el calor. La demanda puede caer ligeramente. Controlá la sobreproducción de masa.";

    var ia = "🧠 *ANÁLISIS FINANCIERO AL CIERRE (EZEIZA)*\n\n" +
      "📈 *SALUD DEL NEGOCIO:*\n" +
      "  • Ingresos del mes: $" + Math.round(ventasMes).toLocaleString('es-AR') + ".\n" +
      "  • Proyección a fin de mes: $" + Math.round(proyeccionMes).toLocaleString('es-AR') + ".\n\n" +
      "💳 *MEDIOS DE PAGO (MES):*\n" +
      "  • El " + digPct + "% de tu facturación entra de forma digital (QR/Tarjeta).\n" +
      "  • El " + efecPct + "% entra en Efectivo físico.\n" +
      "  • Hoy tenés $" + Math.round(eHoy - gastosHoy).toLocaleString('es-AR') + " netos de efectivo en mano tras descontar gastos.\n\n" +
      "📆 *RENDIMIENTO POR DÍAS:*\n" +
      "  • Históricamente, facturás más los días " + dMax + ".\n" +
      "  • El día de menor ingreso es el " + dMin + ".\n\n" +
      "🌦️ *ALERTA CLIMÁTICA:*\n  • " + climaTxt;

    return JSON.stringify({
      error: false, vHoy: Number(vHoy), eHoy: Number(eHoy), tHoy: Number(tHoy), qHoy: Number(qHoy), 
      ventasMes: Number(ventasMes), gastosHoy: Number(gastosHoy), netoEfectivo: Number(eHoy - gastosHoy),
      histDias: hDias, histMontos: hMontos, 
      pagoLabels: ["Efectivo", "Tarjeta", "Virtual / QR"], pagoData: [eMes, tMes, qMes],
      reporteIA: String(ia), metaMensual: Number(metaObj), metaPct: Number(pctMeta)
    });

  } catch(e) {
    return JSON.stringify({
      error: true, reporteIA: "⚠️ Ocurrió un error analizando los datos financieros. Error: " + e.message,
      vHoy: 0, eHoy:0, tHoy:0, qHoy:0, ventasMes: 0, gastosHoy: 0, netoEfectivo: 0,
      histDias: ["Hoy"], histMontos: [0], pagoLabels: ["Sin datos"], pagoData: [1],
      metaMensual: 1000000, metaPct: 0
    });
  }
}
