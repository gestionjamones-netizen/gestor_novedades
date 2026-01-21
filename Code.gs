function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Gestión Operativa')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- CONFIGURACIÓN DE HOJAS Y COLUMNAS ---
const SHEET_BD = "BD PERSONAL";
const SHEET_HIST = "Historico";
const SHEET_CONF = "Configuracion";
const SHEET_ATAJOS = "Config_Atajos";

const COL_BD_CONTRATO = 0; 
const COL_BD_REGISTRO = 1;
const COL_BD_NOMBRE = 3;
const COL_BD_TURNO = 4;
const COL_BD_PUESTO = 5;   

const COL_HIST_ID = 0;
const COL_HIST_REGISTRO = 1;
const COL_HIST_TIPO = 2;
const COL_HIST_CONCEPTO = 3;
const COL_HIST_VALOR = 4;
const COL_HIST_FECHA = 5;
const COL_HIST_SEM_ANO = 6;
const COL_HIST_OBS = 7;
const COL_HIST_TIMESTAMP = 8;
const COL_HIST_BATCH = 9;

// --- CONFIGURACIÓN CENTRALIZADA ---
const PUESTO_ORDER = [
  "Salmueras y mezclas","Embutido","Procesos termicos","Desmolde",
  "Empaque Multivac 1","Empaque Multivac 2","Empaque Rigido",
  "Empaque Flexible","Resistencia de sellado","Mesa de recorte"
];

const CONCEPT_COLORS = {
  "permiso": "#e74c3c", "pago permiso": "#2ecc71", "vacaciones": "#f39c12",
  "incapacidad": "#9b59b6", "calamidad": "#34495e", "sancion": "#7f8c8d", "default": "#95a5a6"
};

// --- API PRINCIPAL ---

function getTurnosList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_BD).getDataRange().getValues();
  data.shift();
  return [...new Set(data.map(r => r[COL_BD_TURNO]).filter(t => t))];
}

function getInitialData(turno, mes, anio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    // 1. CONFIGURACIÓN & CONCEPTOS
    const sheetConf = ss.getSheetByName(SHEET_CONF);
    const dataConf = sheetConf.getDataRange().getValues();
    dataConf.shift();
    const conceptos = dataConf.map(r => ({ tipo: r[0], nombre: r[1] }));
    
    // 1.1 ATAJOS
    let atajos = [];
    const sheetAtajos = ss.getSheetByName(SHEET_ATAJOS);
    if (sheetAtajos && sheetAtajos.getLastRow() > 1) {
      const dataAtajos = sheetAtajos.getDataRange().getValues();
      dataAtajos.shift();
      atajos = dataAtajos.map(r => ({
        nombre: r[0], concepto: r[1], valor: r[2], tipo: r[3]
      })).filter(a => a.nombre);
    }

    // 2. FESTIVOS
    const holidays = getHolidaysForClient(anio);
    
    // 3. PERSONAL
    const sheetBD = ss.getSheetByName(SHEET_BD);
    const dataBD = sheetBD.getDataRange().getValues();
    dataBD.shift();
    const personal = dataBD
      .filter(row => String(row[COL_BD_TURNO]) === String(turno))
      .map(row => ({
        contrato: String(row[COL_BD_CONTRATO]),
        registro: String(row[COL_BD_REGISTRO]),
        nombre: String(row[COL_BD_NOMBRE]),
        puesto: String(row[COL_BD_PUESTO])
      }));
      
    // 4. HISTORICO & BALANCES
    const sheetHist = ss.getSheetByName(SHEET_HIST);
    const lastRow = sheetHist.getLastRow();
    
    let novedades = [];
    let balancesPermisos = {};
    let balancesComp = {};
    let historyWorked = {};
    let historyTakenComp = {};
    
    personal.forEach(p => {
      balancesPermisos[p.registro] = 0;
      balancesComp[p.registro] = 0;
      historyWorked[p.registro] = new Set();
      historyTakenComp[p.registro] = 0;
    });
    
    let masterHolidays = new Set();
    const minYear = 2023, maxYear = new Date().getFullYear() + 1;
    for (let y = minYear; y <= maxYear; y++) {
      let hols = getColombianHolidaysServer(y);
      hols.forEach(d => masterHolidays.add(d));
      let d = new Date(y, 0, 1);
      while (d.getFullYear() === y) {
        if (d.getDay() === 0) masterHolidays.add(Utilities.formatDate(d, tz, "yyyy-MM-dd"));
        d.setDate(d.getDate() + 1);
      }
    }

    if (lastRow > 1) {
      const rawHist = sheetHist.getRange(2, 1, lastRow - 1, sheetHist.getLastColumn()).getValues();
      const listaRegistros = personal.map(p => p.registro);
      const mesInt = parseInt(mes);
      const anioInt = parseInt(anio);
      
      rawHist.forEach(row => {
        let reg = String(row[COL_HIST_REGISTRO]);
        if (!listaRegistros.includes(reg)) return;

        let rawDate = row[COL_HIST_FECHA];
        let fechaStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd") : String(rawDate).substring(0, 10);
        let [fAnio, fMes] = fechaStr.split('-').map(Number);
        
        let concepto = String(row[COL_HIST_CONCEPTO]).trim();
        let conceptoLower = concepto.toLowerCase();
    
        let tipo = String(row[COL_HIST_TIPO]);
        let valor = parseFloat(String(row[COL_HIST_VALOR]).replace(',', '.')) || 0;

        if (conceptoLower.includes("pago permiso")) balancesPermisos[reg] += valor; 
        else if (conceptoLower === "permiso") balancesPermisos[reg] -= valor; 

        if (conceptoLower.includes("compensatorio")) {
             historyTakenComp[reg] += (valor >= 1 ? valor : 1);
        }
        
        if (tipo === 'NOMINA' && masterHolidays.has(fechaStr)) {
           historyWorked[reg].add(fechaStr);
        }

        if (fAnio === anioInt && (fMes - 1) === mesInt) {
            novedades.push({
              id: row[COL_HIST_ID],
              registro: reg,
              tipo: tipo,
              concepto: concepto,
              valor: row[COL_HIST_VALOR],
              fecha: fechaStr,
              obs: row[COL_HIST_OBS],
              batchId: row[COL_HIST_BATCH] || ""
            });
        }
      });
    }

    // 5. CÁLCULO DE COMPENSATORIOS
    personal.forEach(p => {
      let reg = p.registro;
      let uniqueWorkedDates = Array.from(historyWorked[reg]); 
      let monthlyCounts = {};
      uniqueWorkedDates.forEach(dateStr => {
        let key = dateStr.substring(0, 7);
        monthlyCounts[key] = (monthlyCounts[key] || 0) + 1;
      });

      let totalEarned = 0;
      for (const [mesKey, count] of Object.entries(monthlyCounts)) {
        if (count >= 3) {
          totalEarned += (count - 2);
        }
      }
      balancesComp[reg] = totalEarned - historyTakenComp[reg];
    });

    return { 
      success: true, 
      data: { 
        personal, novedades, conceptos, atajos,
        balancesPermisos, balancesComp, holidays,
        config: { puestoOrder: PUESTO_ORDER, colors: CONCEPT_COLORS }
      } 
    };
  } catch (e) {
    return { success: false, error: e.toString() + " stack: " + e.stack };
  }
}

// --- CORE DEL SISTEMA: PROCESAMIENTO ATÓMICO (SOLUCIÓN FIX) ---

function processClientBatch(operations) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { success: false, error: "Servidor ocupado (Lock)" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_HIST);
    const tz = ss.getSpreadsheetTimeZone();

    // 1. LECTURA ÚNICA (SNAPSHOT)
    // Obtenemos todos los datos para mapear IDs a filas reales
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    // Mapas para búsqueda rápida de filas por ID y BatchID
    let idMap = new Map();
    let batchMap = new Map();

    // Empezamos en 1 para saltar headers. 
    // Data Index i corresponde a Fila hoja i + 1.
    for(let i=1; i<data.length; i++) { 
        let r = data[i];
        let id = String(r[COL_HIST_ID]);
        let batch = String(r[COL_HIST_BATCH]);
        
        if(id) idMap.set(id, i);
        if(batch) {
            if(!batchMap.has(batch)) batchMap.set(batch, []);
            batchMap.get(batch).push(i);
        }
    }

    let rowsToDeleteIndices = new Set();
    let rowsToAppend = [];
    let results = [];

    // 2. PROCESAMIENTO LÓGICO EN MEMORIA
    operations.forEach(op => {
      try {
        if (op.action === 'SAVE') {
          // Lógica de guardado: Identificar qué borrar (old) y qué agregar (new)
          let items = Array.isArray(op.data) ? op.data : [op.data];
          
          items.forEach(itemData => {
            if (!itemData || (!itemData.registro && (!itemData.registros || itemData.registros.length === 0))) return;

            // A. Identificar antiguas versiones para borrar
            // FIX: Prioridad absoluta al ID único si existe.
            if (itemData.idAntiguo && idMap.has(String(itemData.idAntiguo))) {
                rowsToDeleteIndices.add(idMap.get(String(itemData.idAntiguo)));
            } 
            // Solo si no hay ID único, usamos BatchID (para borrados masivos legacy)
            else if (itemData.batchIdAntiguo && batchMap.has(String(itemData.batchIdAntiguo))) {
                batchMap.get(String(itemData.batchIdAntiguo)).forEach(idx => rowsToDeleteIndices.add(idx));
            }

            // B. Generar nuevas filas
            let newRows = buildRowsFromItem(itemData, tz);
            rowsToAppend.push(...newRows);
          });
          
          results.push({ id: op.id, success: true });

        } else if (op.action === 'DELETE') {
          // Lógica de borrado explícito
          if (op.id && idMap.has(String(op.id))) {
             rowsToDeleteIndices.add(idMap.get(String(op.id)));
          }
          else if (op.batchId && batchMap.has(String(op.batchId))) {
             batchMap.get(String(op.batchId)).forEach(idx => rowsToDeleteIndices.add(idx));
          } 
          results.push({ id: op.id, success: true });

        } else if (op.action === 'DELETE_DAY') {
          // Lógica de limpieza de día (Escanea data en memoria)
          // Se busca coincidencia de Registro + Fecha
          for(let i=1; i<data.length; i++) {
             let r = data[i];
             // Verificación Registro
             if(String(r[COL_HIST_REGISTRO]) !== String(op.registro)) continue;
             
             // Verificación Fecha
             let rawDate = r[COL_HIST_FECHA];
             let rowFecha = (rawDate instanceof Date) ? 
                Utilities.formatDate(rawDate, tz, "yyyy-MM-dd") : 
                String(rawDate).substring(0, 10);
                
             if(rowFecha === op.fecha) {
                 rowsToDeleteIndices.add(i);
             }
          }
          results.push({ id: op.id, success: true });
        }
      } catch (innerErr) {
        results.push({ id: op.id, success: false, error: innerErr.toString() });
      }
    });

    // 3. EJECUCIÓN FÍSICA (ATÓMICA)
    
    // PASO A: Borrar filas (Orden Descendente OBLIGATORIO para no mover índices)
    let sortedIndices = Array.from(rowsToDeleteIndices).sort((a, b) => b - a);
    sortedIndices.forEach(idx => {
        // idx es índice del array data. Fila hoja es idx + 1
        sheet.deleteRow(idx + 1);
    });

    // PASO B: Agregar nuevas filas
    if(rowsToAppend.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    return { success: true, results: results };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


// --- FUNCIONES LEGACY (Wrappers para processClientBatch) ---
// Mantienen compatibilidad pero usan el nuevo motor robusto

function saveNovedad(payloadInput) {
    const res = processClientBatch([{ action: 'SAVE', data: payloadInput, id: 'legacy_save' }]);
    return res.results[0].success ? { success: true } : { success: false, error: res.results[0].error };
}

function deleteNovedad(id, batchId) {
    const res = processClientBatch([{ action: 'DELETE', id: id, batchId: batchId, id: 'legacy_del' }]);
    return res.results[0].success ? { success: true } : { success: false, error: res.results[0].error };
}

function deleteAllForDay(registro, fechaStrInput) {
    const res = processClientBatch([{ action: 'DELETE_DAY', registro: registro, fecha: fechaStrInput, id: 'legacy_del_day' }]);
    return res.results[0].success ? { success: true } : { success: false, error: res.results[0].error };
}


// --- HELPER LOGICO (SIN ACCESO A HOJA) ---

function buildRowsFromItem(itemData, tz) {
    let generatedRows = [];
    
    let diasSolicitados = 1;
    let valorGuardar = itemData.valor;
    if (valorGuardar) valorGuardar = String(valorGuardar).replace('.', ',');

    const conceptoLower = String(itemData.concepto || "").toLowerCase();
    const esVacaciones = conceptoLower === "vacaciones";
    const esPermisoHoras = ["permiso", "pago permiso"].includes(conceptoLower);

    if (itemData.tipo === 'OTRO' && !esPermisoHoras) {
      diasSolicitados = parseInt(itemData.valor) || 1; 
      valorGuardar = 1;
    } 
    
    let targetRegistros = (itemData.registros && itemData.registros.length > 0) ? itemData.registros : [itemData.registro];
    
    targetRegistros.forEach(regPersona => {
      let newBatchId = Utilities.getUuid(); 

      if(!itemData.fecha) return;
      let parts = itemData.fecha.split('-');
      // Mes es 0-indexado en constructor Date
      let fechaBase = new Date(parts[0], parts[1]-1, parts[2], 12, 0, 0);

      if (esVacaciones) {
        let diasAgregados = 0;
        let fechaCursor = new Date(fechaBase);
        let safety = 0;
  
        while (diasAgregados < diasSolicitados && safety < 60) {
          if (!isHolidayOrSundayServer(fechaCursor, tz)) {
            generatedRows.push(buildRowData(regPersona, itemData, valorGuardar, fechaCursor, tz, newBatchId));
            diasAgregados++;
          }
          fechaCursor.setDate(fechaCursor.getDate() + 1);
          safety++;
        }
      } else {
        for (let i = 0; i < diasSolicitados; i++) {
          let fechaCurrent = new Date(fechaBase);
          fechaCurrent.setDate(fechaBase.getDate() + i);
          generatedRows.push(buildRowData(regPersona, itemData, valorGuardar, fechaCurrent, tz, newBatchId));
        }
      }
    });

    return generatedRows;
}

function buildRowData(reg, payload, valor, dateObj, tz, batchId) {
  const fechaStr = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd");
  // Semana ISO
  const semana = getIsoWeek(dateObj);
  const yearIso = getIsoYear(dateObj, tz);
  const semAnoTurno = `S${semana}-${yearIso}-${payload.turno}`;
  
  return [
    Utilities.getUuid(), 
    String(reg), 
    payload.tipo, 
    payload.concepto, 
    valor, 
    fechaStr, 
    semAnoTurno, 
    payload.obs, 
    new Date(), 
    batchId
  ];
}


// --- UTILIDADES DE FECHA ---

function getIsoWeek(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNo;
}

function getIsoYear(dateObj, tz) {
  let target = new Date(dateObj.getTime());
  let dayNr = target.getDay();
  if (dayNr === 0) dayNr = 7;
  target.setDate(target.getDate() + (4 - dayNr));
  return Utilities.formatDate(target, tz, "yyyy");
}

function getHolidaysForClient(year) {
    return getColombianHolidaysServer(year);
}

function isHolidayOrSundayServer(dateObj, tz) {
  let dateStr = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd");
  let hols = getColombianHolidaysServer(dateObj.getFullYear());
  if (hols.includes(dateStr)) return true;
  if (dateObj.getDay() === 0) return true;
  return false;
}

function getColombianHolidaysServer(year) {
  const fixed = [`${year}-01-01`, `${year}-05-01`, `${year}-07-20`, `${year}-08-07`, `${year}-12-08`, `${year}-12-25`];
  const moveToNextMonday = (dateStr) => {
    let [y, m, d] = dateStr.split('-').map(Number);
    let dt = new Date(y, m-1, d);
    let day = dt.getDay();
    if (day !== 1) dt.setDate(dt.getDate() + ((1 + 7 - day) % 7));
    return Utilities.formatDate(dt, "GMT-5", "yyyy-MM-dd");
  };
  const emiliani = [`${year}-01-06`, `${year}-03-19`, `${year}-06-29`, `${year}-08-15`, `${year}-10-12`, `${year}-11-01`, `${year}-11-11`].map(moveToNextMonday);
  const a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  const easter = new Date(year, month - 1, day);
  
  const addDays = (dObj, days) => {
    let dt = new Date(dObj.valueOf());
    dt.setDate(dt.getDate() + days);
    return Utilities.formatDate(dt, "GMT-5", "yyyy-MM-dd");
  };
  return [...fixed, ...emiliani, addDays(easter, -3), addDays(easter, -2), moveToNextMonday(addDays(easter, 39)), moveToNextMonday(addDays(easter, 60)), moveToNextMonday(addDays(easter, 68))];
}

// --- GENERADOR DE REPORTE ---

function generarReporteSemana(semana, anio, turno, fechaLiqInput) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetHist = ss.getSheetByName(SHEET_HIST);
    let configData = {};
    try { configData = getConfigData(); } catch(e) { }
    const nombreArchivoBase = configData['Nombre_Archivo'] || `Plano_Nomina`;
    const nombreCarpetaDrive = configData['Nombre_Carpeta_Google_Drive'] || `Xchange_Exports`;
    const separador = configData['Separador'] || ';';
    const lenEmpCodigo = 11, lenConcepto = 4, lenValor = 10;

    let fechaLiqStr = "00000000";
    if (fechaLiqInput) fechaLiqStr = String(fechaLiqInput).replace(/-/g, "");
    
    const searchKey = `S${Number(semana)}-${anio}-${turno}`;
    const data = sheetHist.getDataRange().getValues();
    
    let lineOutput = "";
    let countLines = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowTag = String(row[COL_HIST_SEM_ANO]); 
      const rowTipo = String(row[COL_HIST_TIPO]);

      if (rowTag === searchKey && rowTipo === 'NOMINA') {
        let empCodigo = String(row[COL_HIST_REGISTRO]).trim();
        let concepto = String(row[COL_HIST_CONCEPTO]).trim(); 
        let valorRaw = row[COL_HIST_VALOR];
        let fechaNovRaw = row[COL_HIST_FECHA];
        let dcto = 1;
        
        let fechaNovStr = "";
        if (fechaNovRaw instanceof Date) {
          fechaNovStr = Utilities.formatDate(fechaNovRaw, Session.getScriptTimeZone(), "yyyyMMdd");
        } else {
           try { let p = String(fechaNovRaw).split('-');
           if(p.length===3) fechaNovStr = p[0]+p[1]+p[2]; } catch(e){}
        }

        let valorStr = String(valorRaw);
        if (empCodigo.length <= lenEmpCodigo && concepto.length <= lenConcepto && valorStr.length <= lenValor && fechaNovStr.length === 8) {
            const lineaDelimitada = [
             empCodigo, concepto, dcto, valorStr, fechaNovStr, fechaLiqStr
            ].join(separador);
            lineOutput += limpiarTextoHelper(lineaDelimitada) + separador + "\n";
            countLines++;
        }
      }
    }

    if (countLines === 0) return { success: false, error: "No se encontraron registros de NÓMINA para: " + searchKey };

    const folders = DriveApp.getFoldersByName(nombreCarpetaDrive);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(nombreCarpetaDrive);
    
    const usuario = Session.getActiveUser().getEmail().split("@")[0].toUpperCase();
    const timestamp = FormatoFechaHoraHelper();
    const fileName = `${nombreArchivoBase}_S${semana}_${turno}_${usuario}_${timestamp}.txt`;
    const file = folder.createFile(fileName, lineOutput);
    return { success: true, url: file.getUrl(), filename: fileName, count: countLines };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function limpiarTextoHelper(texto) {
  if (!texto) return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ñ/g, "n").replace(/Ñ/g, "N").replace(/[^\x00-\x7F]/g, "");
}

function FormatoFechaHoraHelper() {
  const ahora = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${ahora.getFullYear()}${pad(ahora.getMonth() + 1)}${pad(ahora.getDate())}_${pad(ahora.getHours())}${pad(ahora.getMinutes())}`;
}

// --- MÓDULO DE AUTORIZACIONES (HISTORICO2) ---

const SHEET_PENDING = "Historico2";

function getPendingNovedades(turnoFiltro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Obtener Mapa de Turnos (Registro -> Turno) desde BD
  const sheetBD = ss.getSheetByName(SHEET_BD);
  const dataBD = sheetBD.getDataRange().getValues();
  dataBD.shift();
  
  let mapTurnos = {};
  dataBD.forEach(r => {
    mapTurnos[String(r[COL_BD_REGISTRO])] = String(r[COL_BD_TURNO]); // Mapear Registro a Turno
  });

  // 2. Leer Pendientes
  const sheetPend = ss.getSheetByName(SHEET_PENDING);
  if (!sheetPend || sheetPend.getLastRow() <= 1) return [];
  
  const dataPend = sheetPend.getDataRange().getValues();
  dataPend.shift(); // Quitar headers
  
  // 3. Filtrar por Turno
  let pendientes = [];
  dataPend.forEach(r => {
    let reg = String(r[COL_HIST_REGISTRO]);
    let turnoPersonal = mapTurnos[reg];
    let obs = String(r[COL_HIST_OBS] || ""); // Obtener observación actual
    
    // CONDICIÓN MEJORADA:
    // 1. Coincide el turno
    // 2. Y NO contiene la palabra "RECHAZADO" (ignoramos mayúsculas/minúsculas por seguridad)
    if (turnoPersonal === String(turnoFiltro) && !obs.toUpperCase().includes("RECHAZADO")) {
        pendientes.push({
          id: r[COL_HIST_ID],
          // ... resto de propiedades igual ...
          registro: reg,
          tipo: r[COL_HIST_TIPO],
          concepto: r[COL_HIST_CONCEPTO],
          valor: r[COL_HIST_VALOR],
          fecha: (r[COL_HIST_FECHA] instanceof Date) ? Utilities.formatDate(r[COL_HIST_FECHA], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[COL_HIST_FECHA],
          obs: r[COL_HIST_OBS],
          batchId: r[COL_HIST_BATCH]
        });
    }
  });
  
  return pendientes;
}

function processAuthorizationBatch(decisions) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPend = ss.getSheetByName(SHEET_PENDING);
    const sheetHist = ss.getSheetByName(SHEET_HIST);
    
    // Leer datos actuales para encontrar índices
    const dataPend = sheetPend.getDataRange().getValues();
    let idMap = new Map();
    // i=1 saltando header. row index = i+1
    for(let i=1; i<dataPend.length; i++) {
       idMap.set(String(dataPend[i][COL_HIST_ID]), i + 1);
    }
    
    let rowsToDelete = [];
    let rowsToMove = [];
    let rowsToUpdate = []; // Para rechazos (si decidimos marcar en lugar de borrar)

    decisions.forEach(d => {
      if(!idMap.has(d.id)) return;
      let rowIndex = idMap.get(d.id);
      let rowValues = dataPend[rowIndex - 1]; // array index
      
      if(d.action === 'APPROVE') {
         // Copiar tal cual a Historico
         rowsToMove.push(rowValues);
         rowsToDelete.push(rowIndex);
      } else if (d.action === 'REJECT') {
         // Marcar como rechazado en Historico2
         let currentObs = String(rowValues[COL_HIST_OBS]);
         let newObs = "RECHAZADO: " + currentObs;
         rowsToUpdate.push({rowIndex: rowIndex, colIndex: COL_HIST_OBS + 1, value: newObs});
         // NO lo agregamos a rowsToDelete para que quede evidencia en Historico2
      }
    });

    // 1. Mover a Historico (Append)
    if(rowsToMove.length > 0) {
      sheetHist.getRange(sheetHist.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
    }

    // 2. Actualizar Rechazados (Uno a uno, son pocos)
    rowsToUpdate.forEach(u => {
      sheetPend.getRange(u.rowIndex, u.colIndex).setValue(u.value);
    });

    // 3. Borrar Aprobados de Historico2 (Orden descendente crítico)
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(idx => sheetPend.deleteRow(idx));

    return { success: true };

  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// --- GENERACIÓN DE REPORTE EN SHEETS ---

function generateBalanceSheet(turno, exportData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Limpiamos el nombre del turno para que sea válido como nombre de hoja
  const cleanTurno = turno.toString().replace(/[^a-zA-Z0-9]/g, "_"); 
  const sheetName = "Reporte_" + cleanTurno;
  
  let sheet = ss.getSheetByName(sheetName);
  
  // Si no existe, la creamos. Si existe, borramos su contenido antiguo.
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); 
  }
  
  // 1. Poner Encabezados
  const headers = ["Registro", "Nombre", "Puesto", "Saldo Comp.", "Saldo Perm."];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Estilo de encabezado (Azul y Negrita)
  headerRange.setFontWeight("bold")
             .setBackground("#2980b9")
             .setFontColor("white");
             
  // 2. Volcar Datos (si hay)
  if (exportData && exportData.length > 0) {
    sheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }
  
  // 3. Ajustar anchos de columna automáticamente
  sheet.autoResizeColumns(1, 5);

  // Retornar URL directa a esa hoja (GID)
  return { 
    url: ss.getUrl() + "#gid=" + sheet.getSheetId(),
    name: sheetName
  };
} 