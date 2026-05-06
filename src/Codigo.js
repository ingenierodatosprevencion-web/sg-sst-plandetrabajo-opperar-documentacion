/*
 * Title: Backend Principal SG-SST
 * 
 * Este archivo contiene la lógica transaccional de Google Apps Script. 
 * Actúa como puente entre la base de datos en Google Sheets y la interfaz SPA.
 */

// 1. SERVIR LA APLICACIÓN SPA
function doGet(e) {
  return HtmlService.createTemplateFromFile('App')
    .evaluate()
    .setTitle('SG-SST Opperar | Gestión Integral')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/*
 * Section: Gestión de Tareas
 * Contiene todas las funciones relacionadas con la lectura y actualización del plan de trabajo.
 */

// 2. AUTENTICACIÓN Y PERFIL
function obtenerPerfilUsuario() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUsers = ss.getSheetByName("TBL_USUARIOS");
  const data = sheetUsers.getDataRange().getValues();
  
  let usuarioInfo = {
    email: email,
    nombre: "Usuario Externo",
    rol: "INVITADO",
    existe: false,
    id: ""
  };
  
  for(let i=1; i<data.length; i++){
    if(data[i][2].toString().trim().toLowerCase() == email.toLowerCase()){
      usuarioInfo.nombre = data[i][1];
      usuarioInfo.rol = data[i][3].toString().trim().toUpperCase();
      usuarioInfo.id = data[i][0];
      usuarioInfo.existe = true;
      break;
    }
  }
  return usuarioInfo;
}

// OBTENER LISTAS PARA FILTROS
function obtenerListasFiltros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
     
  const dataAct = ss.getSheetByName("TBL_ACTIVIDADES").getDataRange().getValues();
  const programas = [];
  const pdts = []; // NUEVA LISTA PDT
  
  for(let i=1; i<dataAct.length; i++) {
    if(dataAct[i][1]) programas.push(dataAct[i][1]); // Columna B (Programa)
    if(dataAct[i][7]) pdts.push(dataAct[i][7].toString().trim()); // Columna H (TIPO_PDT)
  }

  const dataUsers = ss.getSheetByName("TBL_USUARIOS").getDataRange().getValues();
  const mapUsuarios = {};
  for(let i = 1; i < dataUsers.length; i++) {
    let nombre = dataUsers[i][1];
    let emailCol = dataUsers[i][2].toString().toLowerCase().trim();
    if(emailCol) mapUsuarios[emailCol] = nombre;
  }

  const dataOps = ss.getSheetByName("TBL_OPERACIONES").getDataRange().getValues();
  const analistasSet = new Set();
  const listaAnalistas = [];

  for(let i=1; i<dataOps.length; i++) {
    let analistasRaw = dataOps[i][3].toString().toLowerCase().trim();
    if (analistasRaw) {
      let analistasArray = analistasRaw.split(',').map(a => a.trim()).filter(a => a !== "");
      analistasArray.forEach(analista => {
        if(!analistasSet.has(analista)) {
          analistasSet.add(analista);
          listaAnalistas.push({ email: analista, nombre: mapUsuarios[analista] || analista });
        }
      });
    }
  }

  listaAnalistas.sort((a, b) => a.nombre.localeCompare(b.nombre));
     
  return { 
    programas: [...new Set(programas)], 
    analistas: listaAnalistas,
    pdts: [...new Set(pdts)] // Retornamos PDTs
  };
}

/*
 * Function: obtenerMisTareas
 * 
 * Recupera y filtra las tareas asignadas a un usuario validando permisos estrictos 
 * (Admin, Analista Principal o Auxiliar asignado).
 * 
 * Parameters:
 *   userEmail - Correo del usuario autenticado.
 *   userRol - Rol del usuario en el sistema.
 *   filtroMes - Mes seleccionado en la interfaz o 'todos'.
 *   filtroOp - ID de la sede/operación seleccionada.
 *   filtroProg - Nombre del programa a filtrar.
 *   filtroAnalista - Correo del analista (solo visible para Administradores).
 *   filtroEstado - Estado actual de la tarea (PLANEADO, TERMINADO, etc.).
 *   filtroAct - Texto para búsqueda parcial en el nombre de la actividad.
 *   filtroPdt - Tipo de Plan de Trabajo (ej. PESV SEDE CENTRAL).
 * 
 * Returns:
 *   Un arreglo de objetos JSON ordenados por mes, listos para ser renderizados por Vue.js.
 */


// OBTENER TAREAS (MIS TAREAS)
function obtenerMisTareas(userEmail, userRol, filtroMes, filtroOp, filtroProg, filtroAnalista, filtroEstado, filtroAct, filtroPdt) {

  const email = userEmail ? userEmail.toLowerCase().trim() : "";
  const esAdmin = (userRol === 'ADMINISTRADOR' || userRol === 'ADMIN');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perfil = obtenerPerfilUsuario();
     
  // ==========================================
  // 1. Crear mapa de nombres de usuarios
  // ==========================================
  const dataUsers = ss.getSheetByName("TBL_USUARIOS").getDataRange().getValues();
  const mapUsuarios = {};
  for(let i = 1; i < dataUsers.length; i++) {
    let nombre = dataUsers[i][1]; 
    let emailCol = dataUsers[i][2].toString().toLowerCase().trim(); 
    let usuarioCol = dataUsers[i][5].toString().toLowerCase().trim(); 
     
    if(emailCol) mapUsuarios[emailCol] = nombre;
    if(usuarioCol) mapUsuarios[usuarioCol] = nombre;
  }

  // ==========================================
  // 2. Mapear Operaciones (Identificando al Analista Principal)
  // ==========================================
  const dataOps = ss.getSheetByName("TBL_OPERACIONES").getDataRange().getValues();
  const mapOpsNames = {};
     
  for(let i=1; i<dataOps.length; i++) {
    let id = dataOps[i][0];
    let analistasRaw = dataOps[i][3].toString().toLowerCase().trim(); // Ej: lmurrego@..., lclopera@...
    let auxiliarEmail = dataOps[i][4].toString().toLowerCase().trim();
    
    // Convertir a lista y tomar el PRIMERO como el Principal
    let listaAnalistas = analistasRaw.split(',').map(a => a.trim()).filter(a => a);
    let analistaPrincipal = listaAnalistas.length > 0 ? listaAnalistas[0] : "";
       
    mapOpsNames[id] = {
      nombre: dataOps[i][2],
      analistaPrincipal: analistaPrincipal, // Guardamos quién es el principal
      auxiliarEmail: auxiliarEmail,         // Guardamos quién es el auxiliar
      analistaNombre: mapUsuarios[analistaPrincipal] || analistaPrincipal,
      auxiliarNombre: mapUsuarios[auxiliarEmail] || auxiliarEmail
    };
  }

  // MAPEO DE ACTIVIDADES (Agregamos la columna H = índice 7)
  const dataAct = ss.getSheetByName("TBL_ACTIVIDADES").getDataRange().getValues();
  const mapAct = {};
  for(let i=1; i<dataAct.length; i++) {
    mapAct[dataAct[i][0]] = { 
      nombre: dataAct[i][2], 
      programa: dataAct[i][1],
      pdt: dataAct[i][7] ? dataAct[i][7].toString().trim() : "Sin Plan" // Columna H
    };
  }

  const dataPlan = ss.getSheetByName("TBL_PLAN_TRABAJO").getDataRange().getValues();
  const tareas = [];
   
  for(let i=1; i<dataPlan.length; i++) {
    let idOp = dataPlan[i][1];
    let mes = dataPlan[i][4];
    let idAct = dataPlan[i][2];
    let estadoActual = dataPlan[i][5]; 
    // NUEVO: Leemos al responsable directo de la Columna J (Índice 9)
    let responsableDirecto = (dataPlan[i][9] || "").toString().toLowerCase().trim(); 

    let opInfo = mapOpsNames[idOp] || { nombre: "Desconocida", analistaPrincipal: "", auxiliarEmail: "" };
    let actInfo = mapAct[idAct] || { nombre: "Desconocida", programa: "" };

    // ==========================================
    // 3. LÓGICA DE PERMISOS ESTRICTOS
    // ==========================================
    // A. ¿Soy el analista al que le asignaron directamente esta tarea?
    let esMiTareaComoAnalista = responsableDirecto.includes(email);
    
    // B. ¿Soy el auxiliar de esta sede Y la tarea es del Analista Principal?
    let esMiTareaComoAuxiliar = (opInfo.auxiliarEmail === email) && (responsableDirecto === opInfo.analistaPrincipal);
    
    // C. Verificamos si tiene acceso
    let tienePermiso = esAdmin || esMiTareaComoAnalista || esMiTareaComoAuxiliar;
    
    if (!tienePermiso) continue; // Si no cumple ninguna, saltar a la siguiente fila

    // Filtros de la interfaz...
    if (filtroPdt && filtroPdt !== 'todos' && actInfo.pdt !== filtroPdt) continue;
    if (filtroMes !== 'todos' && mes != filtroMes) continue;
    if (filtroOp !== 'todas' && idOp != filtroOp) continue;
    if (filtroProg !== 'todos' && actInfo.programa !== filtroProg) continue;
    if (filtroEstado !== 'todos' && estadoActual !== filtroEstado) continue;
    if (filtroAct && !actInfo.nombre.toLowerCase().includes(filtroAct.toLowerCase())) continue;

    if (esAdmin && filtroAnalista !== 'todos') {
      if (responsableDirecto !== filtroAnalista.toLowerCase()) continue;
    }

    tareas.push({
      idRegistro: dataPlan[i][0],
      operacion: opInfo.nombre,
      // Mostramos el nombre de quien es responsable directo de la tarea
      analista: mapUsuarios[responsableDirecto] || responsableDirecto, 
      auxiliar: opInfo.auxiliarNombre,
      actividad: actInfo.nombre,
      programa: actInfo.programa,
      mes: mes,
      estado: estadoActual,
      observacion: dataPlan[i][7],
      evidencia: dataPlan[i][8]
    });
  }  
   
  return tareas.sort((a,b) => a.mes - b.mes);
}

// ======================================================================
// GUARDADO INDIVIDUAL CON LÓGICA DE REPROGRAMACIÓN (NO EJECUTADO)
// ======================================================================
function actualizarEstadoTarea(idReg, estado, obs, evidencia, userRol, userEmail) {  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TBL_PLAN_TRABAJO");  
  const data = sheet.getDataRange().getValues();  
  const esAdmin = (userRol === 'ADMINISTRADOR' || userRol === 'ADMIN');

  // 1. Buscar la tarea actual y guardar sus datos
  let targetIndex = -1;
  let idOp, idAct, mesActual, reprogramaciones;

  for(let i=1; i<data.length; i++) {  
    if(data[i][0] == idReg) {  
      targetIndex = i;
      idOp = data[i][1];
      idAct = data[i][2];
      mesActual = parseInt(data[i][4]);
      reprogramaciones = parseInt(data[i][10]) || 0;
      break;
    }  
  }

  if (targetIndex === -1) return { success: false, error: "Registro no encontrado" };

  // 2. LÓGICA DE REPROGRAMACIÓN
  if (estado === "REPROGRAMADO") {
    
    // Regla A: Diciembre no se toca
    if (mesActual >= 12) {
      return { success: false, reprogramada: false, error: "Las tareas de Diciembre no se pueden reprogramar al próximo año." };
    }

    // Regla B: Límite de 3 intentos para Analistas
    if (!esAdmin && reprogramaciones >= 3) {
      return { success: false, reprogramada: false, error: "Límite de reprogramaciones (3) alcanzado. Debes marcarla como NO EJECUTADO." };
    }

    // Regla C: Anti-Colisión (Tareas mensuales recurrentes)
    let existeProximoMes = false;
    for(let i=1; i<data.length; i++) {
      if (data[i][1] == idOp && data[i][2] == idAct && parseInt(data[i][4]) === (mesActual + 1)) {
        existeProximoMes = true; 
        break;
      }
    }

    if (existeProximoMes) {
      return { success: false, reprogramada: false, error: "Esta actividad es recurrente y ya tienes una programada para el próximo mes. Debes gestionarla como NO EJECUTADO." };
    }

    // Si pasa todas las reglas, hacemos el desplazamiento
    sheet.getRange(targetIndex+1, 5).setValue(mesActual + 1); 
    sheet.getRange(targetIndex+1, 6).setValue("PLANEADO"); // Se reinicia en el nuevo mes
    sheet.getRange(targetIndex+1, 8).setValue(obs);      
    sheet.getRange(targetIndex+1, 9).setValue(evidencia);
    sheet.getRange(targetIndex+1, 11).setValue(reprogramaciones + 1);

    // --> INYECCIÓN DE AUDITORÍA AQUÍ <--
    registrarAuditoria(userEmail, "UPDATE_REPROGRAMADO", "TBL_PLAN_TRABAJO", idReg, `Movido del mes ${mesActual} al ${mesActual + 1}. Obs: ${obs}`);
    
    return { success: true, reprogramada: true, msg: `Tarea reprogramada...` };

  } else {
    // 3. GUARDADO NORMAL (Terminado, Iniciado, No Ejecutado...)
    sheet.getRange(targetIndex+1, 6).setValue(estado);  
    if(estado !== "PLANEADO" && estado !== "REPROGRAMADO") {
      sheet.getRange(targetIndex+1, 7).setValue(new Date()); 
    }
    sheet.getRange(targetIndex+1, 8).setValue(obs);      
    sheet.getRange(targetIndex+1, 9).setValue(evidencia);  
    
    // --> INYECCIÓN DE AUDITORÍA AQUÍ <--
    registrarAuditoria(userEmail, "UPDATE", "TBL_PLAN_TRABAJO", idReg, `Estado cambiado a: ${estado}. Obs: ${obs}`);
    
    return { success: true, reprogramada: false, msg: "Estado actualizado..." };
  }
}

// NUEVA FUNCIÓN: ELIMINAR TAREA (SOLO ADMIN)
// Agregamos userEmail
function eliminarTarea(idReg, userEmail) {  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TBL_PLAN_TRABAJO");  
  const data = sheet.getDataRange().getValues();  
   
  for(let i=1; i<data.length; i++) {  
    if(data[i][0] == idReg) {  
      sheet.deleteRow(i + 1); 
      
      // --> INYECCIÓN DE AUDITORÍA AQUÍ <--
      registrarAuditoria(userEmail, "DELETE", "TBL_PLAN_TRABAJO", idReg, "Se eliminó el registro físicamente");
      
      return { success: true };  
    }  
  }  
  return { success: false };  
}

// DASHBOARD MULTI-NIVEL CON RETRASOS
function obtenerDatosDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPlan = ss.getSheetByName("TBL_PLAN_TRABAJO");
    const sheetAct = ss.getSheetByName("TBL_ACTIVIDADES");
    const sheetOps = ss.getSheetByName("TBL_OPERACIONES");
    const sheetUsers = ss.getSheetByName("TBL_USUARIOS");

    const dataPlan = sheetPlan.getDataRange().getValues();
    const dataAct = sheetAct.getDataRange().getValues();
    const dataOps = sheetOps.getDataRange().getValues();
    const dataUsers = sheetUsers.getDataRange().getValues();

    if (dataPlan.length <= 1) return { kpiGlobal: 0 };

    // 1. Mapear Usuarios (Forzamos Mayúsculas para unificar nombres y evitar duplicados)
    const mapUsuarios = {};
    for(let i = 1; i < dataUsers.length; i++) {
      let nombre = dataUsers[i][1].toString().trim().toUpperCase();
      let emailCol = dataUsers[i][2].toString().toLowerCase().trim();
      let userCol = dataUsers[i][5].toString().toLowerCase().trim();
      if(emailCol) mapUsuarios[emailCol] = nombre;
      if(userCol) mapUsuarios[userCol] = nombre;
    }

    const mapAct = {};
    for(let i=1; i<dataAct.length; i++) {
        mapAct[dataAct[i][0]] = {
            programa: dataAct[i][1],
            pdt: dataAct[i][7] ? dataAct[i][7].toString().trim() : "Sin Plan" // Columna H
        };
    }

    const mapOp = {};
    for(let i=1; i<dataOps.length; i++) mapOp[dataOps[i][0]] = { name: dataOps[i][2] };

    // Acumuladores
    const accGlobal = { p:0, o:0 };
    const accAnalista = {};
    const accOp = {};
    const accProg = {};
    const accMes = {};
    const accPdt = {}; // NUEVO
       
    let totalRetrasos = 0;
    const delayAnalista = {};
    const delayOp = {};
    const delayPdt = {}; // NUEVO

    const puntos = { "TERMINADO": 4, "AVANCE SUPERIOR": 3, "AVANCE PARCIAL": 2, "INICIADO": 1, "NO EJECUTADO": 0, "PLANEADO": 0 };
    const mesActual = new Date().getMonth() + 1;

    for (let i = 1; i < dataPlan.length; i++) {
      let row = dataPlan[i];
      let idOp = row[1];
      let idAct = row[2];
      let mes = row[4];
      let estado = row[5] ? row[5].toString().trim() : "PLANEADO";
      let responsableDirecto = (row[9] || "").toString().toLowerCase().trim(); // COLUMNA J

      let opData = mapOp[idOp] || { name: "Desconocida" };
      let actInfo = mapAct[idAct] || { programa: "Otros", pdt: "Otros" };
      let progName = actInfo.programa;
      let pdtName = actInfo.pdt;
       
      // RESOLVER DUPLICADOS: Si existe en Usuarios usa el nombre bonito, si no, usa lo que haya antes del @
      let analistaName = mapUsuarios[responsableDirecto] || responsableDirecto;
      if(analistaName.includes("@")) analistaName = analistaName.split("@")[0].toUpperCase();

      if (!isNaN(mes) && mes < mesActual && estado !== "TERMINADO") {
         totalRetrasos++;
         if(!delayAnalista[analistaName]) delayAnalista[analistaName] = 0;
         delayAnalista[analistaName]++;
         
         if(!delayOp[opData.name]) delayOp[opData.name] = 0;
         delayOp[opData.name]++;

         if(!delayPdt[pdtName]) delayPdt[pdtName] = 0;
         delayPdt[pdtName]++;
      }

      if (!isNaN(mes) && mes <= mesActual) {
        let base = 4;
        let real = puntos[estado] || 0;
           
        accGlobal.p += base; accGlobal.o += real;

        if(!accAnalista[analistaName]) accAnalista[analistaName] = {p:0, o:0};
        accAnalista[analistaName].p += base; accAnalista[analistaName].o += real;

        if(!accOp[opData.name]) accOp[opData.name] = {p:0, o:0};
        accOp[opData.name].p += base; accOp[opData.name].o += real;

        if(!accProg[progName]) accProg[progName] = {p:0, o:0};
        accProg[progName].p += base; accProg[progName].o += real;

        if(!accPdt[pdtName]) accPdt[pdtName] = {p:0, o:0};
        accPdt[pdtName].p += base; accPdt[pdtName].o += real;

        if(!accMes[mes]) accMes[mes] = {p:0, o:0};
        accMes[mes].p += base; accMes[mes].o += real;
      }  
    }

    const calc = (acc) => Object.keys(acc).map(k => ({ label: k, value: acc[k].p===0 ? 0 : Math.round((acc[k].o/acc[k].p)*100) })).sort((a,b) => b.value - a.value);
    const calcMes = (acc) => Object.keys(acc).map(k => ({ label: k, value: acc[k].p===0 ? 0 : Math.round((acc[k].o/acc[k].p)*100), id: parseInt(k) })).sort((a,b) => a.id - b.id);
    const formatDelay = (acc) => Object.keys(acc).map(k => ({ label: k, count: acc[k] })).sort((a,b) => b.count - a.count);

    return {
        kpiGlobal: accGlobal.p === 0 ? 0 : Math.round((accGlobal.o / accGlobal.p) * 100),
        byAnalista: calc(accAnalista),
        byOp: calc(accOp),
        byProg: calc(accProg),
        byMes: calcMes(accMes),
        byPdt: calc(accPdt), // Retorna PDTs
        delays: {
           total: totalRetrasos,
           byAnalista: formatDelay(delayAnalista),
           byOp: formatDelay(delayOp),
           byPdt: formatDelay(delayPdt) // Retorna Retrasos PDT
        }
    };
  } catch (e) {
    return { error: e.message };
  }
}

// PLANNING Y CONFIG (SIN CAMBIOS)
function obtenerDatosPlaneacion(userEmail, userRol) {
  const perfil = obtenerPerfilUsuario();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Ya no usamos Session ni obtenerPerfilUsuario() aquí
  const email = userEmail ? userEmail.toLowerCase().trim() : "";
  const esAdmin = (userRol === 'ADMINISTRADOR' || userRol === 'ADMIN');

  const dataAct = ss.getSheetByName("TBL_ACTIVIDADES").getDataRange().getValues();
  const actividades = [];
  for(let i=1; i<dataAct.length; i++) {
    if(dataAct[i][0]) actividades.push({ id: dataAct[i][0], nombre: dataAct[i][2], programa: dataAct[i][1] });
  }

  const dataOps = ss.getSheetByName("TBL_OPERACIONES").getDataRange().getValues();
  const misOperaciones = [];
  
  for(let i=1; i<dataOps.length; i++) {
    let idOp = dataOps[i][0];
    let procesoOp = dataOps[i][1]; // <--- NUEVO: Columna B
    let nombreOp = dataOps[i][2];
    let analista = dataOps[i][3].toString().toLowerCase();
    let auxiliar = dataOps[i][4].toString().toLowerCase();

    if (esAdmin || analista === email || auxiliar === email) {
      // Agregamos 'proceso' al objeto para que Vue pueda leerlo
      misOperaciones.push({ id: idOp, nombre: nombreOp, proceso: procesoOp });
    }
  }
  
  return { actividades, misOperaciones };
}

function guardarPlaneacionMasiva(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPlan = ss.getSheetByName("TBL_PLAN_TRABAJO");
  const sheetOps = ss.getSheetByName("TBL_OPERACIONES");
  const dataOps = sheetOps.getDataRange().getValues();
  
  let idsFinales = [];

  if (payload.modoAsignacion === 'individual') {
    idsFinales = payload.operaciones;
  } else {
    // Buscar todas las IDs de operaciones que pertenezcan a los procesos seleccionados
    for (let i = 1; i < dataOps.length; i++) {
      let idOp = dataOps[i][0];
      let procesoOp = dataOps[i][1]; // Columna B
      if (payload.procesosSeleccionados.includes(procesoOp)) {
        idsFinales.push(idOp);
      }
    }
  }

  const nuevasFilas = [];
  idsFinales.forEach(idOp => {
    payload.meses.forEach(mes => {
      // Estructura de 9 columnas para TBL_PLAN_TRABAJO
      nuevasFilas.push([Utilities.getUuid(), idOp, payload.idActividad, payload.anio, mes, "PLANEADO", "", "", ""]);
    });
  });

  if (nuevasFilas.length > 0) {
    sheetPlan.getRange(sheetPlan.getLastRow() + 1, 1, nuevasFilas.length, 9).setValues(nuevasFilas);
  }
  return { success: true, total: nuevasFilas.length };
}

// Agregamos userEmail como parámetro
function crearNuevaActividad(datos, userEmail) {  
  const ss = SpreadsheetApp.getActiveSpreadsheet();  
  const sheet = ss.getSheetByName("TBL_ACTIVIDADES");  
  const lastRow = sheet.getLastRow();  
  const newId = "ACT-" + ("000" + lastRow).slice(-3);  
  
  sheet.appendRow([newId, datos.programa, datos.descripcion, datos.meta || 0.8, datos.recursos || "", datos.tipo || "Operativo", "NO", datos.pdt || ""]);  
  
  // --> INYECCIÓN DE AUDITORÍA <--
  registrarAuditoria(userEmail, "CREATE", "TBL_ACTIVIDADES", newId, `Se creó la actividad: ${datos.descripcion} para el programa ${datos.programa}`);
  
  return { success: true, id: newId };  
}

function obtenerListaProgramas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TBL_PROGRAMAS");
  const data = sheet.getDataRange().getValues();
  const programas = []; for(let i=1; i<data.length; i++) programas.push(data[i][1]);
  return programas;
}

// ======================================================================
// NUEVA FUNCIÓN: VALIDAR LOGIN (USUARIO Y CONTRASEÑA)
// ======================================================================
function validarLogin(usuario, contrasena) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUsers = ss.getSheetByName("TBL_USUARIOS");
  const data = sheetUsers.getDataRange().getValues();
  
  let resultado = {
    existe: false,
    nombre: "",
    rol: "",
    email: "", // Mantener email para filtros de analista
    error: "Credenciales incorrectas"
  };
  
  // Recorremos desde fila 1 (saltando encabezado)
  for(let i=1; i<data.length; i++){
    // Columna F (Indice 5) es USUARIO, Columna G (Indice 6) es CONTRASENA
    // Columna E (Indice 4) es ACTIVO
    let uSheet = data[i][5].toString();
    let pSheet = data[i][6].toString();
    let activo = data[i][4];
    
    // Validación estricta (puedes usar toLowerCase si prefieres ignorar mayúsculas en usuario)
    if(uSheet === usuario && pSheet === contrasena) {
      if(activo === true || activo.toString() === 'true') {
        resultado.existe = true;
        resultado.nombre = data[i][1]; // Nombre
        resultado.email = data[i][2];  // Email (necesario para la lógica de tus filtros)
        resultado.rol = data[i][3].toString().trim().toUpperCase(); // Rol
        resultado.error = "";
      } else {
        resultado.error = "Usuario inactivo. Contacte al administrador.";
      }
      break;
    }
  }
  
  return resultado;
}

// ======================================================================
// GUARDADO MASIVO CON LÓGICA DE REPROGRAMACIÓN (NO EJECUTADO)
// ======================================================================
// ======================================================================
// GUARDADO MASIVO CON LÓGICA DE REPROGRAMACIÓN Y ANTI-COLISIÓN
// ======================================================================
function actualizarTareasMasivo(tareasModificadas, userRol, userEmail) {  
  if (!tareasModificadas || tareasModificadas.length === 0) return { success: true, count: 0, reprogramadas: 0 };  
   
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TBL_PLAN_TRABAJO");  
  const data = sheet.getDataRange().getValues();  
  const esAdmin = (userRol === 'ADMINISTRADOR' || userRol === 'ADMIN');
  
  let actualizadas = 0;  
  let reprogramadas = 0;
  let errores = []; // Para avisarle al frontend si algunas tareas no se pudieron mover
  
  // 1. Crear diccionario de las tareas que llegan del frontend
  const mapTareas = {};  
  tareasModificadas.forEach(t => mapTareas[t.idRegistro] = t);  

  // 2. Crear un SET de ocupación para validación rápida (Anti-Colisión)
  // Guardamos las llaves "ID_OP-ID_ACT-MES" que ya existen en la base de datos
  const ocupacionSet = new Set();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0]) {
      ocupacionSet.add(`${data[i][1]}-${data[i][2]}-${data[i][4]}`);
    }
  }
   
  // 3. Recorrer y aplicar cambios
  for(let i = 1; i < data.length; i++) {  
    let idReg = data[i][0];  
    
    if(mapTareas[idReg]) {  
      let t = mapTareas[idReg];  
      let idOp = data[i][1];
      let idAct = data[i][2];
      let mesActual = parseInt(data[i][4]);
      let reprogramaciones = parseInt(data[i][10]) || 0; // Columna K

      // LÓGICA DE DESPLAZAMIENTO (ESTADO: REPROGRAMADO)
      if (t.estado === "REPROGRAMADO") {
        
        let proximoMes = mesActual + 1;
        
        // Regla A: Diciembre
        if (mesActual >= 12) {
          errores.push(`La tarea de Diciembre (${t.actividad}) no se puede aplazar.`);
          continue; // Saltamos esta tarea, no la modificamos
        }
        
        // Regla B: Límite de Analistas
        else if (!esAdmin && reprogramaciones >= 3) {
          errores.push(`Límite alcanzado (3) para: ${t.actividad}. Marca NO EJECUTADO.`);
          continue;
        }
        
        // Regla C: Anti-Colisión (Si el próximo mes ya tiene esa misma tarea)
        else if (ocupacionSet.has(`${idOp}-${idAct}-${proximoMes}`)) {
          errores.push(`Colisión: Ya tienes la tarea "${t.actividad}" en el mes ${proximoMes}.`);
          continue;
        }
        
        // Si pasa todas las reglas, se reprograma
        else {
          sheet.getRange(i+1, 5).setValue(proximoMes); 
          sheet.getRange(i+1, 6).setValue("PLANEADO"); // Vuelve a estado limpio
          sheet.getRange(i+1, 8).setValue(t.observacion || ""); 
          sheet.getRange(i+1, 11).setValue(reprogramaciones + 1);
          
          // Agregamos la nueva ubicación al SET de ocupación para evitar que otra 
          // tarea modificada masivamente choque con esta que acabamos de mover
          ocupacionSet.add(`${idOp}-${idAct}-${proximoMes}`);
          
          reprogramadas++;
        }
        
      } else {
        // GUARDADO NORMAL (Terminado, Iniciado, No Ejecutado...)
        sheet.getRange(i+1, 6).setValue(t.estado); 
        sheet.getRange(i+1, 8).setValue(t.observacion || ""); 
        
        if(t.estado !== "PLANEADO" && t.estado !== "REPROGRAMADO") {
          sheet.getRange(i+1, 7).setValue(new Date());  
        }
        actualizadas++;  
      }
    }  
  }  
  
  // --> INYECCIÓN DE AUDITORÍA AQUÍ (Justo antes del return) <--
  if (actualizadas > 0 || reprogramadas > 0) {
    let detalleLog = `Masivo: ${actualizadas} actualizadas, ${reprogramadas} reprogramadas.`;
    if (errores.length > 0) detalleLog += ` Hubo ${errores.length} bloqueos anti-colisión.`;
    
    registrarAuditoria(userEmail, "UPDATE_MASIVO", "TBL_PLAN_TRABAJO", "MÚLTIPLES_IDS", detalleLog);
  }

  return { 
    success: true, 
    count: actualizadas, 
    reprogramadas: reprogramadas,
    errores: [...new Set(errores)] 
  };  
}

// ======================================================================
// NUEVA FUNCIÓN: CAMBIAR CONTRASEÑA
// ======================================================================
function cambiarContrasenaUsuario(usuario, passActual, passNueva) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TBL_USUARIOS");
  const data = sheet.getDataRange().getValues();
  
  // Recorremos desde fila 1 (saltando el encabezado)
  for(let i=1; i<data.length; i++){
    let uSheet = data[i][5].toString(); // Columna F (Usuario)
    let pSheet = data[i][6].toString(); // Columna G (Contraseña)
    let activo = data[i][4];            // Columna E (Activo)
    
    if(uSheet === usuario && pSheet === passActual) {
      if(activo === true || activo.toString() === 'true') {
        // Si coincide, sobrescribimos la contraseña en la columna G (índice 7 para getRange)
        sheet.getRange(i+1, 7).setValue(passNueva);
        return { success: true };
      } else {
        return { success: false, error: "Usuario inactivo. Contacte al administrador." };
      }
    }
  }
  return { success: false, error: "El usuario o la contraseña actual son incorrectos." };
}

// ======================================================================
// MOTOR DE AUDITORÍA (LOG DE CAMBIOS)
// ======================================================================
function registrarAuditoria(usuarioEmail, accion, tabla, idRegistro, detalles) {
  try {
    const sheetLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TBL_LOG_AUDITORIA");
    if (!sheetLog) return; // Si la hoja no existe, no hace nada para no romper la app

    sheetLog.appendRow([
      Utilities.getUuid(),
      new Date(),
      usuarioEmail,
      accion,
      tabla,
      idRegistro,
      detalles
    ]);
  } catch (e) {
    console.error("Error al registrar auditoría: ", e);
  }
}