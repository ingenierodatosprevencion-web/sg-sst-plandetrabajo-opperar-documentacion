/*
 * Title: Generador de Plan de Trabajo Masivo
 * 
 * Script independiente para la creación anual de tareas cruzando actividades con sedes.
 */

/*
 * Function: generarPlanDeTrabajoMasivo
 * 
 * Lee el catálogo maestro (TBL_ACTIVIDADES) y genera registros en la tabla 
 * transaccional (TBL_PLAN_TRABAJO) respetando el cruce matricial de responsables.
 * 
 * Lógica de Negocio:
 *   El sistema convierte los analistas de la actividad y los analistas de la sede 
 *   en arreglos separados por comas. Si existe una *intersección* (al menos un 
 *   analista coincide), se genera la tarea asignando al responsable directo.
 * 
 * Returns:
 *   Muestra una alerta nativa (SpreadsheetApp.getUi) con el total de tareas creadas.
 */

function generarPlanDeTrabajoMasivo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const hojaOperaciones = ss.getSheetByName('TBL_OPERACIONES');
  const hojaActividades = ss.getSheetByName('TBL_ACTIVIDADES');
  const hojaPlan = ss.getSheetByName('TBL_PLAN_TRABAJO'); // Hoja productiva
  
  if (!hojaOperaciones || !hojaActividades || !hojaPlan) {
    SpreadsheetApp.getUi().alert('Error: No se encontraron las hojas necesarias.');
    return;
  }

  // Extraer datos omitiendo los encabezados
  const datosOperaciones = hojaOperaciones.getDataRange().getValues().slice(1);
  const datosActividades = hojaActividades.getDataRange().getValues().slice(1);
  
  // Obtenemos llaves existentes (Op-Act-Mes) para evitar duplicados
  const planExistente = hojaPlan.getDataRange().getValues().slice(1)
    .map(r => `${r[1]}-${r[2]}-${r[4]}`);

  const anioProgramacion = 2026; 
  const nuevosRegistros = [];
  
  datosActividades.forEach(actividad => {
    const idActividad = actividad[0];
    
    // TBL_ACTIVIDADES: Columna F (Índice 5) -> ANALISTAS_ASIGNADOS
    const analistasActividadString = (actividad[5] || "").toString().toLowerCase().trim(); 
    // TBL_ACTIVIDADES: Columna G (Índice 6) -> MESES
    const mesesString = (actividad[6] || "").toString().toUpperCase().trim(); 
    
    if (!idActividad || !analistasActividadString) return;

    // 1. Determinar meses aplicables
    let mesesAplicables = [];
    if (mesesString === 'TODOS' || mesesString === '') {
      mesesAplicables = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    } else {
      mesesAplicables = mesesString.split(',').map(m => parseInt(m.trim(), 10)).filter(m => !isNaN(m));
    }
    
    // 2. Convertir la Columna F en una lista filtrada
    const listaAnalistasActividad = analistasActividadString.split(',')
                                      .map(a => a.trim())
                                      .filter(a => a !== "");
    const aplicaATodos = analistasActividadString === 'todos';

    // 3. Recorrer TBL_OPERACIONES buscando coincidencias cruzadas
    datosOperaciones.forEach(operacion => {
      const idOperacion = operacion[0]; // Columna A
      
      // TBL_OPERACIONES: Columna D (Índice 3) -> ANALISTA_ASIGNADO (Ahora puede ser una lista)
      const analistasOperacionString = (operacion[3] || "").toString().toLowerCase().trim(); 
      
      // Convertir la celda de la operación en una lista de analistas
      const listaAnalistasOperacion = analistasOperacionString.split(',')
                                        .map(a => a.trim())
                                        .filter(a => a !== "");
      
      // (Fragmento dentro de CargaMasiva.gs - Reemplazar la sección del .push)
      const hayCoincidencia = aplicaATodos || listaAnalistasOperacion.some(analistaOp => listaAnalistasActividad.includes(analistaOp));
      
      if (hayCoincidencia) {
        // Identificar quién es el responsable exacto para ESTA tarea en esta sede
        let responsableFinal = "";
        if (aplicaATodos) {
            responsableFinal = analistasOperacionString; // Todos los de la sede
        } else {
            // El responsable es la intersección: el que está en la actividad y en la sede
            responsableFinal = listaAnalistasOperacion.find(a => listaAnalistasActividad.includes(a)) || "";
        }

        mesesAplicables.forEach(mes => {
          const llave = `${idOperacion}-${idActividad}-${mes}`;
          if (planExistente.indexOf(llave) === -1) {
            nuevosRegistros.push([
              Utilities.getUuid(), 
              idOperacion,       
              idActividad,       
              anioProgramacion,  
              mes,               
              'PLANEADO',        
              '',                
              '',                
              '',                 
              responsableFinal   // <--- NUEVA COLUMNA J: Se guarda el dueño exacto (ej. jegomez@opperar.com)
            ]);
          }
        });
      }
    });
  });
  
  // Escritura masiva
  if (nuevosRegistros.length > 0) {
    hojaPlan.getRange(hojaPlan.getLastRow() + 1, 1, nuevosRegistros.length, nuevosRegistros[0].length).setValues(nuevosRegistros);
    SpreadsheetApp.getUi().alert(`Éxito: Se generaron ${nuevosRegistros.length} tareas validando múltiples analistas.`);
  } else {
    SpreadsheetApp.getUi().alert('No se generaron registros nuevos (ya existen o no hubo coincidencias cruzadas).');
  }
}