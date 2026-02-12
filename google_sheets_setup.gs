/**
 * ============================================================
 * ENTRENAMIENTO EN CASA — Google Sheets Setup Script
 * ============================================================
 *
 * INSTRUCCIONES:
 * 1. Abre una hoja de Google Sheets nueva
 * 2. Ve a Extensiones > Apps Script
 * 3. Borra el contenido de Code.gs y pega TODO este archivo
 * 4. Guarda (Ctrl+S) y cierra el editor
 * 5. Recarga la hoja — aparecerá el menú "Entrenamiento"
 * 6. Click en "Entrenamiento" > "Configurar hoja completa"
 * 7. ¡Listo! Usa el menú para registrar, limpiar o deshacer
 */

// ── Colores ──────────────────────────────────────────────────
const AZUL_OSCURO  = '#1B2A4A';
const AZUL_MEDIO   = '#2E5090';
const AZUL_CLARO   = '#D6E4F0';
const VERDE        = '#27AE60';
const VERDE_CLARO  = '#D5F5E3';
const NARANJA      = '#E67E22';
const NARANJA_CL   = '#FDEBD0';
const GRIS_CLARO   = '#F2F2F2';
const BLANCO       = '#FFFFFF';
const ROJO         = '#E74C3C';

// ── Rutinas ──────────────────────────────────────────────────
const RUTINAS = {
  'Lunes - Pecho y Tríceps': [
    'Flexiones clásicas',
    'Flexiones diamante',
    'Flexiones declinadas',
    'Fondos en silla',
    'Flexiones abiertas',
    'Extensión tríceps con mancuerna',
  ],
  'Martes - Espalda y Bíceps': [
    'Remo con mancuernas',
    'Dominadas (o banda elástica)',
    'Remo invertido (mesa)',
    'Curl bíceps con mancuerna',
    'Curl martillo',
    'Superman',
  ],
  'Miércoles - Piernas': [
    'Sentadillas',
    'Sentadilla búlgara',
    'Zancadas',
    'Peso muerto rumano',
    'Elevación de talones',
    'Puente de glúteos',
  ],
  'Jueves - Hombros y Core': [
    'Press militar con mancuernas',
    'Elevaciones laterales',
    'Elevaciones frontales',
    'Plancha frontal (seg)',
    'Plancha lateral (seg)',
    'Crunch abdominal',
  ],
  'Viernes - Full Body': [
    'Burpees',
    'Sentadilla con salto',
    'Flexiones',
    'Remo con mancuernas',
    'Zancadas con salto',
    'Mountain climbers',
  ],
};

const DIAS_SEMANA = Object.keys(RUTINAS);

const CLAVES_DIA = {
  'Lunes - Pecho y Tríceps':   'Dia_Lunes',
  'Martes - Espalda y Bíceps': 'Dia_Martes',
  'Miércoles - Piernas':       'Dia_Miercoles',
  'Jueves - Hombros y Core':   'Dia_Jueves',
  'Viernes - Full Body':       'Dia_Viernes',
};

// Rangos óptimos de repeticiones por ejercicio {nombre: [min, max]}
const REPS_RANGES = {
  'Flexiones clásicas': [10, 20],
  'Flexiones diamante': [8, 15],
  'Flexiones declinadas': [8, 15],
  'Fondos en silla': [8, 15],
  'Flexiones abiertas': [10, 20],
  'Extensión tríceps con mancuerna': [10, 15],
  'Remo con mancuernas': [10, 15],
  'Dominadas (o banda elástica)': [5, 12],
  'Remo invertido (mesa)': [8, 15],
  'Curl bíceps con mancuerna': [10, 15],
  'Curl martillo': [10, 15],
  'Superman': [12, 20],
  'Sentadillas': [12, 20],
  'Sentadilla búlgara': [8, 12],
  'Zancadas': [10, 16],
  'Peso muerto rumano': [10, 15],
  'Elevación de talones': [15, 25],
  'Puente de glúteos': [12, 20],
  'Press militar con mancuernas': [8, 12],
  'Elevaciones laterales': [12, 20],
  'Elevaciones frontales': [12, 20],
  'Plancha frontal (seg)': [30, 60],
  'Plancha lateral (seg)': [20, 45],
  'Crunch abdominal': [15, 25],
  'Burpees': [8, 15],
  'Sentadilla con salto': [10, 20],
  'Flexiones': [10, 20],
  'Zancadas con salto': [10, 16],
  'Mountain climbers': [15, 30],
};

// ══════════════════════════════════════════════════════════════
// MENÚ PERSONALIZADO
// ══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏋️ Entrenamiento')
    .addItem('📋 Configurar hoja completa', 'configurarTodo')
    .addSeparator()
    .addItem('▶ Registrar entrada', 'registrarEntrada')
    .addItem('✕ Limpiar campos', 'limpiarCampos')
    .addItem('⟳ Deshacer último', 'deshacerUltimo')
    .addItem('📖 Ver rutina del día', 'cargarEjercicios')
    .addToUi();
}

// ══════════════════════════════════════════════════════════════
// CONFIGURACIÓN COMPLETA
// ══════════════════════════════════════════════════════════════

function configurarTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    'Configurar Entrenamiento',
    'Esto creará todas las hojas necesarias. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // Crear hojas en orden
  crearHojaDatos(ss);
  crearHojaRegistro(ss);
  crearHojaRutinas(ss);
  crearHojaDashboard(ss);
  crearHojaInstrucciones(ss);

  // Eliminar Sheet1 si existe y hay otras hojas
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Activar Registro
  const registro = ss.getSheetByName('Registro');
  if (registro) ss.setActiveSheet(registro);

  ui.alert('¡Configuración completa!',
    'Usa el menú "Entrenamiento" para registrar tus entrenamientos.\n\n' +
    '1. Selecciona Día/Rutina en B6\n' +
    '2. Selecciona Ejercicio en C6\n' +
    '3. Ingresa Serie #, Reps, Peso\n' +
    '4. Menú > Registrar entrada',
    ui.ButtonSet.OK);
}

// ══════════════════════════════════════════════════════════════
// UTILIDADES
// ══════════════════════════════════════════════════════════════

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    // Clear all conditional formatting
    sheet.clearConditionalFormatRules();
    // Remove all named ranges for this sheet
    const namedRanges = ss.getNamedRanges();
    namedRanges.forEach(nr => {
      if (nr.getRange().getSheet().getName() === name) {
        nr.remove();
      }
    });
  } else {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function setHeaderRow(sheet, row, values, bgColor, fontColor) {
  const range = sheet.getRange(row, 1, 1, values.length);
  range.setValues([values]);
  range.setBackground(bgColor);
  range.setFontColor(fontColor || BLANCO);
  range.setFontWeight('bold');
  range.setFontFamily('Calibri');
  range.setFontSize(11);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setWrap(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA: DATOS (oculta)
// ══════════════════════════════════════════════════════════════

function crearHojaDatos(ss) {
  const ws = getOrCreateSheet(ss, 'Datos');

  // Columnas A-E: Días y ejercicios
  const dias = Object.entries(RUTINAS);
  for (let col = 0; col < dias.length; col++) {
    const [dia, ejercicios] = dias[col];
    ws.getRange(1, col + 1).setValue(dia);
    for (let row = 0; row < ejercicios.length; row++) {
      ws.getRange(row + 2, col + 1).setValue(ejercicios[row]);
    }

    // Named range para cada día
    const clave = CLAVES_DIA[dia];
    const lastRow = 1 + ejercicios.length;
    ss.setNamedRange(clave, ws.getRange(2, col + 1, ejercicios.length, 1));
  }

  // Named range para lookup de días
  ss.setNamedRange('Dias_Lookup', ws.getRange(1, 1, 1, dias.length));

  // Columnas H-J: Tabla de ejercicios con rangos de reps
  ws.getRange(1, 8).setValue('Ejercicio');
  ws.getRange(1, 9).setValue('Reps Min');
  ws.getRange(1, 10).setValue('Reps Max');

  // Deduplicar ejercicios manteniendo orden
  const seen = new Set();
  const uniqueExercises = [];
  for (const ejercicios of Object.values(RUTINAS)) {
    for (const ej of ejercicios) {
      if (!seen.has(ej)) {
        seen.add(ej);
        uniqueExercises.push(ej);
      }
    }
  }

  for (let i = 0; i < uniqueExercises.length; i++) {
    const ej = uniqueExercises[i];
    const [rmin, rmax] = REPS_RANGES[ej] || [8, 15];
    ws.getRange(i + 2, 8).setValue(ej);
    ws.getRange(i + 2, 9).setValue(rmin);
    ws.getRange(i + 2, 10).setValue(rmax);
  }

  const lastExRow = 1 + uniqueExercises.length;
  ss.setNamedRange('TablaEjercicios', ws.getRange(1, 8, lastExRow, 3));

  // Ocultar la hoja
  ws.hideSheet();

  return ws;
}

// ══════════════════════════════════════════════════════════════
// HOJA: REGISTRO
// ══════════════════════════════════════════════════════════════

function crearHojaRegistro(ss) {
  const ws = getOrCreateSheet(ss, 'Registro');
  ws.setTabColor(AZUL_MEDIO);

  // Asegurar filas y columnas suficientes
  if (ws.getMaxColumns() < 9) ws.insertColumnsAfter(ws.getMaxColumns(), 9 - ws.getMaxColumns());
  if (ws.getMaxRows() < 200) ws.insertRowsAfter(ws.getMaxRows(), 200 - ws.getMaxRows());

  // Column widths
  ws.setColumnWidth(1, 100);   // Fecha
  ws.setColumnWidth(2, 220);   // Día/Rutina
  ws.setColumnWidth(3, 220);   // Ejercicio
  ws.setColumnWidth(4, 70);    // Serie #
  ws.setColumnWidth(5, 70);    // Reps
  ws.setColumnWidth(6, 90);    // Peso
  ws.setColumnWidth(7, 120);   // Descanso
  ws.setColumnWidth(8, 220);   // Notas
  ws.setColumnWidth(9, 10);    // Auxiliar (oculta)

  // ── Título ──
  ws.getRange('A1:H1').merge().setValue('REGISTRO DE ENTRENAMIENTO EN CASA')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(16).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  ws.getRange('A2:H2').merge()
    .setValue('Ingresa tus datos de cada sesión. Usa el menú Entrenamiento para registrar.')
    .setBackground(AZUL_CLARO).setFontColor('#666666')
    .setFontSize(10).setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  // ── Entrada Rápida título ──
  ws.getRange('A4:H4').merge().setValue('ENTRADA RÁPIDA')
    .setBackground(AZUL_MEDIO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  // ── Labels (fila 5) ──
  const labels = ['Fecha:', 'Día/Rutina:', 'Ejercicio:', 'Serie #:', 'Reps:', 'Peso (kg):', 'Descanso (seg):', 'Notas:'];
  const labelRange = ws.getRange(5, 1, 1, 8);
  labelRange.setValues([labels]);
  labelRange.setBackground(NARANJA_CL).setFontColor('#333333')
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Input row (fila 6) ──
  const inputRange = ws.getRange(6, 1, 1, 8);
  inputRange.setBackground(AZUL_CLARO).setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Fecha automática
  ws.getRange('A6').setFormula('=TODAY()').setNumberFormat('dd/MM/yyyy');

  // ── Validación: Día/Rutina (B6) ──
  const dvDia = SpreadsheetApp.newDataValidation()
    .requireValueInList(DIAS_SEMANA, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona el día de rutina')
    .build();
  ws.getRange('B6').setDataValidation(dvDia);

  // ── Celda auxiliar I6: mapea día → clave de rango ──
  ws.getRange('I6').setFormula(
    '=IF(B6="","",IF(LEFT(B6,5)="Lunes","Dia_Lunes",' +
    'IF(LEFT(B6,6)="Martes","Dia_Martes",' +
    'IF(LEFT(B6,4)="Miér","Dia_Miercoles",' +
    'IF(LEFT(B6,6)="Jueves","Dia_Jueves",' +
    'IF(LEFT(B6,7)="Viernes","Dia_Viernes",""))))))'
  );
  ws.getRange('I6').setFontColor(BLANCO).setFontSize(1);
  // Ocultar columna I
  ws.hideColumns(9);

  // ── Validación dinámica: Ejercicio (C6) ──
  // En Google Sheets, INDIRECT en data validation funciona igual
  const dvEjercicio = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ws.getRange('I6'), true)  // placeholder
    .setAllowInvalid(true)
    .setHelpText('Selecciona un ejercicio de la rutina del día')
    .build();
  // Usamos criterio custom con INDIRECT
  const dvEjercicioCustom = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=NOT(ISERROR(MATCH(C6,INDIRECT(I6),0)))')
    .setAllowInvalid(true)
    .setHelpText('Selecciona ejercicio o escribe uno personalizado')
    .build();
  ws.getRange('C6').setDataValidation(dvEjercicioCustom);

  // ── Validación: Serie # (D6) ──
  const dvSerie = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1','2','3','4','5','6','7','8','9','10'], true)
    .setAllowInvalid(true)
    .setHelpText('Número de serie (1, 2, 3...)')
    .build();
  ws.getRange('D6').setDataValidation(dvSerie);

  // ── Indicador de rango (fila 7) ──
  ws.getRange('A7:C7').merge().setFormula(
    '=IF(C6="","",IF(ISERROR(VLOOKUP(C6,TablaEjercicios,2,FALSE)),' +
    '"","Rango: "&VLOOKUP(C6,TablaEjercicios,2,FALSE)' +
    '&" - "&VLOOKUP(C6,TablaEjercicios,3,FALSE)&" reps"))'
  ).setBackground(AZUL_CLARO).setFontColor('#555555')
    .setFontSize(10).setFontStyle('italic')
    .setHorizontalAlignment('center');

  ws.getRange('D7:G7').merge().setFormula(
    '=IF(OR(C6="",E6=""),"",IF(ISERROR(VLOOKUP(C6,TablaEjercicios,2,FALSE)),' +
    '"",IF(E6>VLOOKUP(C6,TablaEjercicios,3,FALSE),' +
    '"⬆ SUPERA RANGO - SUBE PESO",' +
    'IF(E6<VLOOKUP(C6,TablaEjercicios,2,FALSE),' +
    '"⬇ BAJO RANGO - BAJA PESO",' +
    '"✔ EN RANGO ÓPTIMO"))))'
  ).setBackground(AZUL_CLARO).setFontColor('#333333')
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center');

  // ── Botones (fila 8) - instrucciones visuales ──
  ws.getRange('B8:C8').merge().setValue('▶ REGISTRAR ENTRADA  (menú)')
    .setBackground(VERDE).setFontColor(BLANCO)
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange('E8:F8').merge().setValue('✕ LIMPIAR CAMPOS')
    .setBackground(NARANJA).setFontColor(BLANCO)
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange('G8:H8').merge().setValue('⟳ DESHACER ÚLTIMO')
    .setBackground(ROJO).setFontColor(BLANCO)
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Historial título ──
  ws.getRange('A10:H10').merge().setValue('HISTORIAL DE ENTRENAMIENTOS')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  // ── Headers historial (fila 11) ──
  const headers = ['Fecha', 'Día / Rutina', 'Ejercicio', 'Serie #', 'Reps', 'Peso (kg)', 'Descanso (s)', 'Notas'];
  const headerRange = ws.getRange(11, 1, 1, 8);
  headerRange.setValues([headers]);
  headerRange.setBackground(AZUL_MEDIO).setFontColor(BLANCO)
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center');

  // ── Pre-formato filas 12-200 con bandas alternas ──
  for (let r = 12; r <= 200; r++) {
    const fill = (r % 2 === 0) ? BLANCO : GRIS_CLARO;
    ws.getRange(r, 1, 1, 8).setBackground(fill)
      .setFontFamily('Calibri').setFontSize(11).setFontColor('#333333')
      .setHorizontalAlignment('center');
    ws.getRange(r, 1).setNumberFormat('dd/MM/yyyy');
  }

  // ── Formato condicional: alertas de reps ──
  const repsRange = ws.getRange('E12:E200');

  // Rojo: reps > max → subir peso
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(E12<>"",NOT(ISERROR(VLOOKUP($C12,TablaEjercicios,3,FALSE))),' +
      'E12>VLOOKUP($C12,TablaEjercicios,3,FALSE))')
    .setBackground('#FFCCCC')
    .setFontColor('#CC0000')
    .setBold(true)
    .setRanges([repsRange])
    .build();

  // Amarillo: reps < min → considerar bajar peso
  const ruleLow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(E12<>"",NOT(ISERROR(VLOOKUP($C12,TablaEjercicios,2,FALSE))),' +
      'E12<VLOOKUP($C12,TablaEjercicios,2,FALSE))')
    .setBackground('#FFF3CD')
    .setFontColor('#856404')
    .setBold(true)
    .setRanges([repsRange])
    .build();

  // Verde: reps en rango óptimo
  const ruleOk = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(E12<>"",NOT(ISERROR(VLOOKUP($C12,TablaEjercicios,2,FALSE))),' +
      'E12>=VLOOKUP($C12,TablaEjercicios,2,FALSE),' +
      'E12<=VLOOKUP($C12,TablaEjercicios,3,FALSE))')
    .setBackground('#D4EDDA')
    .setFontColor('#155724')
    .setRanges([repsRange])
    .build();

  ws.setConditionalFormatRules([ruleHigh, ruleLow, ruleOk]);

  // ── Freeze + Filter ──
  ws.setFrozenRows(11);
  ws.getRange('A11:H200').createFilter();

  return ws;
}

// ══════════════════════════════════════════════════════════════
// HOJA: RUTINAS
// ══════════════════════════════════════════════════════════════

function crearHojaRutinas(ss) {
  const ws = getOrCreateSheet(ss, 'Rutinas');
  ws.setTabColor(VERDE);

  if (ws.getMaxColumns() < 7) ws.insertColumnsAfter(ws.getMaxColumns(), 7 - ws.getMaxColumns());

  ws.setColumnWidth(1, 45);
  ws.setColumnWidth(2, 260);
  ws.setColumnWidth(3, 160);
  ws.setColumnWidth(4, 90);
  ws.setColumnWidth(5, 90);
  ws.setColumnWidth(6, 90);
  ws.setColumnWidth(7, 105);

  // Título
  ws.getRange('A1:G1').merge().setValue('RUTINAS SEMANALES - ENTRENAMIENTO EN CASA')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(16).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  ws.getRange('A2:G2').merge()
    .setValue('Personaliza tus rutinas aquí. Ajusta los rangos de reps para recibir alertas automáticas.')
    .setBackground(VERDE_CLARO).setFontColor('#666666')
    .setFontSize(10).setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  let row = 4;
  for (const [dia, ejercicios] of Object.entries(RUTINAS)) {
    // Día header
    ws.getRange(row, 1, 1, 7).merge().setValue(dia)
      .setBackground(AZUL_MEDIO).setFontColor(BLANCO)
      .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('center');
    row++;

    // Sub-headers
    const subHeaders = ['#', 'Ejercicio', 'Grupo Muscular', 'Series Obj.', 'Reps Min', 'Reps Max', 'Peso Obj. (kg)'];
    ws.getRange(row, 1, 1, 7).setValues([subHeaders])
      .setBackground(VERDE).setFontColor(BLANCO)
      .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
      .setHorizontalAlignment('center');
    row++;

    // Ejercicios
    for (let idx = 0; idx < ejercicios.length; idx++) {
      const ej = ejercicios[idx];
      const fill = (idx % 2 === 0) ? GRIS_CLARO : BLANCO;
      const [rmin, rmax] = REPS_RANGES[ej] || [8, 15];

      const rowRange = ws.getRange(row, 1, 1, 7);
      rowRange.setBackground(fill).setFontFamily('Calibri').setFontSize(11)
        .setFontColor('#333333').setHorizontalAlignment('center');

      ws.getRange(row, 1).setValue(idx + 1);
      ws.getRange(row, 2).setValue(ej).setHorizontalAlignment('left');
      // Col 3 (Grupo), Col 4 (Series Obj) = vacíos
      ws.getRange(row, 5).setValue(rmin);
      ws.getRange(row, 6).setValue(rmax);
      row++;
    }

    row++; // espacio entre días
  }

  return ws;
}

// ══════════════════════════════════════════════════════════════
// HOJA: DASHBOARD
// ══════════════════════════════════════════════════════════════

function crearHojaDashboard(ss) {
  const ws = getOrCreateSheet(ss, 'Dashboard');
  ws.setTabColor(NARANJA);

  if (ws.getMaxColumns() < 12) ws.insertColumnsAfter(ws.getMaxColumns(), 12 - ws.getMaxColumns());
  if (ws.getMaxRows() < 55) ws.insertRowsAfter(ws.getMaxRows(), 55 - ws.getMaxRows());

  ws.setColumnWidth(1, 15);
  ws.setColumnWidth(2, 195);
  for (let c = 3; c <= 6; c++) ws.setColumnWidth(c, 105);
  ws.setColumnWidth(7, 15);
  for (let c = 8; c <= 12; c++) ws.setColumnWidth(c, 90);

  // ── Título ──
  ws.getRange('A1:L1').merge().setValue('DASHBOARD DE ENTRENAMIENTO')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(16).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  ws.getRange('A2:L2').merge()
    .setValue('Resumen visual de tu progreso. Los datos se actualizan automáticamente desde la hoja Registro.')
    .setBackground(NARANJA_CL).setFontColor('#666666')
    .setFontSize(10).setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  // ── MÉTRICAS CLAVE ──
  ws.getRange('A4:L4').merge().setValue('MÉTRICAS CLAVE')
    .setBackground(AZUL_MEDIO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  const kpis = [
    {label: 'Total Sesiones', formula: '=COUNTA(Registro!A12:A200)', format: '0', color: VERDE, col: 2},
    {label: 'Último Entreno', formula: '=IF(COUNTA(Registro!A12:A200)>0,MAX(Registro!A12:A200),"-")', format: 'dd/MM/yyyy', color: AZUL_MEDIO, col: 4},
    {label: 'Peso Máx (kg)', formula: '=IF(MAX(Registro!F12:F200)>0,MAX(Registro!F12:F200),"-")', format: '0.0', color: NARANJA, col: 6},
    {label: 'Prom. Reps', formula: '=IFERROR(ROUND(AVERAGE(Registro!E12:E200),1),"-")', format: '0.0', color: AZUL_OSCURO, col: 8},
    {label: 'Total Sets', formula: '=IFERROR(COUNTA(Registro!D12:D200),"-")', format: '#,##0', color: ROJO, col: 10},
  ];

  for (const kpi of kpis) {
    // Label
    ws.getRange(5, kpi.col, 1, 2).merge().setValue(kpi.label)
      .setBackground(kpi.color).setFontColor(BLANCO)
      .setFontSize(9).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('center');

    // Value (2 rows tall)
    ws.getRange(6, kpi.col, 2, 2).merge().setFormula(kpi.formula)
      .setBackground(kpi.color).setFontColor(BLANCO)
      .setFontSize(22).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setNumberFormat(kpi.format);
  }

  // ── VOLUMEN POR DÍA ──
  ws.getRange('B9:F9').merge().setValue('VOLUMEN POR DÍA DE RUTINA')
    .setBackground(AZUL_MEDIO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  const volHeaders = ['Día', 'Sets', 'Total Reps', 'Reps Prom.', 'Peso Prom.'];
  ws.getRange(10, 2, 1, 5).setValues([volHeaders])
    .setBackground(VERDE).setFontColor(BLANCO)
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center');

  for (let i = 0; i < DIAS_SEMANA.length; i++) {
    const r = 11 + i;
    const dia = DIAS_SEMANA[i];
    const short = dia.split(' - ')[0];
    const fill = (i % 2 === 0) ? BLANCO : GRIS_CLARO;

    ws.getRange(r, 2, 1, 5).setBackground(fill).setFontFamily('Calibri').setFontSize(11);
    ws.getRange(r, 2).setValue(short).setFontWeight('bold').setHorizontalAlignment('left');
    ws.getRange(r, 3).setFormula('=COUNTIF(Registro!B$12:B$200,"' + dia + '")').setHorizontalAlignment('center');
    ws.getRange(r, 4).setFormula('=SUMIF(Registro!B$12:B$200,"' + dia + '",Registro!E$12:E$200)').setHorizontalAlignment('center');
    ws.getRange(r, 5).setFormula('=IFERROR(AVERAGEIF(Registro!B$12:B$200,"' + dia + '",Registro!E$12:E$200),0)').setHorizontalAlignment('center');
    ws.getRange(r, 6).setFormula('=IFERROR(AVERAGEIF(Registro!B$12:B$200,"' + dia + '",Registro!F$12:F$200),0)').setHorizontalAlignment('center');
  }

  // ── Gráfico de barras ──
  const chartBar = ws.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(ws.getRange('B10:C15'))
    .setPosition(9, 8, 0, 0)
    .setOption('title', 'Sets por Día')
    .setOption('legend', {position: 'none'})
    .setOption('width', 450)
    .setOption('height', 280)
    .build();
  ws.insertChart(chartBar);

  // ── INSIGHTS ──
  ws.getRange('B17:F17').merge().setValue('INSIGHTS')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  const insights = [
    ['Días sin entrenar', '=IF(COUNTA(Registro!A12:A200)>0,TODAY()-MAX(Registro!A12:A200),"-")'],
    ['Sesiones esta semana', '=COUNTIFS(Registro!A12:A200,">="&(TODAY()-WEEKDAY(TODAY(),2)+1),Registro!A12:A200,"<="&TODAY())'],
    ['Sesiones este mes', '=COUNTIFS(Registro!A12:A200,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Registro!A12:A200,"<="&TODAY())'],
    ['Ejercicio más frecuente', '=IFERROR(INDEX(Registro!C12:C200,MATCH(MAX(COUNTIF(Registro!C12:C200,Registro!C12:C200)),COUNTIF(Registro!C12:C200,Registro!C12:C200),0)),"-")'],
    ['Total reps acumuladas', '=IFERROR(SUM(Registro!E12:E200),"-")'],
    ['Reps max en 1 set', '=IFERROR(MAX(Registro!E12:E200),"-")'],
  ];

  for (let i = 0; i < insights.length; i++) {
    const r = 18 + i;
    const fill = (i % 2 === 0) ? BLANCO : GRIS_CLARO;
    ws.getRange(r, 2).setValue(insights[i][0])
      .setBackground(fill).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('left');
    ws.getRange(r, 3, 1, 4).merge().setFormula(insights[i][1])
      .setBackground(fill).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('center');
  }

  // ── Gráfico circular ──
  const chartPie = ws.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(ws.getRange('B10:C15'))
    .setPosition(22, 8, 0, 0)
    .setOption('title', 'Distribución por Rutina')
    .setOption('pieSliceText', 'percentage')
    .setOption('width', 450)
    .setOption('height', 280)
    .build();
  ws.insertChart(chartPie);

  // ── ALERTAS DE PROGRESO ──
  ws.getRange('B26:F26').merge().setValue('⚠ ALERTAS DE PROGRESO')
    .setBackground(ROJO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  const alertas = [
    ['Registros SOBRE rango (subir peso)',
     '=IFERROR(SUMPRODUCT((Registro!E12:E200<>"")*NOT(ISERROR(VLOOKUP(Registro!C12:C200,TablaEjercicios,3,FALSE)))*(Registro!E12:E200>VLOOKUP(Registro!C12:C200,TablaEjercicios,3,FALSE))),0)'],
    ['Registros BAJO rango (bajar peso)',
     '=IFERROR(SUMPRODUCT((Registro!E12:E200<>"")*NOT(ISERROR(VLOOKUP(Registro!C12:C200,TablaEjercicios,2,FALSE)))*(Registro!E12:E200<VLOOKUP(Registro!C12:C200,TablaEjercicios,2,FALSE))),0)'],
    ['% en rango óptimo',
     '=IFERROR(TEXT(1-((SUMPRODUCT((Registro!E12:E200<>"")*NOT(ISERROR(VLOOKUP(Registro!C12:C200,TablaEjercicios,3,FALSE)))*(Registro!E12:E200>VLOOKUP(Registro!C12:C200,TablaEjercicios,3,FALSE)))+SUMPRODUCT((Registro!E12:E200<>"")*NOT(ISERROR(VLOOKUP(Registro!C12:C200,TablaEjercicios,2,FALSE)))*(Registro!E12:E200<VLOOKUP(Registro!C12:C200,TablaEjercicios,2,FALSE))))/MAX(COUNTA(Registro!E12:E200),1)),"0%"),"-")'],
  ];

  for (let i = 0; i < alertas.length; i++) {
    const r = 27 + i;
    const fill = (i % 2 === 0) ? BLANCO : GRIS_CLARO;
    ws.getRange(r, 2).setValue(alertas[i][0])
      .setBackground(fill).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('left');
    ws.getRange(r, 3, 1, 4).merge().setFormula(alertas[i][1])
      .setBackground(fill).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('center');
  }

  // ── ÚLTIMOS 10 REGISTROS ──
  ws.getRange('B31:F31').merge().setValue('ÚLTIMOS 10 REGISTROS')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  const lastHeaders = ['Fecha', 'Rutina', 'Ejercicio', 'Reps', 'Peso (kg)'];
  ws.getRange(32, 2, 1, 5).setValues([lastHeaders])
    .setBackground(VERDE).setFontColor(BLANCO)
    .setFontWeight('bold').setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center');

  const colsMap = {2: 'A', 3: 'B', 4: 'C', 5: 'E', 6: 'F'};
  for (let i = 0; i < 10; i++) {
    const r = 33 + i;
    const fill = (i % 2 === 0) ? BLANCO : GRIS_CLARO;
    ws.getRange(r, 2, 1, 5).setBackground(fill).setFontFamily('Calibri').setHorizontalAlignment('center');

    for (const [c, colLetter] of Object.entries(colsMap)) {
      ws.getRange(r, parseInt(c)).setFormula(
        '=IFERROR(INDEX(Registro!' + colLetter + '$12:' + colLetter + '$200,COUNTA(Registro!A$12:A$200)-' + i + '),"")'
      );
    }
    ws.getRange(r, 2).setNumberFormat('dd/MM/yyyy');
  }

  // Freeze
  ws.setFrozenRows(3);

  return ws;
}

// ══════════════════════════════════════════════════════════════
// HOJA: INSTRUCCIONES
// ══════════════════════════════════════════════════════════════

function crearHojaInstrucciones(ss) {
  const ws = getOrCreateSheet(ss, 'Instrucciones');
  ws.setTabColor('#95A5A6');

  ws.setColumnWidth(1, 35);
  ws.setColumnWidth(2, 580);

  // Título
  ws.getRange('A1:B1').merge().setValue('INSTRUCCIONES DE USO')
    .setBackground(AZUL_OSCURO).setFontColor(BLANCO)
    .setFontSize(16).setFontWeight('bold').setFontFamily('Calibri')
    .setHorizontalAlignment('center');

  const instrucciones = [
    ['HOJA \'REGISTRO\'', [
      'Cada fila = UN SET de un ejercicio (Serie 1: 15 reps, Serie 2: 12 reps, etc.).',
      '1. La FECHA se llena automáticamente con el día de hoy.',
      '2. Selecciona el DÍA/RUTINA y luego el EJERCICIO del desplegable dinámico.',
      '3. Ingresa el SERIE # (1, 2, 3...), REPS y PESO de ese set.',
      '4. Al registrar, el Serie # sube automáticamente para el siguiente set.',
      '5. La fila 7 muestra si tus reps están EN RANGO, SOBRE o BAJO el óptimo.',
      '6. En el historial: VERDE = en rango, ROJO = sube peso, AMARILLO = baja peso.',
      '7. Ve a menú Entrenamiento > Registrar entrada para guardar el set.',
      '8. Usa Limpiar campos o Deshacer último según necesites.',
    ]],
    ['HOJA \'RUTINAS\'', [
      'Rutinas predefinidas para 5 días con rangos de repeticiones óptimas.',
      'Columnas Reps Min y Reps Max definen el rango ideal por ejercicio.',
      'Si superas Reps Max consistentemente, es hora de SUBIR PESO.',
      'Puedes editar los rangos para personalizar tus objetivos.',
    ]],
    ['HOJA \'DASHBOARD\'', [
      'Dashboard visual con KPIs, gráficos e insights de tu progreso.',
      '5 tarjetas de métricas clave: sesiones, último entreno, peso máx, etc.',
      'Gráfico de barras: sesiones por día de rutina.',
      'Gráfico circular: distribución porcentual de entrenamientos.',
      'Sección de insights: días sin entrenar, sesiones semanales/mensuales.',
      'Los últimos 10 registros se muestran al final.',
    ]],
    ['SISTEMA DE ALERTAS', [
      'Cada ejercicio tiene un rango óptimo de repeticiones (Reps Min - Reps Max).',
      'VERDE en la columna Reps = estás dentro del rango óptimo.',
      'ROJO = superaste el máximo → SUBE EL PESO en tu próximo entrenamiento.',
      'AMARILLO = estás por debajo del mínimo → considera BAJAR PESO.',
      'El Dashboard muestra cuántos registros están fuera de rango.',
    ]],
    ['TIPS', [
      'Registra CADA SET por separado: Serie 1 (15 reps), Serie 2 (12 reps), etc.',
      'El Serie # sube solo al registrar — perfecto para múltiples sets seguidos.',
      'Revisa el Dashboard regularmente para ver alertas y tu progreso.',
      '¡La constancia es la clave! Intenta entrenar al menos 3-4 días por semana.',
    ]],
  ];

  let row = 3;
  for (const [titulo, lineas] of instrucciones) {
    ws.getRange(row, 1, 1, 2).merge().setValue(titulo)
      .setBackground(AZUL_MEDIO).setFontColor(BLANCO)
      .setFontSize(12).setFontWeight('bold').setFontFamily('Calibri')
      .setHorizontalAlignment('center');
    row++;

    for (const linea of lineas) {
      ws.getRange(row, 1).setValue('•').setHorizontalAlignment('center')
        .setBackground(BLANCO).setFontFamily('Calibri').setFontSize(11);
      ws.getRange(row, 2).setValue(linea).setHorizontalAlignment('left')
        .setBackground(BLANCO).setFontFamily('Calibri').setFontSize(11);
      row++;
    }
    row++;
  }

  return ws;
}

// ══════════════════════════════════════════════════════════════
// MACROS: Registrar / Limpiar / Deshacer
// ══════════════════════════════════════════════════════════════

function registrarEntrada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Registro');
  const ui = SpreadsheetApp.getUi();

  if (!ws) {
    ui.alert('Error', 'No se encontró la hoja "Registro". Ejecuta primero "Configurar hoja completa".', ui.ButtonSet.OK);
    return;
  }

  // Buscar siguiente fila vacía
  let nextRow = 12;
  while (ws.getRange(nextRow, 1).getValue() !== '' && nextRow <= 200) {
    nextRow++;
  }

  if (nextRow > 200) {
    ui.alert('Historial lleno', 'El historial está lleno (200 registros). Copia los datos a otro archivo.', ui.ButtonSet.OK);
    return;
  }

  // Validar ejercicio
  const ejercicio = ws.getRange('C6').getValue();
  if (!ejercicio || String(ejercicio).trim() === '') {
    ui.alert('Falta información', 'Por favor ingresa al menos el nombre del ejercicio.', ui.ButtonSet.OK);
    return;
  }

  // Copiar datos de fila 6 a historial
  for (let col = 1; col <= 8; col++) {
    const val = ws.getRange(6, col).getValue();
    ws.getRange(nextRow, col).setValue(val);
  }

  // Formato fecha
  ws.getRange(nextRow, 1).setNumberFormat('dd/MM/yyyy');

  // Auto-incrementar Serie #
  const serieActual = ws.getRange('D6').getValue();
  const nextSerie = (typeof serieActual === 'number' && serieActual > 0)
    ? serieActual + 1
    : 2;
  ws.getRange('D6').setValue(nextSerie);

  // Limpiar solo reps, peso, descanso y notas (mantener ejercicio y día)
  ws.getRange('E6').setValue('');
  ws.getRange('F6').setValue('');
  ws.getRange('G6').setValue('');
  ws.getRange('H6').setValue('');

  const registeredSerie = ws.getRange(nextRow, 4).getValue();
  ui.alert('Registrado',
    'Serie ' + registeredSerie + ' de ' + ejercicio + ' registrada.\n' +
    'Siguiente: Serie ' + nextSerie,
    ui.ButtonSet.OK);
}

function limpiarCampos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Registro');

  if (!ws) return;

  ws.getRange('A6').setFormula('=TODAY()');
  ws.getRange('B6').setValue('');
  ws.getRange('C6').setValue('');
  ws.getRange('D6').setValue('');
  ws.getRange('E6').setValue('');
  ws.getRange('F6').setValue('');
  ws.getRange('G6').setValue('');
  ws.getRange('H6').setValue('');

  SpreadsheetApp.getUi().alert('Listo', 'Campos limpiados.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function deshacerUltimo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Registro');
  const ui = SpreadsheetApp.getUi();

  if (!ws) return;

  // Encontrar última fila con datos
  let lastRow = 11;
  while (ws.getRange(lastRow + 1, 1).getValue() !== '' && lastRow < 200) {
    lastRow++;
  }

  if (lastRow < 12) {
    ui.alert('Sin registros', 'No hay registros para deshacer.', ui.ButtonSet.OK);
    return;
  }

  const fecha = ws.getRange(lastRow, 1).getValue();
  const ejercicio = ws.getRange(lastRow, 3).getValue();

  const resp = ui.alert('Confirmar',
    '¿Eliminar el último registro?\n' +
    'Fecha: ' + fecha + '\n' +
    'Ejercicio: ' + ejercicio,
    ui.ButtonSet.YES_NO);

  if (resp === ui.Button.YES) {
    for (let col = 1; col <= 8; col++) {
      ws.getRange(lastRow, col).setValue('');
    }
    ui.alert('Deshecho', 'Último registro eliminado.', ui.ButtonSet.OK);
  }
}

function cargarEjercicios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Registro');
  const wsRut = ss.getSheetByName('Rutinas');
  const ui = SpreadsheetApp.getUi();

  if (!ws || !wsRut) return;

  const diaSeleccionado = String(ws.getRange('B6').getValue()).trim();

  if (!diaSeleccionado) {
    ui.alert('Selecciona un día', 'Primero selecciona un Día/Rutina en la celda B6.', ui.ButtonSet.OK);
    return;
  }

  // Buscar día en Rutinas
  let found = -1;
  for (let r = 1; r <= 100; r++) {
    const val = String(wsRut.getRange(r, 1).getValue());
    if (val.indexOf(diaSeleccionado) !== -1) {
      found = r;
      break;
    }
  }

  if (found === -1) {
    ui.alert('No encontrado', 'No se encontró la rutina: ' + diaSeleccionado, ui.ButtonSet.OK);
    return;
  }

  let msg = 'Ejercicios para ' + diaSeleccionado + ':\n\n';
  let startRow = found + 2;
  let ej = 0;
  while (wsRut.getRange(startRow + ej, 2).getValue() !== '' && ej < 20) {
    msg += (ej + 1) + '. ' + wsRut.getRange(startRow + ej, 2).getValue() + '\n';
    ej++;
  }

  ui.alert('Rutina del día', msg, ui.ButtonSet.OK);
}
