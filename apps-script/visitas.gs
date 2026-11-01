// Contador de aperturas de DiliApp.
// Proyecto de Apps Script aparte ("DiliApp Visitas"), separado del backend de
// diligenciadores para no redesplegar ese. Publicado como app web:
// ejecutar como yo, acceso cualquiera. Cada POST anota una apertura en la
// pestaña "Visitas" del mismo Sheet, con resumen por semana y por mes.

const SHEET_ID = '1_rp7MvHx7-iftbfsAVWTHSSGSjFip-fa018SvQfGLMA';
const TZ = 'America/Argentina/Buenos_Aires';
const MAX_FILAS = 100000; // tope contra llenado malicioso de la hoja

function doPost() {
  try {
    const hoja = hojaVisitas_();
    if (hoja.getLastRow() >= MAX_FILAS) {
      return ContentService.createTextOutput('ERROR:limite');
    }
    const ahora = new Date();
    const diaSemana = Number(Utilities.formatDate(ahora, TZ, 'u')); // 1 = lunes
    const lunes = new Date(ahora.getTime() - (diaSemana - 1) * 86400000);
    hoja.appendRow([
      Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy HH:mm'),
      Utilities.formatDate(lunes, TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(ahora, TZ, 'yyyy-MM'),
    ]);
  } catch (err) {
    console.error('doPost error:', err);
  }
  return ContentService.createTextOutput('OK');
}

function hojaVisitas_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hoja = ss.getSheetByName('Visitas');
  if (hoja) return hoja;

  hoja = ss.insertSheet('Visitas');
  // Texto plano para que Sheets no convierta "2026-11" en fecha
  hoja.getRange('A:C').setNumberFormat('@');
  hoja.getRange('A1:C1').setValues([['Fecha y hora', 'Semana (lunes)', 'Mes']]);
  hoja.getRange('E1').setFormula(
    '=QUERY(B:B, "select B, count(B) where B is not null group by B order by B desc ' +
    'label B \'Semana (lunes)\', count(B) \'Aperturas\'", 1)'
  );
  hoja.getRange('H1').setFormula(
    '=QUERY(C:C, "select C, count(C) where C is not null group by C order by C desc ' +
    'label C \'Mes\', count(C) \'Aperturas\'", 1)'
  );
  hoja.getRange('A1:I1').setFontWeight('bold');
  hoja.setFrozenRows(1);
  return hoja;
}

// Correr una vez a mano desde el editor: crea la pestaña y pide la autorización.
function crearHoja() {
  hojaVisitas_();
}
