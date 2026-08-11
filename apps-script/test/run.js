'use strict';
// Pruebas con datos SINTÉTICOS (nombres/montos inventados) — seguras para
// un repo público. Para probar contra tus datos reales, ver
// test-real-data.js (no está en git, ver .gitignore).
const { MockSpreadsheet, buildSandbox, loadCode } = require('./mock-gas');

let pass = 0, fail = 0;
function check(label, cond, extra) {
  if (cond) { pass++; console.log('  OK   ' + label); }
  else { fail++; console.log('  FAIL ' + label + (extra !== undefined ? '  -> ' + JSON.stringify(extra) : '')); }
}
function section(title) { console.log('\n== ' + title + ' =='); }

// Reproduce a propósito las mismas rarezas estructurales de la planilla
// real: filas en blanco antes del encabezado, encabezado "Tenant" repetido
// dos veces en Tenants, y una fila futura sin monto en To Landlord.
function syntheticTenants() {
  return [
    [],
    ['Room 1', 'Small', 280, 'Single'],
    ['Room 2', 'Big', 390, 'Couple'],
    [],
    ['Date', 'Day Rent', 'Amount', 'Payment Method', 'Type', 'Detail', 'Room', 'Tenant', '', 'Tenant', 'Rent', 'Bond Held', 'Internet', 'Water', 'Electricity', 'Gas'],
    [new Date(2026, 0, 5), 'Monday', 200, 'Transfer', 'Rent', 'semana 1', 1, 'Ana Test', '', 'Ana Test', 200, 0, 0, 0, 0, 0],
    [new Date(2026, 0, 12), 'Monday', 300, 'Transfer', 'Rent', 'semana 2', 2, 'Beto Test', '', '', '', '', '', '', '', '']
  ];
}
function syntheticLandlord() {
  return [
    [],
    ['Date', 'Day', 'Amount', 'Payment Method', 'Type', 'Detail'],
    [new Date(2026, 0, 1), 'Thursday', 500, 'Transfer', 'Rent', 'semana 1'],
    [new Date(2026, 0, 15), '', '', '', '', ''] // fila futura precargada, sin monto
  ];
}
function syntheticExpenses() {
  return [
    [],
    ['Date', 'Amount', 'Payment Method', 'Type', 'Detail', '', 'House', 100],
    [new Date(2026, 0, 3), 100, 'Card', 'House', 'artículos de cocina']
  ];
}

function freshSandbox() {
  const ss = new MockSpreadsheet();
  ss._seed('Tenants', syntheticTenants());
  ss._seed('To Landlord', syntheticLandlord());
  ss._seed('Expenses', syntheticExpenses());
  const sandbox = loadCode(buildSandbox(ss));
  return { ss, sandbox };
}

function callDoGet(sandbox, params) {
  return JSON.parse(sandbox.doGet({ parameter: params }).getContent());
}
function callDoPost(sandbox, body) {
  return JSON.parse(sandbox.doPost({ postData: { contents: JSON.stringify(body) } }).getContent());
}

const SECRET = '5EGVxQhUJ2RsQ10dFUEkHAuM';

// ---------------------------------------------------------------
section('1. Resumen contra datos sintéticos');
{
  const { sandbox } = freshSandbox();
  const res = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('ok:true', res.ok === true, res);
  check('Ana Test pagó 200', res.summary.tenants.paidByTenant['Ana Test'] === 200);
  check('Beto Test pagó 300', res.summary.tenants.paidByTenant['Beto Test'] === 300);
  check('auto total House es número (100)', res.summary.autoTotals.expenses.House === 100);
  check('Pieza 1 total 200', res.summary.byRoom['Pieza 1'].total === 200);
  check('Pieza 2 total 300', res.summary.byRoom['Pieza 2'].total === 300);
}

// ---------------------------------------------------------------
section('2. Registro de arrendatarios: crear, editar (upsert), no duplicar');
{
  const { ss, sandbox } = freshSandbox();
  check('pestaña Arrendatarios no existe todavía', ss.getSheetByName('Arrendatarios') === null);

  const r1 = callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Carla Test', room: 'Pieza 2', paysUtilities: true, bondAmount: 500 });
  check('primera ficha: ok', r1.ok === true, r1);
  check('primera ficha: updated=false (creación)', r1.created[0].updated === false);
  check('pestaña Arrendatarios se creó sola', ss.getSheetByName('Arrendatarios') !== null);

  const r2 = callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Carla Test', room: 'Pieza 2', paysUtilities: true, rent: 450 });
  check('segunda vez mismo nombre: updated=true (edición)', r2.created[0].updated === true);
  check('sigue habiendo solo 1 fila (no duplicó)', ss.getSheetByName('Arrendatarios').getLastRow() === 2);

  const listed = callDoGet(sandbox, { secret: SECRET, tenants: '1' });
  check('la renta editada quedó guardada (450)', listed.tenants[0]['Renta'] === 450, listed.tenants[0]);
}

// ---------------------------------------------------------------
section('3. Subida de archivo (foto de ID) — no toca Drive real, solo el mock');
{
  const { sandbox } = freshSandbox();
  const tinyPngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
  const res = callDoPost(sandbox, {
    secret: SECRET, action: 'addTenantProfile', name: 'Dana Test', room: 'Pieza 1', paysUtilities: false,
    idPhoto: { base64: tinyPngBase64, mimeType: 'image/png', filename: 'id.png' }
  });
  check('ok', res.ok === true, res);
  const listed = callDoGet(sandbox, { secret: SECRET, tenants: '1' });
  const foto = listed.tenants[0]['Foto ID'];
  check('quedó guardado un link de Drive (mock) en Foto ID', typeof foto === 'string' && foto.indexOf('drive.google.com/mock') !== -1, foto);
  check('DriveApp.createFile se llamó una vez', sandbox.__driveLog.length === 1, sandbox.__driveLog);
}

// ---------------------------------------------------------------
section('4. División de cuenta compartida — con y sin gente en el registro');
{
  const { sandbox } = freshSandbox();

  const r1 = callDoPost(sandbox, { secret: SECRET, action: 'addExpense', date: '2026-08-09', amount: 200, type: 'Electricity' });
  check('4a ok (fallback sin registro)', r1.ok === true, r1);
  check('4a crea 2 registros', r1.created.length === 2, r1.created);
  check('4a pendiente $100 (50%) para el fallback', r1.created[1].amount === 100);

  callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Ana Test', room: 'Pieza 2', paysUtilities: true });
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Beto Test', room: 'Pieza 2', paysUtilities: true });
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Carla Test', room: 'Pieza 1', paysUtilities: false });

  const r2 = callDoPost(sandbox, { secret: SECRET, action: 'addExpense', date: '2026-08-09', amount: 200, type: 'Water' });
  check('4b crea 3 registros (2 personas pagan servicios)', r2.created.length === 3, r2.created);
  const pend = r2.created.slice(1);
  check('4b cada pendiente es $50 (100/2)', pend.every((p) => p.amount === 50), pend);
}

// ---------------------------------------------------------------
section('5. To Landlord: completa la fila prellenada de la fecha, no duplica');
{
  const { ss, sandbox } = freshSandbox();
  const before = ss.getSheetByName('To Landlord').getLastRow();
  const res = callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2026-01-15', amount: 500, type: 'Rent' });
  check('ok', res.ok === true, res);
  check('NO agregó fila nueva (completó la prellenada del 15/01)', ss.getSheetByName('To Landlord').getLastRow() === before, { before });

  const before2 = ss.getSheetByName('To Landlord').getLastRow();
  callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2099-01-01', amount: 999, type: 'Other' });
  check('fecha sin fila prellenada SÍ agrega una fila nueva', ss.getSheetByName('To Landlord').getLastRow() === before2 + 1);
}

// ---------------------------------------------------------------
section('6. Resumen nuevo: renta/servicios pagados por tipo, bond a devolver, renta esperada, historial completo');
{
  const { ss, sandbox } = freshSandbox();

  // paidByType: suma de Rent pagado por TODOS los tenants (Ana 200 + Beto 300 del seed sintético).
  const res1 = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('paidByType.Rent = 500 (Ana + Beto)', res1.summary.tenants.paidByType['Rent'] === 500, res1.summary.tenants.paidByType);
  check('paidByTypePerTenant.Ana Test.Rent = 200', res1.summary.tenants.paidByTypePerTenant['Ana Test']['Rent'] === 200);

  // landlord.expected: un solo pago sembrado el 2026-01-01 -> sinceDate debe coincidir.
  check('landlord.expected.sinceDate = 01/01/2026', res1.summary.landlord.expected.sinceDate === '01/01/2026', res1.summary.landlord.expected);
  check('landlord.expected.amount es número >= 0', typeof res1.summary.landlord.expected.amount === 'number' && res1.summary.landlord.expected.amount >= 0);

  // bond a devolver: ficha con Bond Monto 500, luego un Refund de 150 -> quedan 350.
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Elena Test', room: 'Pieza 1', bondAmount: 500 });
  const res2a = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('bond.byTenant.Elena Test = 500 antes de cualquier refund', res2a.summary.bond.byTenant['Elena Test'] === 500, res2a.summary.bond);

  callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-02-01', amount: 150, tenant: 'Elena Test', type: 'Refund' });
  const res2b = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('bond.byTenant.Elena Test = 350 después del refund de 150', res2b.summary.bond.byTenant['Elena Test'] === 350, res2b.summary.bond);
  check('bond.totalHeld incluye los 350 de Elena', res2b.summary.bond.totalHeld >= 350, res2b.summary.bond);

  // full=1 debe traer TODAS las filas de Tenants, no solo las últimas 5.
  for (let i = 0; i < 6; i++) {
    callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-03-0' + (i + 1), amount: 10 + i, tenant: 'Fill Test', type: 'Rent' });
  }
  const shortList = callDoGet(sandbox, { secret: SECRET, recent: '2' });
  const fullList = callDoGet(sandbox, { secret: SECRET, full: '1' });
  check('recent=2 trae solo 2 filas de Tenants', shortList.recent['Tenants'].length === 2, shortList.recent['Tenants'].length);
  check('full=1 trae más filas que recent=2', fullList.recent['Tenants'].length > shortList.recent['Tenants'].length, fullList.recent['Tenants'].length);
}

// ---------------------------------------------------------------
section('7. Borrar ficha por ID (no por nombre) — deja intacta la otra ficha con el mismo nombre');
{
  const { ss, sandbox } = freshSandbox();
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Juan Plata', room: 'Pieza 2', bondAmount: 100 });
  // Como upsertRegistryRow matchea por nombre, para simular el bug real (dos
  // fichas con el mismo nombre, ej. por un espacio invisible distinto) se
  // agrega la segunda fila directo al mock, saltándose el upsert.
  const sheet = ss.getSheetByName('Arrendatarios');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row2 = headers.map((h) => (h === 'ID' ? 'dup-0002' : h === 'Nombre' ? 'Juan Plata' : h === 'Bond Monto' ? 200 : ''));
  sheet.appendRow(row2);

  const before = callDoGet(sandbox, { secret: SECRET, tenants: '1' });
  check('hay 2 fichas "Juan Plata" antes de borrar', before.tenants.filter((t) => t['Nombre'] === 'Juan Plata').length === 2, before.tenants);
  const dupId = before.tenants[1]['ID'];

  const del = callDoPost(sandbox, { secret: SECRET, action: 'deleteTenantProfile', id: dupId });
  check('deleteTenantProfile: ok', del.ok === true, del);

  const after = callDoGet(sandbox, { secret: SECRET, tenants: '1' });
  check('queda solo 1 ficha "Juan Plata"', after.tenants.filter((t) => t['Nombre'] === 'Juan Plata').length === 1, after.tenants);
  check('quedó la ficha correcta (Bond Monto 100, no 200)', after.tenants[0]['Bond Monto'] === 100, after.tenants[0]);

  const delAgain = callDoPost(sandbox, { secret: SECRET, action: 'deleteTenantProfile', id: dupId });
  check('borrar un ID ya borrado devuelve ok:false con error claro', delAgain.ok === false && !!delAgain.error, delAgain);
}

// ---------------------------------------------------------------
section('8. Renta esperada ancla en el primer pago tipo Rent, no en un Bond anterior');
{
  // OJO: freshSandbox() trae un pago "Rent" sintético precargado el
  // 2026-01-01 (ver syntheticLandlord) — para esta prueba se necesita un
  // "To Landlord" realmente vacío, si no ese pago sintético (no el Bond de
  // esta prueba) sería el que falsamente "pasa" la aserción.
  const ss = new MockSpreadsheet();
  ss._seed('Tenants', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail', 'Room', 'Tenant']]);
  ss._seed('To Landlord', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  ss._seed('Expenses', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  const sandbox = loadCode(buildSandbox(ss));
  // Bond pagado primero (2026-01-01), renta real empieza después (2026-02-24).
  callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2026-01-01', amount: 1600, type: 'Bond' });
  callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2026-02-24', amount: 1600, type: 'Rent' });
  const res = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('landlord.expected.sinceDate = 24/02/2026 (ignora el Bond del 01/01)', res.summary.landlord.expected.sinceDate === '24/02/2026', res.summary.landlord.expected);
}

// ---------------------------------------------------------------
section('9. rowToObject normaliza fechas tipo Date real a texto dd/MM/yyyy (evita timestamps con hora)');
{
  const { ss, sandbox } = freshSandbox();
  // Simula una fila donde la fecha se guardó como Date real (ej: tecleada a
  // mano en Sheets), no como el texto que escribe appendRow.
  ss._seed('Expenses', [
    [],
    ['Date', 'Amount', 'Payment Method', 'Type', 'Detail'],
    [new Date(2026, 1, 24, 15, 30, 0), 50, 'Card', 'House', 'con hora en la celda']
  ]);
  const sandbox2 = loadCode(buildSandbox(ss));
  const res = callDoGet(sandbox2, { secret: SECRET, full: '1' });
  const dateStr = res.recent['Expenses'][0]['Date'];
  check('fecha queda como "24/02/2026", sin hora ni "T"/"GMT"', dateStr === '24/02/2026', dateStr);
}

// ---------------------------------------------------------------
section('10. Lo mismo pero para "Bond Fecha" / "Fecha Inicio" del registro (bug real encontrado en la planilla de Nicolás)');
{
  const { ss, sandbox } = freshSandbox();
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantProfile', name: 'Carla Test', room: 'Pieza 1', bondAmount: 500 });
  // addTenantProfile siempre escribe texto — para reproducir el caso real
  // (fichas viejas con la fecha tecleada directo en Sheets como Date real)
  // se sobreescribe la celda a mano con un objeto Date.
  const sheet = ss.getSheetByName('Arrendatarios');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const bondFechaCol = headers.indexOf('Bond Fecha') + 1;
  const fechaInicioCol = headers.indexOf('Fecha Inicio') + 1;
  sheet.getRange(2, bondFechaCol).setValue(new Date(2026, 2, 13, 16, 0, 0));
  sheet.getRange(2, fechaInicioCol).setValue(new Date(2026, 7, 10, 16, 0, 0));

  const res = callDoGet(sandbox, { secret: SECRET, tenants: '1' });
  check('Bond Fecha = "13/03/2026" (no timestamp ISO)', res.tenants[0]['Bond Fecha'] === '13/03/2026', res.tenants[0]['Bond Fecha']);
  check('Fecha Inicio = "10/08/2026" (no timestamp ISO)', res.tenants[0]['Fecha Inicio'] === '10/08/2026', res.tenants[0]['Fecha Inicio']);
}

// ---------------------------------------------------------------
section('11. Editar un registro (gasto) — cambia los valores en la fila correcta, no crea una fila nueva');
{
  const { ss, sandbox } = freshSandbox();
  const before = ss.getSheetByName('Expenses').getLastRow();
  callDoPost(sandbox, { secret: SECRET, action: 'addExpense', date: '2026-08-01', amount: 77, type: 'House', detail: 'original', autoSplit: false });
  const afterAdd = ss.getSheetByName('Expenses').getLastRow();
  check('agregar el gasto sumó exactamente 1 fila', afterAdd === before + 1, { before, afterAdd });

  const full = callDoGet(sandbox, { secret: SECRET, full: '1' });
  const row = full.recent['Expenses'].find((r) => r['Detail'] === 'original');
  check('la fila nueva trae __row', typeof row.__row === 'number', row);

  const edit = callDoPost(sandbox, {
    secret: SECRET, action: 'editRecord', sheet: 'Expenses', row: row.__row,
    expectedAmount: row['Amount'], expectedDate: row['Date'],
    date: '2026-08-02', amount: 99, type: 'Internet', detail: 'corregido', paymentMethod: 'Cash'
  });
  check('editRecord: ok', edit.ok === true, edit);
  check('NO agregó una fila nueva al editar', ss.getSheetByName('Expenses').getLastRow() === afterAdd, ss.getSheetByName('Expenses').getLastRow());

  const full2 = callDoGet(sandbox, { secret: SECRET, full: '1' });
  const edited = full2.recent['Expenses'].find((r) => r.__row === row.__row);
  check('el monto quedó en 99', edited['Amount'] === 99, edited);
  check('el detalle quedó en "corregido"', edited['Detail'] === 'corregido', edited);
  check('el tipo quedó en Internet', edited['Type'] === 'Internet', edited);
  check('la fecha quedó en 02/08/26', edited['Date'] === '02/08/26', edited['Date']);
}

// ---------------------------------------------------------------
section('12. Borrar un registro (deleteRecord) — vacía la fila sin eliminarla ni afectar otras filas');
{
  const { ss, sandbox } = freshSandbox();
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-08-05', amount: 55, tenant: 'Borrar Test', type: 'Rent', room: '1' });
  // Se agrega un SEGUNDO pago después, para que la fila que se va a borrar no
  // quede como la última de la hoja — así se puede verificar que
  // sheet.getLastRow() no cambia (una fila vaciada sigue existiendo, solo
  // que en Sheets/el mock una fila en blanco AL FINAL de la hoja hace que
  // getLastRow() "encoja" — comportamiento real de Sheets, no es lo que se
  // quiere probar acá).
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-08-06', amount: 66, tenant: 'Despues Test', type: 'Rent', room: '2' });
  const rowCountBefore = ss.getSheetByName('Tenants').getLastRow();

  const full = callDoGet(sandbox, { secret: SECRET, full: '1' });
  const row = full.recent['Tenants'].find((r) => r['Tenant'] === 'Borrar Test');

  const del = callDoPost(sandbox, {
    secret: SECRET, action: 'deleteRecord', sheet: 'Tenants', row: row.__row,
    expectedAmount: row['Amount'], expectedDate: row['Date']
  });
  check('deleteRecord: ok', del.ok === true, del);
  check('NO cambió la cantidad de filas de la hoja (no se movieron otras filas)', ss.getSheetByName('Tenants').getLastRow() === rowCountBefore, ss.getSheetByName('Tenants').getLastRow());

  const summary = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('el pago borrado ya no cuenta en paidByTenant', !summary.summary.tenants.paidByTenant['Borrar Test'], summary.summary.tenants.paidByTenant);
  check('el pago posterior (Despues Test) no se vio afectado', summary.summary.tenants.paidByTenant['Despues Test'] === 66, summary.summary.tenants.paidByTenant);

  // No debe tocar la fila del resumen lateral de Room 2 (Ana Test es quien
  // queda en esa columna en el seed sintético — ver syntheticTenants()).
  check('el resumen lateral de Tenants (Room 2) sigue intacto tras borrar otra fila', summary.summary.autoTotals.tenantsSideSnapshot && summary.summary.autoTotals.tenantsSideSnapshot.tenant === 'Ana Test', summary.summary.autoTotals.tenantsSideSnapshot);
}

// ---------------------------------------------------------------
section('13. editRecord/deleteRecord rechazan la operación si la fila cambió desde que se cargó la pantalla');
{
  const { ss, sandbox } = freshSandbox();
  callDoPost(sandbox, { secret: SECRET, action: 'addExpense', date: '2026-08-01', amount: 40, type: 'House', autoSplit: false });
  const full = callDoGet(sandbox, { secret: SECRET, full: '1' });
  const row = full.recent['Expenses'].find((r) => r['Amount'] === 40);

  const edit = callDoPost(sandbox, {
    secret: SECRET, action: 'editRecord', sheet: 'Expenses', row: row.__row,
    expectedAmount: 999, // monto incorrecto a propósito — no coincide con lo que hay en la fila
    amount: 1
  });
  check('editRecord con expectedAmount equivocado: ok:false', edit.ok === false && !!edit.error, edit);

  const full2 = callDoGet(sandbox, { secret: SECRET, full: '1' });
  const stillThere = full2.recent['Expenses'].find((r) => r.__row === row.__row);
  check('el monto NO cambió (el rechazo evitó la edición)', stillThere['Amount'] === 40, stillThere);
}

// ---------------------------------------------------------------
section('14. Nota manual "Rent Paid from Nico" en To Landlord — se lee tal cual, sin importar de qué lado cae el valor');
{
  // Reproduce el layout real encontrado en la planilla de Nicolás: columnas
  // extra en "To Landlord" donde anota notas sueltas, una por fila, con el
  // valor a veces a la izquierda y a veces a la derecha de la etiqueta.
  const ss = new MockSpreadsheet();
  ss._seed('Tenants', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail', 'Room', 'Tenant']]);
  ss._seed('To Landlord', [
    [],
    ['Date', 'Day', 'Amount', 'Payment Method', 'Type', 'Detail', 'Total Rent Paid'],
    [new Date(2026, 2, 17), 'Tuesday', 1600, 'Transfer', 'Rent', '', 'Room1', 6460], // valor a la DERECHA de la etiqueta
    [new Date(2026, 3, 14), 'Tuesday', 1600, 'Transfer', 'Rent', '', 3360, 'Rent Paid from Nico'] // valor a la IZQUIERDA de la etiqueta
  ]);
  ss._seed('Expenses', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  const sandbox = loadCode(buildSandbox(ss));

  const res = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  const note = res.summary.landlord.rentPaidFromNicoNote;
  check('rentPaidFromNicoNote no es null', note != null, note);
  check('valor = 3360 (lo encontró aunque estaba a la izquierda)', note && note.value === 3360, note);
}

// ---------------------------------------------------------------
section('15. Rent Paid from Nico calculado EN VIVO (confirmado por Nicolás: = renta pagada al arrendador − renta pagada por tenants)');
{
  // Sandbox realmente vacío (no freshSandbox()) — el seed sintético trae de
  // fábrica $500 de Rent tanto en Landlord como en Tenants, lo que
  // distorsionaría los valores absolutos que se están verificando acá.
  const ss = new MockSpreadsheet();
  ss._seed('Tenants', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail', 'Room', 'Tenant']]);
  ss._seed('To Landlord', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  ss._seed('Expenses', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  const sandbox = loadCode(buildSandbox(ss));
  // Reproduce los números reales verificados en producción el 2026-08-11:
  // landlord.byType.Rent = 19200, tenants.paidByType.Rent = 15840 -> 3360.
  for (let i = 0; i < 12; i++) {
    callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2026-0' + (1 + (i % 9)) + '-0' + (1 + (i % 9)), amount: 1600, type: 'Rent' });
  }
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-05-01', amount: 15840, tenant: 'Real Test', type: 'Rent' });

  const res = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  check('landlord.byType.Rent = 19200 (12 x 1600)', res.summary.landlord.byType['Rent'] === 19200, res.summary.landlord.byType);
  check('tenants.paidByType.Rent = 15840', res.summary.tenants.paidByType['Rent'] === 15840, res.summary.tenants.paidByType);
  check('landlord.rentPaidFromNico = 3360 (19200 - 15840)', res.summary.landlord.rentPaidFromNico === 3360, res.summary.landlord.rentPaidFromNico);
}

// ---------------------------------------------------------------
section('16. Margen neto sin bond — el bond no debe distorsionar el flujo de caja real');
{
  const ss = new MockSpreadsheet();
  ss._seed('Tenants', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail', 'Room', 'Tenant']]);
  ss._seed('To Landlord', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  ss._seed('Expenses', [[], ['Date', 'Amount', 'Payment Method', 'Type', 'Detail']]);
  const sandbox = loadCode(buildSandbox(ss));

  // Renta real: tenant paga 1000, Nico paga 800 al landlord, gasto de casa 100.
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-05-01', amount: 1000, tenant: 'Margin Test', type: 'Rent' });
  callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2026-05-01', amount: 800, type: 'Rent' });
  callDoPost(sandbox, { secret: SECRET, action: 'addExpense', date: '2026-05-01', amount: 100, type: 'House', autoSplit: false });
  // Bond: tenant deposita 500, pero Nico todavía no le ha pasado todo al
  // landlord (solo 300) — a propósito DISTINTO en cada lado, para probar
  // que excluir el bond realmente cambia el resultado (si fueran iguales
  // se cancelarían solos y el test no probaría nada).
  callDoPost(sandbox, { secret: SECRET, action: 'addTenantPayment', date: '2026-05-01', amount: 500, tenant: 'Margin Test', type: 'Bond Held' });
  callDoPost(sandbox, { secret: SECRET, action: 'addLandlordPayment', date: '2026-05-01', amount: 300, type: 'Bond' });

  const res = callDoGet(sandbox, { secret: SECRET, summary: '1' });
  const bp = res.summary.bigPicture;
  check('margenNeto (con bond) = 300 (1500 - 1100 - 100)', bp.margenNeto === 300, bp);
  check('margenNetoSinBond = 100 (1000 - 800 - 100, el bond no debe aparecer)', bp.margenNetoSinBond === 100, bp);
}

// ---------------------------------------------------------------
console.log('\n' + '='.repeat(50));
console.log(pass + ' OK, ' + fail + ' FAIL');
if (fail) process.exit(1);

console.log('\nNota: esto valida la lógica con datos inventados. Para probar');
console.log('contra tus datos reales de forma segura (no se sube a git),');
console.log('pide una foto de datos actual y corre test-real-data.js.');
