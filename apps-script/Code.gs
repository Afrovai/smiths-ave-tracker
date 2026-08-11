/**
 * 10 Smiths Avenue, Redcliffe — backend para la app de registro rápido.
 *
 * Qué hace:
 *  - Recibe POST desde la app web (GitHub Pages) y agrega filas nuevas a las
 *    pestañas "Expenses", "Tenants" o "To Landlord" de la planilla.
 *  - Mantiene una pestaña nueva "Arrendatarios" (la crea sola si no existe)
 *    con la ficha de cada persona: pieza, renta, frecuencia de pago, si paga
 *    servicios, bond, foto de ID y comprobante de ingreso (subidos a una
 *    carpeta privada de Drive, solo visibles para el dueño de la cuenta).
 *  - Si un gasto es de un tipo "compartible" (luz, agua, gas, internet), la
 *    divide entre quien(es) estén marcados como "Paga Servicios: Sí" en el
 *    registro de arrendatarios (ya no depende de un nombre fijo en el código).
 *  - En "To Landlord" (que ya trae fechas futuras precargadas con el monto
 *    vacío), si existe una fila con esa misma fecha y sin monto, la completa
 *    ahí en vez de agregar una fila nueva y duplicada.
 *  - NO toca ni recalcula las columnas/celdas de totales automáticos que ya
 *    existen en tu planilla (los totales por categoría en "Expenses", el
 *    resumen lateral de "Tenants") — el dashboard los LEE tal cual están,
 *    no los recalcula por su cuenta.
 *
 * Instrucciones de despliegue: ver apps-script/README.md en este mismo repo.
 */

// ID de la planilla "10 Smiths Avenue, Redcliffe"
const SHEET_ID = '1Z8asY7OdehqG5yCrcw2uB9xP9ZE7Zb16aXpasg4W1tA';

// Debe coincidir EXACTO con el secreto que configures en la app web.
const SECRET = '5EGVxQhUJ2RsQ10dFUEkHAuM';

// Tipos de gasto que se dividen automáticamente entre quienes "pagan
// servicios" en el registro de arrendatarios, y qué fracción del total
// les corresponde a ellos en conjunto (se reparte en partes iguales entre
// todos los que pagan servicios).
const SPLIT_RULES = {
  'Electricity': 0.5,
  'Water': 0.5,
  'Gas': 0.5,
  'Internet': 0.5
};

// Fallback SOLO si el registro de "Arrendatarios" está vacío todavía (nadie
// marcado como "Paga Servicios: Sí") — así la app sigue funcionando igual
// que antes mientras cargas las fichas por primera vez.
const DEFAULT_SPLIT_TENANT = 'Juan Plata';

// Piezas de la casa — renta de referencia y frecuencia habitual. Editable
// por ficha individual en el registro (esto es solo el valor por defecto).
const ROOMS = {
  'Pieza 1': { rentPerWeek: 280, frequency: 'Mensual' },
  'Pieza 2': { rentPerWeek: 390, frequency: 'Quincenal' }
};
const ROOM_TO_NUMBER = { 'Pieza 1': 1, 'Pieza 2': 2 };
const NUMBER_TO_ROOM_LABEL = { '1': 'Pieza 1', '2': 'Pieza 2', 1: 'Pieza 1', 2: 'Pieza 2' };
const HOUSE_RENT_PER_WEEK = 800; // total a pagar al arrendador, se paga quincenal

const REGISTRY_SHEET = 'Arrendatarios';
const REGISTRY_HEADERS = [
  'ID', 'Nombre', 'Pieza', 'Renta', 'Frecuencia', 'Paga Servicios',
  'Bond Monto', 'Bond Fecha', 'Bond Detalle', 'Foto ID', 'Comprobante Ingreso',
  'Fecha Inicio', 'Estado'
];
const DOCS_FOLDER_NAME = '10 Smiths Avenue - Documentos (privado)';

function doPost(e) {
  // Evita que dos solicitudes casi simultáneas (ej: doble click en "Guardar"
  // antes de que la primera termine) lean el mismo estado "no existe todavía"
  // y ambas terminen creando una ficha duplicada en Arrendatarios.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut({ ok: false, error: 'El sistema está ocupado con otra solicitud, intenta de nuevo en unos segundos.' });
  }
  try {
    if (!e.postData || !e.postData.contents) {
      return jsonOut({ ok: false, error: 'Sin datos en la solicitud' });
    }
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) {
      return jsonOut({ ok: false, error: 'No autorizado' });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const created = [];

    if (body.action === 'addExpense') {
      requireFields(body, ['date', 'amount', 'type']);
      appendRow(ss, 'Expenses', {
        'Date': body.date,
        'Amount': Number(body.amount),
        'Payment Method': body.paymentMethod || '',
        'Type': body.type,
        'Detail': body.detail || ''
      });
      created.push({ sheet: 'Expenses', amount: Number(body.amount) });

      const frac = SPLIT_RULES[body.type];
      if (frac && body.autoSplit !== false) {
        const pct = body.splitPct != null ? Number(body.splitPct) : frac;
        const totalShare = round2(Number(body.amount) * pct);
        let payers = getBillPayers(ss);
        if (!payers.length) {
          // Registro vacío todavía — modo compatible con la versión anterior.
          payers = [{ name: body.splitTenant || DEFAULT_SPLIT_TENANT, room: '' }];
        }
        const perPayerShare = round2(totalShare / payers.length);
        payers.forEach(function (p) {
          appendRow(ss, 'Tenants', {
            'Date': body.date,
            'Amount': perPayerShare,
            'Payment Method': 'Pending',
            'Type': body.type,
            'Detail': 'Pendiente: parte de cuenta de ' + body.type +
                      ' ($' + body.amount + ' del ' + formatIsoDate(body.date, 'dd/MM/yyyy') + ')' +
                      (p.room ? ' — ' + p.room : ''),
            'Tenant': p.name,
            'Room': p.room ? String(ROOM_TO_NUMBER[p.room] || '') : ''
          });
          created.push({ sheet: 'Tenants', amount: perPayerShare, tenant: p.name, pending: true });
        });
      }

    } else if (body.action === 'addTenantPayment') {
      requireFields(body, ['date', 'amount', 'tenant']);
      appendRow(ss, 'Tenants', {
        'Date': body.date,
        'Amount': Number(body.amount),
        'Payment Method': body.paymentMethod || '',
        'Type': body.type || 'Rent',
        'Detail': body.detail || '',
        'Room': body.room || '',
        'Tenant': body.tenant
      });
      created.push({ sheet: 'Tenants', amount: Number(body.amount), tenant: body.tenant });

    } else if (body.action === 'addLandlordPayment') {
      requireFields(body, ['date', 'amount']);
      appendRow(ss, 'To Landlord', {
        'Date': body.date,
        'Amount': Number(body.amount),
        'Payment Method': body.paymentMethod || '',
        'Type': body.type || 'Rent',
        'Detail': body.detail || ''
      });
      created.push({ sheet: 'To Landlord', amount: Number(body.amount) });

    } else if (body.action === 'addTenantProfile') {
      requireFields(body, ['name', 'room']);
      const sheet = ensureRegistrySheet(ss);
      const fields = {
        'Nombre': body.name,
        'Pieza': body.room,
        'Renta': body.rent != null && body.rent !== '' ? Number(body.rent) : (ROOMS[body.room] ? ROOMS[body.room].rentPerWeek : ''),
        'Frecuencia': body.frequency || (ROOMS[body.room] ? ROOMS[body.room].frequency : ''),
        'Paga Servicios': body.paysUtilities ? 'Sí' : 'No',
        'Bond Monto': body.bondAmount != null && body.bondAmount !== '' ? Number(body.bondAmount) : '',
        'Bond Fecha': body.bondDate ? formatIsoDate(body.bondDate, 'dd/MM/yyyy') : '',
        'Bond Detalle': body.bondDetail || '',
        'Fecha Inicio': body.startDate ? formatIsoDate(body.startDate, 'dd/MM/yyyy') : formatIsoDate(new Date().toISOString().slice(0, 10), 'dd/MM/yyyy'),
        'Estado': body.status || 'Activo'
      };
      if (body.idPhoto && body.idPhoto.base64) {
        fields['Foto ID'] = saveUploadedFile(body.idPhoto, body.name + ' - ID');
      }
      if (body.incomeProof && body.incomeProof.base64) {
        fields['Comprobante Ingreso'] = saveUploadedFile(body.incomeProof, body.name + ' - Ingreso');
      }
      const result = upsertRegistryRow(sheet, body.name, fields);
      created.push({ sheet: 'Arrendatarios', tenant: body.name, updated: result.updated });

    } else if (body.action === 'deleteTenantProfile') {
      requireFields(body, ['id']);
      const sheet = ss.getSheetByName(REGISTRY_SHEET);
      const result = sheet ? deleteRegistryRowById(sheet, body.id) : { deleted: false };
      if (!result.deleted) {
        return jsonOut({ ok: false, error: 'No se encontró esa ficha (puede que ya se haya borrado).' });
      }
      created.push({ sheet: 'Arrendatarios', deletedId: body.id });

    } else if (body.action === 'editRecord' || body.action === 'deleteRecord') {
      requireFields(body, ['sheet', 'row']);
      const editable = ['Expenses', 'Tenants', 'To Landlord'];
      if (editable.indexOf(body.sheet) === -1) {
        return jsonOut({ ok: false, error: 'Esa pestaña no se puede editar desde la app: ' + body.sheet });
      }
      const targetSheet = ss.getSheetByName(body.sheet);
      if (!targetSheet) return jsonOut({ ok: false, error: 'No existe la pestaña ' + body.sheet });

      const result = body.action === 'deleteRecord'
        ? clearRecordRow(targetSheet, body)
        : editRecordRow(targetSheet, body);
      if (!result.ok) return jsonOut({ ok: false, error: result.error });
      created.push({ sheet: body.sheet, row: body.row, action: body.action });

    } else {
      return jsonOut({ ok: false, error: 'Acción desconocida: ' + body.action });
    }

    return jsonOut({ ok: true, created: created });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * GET ?secret=...&recent=5   -> últimas filas de cada pestaña.
 * GET ?secret=...&full=1     -> TODAS las filas de cada pestaña (para el
 *                               historial completo de cada tab de la app).
 * GET ?secret=...&summary=1  -> totales agregados para el dashboard.
 * GET ?secret=...&tenants=1  -> fichas del registro de arrendatarios.
 */
function doGet(e) {
  try {
    const secret = e.parameter.secret;
    if (secret !== SECRET) return jsonOut({ ok: false, error: 'No autorizado' });

    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (e.parameter.summary) {
      return jsonOut({ ok: true, summary: computeSummary(ss) });
    }
    if (e.parameter.tenants) {
      return jsonOut({ ok: true, tenants: getRegistryRows(ss), rooms: ROOMS });
    }

    const n = e.parameter.full ? Number.MAX_SAFE_INTEGER : Number(e.parameter.recent || 5);
    const out = {};
    ['Expenses', 'Tenants', 'To Landlord'].forEach(function (name) {
      out[name] = lastRows(ss, name, n);
    });
    return jsonOut({ ok: true, recent: out });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ---------- registro de arrendatarios ----------

function ensureRegistrySheet(ss) {
  let sheet = ss.getSheetByName(REGISTRY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_SHEET);
    sheet.getRange(1, 1, 1, REGISTRY_HEADERS.length).setValues([REGISTRY_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Si ya existe una ficha con ese nombre (sin importar mayúsculas), la
// actualiza en vez de crear una duplicada — así "Renta"/"Frecuencia"/etc.
// quedan editables solo con volver a enviar el formulario con el mismo nombre.
function upsertRegistryRow(sheet, name, fields) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const nameIdx = headers.indexOf('Nombre');
  let targetRow = -1;

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][nameIdx]).trim().toLowerCase() === name.trim().toLowerCase()) {
        targetRow = 2 + i;
        break;
      }
    }
  }

  if (targetRow === -1) {
    const row = headers.map(function (h) {
      if (h === 'ID') return Utilities.getUuid().slice(0, 8);
      return Object.prototype.hasOwnProperty.call(fields, h) ? fields[h] : '';
    });
    sheet.appendRow(row);
    return { updated: false };
  }

  headers.forEach(function (h, c) {
    if (Object.prototype.hasOwnProperty.call(fields, h)) {
      sheet.getRange(targetRow, c + 1).setValue(fields[h]);
    }
  });
  return { updated: true };
}

// Borra UNA ficha por su "ID" (no por nombre) — así, si por error quedaron
// dos fichas con el mismo nombre (ej: doble click antes de este fix), se
// puede borrar específicamente la que sobra sin arriesgar la otra.
function deleteRegistryRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { deleted: false };
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idIdx = headers.indexOf('ID');
  if (idIdx === -1) return { deleted: false };
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idIdx]).trim() === String(id).trim()) {
      sheet.deleteRow(2 + i);
      return { deleted: true };
    }
  }
  return { deleted: false };
}

function getRegistryRows(ss) {
  const sheet = ss.getSheetByName(REGISTRY_SHEET);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(function (row) { return rowToObject(headers, row); });
}

function getActiveTenants(ss) {
  return getRegistryRows(ss).filter(function (r) { return r['Estado'] === 'Activo'; });
}

function getBillPayers(ss) {
  return getActiveTenants(ss)
    .filter(function (r) { return r['Paga Servicios'] === 'Sí'; })
    .map(function (r) { return { name: r['Nombre'], room: r['Pieza'] }; });
}

// ---------- archivos (foto de ID / comprobante de ingreso) ----------

// Los archivos quedan en una carpeta de Drive SIN compartir — visibles solo
// para tu cuenta de Google, igual que cualquier archivo tuyo por defecto.
function getDocsFolder() {
  const it = DriveApp.getFoldersByName(DOCS_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DOCS_FOLDER_NAME);
}

function saveUploadedFile(fileObj, namePrefix) {
  const bytes = Utilities.base64Decode(fileObj.base64);
  const blob = Utilities.newBlob(bytes, fileObj.mimeType || 'image/jpeg', namePrefix + ' - ' + (fileObj.filename || 'archivo'));
  const file = getDocsFolder().createFile(blob);
  return file.getUrl();
}

/**
 * Junta todos los movimientos reales (sin las filas futuras vacías de
 * "To Landlord") y arma los totales para el dashboard: cuánto se le ha
 * pagado al arrendador, cuánto se ha gastado en la casa (por tipo), cuánto
 * ha pagado/debe cada inquilino, lo mismo por pieza, y además LEE tal cual
 * (sin recalcular) los totales automáticos que ya existen en tu planilla.
 */
function computeSummary(ss) {
  const expenses = allRows(ss, 'Expenses').filter(function (r) { return r['Amount'] !== '' && r['Amount'] != null; });
  const tenantRows = allRows(ss, 'Tenants').filter(function (r) { return r['Amount'] !== '' && r['Amount'] != null; });
  const landlordRows = allRows(ss, 'To Landlord').filter(function (r) { return r['Amount'] !== '' && r['Amount'] != null; });

  const expenseTotal = sumField(expenses, 'Amount');
  const expenseByType = sumByKey(expenses, 'Type', 'Amount');

  const landlordTotal = sumField(landlordRows, 'Amount');
  const landlordByType = sumByKey(landlordRows, 'Type', 'Amount');

  const paidByTenant = {};
  const pendingByTenant = {};
  const byTypePerTenant = {};
  // Igual que byTypePerTenant pero SOLO lo efectivamente pagado (no lo
  // pendiente) — es lo que se muestra en el resumen por arrendatario.
  const paidByTypePerTenant = {};
  // Lo mismo pero sumado entre todos los arrendatarios (ej: renta total
  // pagada por los tenants de la casa, para la pestaña Landlord).
  const paidByType = {};
  const utilityChargedToTenants = { Internet: 0, Water: 0, Electricity: 0, Gas: 0 };
  // Solo movimientos YA PAGADOS (no pendientes) — "cuánto me ha dado" cada
  // pieza, no cuánto le falta por pagar.
  const byRoom = {};

  tenantRows.forEach(function (r) {
    const tenant = r['Tenant'];
    const amt = Number(r['Amount']) || 0;
    const type = r['Type'] || 'Other';
    const isPending = r['Payment Method'] === 'Pending';

    if (tenant) {
      byTypePerTenant[tenant] = byTypePerTenant[tenant] || {};
      byTypePerTenant[tenant][type] = round2((byTypePerTenant[tenant][type] || 0) + amt);

      if (isPending) {
        pendingByTenant[tenant] = round2((pendingByTenant[tenant] || 0) + amt);
      } else {
        paidByTenant[tenant] = round2((paidByTenant[tenant] || 0) + amt);
        paidByTypePerTenant[tenant] = paidByTypePerTenant[tenant] || {};
        paidByTypePerTenant[tenant][type] = round2((paidByTypePerTenant[tenant][type] || 0) + amt);
        paidByType[type] = round2((paidByType[type] || 0) + amt);
      }
    }

    if (utilityChargedToTenants.hasOwnProperty(type)) {
      utilityChargedToTenants[type] = round2(utilityChargedToTenants[type] + amt);
    }

    if (!isPending) {
      const roomLabel = NUMBER_TO_ROOM_LABEL[r['Room']] || (r['Room'] ? ('Pieza ' + r['Room']) : 'Sin pieza asignada');
      byRoom[roomLabel] = byRoom[roomLabel] || { total: 0, byType: {} };
      byRoom[roomLabel].total = round2(byRoom[roomLabel].total + amt);
      byRoom[roomLabel].byType[type] = round2((byRoom[roomLabel].byType[type] || 0) + amt);
    }
  });

  const totalPaidByTenants = round2(sumValues(paidByTenant));
  const totalPendingFromTenants = round2(sumValues(pendingByTenant));

  const expensesSheet = ss.getSheetByName('Expenses');
  const expensesAutoTotals = expensesSheet
    ? readSideTotals(expensesSheet, ['House', 'Internet', 'Water', 'Electricity', 'Gas', 'Total'])
    : {};

  const tenantsSheet = ss.getSheetByName('Tenants');
  const tenantsSnapshot = tenantsSheet ? readTenantsSideSnapshot(tenantsSheet) : null;

  const landlordSheet = ss.getSheetByName('To Landlord');
  const rentPaidFromNicoNote = landlordSheet ? readLandlordManualNote(landlordSheet, 'Rent Paid from Nico') : null;

  // "Rent Paid from Nico" = renta que Nico le paga al arrendador MENOS la
  // renta que los arrendatarios ya le pagaron a él — confirmado por
  // Nicolás: es literalmente la fórmula J2-J3 de su hoja "To Landlord".
  // Antes solo se leía como nota manual (se desactualizaba); ahora se
  // calcula en vivo con los mismos totales que ya arma este resumen.
  const rentPaidFromNico = round2((landlordByType['Rent'] || 0) - (paidByType['Rent'] || 0));

  const registryRows = getRegistryRows(ss);

  return {
    landlord: {
      total: round2(landlordTotal), byType: landlordByType, expected: computeLandlordExpected(landlordRows),
      rentPaidFromNico: rentPaidFromNico,
      rentPaidFromNicoNote: rentPaidFromNicoNote
    },
    expenses: { total: round2(expenseTotal), byType: expenseByType },
    tenants: {
      paidByTenant: paidByTenant,
      pendingByTenant: pendingByTenant,
      byTypePerTenant: byTypePerTenant,
      paidByTypePerTenant: paidByTypePerTenant,
      paidByType: paidByType,
      totalPaid: totalPaidByTenants,
      totalPending: totalPendingFromTenants
    },
    bond: computeBond(tenantRows, registryRows),
    utilityChargedToTenants: utilityChargedToTenants,
    byRoom: byRoom,
    registry: registryRows,
    // Tal cual están en tu Excel — no recalculados por esta app.
    autoTotals: {
      expenses: expensesAutoTotals,
      tenantsSideSnapshot: tenantsSnapshot
    },
    bigPicture: {
      cobradoATenants: totalPaidByTenants,
      pagadoALandlord: round2(landlordTotal),
      gastadoEnCasa: round2(expenseTotal),
      margenNeto: round2(totalPaidByTenants - landlordTotal - expenseTotal)
    }
  };
}

// Bond que aún se le debe devolver a cada arrendatario: lo que se anotó al
// crear su ficha ("Bond Monto") menos cualquier pago tipo "Refund" que se
// le haya registrado desde entonces en Tenants.
function computeBond(tenantRows, registryRows) {
  const refundsByTenant = {};
  tenantRows.forEach(function (r) {
    if (r['Type'] === 'Refund' && r['Tenant']) {
      refundsByTenant[r['Tenant']] = round2((refundsByTenant[r['Tenant']] || 0) + (Number(r['Amount']) || 0));
    }
  });
  const byTenant = {};
  let totalHeld = 0;
  registryRows.forEach(function (t) {
    const held = Number(t['Bond Monto']) || 0;
    const refunded = refundsByTenant[t['Nombre']] || 0;
    const remaining = round2(held - refunded);
    byTenant[t['Nombre']] = remaining;
    totalHeld = round2(totalHeld + remaining);
  });
  return { totalHeld: totalHeld, byTenant: byTenant };
}

// Cuánto DEBERÍA llevar pagado Nicolás al arrendador según la tarifa fija
// ($1.600 cada 2 semanas), calculado desde la fecha del PRIMER PAGO DE
// RENTA (Type = "Rent") hasta hoy — no desde el primer pago de cualquier
// tipo, porque el primer pago real fue el bond (que no es parte de la
// tarifa recurrente) y eso adelantaba la fecha de inicio incorrectamente.
function computeLandlordExpected(landlordRows) {
  const rentRows = landlordRows.filter(function (r) { return r['Type'] === 'Rent'; });
  const dates = rentRows.map(function (r) { return parseSheetDate(r['Date']); }).filter(function (d) { return d; });
  if (!dates.length) return null;
  const minDate = new Date(Math.min.apply(null, dates.map(function (d) { return d.getTime(); })));
  const today = new Date();
  const days = Math.max(0, Math.round((today.getTime() - minDate.getTime()) / 86400000));
  const fortnights = days / 14;
  return {
    amount: round2(fortnights * HOUSE_RENT_PER_WEEK * 2),
    sinceDate: Utilities.formatDate(minDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')
  };
}

// Busca en TODA la hoja pares "etiqueta -> valor" (la etiqueta en una
// celda, el valor en la celda inmediatamente a la derecha) para las
// etiquetas dadas — así encuentra los totales automáticos de "Expenses"
// sin depender de en qué columna exacta estén.
//
// Ojo: "House"/"Internet"/"Water"/etc. también aparecen como texto en la
// columna "Type" de cada gasto, con el Detalle (texto) al lado — para no
// confundir eso con el total real, solo cuenta como match si la celda de
// al lado es un NÚMERO (el total real lo es; el detalle de un gasto no).
function readSideTotals(sheet, labels) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 2) return {};
  const grid = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const out = {};
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length - 1; c++) {
      const label = String(grid[r][c]).trim();
      if (labels.indexOf(label) !== -1 && typeof grid[r][c + 1] === 'number') {
        out[label] = grid[r][c + 1];
      }
    }
  }
  return out;
}

// "Tenants" tiene una segunda columna "Tenant" + Rent/Bond Held/Internet/
// Water/Electricity/Gas que tú vas anotando a mano de vez en cuando como
// resumen del Room 2 — se lee el último valor no vacío de esas columnas,
// tal cual está, y se etiqueta claramente como "lo que ya llevas anotado".
function readTenantsSideSnapshot(sheet) {
  const headerRow = findHeaderRow(sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  let firstTenantIdx = -1, secondTenantIdx = -1;
  headers.forEach(function (h, i) {
    if (h === 'Tenant') {
      if (firstTenantIdx === -1) firstTenantIdx = i; else if (secondTenantIdx === -1) secondTenantIdx = i;
    }
  });
  if (secondTenantIdx === -1) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return null;
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][secondTenantIdx]) {
      const snap = { tenant: data[i][secondTenantIdx] };
      ['Rent', 'Bond Held', 'Internet', 'Water', 'Electricity', 'Gas'].forEach(function (label) {
        const idx = headers.indexOf(label);
        if (idx !== -1) snap[label] = data[i][idx];
      });
      return snap;
    }
  }
  return null;
}

// "To Landlord" tiene un par de columnas extra donde Nicolás anota a mano,
// de vez en cuando, algunas notas sueltas ("Rent Paid from Nico", "Room1",
// "Room2", etc. — cada una en SU PROPIA fila, no todas juntas). No sabemos
// en qué columna exacta cae el valor respecto a la etiqueta, así que se
// escanea la fila completa y se toma el número que esté inmediatamente al
// lado (izquierda o derecha) de la etiqueta buscada. Se lee TAL CUAL,
// nunca se recalcula — es una nota manual de Nicolás, no algo que la app
// sepa reproducir.
function readLandlordManualNote(sheet, label) {
  const headerRow = findHeaderRow(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow <= headerRow) return null;
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const dateIdx = headers.indexOf('Date');
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c]).trim() !== label) continue;
      const left = c > 0 ? row[c - 1] : null;
      const right = c < row.length - 1 ? row[c + 1] : null;
      const val = typeof left === 'number' ? left : (typeof right === 'number' ? right : null);
      if (val == null) continue;
      return { value: round2(val), date: dateIdx !== -1 ? formatDateValue(row[dateIdx]) : null };
    }
  }
  return null;
}

function formatDateValue(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return v || null;
}

function sumField(rows, field) {
  return rows.reduce(function (s, r) { return s + (Number(r[field]) || 0); }, 0);
}

function sumByKey(rows, keyField, valField) {
  const out = {};
  rows.forEach(function (r) {
    const k = r[keyField] || 'Other';
    out[k] = round2((out[k] || 0) + (Number(r[valField]) || 0));
  });
  return out;
}

function sumValues(obj) {
  return Object.keys(obj).reduce(function (s, k) { return s + obj[k]; }, 0);
}

function allRows(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const headerRow = findHeaderRow(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  if (lastRow <= headerRow) return [];
  const values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  return values.map(function (row, i) {
    const obj = rowToObject(headers, row);
    obj.__row = headerRow + 1 + i; // fila real en la hoja — usada para editar/borrar
    return obj;
  });
}

// ---------- helpers ----------

/**
 * Las tres pestañas originales tienen filas en blanco / de definición ANTES
 * de la fila real de encabezados — no está en la fila 1 en ninguna de las
 * tres. Se busca la fila cuya primera celda diga exactamente "Date".
 */
function findHeaderRow(sheet) {
  const scanRows = Math.min(sheet.getLastRow(), 20);
  if (scanRows < 1) throw new Error('Pestaña vacía: ' + sheet.getName());
  const colA = sheet.getRange(1, 1, scanRows, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() === 'Date') return i + 1; // 1-indexado
  }
  throw new Error('No se encontró la fila de encabezados ("Date" en columna A) en ' + sheet.getName());
}

// Formato de fecha que usa cada pestaña en sus filas existentes — se usa
// solo para ESCRIBIR una fecha consistente con lo que ya hay en la columna.
const DATE_FORMAT = {
  'Tenants': 'dd/MM/yy',
  'Expenses': 'dd/MM/yy',
  'To Landlord': 'dd/MM/yyyy'
};

function appendRow(ss, sheetName, valuesByHeader) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('No existe la pestaña: ' + sheetName);
  const headerRow = findHeaderRow(sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  const values = Object.assign({}, valuesByHeader);
  if (values['Date']) {
    const fmt = DATE_FORMAT[sheetName] || 'dd/MM/yyyy';
    values['Date'] = formatIsoDate(values['Date'], fmt);
  }

  const dateColIdx = headers.indexOf('Date');
  const amountColIdx = headers.indexOf('Amount');

  // Si ya existe una fila para esta fecha con el monto vacío (el caso de
  // "To Landlord", que trae fechas futuras precargadas), se completa esa
  // fila en vez de agregar una nueva y duplicada.
  if (dateColIdx !== -1 && amountColIdx !== -1) {
    const lastRow = sheet.getLastRow();
    if (lastRow > headerRow) {
      const body = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
      const wanted = normalizeDateForCompare(values['Date']);
      for (let i = 0; i < body.length; i++) {
        const cellDate = normalizeDateForCompare(body[i][dateColIdx]);
        const amountEmpty = body[i][amountColIdx] === '' || body[i][amountColIdx] === null;
        if (wanted && cellDate === wanted && amountEmpty) {
          const targetRow = headerRow + 1 + i;
          const seenFill = {};
          headers.forEach(function (h, c) {
            if (h && Object.prototype.hasOwnProperty.call(values, h) && !seenFill[h]) {
              seenFill[h] = true;
              sheet.getRange(targetRow, c + 1).setValue(values[h]);
            }
          });
          return;
        }
      }
    }
  }

  // "Tenants" repite el encabezado "Tenant" dos veces (la segunda es parte
  // del resumen lateral de Room 2, no del registro de este pago) — solo se
  // escribe en la PRIMERA columna que tenga cada nombre de encabezado.
  const seenWrite = {};
  const row = headers.map(function (h) {
    if (h && Object.prototype.hasOwnProperty.call(values, h) && !seenWrite[h]) {
      seenWrite[h] = true;
      return values[h];
    }
    return '';
  });
  sheet.appendRow(row);
}

// Verifica que la fila que llegó desde la app (identificada por número de
// fila, no por ID — estas pestañas no tienen columna ID) siga siendo la
// misma que el usuario vio en pantalla, comparando Amount/Date actuales
// contra lo que la app tenía cargado. Evita editar/borrar la fila
// equivocada si algo cambió el orden de las filas entre que se cargó la
// pantalla y que se guardó la edición.
function validateRecordRow(sheet, body) {
  const headerRow = findHeaderRow(sheet);
  const rowNum = Number(body.row);
  if (!rowNum || rowNum <= headerRow || rowNum > sheet.getLastRow()) {
    return { ok: false, error: 'Esa fila ya no existe — actualiza la pantalla e intenta de nuevo.' };
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  if (body.expectedAmount != null || body.expectedDate) {
    const current = rowToObject(headers, sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0]);
    const amountOk = body.expectedAmount == null || Number(current['Amount']) === Number(body.expectedAmount);
    const dateOk = !body.expectedDate || current['Date'] === body.expectedDate;
    if (!amountOk || !dateOk) {
      return { ok: false, error: 'Esa fila cambió desde que se cargó la pantalla — actualiza e intenta de nuevo.' };
    }
  }
  return { ok: true, rowNum: rowNum, headers: headers, lastCol: lastCol };
}

function editRecordRow(sheet, body) {
  const v = validateRecordRow(sheet, body);
  if (!v.ok) return v;

  const values = {};
  if (body.date) values['Date'] = formatIsoDate(body.date, DATE_FORMAT[sheet.getName()] || 'dd/MM/yyyy');
  if (body.amount != null && body.amount !== '') values['Amount'] = Number(body.amount);
  if (body.paymentMethod != null) values['Payment Method'] = body.paymentMethod;
  if (body.type != null) values['Type'] = body.type;
  if (body.detail != null) values['Detail'] = body.detail;
  if (body.tenant != null) values['Tenant'] = body.tenant;
  if (body.room != null) values['Room'] = body.room;

  // Igual que appendRow: "Tenants" repite "Tenant" dos veces — solo se
  // escribe en la PRIMERA columna con ese nombre, para no pisar el resumen
  // lateral de Room 2 que vive en la segunda.
  const seenWrite = {};
  v.headers.forEach(function (h, c) {
    if (h && Object.prototype.hasOwnProperty.call(values, h) && !seenWrite[h]) {
      seenWrite[h] = true;
      sheet.getRange(v.rowNum, c + 1).setValue(values[h]);
    }
  });
  return { ok: true };
}

// "Borrar" no elimina la fila (sheet.deleteRow correría el riesgo de
// arrastrar filas de abajo hacia arriba y, en "Tenants", de destruir datos
// del resumen lateral de Room 2 que vive en las mismas filas que el
// registro) — en vez de eso, VACÍA solo las columnas del registro. Las
// funciones de resumen ya ignoran filas con Amount vacío.
function clearRecordRow(sheet, body) {
  const v = validateRecordRow(sheet, body);
  if (!v.ok) return v;

  const clearable = ['Date', 'Amount', 'Payment Method', 'Type', 'Detail', 'Tenant', 'Room'];
  const seenWrite = {};
  v.headers.forEach(function (h, c) {
    if (h && clearable.indexOf(h) !== -1 && !seenWrite[h]) {
      seenWrite[h] = true;
      sheet.getRange(v.rowNum, c + 1).setValue('');
    }
  });
  return { ok: true };
}

// "2026-08-04" (input type=date) -> "04/08/2026" o "04/08/26" según fmt.
function formatIsoDate(isoStr, fmt) {
  const m = String(isoStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return isoStr;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), fmt);
}

// Convierte un valor de fecha (Date real de Sheets, o texto dd/mm/yy[yy])
// a una forma canónica "dd/mm/yyyy" para poder comparar sin importar si la
// pestaña usa año de 2 o 4 dígitos, o si la celda quedó como texto o Date.
function normalizeDateForCompare(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!m) return null;
  const dd = m[1].padStart(2, '0');
  const mm = m[2].padStart(2, '0');
  const yyyy = m[3].length === 2 ? '20' + m[3] : m[3];
  return dd + '/' + mm + '/' + yyyy;
}

// Convierte un valor de fecha de la planilla (Date real, o texto
// dd/mm/yy[yy]) a un objeto Date de JS — para poder comparar/restar fechas.
// A diferencia de normalizeDateForCompare (que devuelve texto canónico para
// IGUALAR fechas), esta devuelve un Date real para poder calcular DÍAS
// TRANSCURRIDOS (usado en computeLandlordExpected).
function parseSheetDate(value) {
  if (value instanceof Date) return value;
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!m) return null;
  const dd = Number(m[1]);
  const mm = Number(m[2]) - 1;
  const yyyy = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  return new Date(yyyy, mm, dd);
}

function lastRows(ss, sheetName, n) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const headerRow = findHeaderRow(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  if (lastRow <= headerRow) return [];
  const startRow = Math.max(headerRow + 1, lastRow - n + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
  return values.map(function (row, i) {
    const obj = rowToObject(headers, row);
    obj.__row = startRow + i; // fila real en la hoja — usada para editar/borrar
    return obj;
  }).reverse();
}

// Se queda con la PRIMERA columna que tenga cada nombre de encabezado
// (necesario porque "Tenants" repite "Tenant" dos veces).
function rowToObject(headers, row) {
  const obj = {};
  headers.forEach(function (h, i) {
    if (h && !(h in obj)) obj[h] = row[i];
  });
  // Cualquier celda que haya quedado como Date real de Sheets (en vez del
  // texto que escribe esta app) — no solo "Date", también "Bond Fecha" /
  // "Fecha Inicio" del registro de arrendatarios — se normaliza a
  // dd/MM/yyyy. Si no, JSON.stringify la convierte en un timestamp completo
  // (con hora y zona) y la app debe mostrar solo día/mes/año.
  Object.keys(obj).forEach(function (k) {
    if (obj[k] instanceof Date) {
      obj[k] = Utilities.formatDate(obj[k], Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
  });
  return obj;
}

function requireFields(body, fields) {
  fields.forEach(function (f) {
    if (body[f] === undefined || body[f] === null || body[f] === '') {
      throw new Error('Falta el campo: ' + f);
    }
  });
}

function round2(n) { return Math.round(n * 100) / 100; }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
