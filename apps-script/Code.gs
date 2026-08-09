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

    } else {
      return jsonOut({ ok: false, error: 'Acción desconocida: ' + body.action });
    }

    return jsonOut({ ok: true, created: created });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/**
 * GET ?secret=...&recent=5   -> últimas filas de cada pestaña.
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

    const n = Number(e.parameter.recent || 5);
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
  const utilityChargedToTenants = { Internet: 0, Water: 0, Electricity: 0, Gas: 0 };
  const byRoom = {};

  tenantRows.forEach(function (r) {
    const tenant = r['Tenant'];
    const amt = Number(r['Amount']) || 0;
    const type = r['Type'] || 'Other';
    const isPending = r['Payment Method'] === 'Pending';

    if (tenant) {
      if (isPending) {
        pendingByTenant[tenant] = round2((pendingByTenant[tenant] || 0) + amt);
      } else {
        paidByTenant[tenant] = round2((paidByTenant[tenant] || 0) + amt);
      }
      byTypePerTenant[tenant] = byTypePerTenant[tenant] || {};
      byTypePerTenant[tenant][type] = round2((byTypePerTenant[tenant][type] || 0) + amt);
    }

    if (utilityChargedToTenants.hasOwnProperty(type)) {
      utilityChargedToTenants[type] = round2(utilityChargedToTenants[type] + amt);
    }

    const roomLabel = NUMBER_TO_ROOM_LABEL[r['Room']] || (r['Room'] ? ('Pieza ' + r['Room']) : 'Sin pieza asignada');
    byRoom[roomLabel] = byRoom[roomLabel] || { total: 0, byType: {} };
    byRoom[roomLabel].total = round2(byRoom[roomLabel].total + amt);
    byRoom[roomLabel].byType[type] = round2((byRoom[roomLabel].byType[type] || 0) + amt);
  });

  const totalPaidByTenants = round2(sumValues(paidByTenant));
  const totalPendingFromTenants = round2(sumValues(pendingByTenant));

  const expensesSheet = ss.getSheetByName('Expenses');
  const expensesAutoTotals = expensesSheet
    ? readSideTotals(expensesSheet, ['House', 'Internet', 'Water', 'Electricity', 'Gas', 'Total'])
    : {};

  const tenantsSheet = ss.getSheetByName('Tenants');
  const tenantsSnapshot = tenantsSheet ? readTenantsSideSnapshot(tenantsSheet) : null;

  return {
    landlord: { total: round2(landlordTotal), byType: landlordByType },
    expenses: { total: round2(expenseTotal), byType: expenseByType },
    tenants: {
      paidByTenant: paidByTenant,
      pendingByTenant: pendingByTenant,
      byTypePerTenant: byTypePerTenant,
      totalPaid: totalPaidByTenants,
      totalPending: totalPendingFromTenants
    },
    utilityChargedToTenants: utilityChargedToTenants,
    byRoom: byRoom,
    registry: getRegistryRows(ss),
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
  return values.map(function (row) { return rowToObject(headers, row); });
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
  return values.reverse().map(function (row) { return rowToObject(headers, row); });
}

// Se queda con la PRIMERA columna que tenga cada nombre de encabezado
// (necesario porque "Tenants" repite "Tenant" dos veces).
function rowToObject(headers, row) {
  const obj = {};
  headers.forEach(function (h, i) {
    if (h && !(h in obj)) obj[h] = row[i];
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
