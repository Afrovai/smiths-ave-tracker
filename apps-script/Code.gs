/**
 * 10 Smiths Avenue, Redcliffe — backend para la app de registro rápido.
 *
 * Qué hace:
 *  - Recibe POST desde la app web (GitHub Pages) y agrega filas nuevas a las
 *    pestañas "Expenses", "Tenants" o "To Landlord" de la planilla.
 *  - Si un gasto es de un tipo "compartible" (luz, agua, gas, internet),
 *    calcula automáticamente el monto que le corresponde pagar al inquilino
 *    del Room 2 y crea un cobro pendiente en "Tenants".
 *  - NO toca ni recalcula las columnas de totales acumulados que ya existen
 *    más a la derecha en "Tenants" — solo agrega filas nuevas usando los
 *    encabezados de la fila 1 de cada pestaña, así nunca pisa datos.
 *
 * Instrucciones de despliegue: ver apps-script/README.md en este mismo repo.
 */

// ID de la planilla "10 Smiths Avenue, Redcliffe"
const SHEET_ID = '1Z8asY7OdehqG5yCrcw2uB9xP9ZE7Zb16aXpasg4W1tA';

// Debe coincidir EXACTO con el secreto que configures en la app web.
// Cámbialo si quieres (y actualízalo también en la app) — es la única
// protección del endpoint, ya que el Web App queda con acceso "Anyone".
const SECRET = '5EGVxQhUJ2RsQ10dFUEkHAuM';

// Tipos de gasto que se dividen automáticamente con el inquilino del Room 2,
// y qué fracción le corresponde a él. Ajusta libremente.
const SPLIT_RULES = {
  'Electricity': 0.5,
  'Water': 0.5,
  'Gas': 0.5,
  'Internet': 0.5
};

// Inquilino por defecto al que se le carga la mitad de las cuentas — cambia
// esto cuando cambie el inquilino del Room 2 (o mándalo desde la app).
const DEFAULT_SPLIT_TENANT = 'Juan Plata';

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
        const tenant = body.splitTenant || DEFAULT_SPLIT_TENANT;
        const pct = body.splitPct != null ? Number(body.splitPct) : frac;
        const share = round2(Number(body.amount) * pct);
        appendRow(ss, 'Tenants', {
          'Date': body.date,
          'Amount': share,
          'Payment Method': 'Pending',
          'Type': body.type,
          'Detail': 'Pendiente: ' + Math.round(pct * 100) + '% de cuenta de ' + body.type +
                    ' ($' + body.amount + ' del ' + body.date + ')',
          'Tenant': tenant
        });
        created.push({ sheet: 'Tenants', amount: share, tenant: tenant, pending: true });
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

    } else {
      return jsonOut({ ok: false, error: 'Acción desconocida: ' + body.action });
    }

    return jsonOut({ ok: true, created: created });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/**
 * GET ?secret=...&recent=5  -> últimas filas de cada pestaña, para que la
 * app muestre "actividad reciente" y sirva de confirmación visual.
 */
function doGet(e) {
  try {
    const secret = e.parameter.secret;
    if (secret !== SECRET) return jsonOut({ ok: false, error: 'No autorizado' });

    const n = Number(e.parameter.recent || 5);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const out = {};
    ['Expenses', 'Tenants', 'To Landlord'].forEach(function (name) {
      out[name] = lastRows(ss, name, n);
    });
    return jsonOut({ ok: true, recent: out });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ---------- helpers ----------

/**
 * Las tres pestañas tienen filas en blanco / de definición (ej. "Room 1,
 * Small, 280...") ANTES de la fila real de encabezados — no está en la
 * fila 1 en ninguna de las tres. En vez de asumir un número de fila fijo
 * (frágil, y distinto según la pestaña), se busca la fila cuya primera
 * celda diga exactamente "Date": esa es la fila de encabezados real.
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

function appendRow(ss, sheetName, valuesByHeader) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('No existe la pestaña: ' + sheetName);
  const headerRow = findHeaderRow(sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const row = headers.map(function (h) {
    return Object.prototype.hasOwnProperty.call(valuesByHeader, h) ? valuesByHeader[h] : '';
  });
  sheet.appendRow(row);
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
  return values.reverse().map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { if (h) obj[h] = row[i]; });
    return obj;
  });
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
