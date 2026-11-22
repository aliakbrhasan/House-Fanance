const SECRET_PIN = '1234'; // يمكنك تغيير الرمز السري هنا
const DATA_SHEET_NAME = 'App_Data';
const CAT_SHEET_NAME = 'App_Categories';
const OBL_SHEET_NAME = 'App_Obligations';
const OBL_PAY_SHEET_NAME = 'App_Obligation_Payments';
const DEFAULT_EXPENSE_CATS = ['المنزل', 'تمتم', 'غصون', 'علي', 'السيارة والمواصلات', 'صحة وعناية', 'هدنة', 'الادخار (السلف)', 'السفر', 'الديون'];
const DEFAULT_INCOME_CATS = ['راتب علي', 'راتب غصون', 'بيج تمتم', 'مدخرات (زين كاش_نيو)', 'القبس', 'اعمال اخرى', 'استدانة'];

function doGet(e) {
  try {
    const pin = (e && e.parameter && e.parameter.pin) ? String(e.parameter.pin).trim() : '';
    if (pin !== SECRET_PIN) {
      return jsonResponse({ status: 'error', message: 'الرمز السري غير صحيح' });
    }

    const sheet = getOrInitDataSheet_();
    const values = sheet.getDataRange().getValues();
    const transactions = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      const dateStr = formatDateCell_(row[1]);
      const monthStr = row[2] ? String(row[2]).trim() : (dateStr ? dateStr.substring(0, 7) : '');
      transactions.push({
        id: String(row[0]),
        date: dateStr,
        month: monthStr,
        type: String(row[3] || 'expense'),
        amount: Number(row[4]) || 0,
        category: String(row[5] || '').trim(),
        description: String(row[6] || '').trim()
      });
    }

    transactions.sort((a, b) => (b.date > a.date ? 1 : b.date < a.date ? -1 : 0));
    const categories = getOrInitCategories_();
    const obligations = getObligations_();
    return jsonResponse({ status: 'success', transactions, categories, obligations });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (String(body.pin || '').trim() !== SECRET_PIN) {
      return jsonResponse({ status: 'error', message: 'الرمز السري غير صحيح' });
    }

    const sheet = getOrInitDataSheet_();

    if (body.action === 'add') {
      const amount = Number(body.amount);
      if (!amount || amount <= 0) {
        return jsonResponse({ status: 'error', message: 'المبلغ غير صالح' });
      }
      const dateStr = body.date ? String(body.date).trim() : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const monthStr = dateStr.substring(0, 7);
      const tx = {
        id: 'TX-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
        date: dateStr,
        month: monthStr,
        type: body.type === 'income' ? 'income' : 'expense',
        amount: amount,
        category: String(body.category || '').trim(),
        description: String(body.description || '').trim()
      };

      sheet.appendRow([
        tx.id,
        tx.date,
        tx.month,
        tx.type,
        tx.amount,
        tx.category,
        tx.description,
        new Date().toISOString()
      ]);

      return jsonResponse({ status: 'success', transaction: tx });
    }

    if (body.action === 'delete') {
      const targetId = String(body.id || '').trim();
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === targetId) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ status: 'success' });
        }
      }
      return jsonResponse({ status: 'error', message: 'العملية غير موجودة' });
    }

    if (body.action === 'save_categories' && body.categories) {
      saveCategoriesToSheet_(body.categories);
      return jsonResponse({ status: 'success' });
    }

    if (body.action === 'add_obligation' && body.obligation) {
      addObligationToSheet_(body.obligation);
      return jsonResponse({ status: 'success' });
    }

    if (body.action === 'add_obligation_payment' && body.payment) {
      const pay = body.payment;
      addObligationPaymentToSheet_(pay);
      let loggedTx = null;
      if (body.alsoLogExpense) {
        const dateStr = pay.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        loggedTx = {
          id: 'TX-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
          date: dateStr,
          month: dateStr.substring(0, 7),
          type: 'expense',
          amount: Number(pay.amount) || 0,
          category: String(body.expenseCategory || 'الديون').trim(),
          description: String(body.expenseDescription || '').trim()
        };
        sheet.appendRow([loggedTx.id, loggedTx.date, loggedTx.month, loggedTx.type, loggedTx.amount, loggedTx.category, loggedTx.description, new Date().toISOString()]);
      }
      return jsonResponse({ status: 'success', payment: pay, transaction: loggedTx });
    }

    if (body.action === 'delete_obligation' && body.id) {
      deleteObligationFromSheet_(String(body.id));
      return jsonResponse({ status: 'success' });
    }

    if (body.action === 'delete_obligation_payment' && body.paymentId) {
      deleteObligationPaymentFromSheet_(String(body.paymentId));
      return jsonResponse({ status: 'success' });
    }

    return jsonResponse({ status: 'error', message: 'إجراء غير مدعوم' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function getOrInitDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME, 0);
    sheet.appendRow(['ID', 'Date', 'Month', 'Type', 'Amount', 'Category', 'Description', 'CreatedAt']);
    sheet.setFrozenRows(1);
    migrateHistoricalSheets_(ss, sheet);
  }
  return sheet;
}

function getOrInitCategories_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CAT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CAT_SHEET_NAME);
    sheet.appendRow(['Type', 'Name']);
    sheet.setFrozenRows(1);
    const rows = [];
    DEFAULT_EXPENSE_CATS.forEach(c => rows.push(['expense', c]));
    DEFAULT_INCOME_CATS.forEach(c => rows.push(['income', c]));
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    return { expense: DEFAULT_EXPENSE_CATS, income: DEFAULT_INCOME_CATS };
  }
  const values = sheet.getDataRange().getValues();
  const expense = [];
  const income = [];
  for (let i = 1; i < values.length; i++) {
    const t = String(values[i][0] || '').trim();
    const n = String(values[i][1] || '').trim();
    if (!n) continue;
    if (t === 'income') income.push(n);
    else if (t === 'expense') expense.push(n);
  }
  return {
    expense: expense.length ? expense : DEFAULT_EXPENSE_CATS,
    income: income.length ? income : DEFAULT_INCOME_CATS
  };
}

function saveCategoriesToSheet_(cats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CAT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CAT_SHEET_NAME);
    sheet.setFrozenRows(1);
  }
  sheet.clearContents();
  const rows = [['Type', 'Name']];
  (cats.expense || []).forEach(c => { const v = String(c || '').trim(); if (v) rows.push(['expense', v]); });
  (cats.income || []).forEach(c => { const v = String(c || '').trim(); if (v) rows.push(['income', v]); });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function getOrCreateOblSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let oblSheet = ss.getSheetByName(OBL_SHEET_NAME);
  if (!oblSheet) {
    oblSheet = ss.insertSheet(OBL_SHEET_NAME);
    oblSheet.appendRow(['ID', 'Kind', 'Title', 'Party', 'TotalAmount', 'PaidBefore', 'DefaultPayment', 'Notes', 'CreatedAt']);
    oblSheet.setFrozenRows(1);
  }
  let paySheet = ss.getSheetByName(OBL_PAY_SHEET_NAME);
  if (!paySheet) {
    paySheet = ss.insertSheet(OBL_PAY_SHEET_NAME);
    paySheet.appendRow(['PaymentID', 'ObligationID', 'Amount', 'Date', 'PaidAt', 'Note']);
    paySheet.setFrozenRows(1);
  }
  return { oblSheet, paySheet };
}

function getObligations_() {
  const { oblSheet, paySheet } = getOrCreateOblSheets_();
  const oblRows = oblSheet.getDataRange().getValues();
  const payRows = paySheet.getDataRange().getValues();
  const payMap = {};
  for (let i = 1; i < payRows.length; i++) {
    const r = payRows[i];
    if (!r[0] || !r[1]) continue;
    const oblId = String(r[1]);
    if (!payMap[oblId]) payMap[oblId] = [];
    payMap[oblId].push({
      id: String(r[0]),
      obligationId: oblId,
      amount: Number(r[2]) || 0,
      date: formatDateCell_(r[3]),
      paidAt: String(r[4] || ''),
      note: String(r[5] || '')
    });
  }
  const list = [];
  for (let i = 1; i < oblRows.length; i++) {
    const r = oblRows[i];
    if (!r[0]) continue;
    const id = String(r[0]);
    list.push({
      id,
      kind: String(r[1] || 'installment'),
      title: String(r[2] || ''),
      party: String(r[3] || ''),
      totalAmount: Number(r[4]) || 0,
      paidBefore: Number(r[5]) || 0,
      defaultPayment: Number(r[6]) || 0,
      notes: String(r[7] || ''),
      createdAt: String(r[8] || ''),
      payments: payMap[id] || []
    });
  }
  return list.reverse();
}

function addObligationToSheet_(o) {
  const { oblSheet } = getOrCreateOblSheets_();
  oblSheet.appendRow([
    String(o.id || ''),
    String(o.kind || 'installment'),
    String(o.title || ''),
    String(o.party || ''),
    Number(o.totalAmount) || 0,
    Number(o.paidBefore) || 0,
    Number(o.defaultPayment) || 0,
    String(o.notes || ''),
    String(o.createdAt || new Date().toISOString())
  ]);
}

function addObligationPaymentToSheet_(p) {
  const { paySheet } = getOrCreateOblSheets_();
  paySheet.appendRow([
    String(p.id || ''),
    String(p.obligationId || ''),
    Number(p.amount) || 0,
    String(p.date || ''),
    String(p.paidAt || ''),
    String(p.note || '')
  ]);
}

function deleteObligationFromSheet_(id) {
  const { oblSheet, paySheet } = getOrCreateOblSheets_();
  const oVals = oblSheet.getDataRange().getValues();
  for (let i = oVals.length - 1; i >= 1; i--) {
    if (String(oVals[i][0]) === id) oblSheet.deleteRow(i + 1);
  }
  const pVals = paySheet.getDataRange().getValues();
  for (let i = pVals.length - 1; i >= 1; i--) {
    if (String(pVals[i][1]) === id) paySheet.deleteRow(i + 1);
  }
}

function deleteObligationPaymentFromSheet_(payId) {
  const { paySheet } = getOrCreateOblSheets_();
  const pVals = paySheet.getDataRange().getValues();
  for (let i = pVals.length - 1; i >= 1; i--) {
    if (String(pVals[i][0]) === payId) paySheet.deleteRow(i + 1);
  }
}

function migrateHistoricalSheets_(ss, dataSheet) {
  const sheets = ss.getSheets();
  const rowsToInsert = [];
  const nowIso = new Date().toISOString();
  let counter = 1;

  sheets.forEach(sh => {
    const name = sh.getName().trim();
    const match = name.match(/^(\d{1,2})\/(\d{2})\s+transactions/i);
    if (!match) return;

    const sheetMonthNum = Number(match[1]);
    const sheetYear = 2000 + Number(match[2]);
    const fallbackMonth = sheetYear + '-' + String(sheetMonthNum).padStart(2, '0');
    const fallbackDate = fallbackMonth + '-01';

    const data = sh.getDataRange().getValues();
    let headerRow = -1;
    let expCols = null;
    let incCols = null;

    for (let r = 0; r < Math.min(data.length, 10); r++) {
      const rowLower = data[r].map(c => String(c || '').trim().toLowerCase());
      const dateIndexes = [];
      rowLower.forEach((val, idx) => {
        if (val === 'date' && rowLower[idx + 1] === 'amount') {
          dateIndexes.push(idx);
        }
      });
      if (dateIndexes.length >= 1) {
        headerRow = r;
        expCols = dateIndexes[0];
        if (dateIndexes.length >= 2) incCols = dateIndexes[1];
        break;
      }
    }

    if (headerRow === -1) return;

    for (let r = headerRow + 1; r < data.length; r++) {
      const row = data[r];

      if (expCols !== null) {
        const parsed = parseRowBlock_(row, expCols, 'expense', sheetMonthNum, fallbackDate, fallbackMonth);
        if (parsed) {
          rowsToInsert.push(['MIG-' + (counter++), parsed.date, parsed.month, 'expense', parsed.amount, parsed.category, parsed.description, nowIso]);
        }
      }

      if (incCols !== null) {
        const parsed = parseRowBlock_(row, incCols, 'income', sheetMonthNum, fallbackDate, fallbackMonth);
        if (parsed) {
          rowsToInsert.push(['MIG-' + (counter++), parsed.date, parsed.month, 'income', parsed.amount, parsed.category, parsed.description, nowIso]);
        }
      }
    }
  });

  if (rowsToInsert.length > 0) {
    dataSheet.getRange(2, 1, rowsToInsert.length, 8).setValues(rowsToInsert);
  }
}

function parseRowBlock_(row, startCol, type, sheetMonthNum, fallbackDate, fallbackMonth) {
  const rawDate = row[startCol];
  const rawAmount = row[startCol + 1];
  const desc = String(row[startCol + 2] || '').trim();
  const cat = String(row[startCol + 3] || '').trim();

  const amount = typeof rawAmount === 'number'
    ? rawAmount
    : Number(String(rawAmount || '').replace(/[^0-9.-]/g, ''));

  if (!amount || amount <= 0) return null;
  if (!desc && !cat) return null;

  const dateStr = parseFlexibleDate_(rawDate, sheetMonthNum) || fallbackDate;
  const monthStr = dateStr ? dateStr.substring(0, 7) : fallbackMonth;
  const cleanCat = cat || (type === 'expense' ? 'المنزل' : 'اعمال اخرى');

  return { date: dateStr, month: monthStr, amount, category: cleanCat, description: desc };
}

function parseFlexibleDate_(val, sheetMonthNum) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return '';
  const a = Number(m[1]);
  const b = Number(m[2]);
  const y = Number(m[3]);
  let day = a, month = b;
  if (a <= 12 && b > 12) {
    month = a; day = b;
  } else if (a === sheetMonthNum && b !== sheetMonthNum) {
    month = a; day = b;
  }
  return y + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
}

function formatDateCell_(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
