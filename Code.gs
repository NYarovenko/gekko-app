/**
 * ГЕККО І КО — Apps Script backend
 * Обробляє запити від веб-додатку (десктоп + мобільний)
 */

const SHEET_USERS = 'Users';
const SHEET_TICKETS = 'Tickets';
const SHEET_SETTINGS = 'Settings';
const SHEET_MATERIALS = 'Materials';
const SHEET_SERVICE_HISTORY = 'ServiceHistory';

// ===================== ІНІЦІАЛІЗАЦІЯ (запустити один раз вручну) =====================
// ===================== МІГРАЦІЯ: додати стовпці порядку до існуючих даних (запустити один раз) =====================
function migrateMaterialsOrder() {
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 1) return;

  const headers = data[0];
  if (headers.indexOf('cat_order') !== -1 && headers.indexOf('item_order') !== -1) {
    Logger.log('Стовпці порядку вже існують');
    return;
  }

  const newHeaders = ['login', 'category', 'item_name', 'cat_order', 'item_order'];
  sheet.getRange(1, 1, 1, 5).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, 5).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');

  const catOrderMap = {};
  const itemOrderMap = {};

  for (let i = 1; i < data.length; i++) {
    const login = data[i][0];
    const category = data[i][1];
    if (!login || !category) continue;

    if (!catOrderMap[login]) catOrderMap[login] = {};
    if (!(category in catOrderMap[login])) {
      catOrderMap[login][category] = Object.keys(catOrderMap[login]).length;
    }
    const catOrder = catOrderMap[login][category];

    if (!itemOrderMap[login]) itemOrderMap[login] = {};
    if (!itemOrderMap[login][category]) itemOrderMap[login][category] = 0;
    const itemOrder = itemOrderMap[login][category]++;

    sheet.getRange(i + 1, 4).setValue(catOrder);
    sheet.getRange(i + 1, 5).setValue(itemOrder);
  }

  Logger.log('Міграція завершена');
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Видаляємо стандартний "Sheet1" якщо є
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Аркуш1');
  
  // --- Users ---
  let wsUsers = ss.getSheetByName(SHEET_USERS);
  if (!wsUsers) wsUsers = ss.insertSheet(SHEET_USERS);
  wsUsers.clear();
  wsUsers.getRange(1, 1, 1, 4).setValues([['login', 'password', 'display_name', 'role']]);
  wsUsers.getRange(2, 1, 5, 4).setValues([
    ['garkusha', '0md2hc', 'Гаркуша', 'executor'],
    ['yarovenko', 'kt26yt', 'Яровенко', 'executor'],
    ['yasnogor', '4lg6ob', 'Ясногор', 'executor'],
    ['brigada_uh', 'szuqiu', 'бригада УХ', 'executor'],
    ['dimonchik', '8mcc90', 'адмін Дімончік', 'admin'],
  ]);
  wsUsers.getRange(1, 1, 1, 4).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  wsUsers.setColumnWidths(1, 4, 140);

  // --- Tickets ---
  let wsTickets = ss.getSheetByName(SHEET_TICKETS);
  if (!wsTickets) wsTickets = ss.insertSheet(SHEET_TICKETS);
  wsTickets.clear();
  const ticketHeaders = ['id','num','prio','zayav','date','name','phone','phones','addr','addrRaw',
    'problem','resp','hours','service','comment','isFreenet','account','plannedDate','assigned_to','status','hidden_by'];
  wsTickets.getRange(1, 1, 1, ticketHeaders.length).setValues([ticketHeaders]);
  wsTickets.getRange(1, 1, 1, ticketHeaders.length).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  wsTickets.setColumnWidths(1, ticketHeaders.length, 130);

  // --- Settings ---
  let wsSettings = ss.getSheetByName(SHEET_SETTINGS);
  if (!wsSettings) wsSettings = ss.insertSheet(SHEET_SETTINGS);
  wsSettings.clear();
  wsSettings.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  wsSettings.getRange(2, 1, 2, 2).setValues([
    ['last_import', ''],
    ['imported_by', ''],
  ]);
  wsSettings.getRange(1, 1, 1, 2).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');

  // Видаляємо порожній стандартний лист, якщо є інші листи
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // --- Materials ---
  let wsMaterials = ss.getSheetByName(SHEET_MATERIALS);
  if (!wsMaterials) wsMaterials = ss.insertSheet(SHEET_MATERIALS);
  wsMaterials.clear();
  wsMaterials.getRange(1, 1, 1, 5).setValues([['login', 'category', 'item_name', 'cat_order', 'item_order']]);
  wsMaterials.getRange(1, 1, 1, 5).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  wsMaterials.setColumnWidths(1, 5, 160);

  // --- ServiceHistory (історія обслуговування по особовому рахунку) ---
  let wsHistory = ss.getSheetByName(SHEET_SERVICE_HISTORY);
  if (!wsHistory) wsHistory = ss.insertSheet(SHEET_SERVICE_HISTORY);
  wsHistory.clear();
  const historyHeaders = ['account', 'ticket_id', 'ticket_date', 'assigned_to', 'problem', 'comment', 'comment_date'];
  wsHistory.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
  wsHistory.getRange(1, 1, 1, historyHeaders.length).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  wsHistory.setColumnWidths(1, historyHeaders.length, 160);
  // Текстовий формат для дат, щоб уникнути авто-конвертації в Date object
  wsHistory.getRange(2, 3, 1000, 1).setNumberFormat('@STRING@');
  wsHistory.getRange(2, 7, 1000, 1).setNumberFormat('@STRING@');

  Logger.log('Setup complete!');
}

// ===================== КОМЕНТАРІ ВИКОНАВЦЯ (ServiceHistory) =====================

// Додати коментар до заявки. Записує в ServiceHistory (історія по особовому рахунку)
// і оновлює агрегований стовпець comment_from_executor у Tickets (для відображення в адміна)
function handleAddComment(ticketId, login, commentText) {
  if (!ticketId || !login || !commentText || !commentText.trim()) {
    return { ok: false, error: 'Не всі поля заповнені' };
  }

  const ticketsSheet = getSheet(SHEET_TICKETS);
  const ticketRowIdx = findRowIndexById(ticketsSheet, 'id', ticketId);
  if (ticketRowIdx === -1) return { ok: false, error: 'Заявку не знайдено' };

  const headers = ticketsSheet.getRange(1, 1, 1, ticketsSheet.getLastColumn()).getValues()[0];
  const rowValues = ticketsSheet.getRange(ticketRowIdx, 1, 1, headers.length).getValues()[0];
  const ticketObj = {};
  headers.forEach((h, i) => { ticketObj[h] = rowValues[i]; });

  const now = new Date();
  const commentDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  const trimmedComment = commentText.trim();

  // 1. Записуємо в ServiceHistory
  const historySheet = getSheet(SHEET_SERVICE_HISTORY);
  historySheet.appendRow([
    ticketObj.account || '',
    ticketId,
    ticketObj.date || '',
    ticketObj.assigned_to || login,
    ticketObj.problem || '',
    trimmedComment,
    commentDate
  ]);
  const lastRow = historySheet.getLastRow();
  historySheet.getRange(lastRow, 3).setNumberFormat('@STRING@');
  historySheet.getRange(lastRow, 7).setNumberFormat('@STRING@');

  // 2. Оновлюємо агрегований стовпець в Tickets (накопичувальний)
  let commentCol = headers.indexOf('comment_from_executor') + 1;
  if (commentCol === 0) {
    commentCol = headers.length + 1;
    ticketsSheet.getRange(1, commentCol).setValue('comment_from_executor');
    ticketsSheet.getRange(1, commentCol).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  const existingComments = ticketsSheet.getRange(ticketRowIdx, commentCol).getValue();
  const newEntry = '[' + commentDate + ', ' + login + ']: ' + trimmedComment;
  const updatedComments = existingComments ? (existingComments + '\n' + newEntry) : newEntry;
  ticketsSheet.getRange(ticketRowIdx, commentCol).setValue(updatedComments);

  return { ok: true, commentDate: commentDate };
}

// Отримати всі коментарі (історію) по заявці
function handleGetComments(ticketId) {
  if (!ticketId) return { ok: false, error: 'Не вказано ticketId' };
  const sheet = getSheet(SHEET_SERVICE_HISTORY);
  const rows = sheetToObjects(sheet).filter(r => String(r.ticket_id) === String(ticketId));
  return { ok: true, comments: rows };
}

// ===================== ОКРЕМА МІГРАЦІЯ: додати лист ServiceHistory без зачіпання інших даних =====================
function setupServiceHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsHistory = ss.getSheetByName(SHEET_SERVICE_HISTORY);
  if (!wsHistory) {
    wsHistory = ss.insertSheet(SHEET_SERVICE_HISTORY);
    const historyHeaders = ['account', 'ticket_id', 'ticket_date', 'assigned_to', 'problem', 'comment', 'comment_date'];
    wsHistory.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
    wsHistory.getRange(1, 1, 1, historyHeaders.length).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
    wsHistory.setColumnWidths(1, historyHeaders.length, 160);
    wsHistory.getRange(2, 3, 1000, 1).setNumberFormat('@STRING@');
    wsHistory.getRange(2, 7, 1000, 1).setNumberFormat('@STRING@');
    Logger.log('Лист ServiceHistory створено');
  } else {
    Logger.log('Лист ServiceHistory вже існує');
  }
}

// ===================== HELPERS =====================
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function findRowIndexById(sheet, idColName, idValue) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(idColName);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(idValue)) return i + 1; // 1-indexed row
  }
  return -1;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== MAIN ENTRY POINTS =====================
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const action = e.parameter.action;
    let payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (err) {}
    }

    switch (action) {
      case 'login':
        return jsonResponse(handleLogin(e.parameter.login, e.parameter.password));
      case 'getTickets':
        return jsonResponse(handleGetTickets(e.parameter.login));
      case 'importTickets':
        return jsonResponse(handleImportTickets(payload.tickets, e.parameter.login));
      case 'assignTicket':
        return jsonResponse(handleAssignTicket(e.parameter.ticketId, e.parameter.assignTo));
      case 'hideTicket':
        return jsonResponse(handleHideTicket(e.parameter.ticketId, e.parameter.login));
      case 'getUsers':
        return jsonResponse(handleGetUsers());
      case 'getMaterials':
        return jsonResponse(handleGetMaterials(e.parameter.login));
      case 'addMaterial':
        return jsonResponse(handleAddMaterial(e.parameter.login, e.parameter.category, e.parameter.itemName));
      case 'deleteMaterial':
        return jsonResponse(handleDeleteMaterial(e.parameter.login, e.parameter.category, e.parameter.itemName));
      case 'deleteCategory':
        return jsonResponse(handleDeleteCategory(e.parameter.login, e.parameter.category));
      case 'renameCategory':
        return jsonResponse(handleRenameCategory(e.parameter.login, e.parameter.oldName, e.parameter.newName));
      case 'editMaterial':
        return jsonResponse(handleEditMaterial(e.parameter.login, e.parameter.category, e.parameter.oldName, e.parameter.newName));
      case 'reorderItems':
        return jsonResponse(handleReorderItems(payload.login, payload.category, payload.orderedNames));
      case 'reorderCategories':
        return jsonResponse(handleReorderCategories(payload.login, payload.orderedCategories));
      case 'addComment':
        return jsonResponse(handleAddComment(payload.ticketId, payload.login, payload.comment));
      case 'getComments':
        return jsonResponse(handleGetComments(e.parameter.ticketId));
      default:
        return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ===================== ACTIONS =====================

// Логін
function handleLogin(login, password) {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet);
  const user = users.find(u => u.login === login && String(u.password) === String(password));
  if (!user) return { ok: false, error: 'Невірний логін або пароль' };
  return {
    ok: true,
    user: { login: user.login, display_name: user.display_name, role: user.role }
  };
}

// Отримати список користувачів (для адмінки — призначення)
function handleGetUsers() {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet).map(u => ({
    login: u.login, display_name: u.display_name, role: u.role
  }));
  return { ok: true, users };
}

// Отримати заявки. Якщо login переданий і роль executor — тільки призначені йому.
// Адмін бачить все.
function handleGetTickets(login) {
  const ticketsSheet = getSheet(SHEET_TICKETS);
  let tickets = sheetToObjects(ticketsSheet);

  if (login) {
    const usersSheet = getSheet(SHEET_USERS);
    const users = sheetToObjects(usersSheet);
    const user = users.find(u => u.login === login);
    if (user && user.role === 'executor') {
      tickets = tickets.filter(t => t.assigned_to === login);
    }
  }
  // Прибираємо приховані (тільки для виконавця, що сам приховав)
  // hidden_by — список логінів через кому, хто приховав цю заявку
  if (login) {
    tickets = tickets.filter(t => {
      const hiddenBy = String(t.hidden_by || '').split(',').map(s => s.trim());
      return !hiddenBy.includes(login);
    });
  }
  return { ok: true, tickets };
}

// Імпорт заявок з xls (заміняє весь лист Tickets, зберігаючи assigned_to/status/hidden_by для заявок з тим самим id)
function handleImportTickets(newTickets, importedBy) {
  if (!newTickets || !Array.isArray(newTickets)) {
    return { ok: false, error: 'Немає даних для імпорту' };
  }
  const sheet = getSheet(SHEET_TICKETS);
  const oldTickets = sheetToObjects(sheet);
  const oldById = {};
  oldTickets.forEach(t => { oldById[t.id] = t; });

  const headers = ['id','num','prio','zayav','date','name','phone','phones','addr','addrRaw',
    'problem','resp','hours','service','comment','isFreenet','account','plannedDate','assigned_to','status','hidden_by'];

  const rows = newTickets.map(t => {
    const old = oldById[t.id] || {};
    return [
      t.id, t.num, t.prio, t.zayav, t.date, t.name, t.phone, t.phones || '', t.addr, t.addrRaw,
      t.problem, t.resp, t.hours, t.service, t.comment, t.isFreenet,
      t.account || '', t.plannedDate || '',
      old.assigned_to || '', old.status || 'new', old.hidden_by || ''
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  if (rows.length > 0) {
    // Примушуємо стовпці з датами (date, plannedDate) зберігатись як ТЕКСТ,
    // інакше Google Sheets автоматично конвертує "27.06.2026 21:00" в Date object,
    // який при зчитуванні повертається як ISO-рядок (2026-06-27T21:00:00.000Z)
    const dateColIdx = headers.indexOf('date') + 1;
    const plannedDateColIdx = headers.indexOf('plannedDate') + 1;
    sheet.getRange(2, dateColIdx, rows.length, 1).setNumberFormat('@STRING@');
    sheet.getRange(2, plannedDateColIdx, rows.length, 1).setNumberFormat('@STRING@');

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Update settings
  const settingsSheet = getSheet(SHEET_SETTINGS);
  const settingsData = settingsSheet.getDataRange().getValues();
  for (let i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] === 'last_import') settingsSheet.getRange(i + 1, 2).setValue(new Date().toISOString());
    if (settingsData[i][0] === 'imported_by') settingsSheet.getRange(i + 1, 2).setValue(importedBy || '');
  }

  return { ok: true, count: rows.length };
}

// Призначити заявку виконавцю (адмін)
function handleAssignTicket(ticketId, assignTo) {
  const sheet = getSheet(SHEET_TICKETS);
  const rowIdx = findRowIndexById(sheet, 'id', ticketId);
  if (rowIdx === -1) return { ok: false, error: 'Заявку не знайдено' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const assignedCol = headers.indexOf('assigned_to') + 1;
  sheet.getRange(rowIdx, assignedCol).setValue(assignTo || '');
  return { ok: true };
}

// Сховати заявку (виконавець свайпнув картку)
function handleHideTicket(ticketId, login) {
  const sheet = getSheet(SHEET_TICKETS);
  const rowIdx = findRowIndexById(sheet, 'id', ticketId);
  if (rowIdx === -1) return { ok: false, error: 'Заявку не знайдено' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hiddenCol = headers.indexOf('hidden_by') + 1;
  const current = sheet.getRange(rowIdx, hiddenCol).getValue();
  const list = String(current || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!list.includes(login)) list.push(login);
  sheet.getRange(rowIdx, hiddenCol).setValue(list.join(','));
  return { ok: true };
}

// ===================== MATERIALS (особисті матеріали виконавця) =====================

// Отримати всі матеріали користувача, згруповані по категоріях, з урахуванням порядку
function handleGetMaterials(login) {
  if (!login) return { ok: false, error: 'Не вказано login' };
  const sheet = getSheet(SHEET_MATERIALS);
  const rows = sheetToObjects(sheet).filter(r => r.login === login);

  // Визначаємо порядок категорій (мінімальний cat_order для кожної категорії)
  const catOrderMap = {};
  rows.forEach(r => {
    const co = parseInt(r.cat_order);
    const order = isNaN(co) ? 9999 : co;
    if (!(r.category in catOrderMap) || order < catOrderMap[r.category]) {
      catOrderMap[r.category] = order;
    }
  });

  const categoriesList = Object.keys(catOrderMap).sort((a, b) => catOrderMap[a] - catOrderMap[b]);

  const categories = {};
  categoriesList.forEach(cat => {
    const items = rows
      .filter(r => r.category === cat)
      .sort((a, b) => {
        const ao = parseInt(a.item_order); const bo = parseInt(b.item_order);
        return (isNaN(ao) ? 9999 : ao) - (isNaN(bo) ? 9999 : bo);
      })
      .map(r => r.item_name);
    categories[cat] = items;
  });

  return { ok: true, categories: categories };
}

// Додати матеріал (і, за потреби, нову категорію — створюється автоматично)
function handleAddMaterial(login, category, itemName) {
  if (!login || !category || !itemName) return { ok: false, error: 'Не всі поля заповнені' };
  const sheet = getSheet(SHEET_MATERIALS);
  const rows = sheetToObjects(sheet);
  const userRows = rows.filter(r => r.login === login);

  const exists = userRows.some(r => r.category === category && r.item_name === itemName);
  if (exists) return { ok: false, error: 'Така позиція вже є' };

  // cat_order: якщо категорія вже існує — беремо її order, інакше — наступний вільний
  const sameCategoryRows = userRows.filter(r => r.category === category);
  let catOrder;
  if (sameCategoryRows.length > 0) {
    catOrder = parseInt(sameCategoryRows[0].cat_order) || 0;
  } else {
    const maxCatOrder = userRows.reduce((max, r) => Math.max(max, parseInt(r.cat_order) || 0), -1);
    catOrder = maxCatOrder + 1;
  }

  // item_order: наступний вільний всередині категорії
  const maxItemOrder = sameCategoryRows.reduce((max, r) => Math.max(max, parseInt(r.item_order) || 0), -1);
  const itemOrder = maxItemOrder + 1;

  sheet.appendRow([login, category, itemName, catOrder, itemOrder]);
  return { ok: true };
}

// Видалити одну позицію матеріалу
function handleDeleteMaterial(login, category, itemName) {
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === login && data[i][1] === category && data[i][2] === itemName) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Позицію не знайдено' };
}

// Видалити цілу категорію (всі позиції в ній) для користувача
function handleDeleteCategory(login, category) {
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === login && data[i][1] === category) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// Перейменувати категорію
function handleRenameCategory(login, oldName, newName) {
  if (!newName || !newName.trim()) return { ok: false, error: 'Нова назва не може бути порожньою' };
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === login && data[i][1] === oldName) {
      sheet.getRange(i + 1, 2).setValue(newName);
    }
  }
  return { ok: true };
}

// Редагувати назву позиції (зберігаючи її порядок)
function handleEditMaterial(login, category, oldName, newName) {
  if (!newName || !newName.trim()) return { ok: false, error: 'Назва не може бути порожньою' };
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('item_name') + 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === login && data[i][1] === category && data[i][2] === newName.trim()) {
      return { ok: false, error: 'Така позиція вже є' };
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === login && data[i][1] === category && data[i][2] === oldName) {
      sheet.getRange(i + 1, nameCol).setValue(newName.trim());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Позицію не знайдено' };
}

// Зберегти новий порядок позицій всередині категорії
function handleReorderItems(login, category, orderedNames) {
  if (!login || !category || !orderedNames || !Array.isArray(orderedNames)) {
    return { ok: false, error: 'Невірні дані' };
  }
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const itemOrderCol = headers.indexOf('item_order') + 1;

  const orderMap = {};
  orderedNames.forEach((name, idx) => { orderMap[name] = idx; });

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === login && data[i][1] === category) {
      const itemName = data[i][2];
      if (itemName in orderMap) {
        sheet.getRange(i + 1, itemOrderCol).setValue(orderMap[itemName]);
      }
    }
  }
  return { ok: true };
}

// Зберегти новий порядок категорій
function handleReorderCategories(login, orderedCategories) {
  if (!login || !orderedCategories || !Array.isArray(orderedCategories)) {
    return { ok: false, error: 'Невірні дані' };
  }
  const sheet = getSheet(SHEET_MATERIALS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const catOrderCol = headers.indexOf('cat_order') + 1;

  const orderMap = {};
  orderedCategories.forEach((cat, idx) => { orderMap[cat] = idx; });

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === login) {
      const category = data[i][1];
      if (category in orderMap) {
        sheet.getRange(i + 1, catOrderCol).setValue(orderMap[category]);
      }
    }
  }
  return { ok: true };
}

// ===================== ЗАПОВНИТИ БАЗОВИЙ ПЕРЕЛІК МАТЕРІАЛІВ (запустити вручну один раз) =====================
function seedDefaultMaterials() {
  const sheet = getSheet(SHEET_MATERIALS);

  // Базовий перелік
  const defaultMaterials = {
    'Кабель': [
      'FTTH UMM 1F G657A1',
      'ок-2',
      'Finmark UT004-SM-16-1KN',
      'ок-8',
      'ок-12',
      'ок-24'
    ],
    'Гільзи': [
      'гильзы FTTH 60mm',
      'гильзы 45mm'
    ],
    'Патчкорд': [
      'Патчкорд SC/UPC-SC/UPC SM-3.0 3m'
    ],
    'Зажими': [
      'натяжной зажим H-15',
      'натяжной зажим H-3'
    ]
  };

  // Усі логіни виконавців (без адміна — йому склад не потрібен)
  const usersSheet = getSheet(SHEET_USERS);
  const allUsers = sheetToObjects(usersSheet);
  const executorLogins = allUsers.filter(u => u.role === 'executor').map(u => u.login);

  // Уникаємо дублікатів — перевіряємо що вже є
  const existing = sheetToObjects(sheet);
  const existsSet = new Set(existing.map(r => r.login + '|||' + r.category + '|||' + r.item_name));

  const rowsToAdd = [];
  executorLogins.forEach(login => {
    Object.keys(defaultMaterials).forEach((category, catIdx) => {
      defaultMaterials[category].forEach((item, itemIdx) => {
        const key = login + '|||' + category + '|||' + item;
        if (!existsSet.has(key)) {
          rowsToAdd.push([login, category, item, catIdx, itemIdx]);
        }
      });
    });
  });

  if (rowsToAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
  }

  Logger.log('Додано рядків: ' + rowsToAdd.length + ' для користувачів: ' + executorLogins.join(', '));
}
