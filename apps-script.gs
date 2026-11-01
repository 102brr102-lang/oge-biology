// БиоХак ОГЭ — сбор результатов учеников в Google-таблицу.
// 1) Создайте Google-таблицу -> Расширения -> Apps Script, вставьте этот код.
// 2) Замените KEY на свой секретный ключ (любая строка, например: lisichka-2026).
// 3) Начать развертывание -> Новое развертывание -> Веб-приложение:
//    "Выполнять от имени": Я, "Доступ": Все. Скопируйте адрес веб-приложения (заканчивается на /exec).
// 4) В панели учителя БиоХак ОГЭ: «Сбор результатов» -> вставьте адрес и ключ.

const KEY = 'ЗАМЕНИТЕ_НА_СВОЙ_КЛЮЧ';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Результаты');
  if (!sh) {
    sh = ss.insertSheet('Результаты');
    sh.appendRow(['Имя', 'Обновлено', 'Тренировка, %', 'Зачтено линий', 'Средний зачёт, %', 'Данные (JSON)']);
  }
  return sh;
}

// Ученик отправляет свой прогресс (одна строка на ученика, обновляется)
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return ContentService.createTextOutput('bad'); }
  if (!d || !d.name) return ContentService.createTextOutput('bad');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow();
    const names = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => r[0]) : [];
    const i = names.indexOf(d.name);
    const row = [d.name, new Date(d.ts || Date.now()), Math.round((d.trainPct || 0) * 100), d.passed || 0, Math.round((d.avgPct || 0) * 100), JSON.stringify(d)];
    if (i >= 0) sh.getRange(i + 2, 1, 1, 6).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}

// Учитель читает результаты (нужен ключ)
function doGet(e) {
  const out = (o) => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (!e.parameter || e.parameter.key !== KEY) return out({ error: 'key' });
  const sh = sheet_();
  const n = sh.getLastRow();
  const res = [];
  if (n > 1) sh.getRange(2, 6, n - 1, 1).getValues().forEach(r => { try { res.push(JSON.parse(r[0])); } catch (err) {} });
  return out(res);
}
