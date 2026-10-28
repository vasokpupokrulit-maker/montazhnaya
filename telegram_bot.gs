/**
 * Монтажная → Telegram
 * Присылает в Telegram сообщение о каждой новой задаче из вкладки «Задачи».
 *
 * Установка (один раз):
 * 1. Расширения → Apps Script, вставить весь этот код вместо того, что там есть, сохранить.
 * 2. Слева «Настройки проекта» (шестерёнка) → «Свойства скрипта» → добавить
 *    свойство TELEGRAM_TOKEN со значением токена от @BotFather.
 * 3. Написать боту в Telegram любое сообщение (или добавить бота в группу и написать там).
 * 4. Вверху выбрать функцию setup и нажать «Выполнить», разрешить доступ.
 *    Бот пришлёт «Бот подключён». Дальше всё работает само.
 */

const SHEET_NAME = 'Задачи';
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/16e-TyCsfDIs6V1PEPcdWJkQosVCYE6ZmqKBcz_ZIsMw/edit';
const OFFICE_URL = 'https://vasokpupokrulit-maker.github.io/montazhnaya/';
const ME = 'Никита'; // задачи на этого человека помечаются отдельно

// Колонки вкладки «Задачи» (A = 0)
const COL = { id: 0, branch: 1, project: 2, title: 3, type: 4, from: 5, who: 6, prio: 7, due: 8, status: 9, src: 11, format: 13 };

/** Запустить один раз: находит чат, отмечает старые задачи как известные, ставит проверку раз в минуту. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TELEGRAM_TOKEN')) {
    throw new Error('Сначала добавь свойство скрипта TELEGRAM_TOKEN (Настройки проекта → Свойства скрипта).');
  }
  if (!props.getProperty('CHAT_ID')) {
    const chatId = findChatId_();
    if (!chatId) throw new Error('Бот пока не видит ни одного чата. Напиши боту любое сообщение в Telegram и запусти setup ещё раз.');
    props.setProperty('CHAT_ID', String(chatId));
  }

  // Всё, что уже есть в таблице, считаем известным, чтобы не прислать пачку старых задач.
  const seen = {};
  readTasks_().forEach(t => { seen[t.id] = true; });
  props.setProperty('SEEN', JSON.stringify(seen));
  props.deleteProperty('PENDING');

  ScriptApp.getProjectTriggers()
    .filter(tr => tr.getHandlerFunction() === 'checkNewTasks')
    .forEach(tr => ScriptApp.deleteTrigger(tr));
  ScriptApp.newTrigger('checkNewTasks').timeBased().everyMinutes(1).create();

  send_('✈️ Бот «Монтажная» подключён. Буду писать сюда о каждой новой задаче.');
}

/** Запускается раз в минуту. Задача отправляется, когда её строка не менялась между двумя проверками. */
function checkNewTasks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const seen = JSON.parse(props.getProperty('SEEN') || '{}');
    const pending = JSON.parse(props.getProperty('PENDING') || '{}');
    const nextPending = {};

    readTasks_().forEach(t => {
      if (seen[t.id]) return;
      // Ждём, пока человек допишет строку: отправляем, только если за минуту ничего не поменялось.
      if (pending[t.id] === t.hash) {
        send_(format_(t));
        seen[t.id] = true;
      } else {
        nextPending[t.id] = t.hash;
      }
    });

    props.setProperty('SEEN', JSON.stringify(seen));
    props.setProperty('PENDING', JSON.stringify(nextPending));
  } finally {
    lock.releaseLock();
  }
}

/** Отправить себе тестовое сообщение. */
function testMessage() {
  send_('Проверка связи: бот «Монтажная» на месте 🍺');
}

/** Отключить уведомления. */
function stop() {
  ScriptApp.getProjectTriggers().forEach(tr => ScriptApp.deleteTrigger(tr));
}

// ---------- внутреннее ----------

function readTasks_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 16).getDisplayValues();
  return rows
    .map((r, i) => ({ r: r, row: i + 2 }))
    .filter(x => String(x.r[COL.title]).trim() && String(x.r[COL.who]).trim())
    .filter(x => !/^пример/i.test(String(x.r[COL.project]).trim()))
    .map(x => {
      const r = x.r.map(v => String(v).trim());
      return {
        id: r[COL.id] || ('row' + x.row), row: x.row,
        branch: r[COL.branch], project: r[COL.project], title: r[COL.title], type: r[COL.type],
        from: r[COL.from], who: r[COL.who], prio: r[COL.prio], due: r[COL.due],
        status: r[COL.status], src: r[COL.src], format: r[COL.format],
        hash: r.join('\u0001')
      };
    });
}

function format_(t) {
  const lines = [];
  const hot = t.prio === 'Горит' ? ' 🔥' : '';
  lines.push(`🆕 <b>Новая задача ${esc_(t.id)}</b> · ${esc_(t.branch || 'Монтаж')}${hot}`);
  lines.push(`<b>${esc_(t.title)}</b>`);
  if (t.project) lines.push(`Проект: ${esc_(t.project)}`);
  if (t.type) lines.push(`Тип: ${esc_(t.type)}`);
  lines.push(`Исполнитель: ${esc_(t.who)}${t.who === ME ? ' 👈 это тебе' : ''}`);
  if (t.from) lines.push(`Поставил(а): ${esc_(t.from)}`);
  if (t.due) lines.push(`Дедлайн: ${esc_(t.due)}`);
  if (t.prio && t.prio !== 'Обычный') lines.push(`Приоритет: ${esc_(t.prio)}`);
  if (t.format) lines.push(`Формат: ${esc_(t.format)}`);
  if (/^https?:\/\//.test(t.src)) lines.push(`<a href="${esc_(t.src)}">Исходники</a>`);
  lines.push('');
  lines.push(`<a href="${SHEET_URL}">Таблица</a> · <a href="${OFFICE_URL}">Офис</a>`);
  return lines.join('\n');
}

function send_(html) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_TOKEN');
  const chatIds = String(props.getProperty('CHAT_ID') || '').split(',').map(s => s.trim()).filter(Boolean);
  chatIds.forEach(chatId => {
    const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: html, parse_mode: 'HTML', disable_web_page_preview: true }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) console.error('Telegram ответил ' + res.getResponseCode() + ': ' + res.getContentText());
  });
}

function findChatId_() {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/getUpdates`, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Токен не подходит. Проверь TELEGRAM_TOKEN.');
  const updates = JSON.parse(res.getContentText()).result || [];
  for (let i = updates.length - 1; i >= 0; i--) {
    const m = updates[i].message || updates[i].channel_post || updates[i].my_chat_member;
    if (m && m.chat) return m.chat.id;
  }
  return null;
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
