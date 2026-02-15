// Этот файл является частью time-to-table //
// SPDX-License-Identifier: GPL-3.0-or-later //

"use strict";

// === БЕЗОПАСНОЕ ЛОГИРОВАНИЕ ===
// В production режиме ограничиваем вывод чувствительной информации
const isDevelopment = () => {
    try {
        return globalThis.location?.hostname === 'localhost' || 
               globalThis.location?.hostname === 'tauri.localhost' ||
               globalThis.__TAURI_INTERNALS__?.postMessage !== undefined;
    } catch {
        return false;
    }
};

// Безопасное логирование ошибок - в production выводит только сообщение без деталей
function safeLogError(message, error) {
    if (isDevelopment()) {
        console.error(message, error);
    } else {
        // В production только общее сообщение без stack trace
        console.error(message);
    }
}

// Безопасное debug логирование - только в dev режиме
function safeDebug(...args) {
    if (isDevelopment() && console.debug) {
        console.debug(...args);
    }
}

// === TAURI API ===
let tauriDialog = null;
let tauriInvoke = null;

// Инициализация Tauri API после загрузки
async function initTauriApi() {
    if (globalThis.__TAURI__) {
        try {
            // В Tauri v2 модули доступны через __TAURI__
            tauriDialog = globalThis.__TAURI__.dialog;
            tauriInvoke = globalThis.__TAURI__.core.invoke;
            safeDebug('Tauri API доступен');
        } catch (e) {
            safeLogError('Tauri API init error:', e);
        }
    }
}

// Безопасная запись файла через Rust команду
async function saveFileSecure(path, content) {
    if (tauriInvoke) {
        return await tauriInvoke('save_file_secure', { path, content });
    }
    throw new Error('Tauri недоступен');
}

// Вызываем при загрузке с задержкой для гарантии загрузки Tauri
globalThis.addEventListener('DOMContentLoaded', () => {
    setTimeout(initTauriApi, 100);
});

// === ФУНКЦИИ БЕЗОПАСНОСТИ ===
function sanitizeInput(str, maxLength = 500) {
    if (typeof str !== 'string') return '';
    return str.substring(0, maxLength).trim();
}

// Строгая санитизация для названий/описаний: допускаем буквы (лат/кириллица), цифры, запятую, точку, символ № и пробел
function sanitizeStrict(str, maxLength = 500) {
    if (typeof str !== 'string') return '';
    const cleaned = String(str).replaceAll(/[^A-Za-z\u0400-\u04FF0-9,.№ _/-]+/g, '');
    return cleaned.substring(0, maxLength);
}

// Удаляет ведущий порядковый префикс вида "1) ", "2) " и т.п.
function stripOrdinalPrefix(str) {
    if (typeof str !== 'string') return '';
    return str.replace(/^\s*\d+\)\s*/, '');
}

// Возвращает hex-строку из `bytes` случайных байт, используя crypto.getRandomValues при наличии.
// Откат на crypto.randomUUID() (без дефисов) или hex-строку из timestamp+счётчик.
function secureRandomHex(bytes = 8) {
    try {
        if (globalThis.crypto?.getRandomValues) {
            const arr = new Uint8Array(bytes);
            globalThis.crypto.getRandomValues(arr);
            return Array.from(arr).map(b => b.toString(16).padStart(2, '0')).join('');
        }
        if (globalThis.crypto?.randomUUID) {
            return globalThis.crypto.randomUUID().replaceAll('-', '');
        }
    } catch (e) {
        console.debug?.('secureRandomHex crypto error:', e?.message);
    }
    // Крайний запасной вариант: timestamp + performance + счётчик.
    secureRandomHex._counter = (secureRandomHex._counter || 0) + 1;
    const nowHex = Date.now().toString(16);
    const perfHex = performance?.now ? Math.floor(performance.now()).toString(16) : '0';
    return nowHex + perfHex + secureRandomHex._counter.toString(16);
}

// Обёртки localStorage с логированием ошибок.
// localStorage синхронный и однопоточный в рамках одного origin (Tauri webview).
async function safeLocalStorageSet(key, value) {
    try {
        localStorage.setItem(key, value);
    } catch (e) {
        safeLogError('localStorage set error:', e);
        throw e;
    }
}

async function safeLocalStorageRemove(key) {
    try {
        localStorage.removeItem(key);
    } catch (e) {
        safeLogError('localStorage remove error:', e);
    }
}

// Защита от Excel-инъекций: если текст начинается с символов формулы,
// добавляем ведущую апостроф-кавычку, чтобы Excel воспринимал это как текст.
function excelSanitizeCell(str) {
    if (typeof str !== 'string') return '';
    if (str.length === 0) return '';
    const first = str[0];
    if (['=', '+', '-', '@'].includes(first)) return "'" + str;
    return str;
}

// Санитизация числового ввода: допускаем до 5 цифр в целой части и до 2 цифр в дробной.
function sanitizeDecimalInput(raw) {
    if (raw === null || raw === undefined) return '';
    let s = String(raw);
    // Оставляем только цифры и разделители . и ,
    s = s.replaceAll(/[^0-9.,]/g, '');
    // Найдём первый разделитель
    const m = /[.,]/.exec(s);
    if (!m) {
        // Только целая часть, обрезаем до 5 цифр
        return s.slice(0, 5);
    }
    const sep = m[0];
    const idx = s.indexOf(sep);
    let intPart = s.slice(0, idx).replaceAll(/[.,]/g, '').slice(0, 5);
    let fracPart = s.slice(idx + 1).replaceAll(/[.,]/g, '').slice(0, 2);
    // Если дробная часть ещё пустая — возвращаем с точкой, чтобы пользователь мог продолжить вводить
    if (fracPart.length === 0) return intPart + '.';
    // Нормализуем разделитель на точку для дальнейшего парсинга
    return intPart + '.' + fracPart; // используем точку внутренно для парсинга
}

function validateNumber(value, min, max) {
    const num = Number.parseInt(value, 10);
    if (Number.isNaN(num)) return min;
    return Math.max(min, Math.min(max, num));
}

function formatDateISO(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function formatTimeHMS(date) {
    const h = String(date.getHours()).padStart(2, '0');
    const m = String(date.getMinutes()).padStart(2, '0');
    const s = String(date.getSeconds()).padStart(2, '0');
    return `${h}:${m}:${s}`;
}

// === ОБЩИЕ ХЕЛПЕРЫ  ===
// Универсальный диалог подтверждения (Tauri + fallback)
async function confirmAction(message, title = 'Подтверждение') {
    if (tauriDialog?.confirm) {
        try {
            return await tauriDialog.confirm(message, { title, kind: 'warning' });
        } catch (e) {
            safeLogError('Tauri confirm error:', e);
            return globalThis.confirm(message);
        }
    }
    return globalThis.confirm(message);
}

// Универсальное сообщение (Tauri + fallback)
async function showMessage(message, title = 'Информация', kind = 'info') {
    if (tauriDialog?.message) {
        try {
            await tauriDialog.message(message, { title, kind });
            return;
        } catch (e) { safeLogError('Tauri message error:', e); }
    }
    alert(message);
}

// Диалог конфликта при импорте — 3 варианта: 'all' | 'new' | 'cancel'
function showImportConflictDialog(conflicts) {
    return new Promise((resolve) => {
        const modal = document.getElementById('importConflictModal');
        const textEl = document.getElementById('importConflictText');
        const overwriteBtn = document.getElementById('importConflictOverwriteBtn');
        const newOnlyBtn = document.getElementById('importConflictNewOnlyBtn');
        const cancelBtn = document.getElementById('importConflictCancelBtn');
        const closeBtn = document.getElementById('closeImportConflictModal');

        const names = conflicts.map(k => '  • ' + k.replaceAll('z7_card_', '')).join('\n');
        textEl.textContent = `Следующие техкарты уже существуют (${conflicts.length} шт.):\n\n${names}`;

        function cleanup(result) {
            modal.classList.remove('active');
            overwriteBtn.removeEventListener('click', onOverwrite);
            newOnlyBtn.removeEventListener('click', onNewOnly);
            cancelBtn.removeEventListener('click', onCancel);
            closeBtn.removeEventListener('click', onCancel);
            resolve(result);
        }

        function onOverwrite() { cleanup('all'); }
        function onNewOnly()   { cleanup('new'); }
        function onCancel()    { cleanup('cancel'); }

        overwriteBtn.addEventListener('click', onOverwrite);
        newOnlyBtn.addEventListener('click', onNewOnly);
        cancelBtn.addEventListener('click', onCancel);
        closeBtn.addEventListener('click', onCancel);

        modal.classList.add('active');
    });
}

// Определяет суффикс единицы измерения для заголовка таблицы
function getHeaderUnitSuffix(rows) {
    const uniqueUnits = [...new Set(rows.map(r => r.unit || 'min'))];
    if (uniqueUnits.length === 1) {
        if (uniqueUnits[0] === 'min') return ' (мин)';
        if (uniqueUnits[0] === 'hour') return ' (час)';
    }
    return '';
}

// Создаёт DOM-элемент Z7 таблицы
function createZ7TableElement(z7Lines) {
    const z7Table = createEl('table', { className: 'history-z7', style: 'width:100%; border-collapse:collapse;' });
    const z7Head = createEl('thead');
    const thZ7 = createEl('th', { className: 'z7-header-common', colspan: '12' }, 'Z7');
    const z7HeadTr = createEl('tr');
    z7HeadTr.append(thZ7);
    z7Head.append(z7HeadTr);
    const z7Body = createEl('tbody');
    const z7Tr = createEl('tr');
    const z7Td = createEl('td');
    z7Lines.forEach(line => z7Td.append(createEl('div', { className: 'z7-line-item' }, line)));
    z7Tr.append(z7Td);
    z7Body.append(z7Tr);
    z7Table.append(z7Head, z7Body);
    return z7Table;
}

// Ключ localStorage для пользовательских умолчаний (Настройки)
const DEFAULTS_KEY = 'z7_defaults';

// Возвращает '#000000' или '#FFFFFF' — цвет текста с достаточным контрастом (WCAG) для заданного фона
function getContrastColor(hex) {
    const r = parseInt(hex.slice(1, 3), 16) / 255;
    const g = parseInt(hex.slice(3, 5), 16) / 255;
    const b = parseInt(hex.slice(5, 7), 16) / 255;
    const lin = c => c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    const L = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
    return L > 0.179 ? '#000000' : '#FFFFFF';
}

// Встроенные умолчания
function _builtinDefaults() {
    return {
        chainMode: true,
        timeMode: 'total',
        statusBefore: 'замечаний нет',
        workExtra: 'нет',
        devRec: 'нет',
        sortMode: 'sequential',
        theme: 'light',
        excelColors: {
            locked:    '#F7A1A1', // Заблокированные ячейки
            editable:  '#FFFFFF', // Редактируемые ячейки
            header:    '#b98109', // Заголовки таблиц
            author:    '#EDF7ED', // Строка настроек/автора
            confirmed: '#D5F5D5', // Статус «Подтверждено»
            pdtv:      '#FFF9C4'  // ПДТВ (авто-формула)
        }
    };
}

// === ТЕМА ===
// Применяет тему мгновенно, добавляя класс к body и переключая тему окна через Tauri API
function applyTheme(theme) {
    if (theme === 'dark') {
        document.body.classList.add('dark');
    } else {
        document.body.classList.remove('dark');
    }
    try {
        const win = globalThis.__TAURI__?.webviewWindow?.getCurrentWebviewWindow?.();
        if (win?.setTheme) {
            win.setTheme(theme === 'dark' ? 'dark' : 'light').catch(() => {});
        }
    } catch (e) { console.debug?.('setTheme error', e?.message); }
}

// Применяем тему как можно раньше, чтобы избежать вспышки
(function earlyApplyTheme() {
    try {
        const raw = localStorage.getItem(DEFAULTS_KEY);
        if (raw) {
            const d = safeJsonParse(raw);
            if (d?.theme === 'dark') {
                document.body.classList.add('dark');
                // Заголовок окна переключим после инициализации Tauri
                globalThis.addEventListener('DOMContentLoaded', () => {
                    setTimeout(() => {
                        try {
                            const win = globalThis.__TAURI__?.webviewWindow?.getCurrentWebviewWindow?.();
                            if (win?.setTheme) win.setTheme('dark').catch(() => {});
                        } catch (error_) {
                            console.debug?.('earlyApplyTheme setTheme error', error_?.message);
                        }
                    }, 150);
                });
            }
        }
    } catch (error_) {
        console.debug?.('earlyApplyTheme error', error_?.message);
    }
})();

// Загружает пользовательские умолчания (или встроенные, если не заданы)
function getUserDefaults() {
    try {
        const raw = localStorage.getItem(DEFAULTS_KEY);
        if (raw) {
            const d = safeJsonParse(raw);
            if (d && typeof d === 'object') return Object.assign(_builtinDefaults(), d);
        }
    } catch (e) { console.debug?.('getUserDefaults error', e?.message); }
    return _builtinDefaults();
}

// Общие значения по умолчанию для формы (учитывают пользовательские настройки)
function getFormDefaults() {
    const _todayStr = formatDateISO(new Date());
    const ud = getUserDefaults();
    return {
        totalOps: 1, workerCount: 1, startDate: _todayStr, startTime: '08:00:00',
        chainMode: ud.chainMode, lunchStart: '12:00', lunchStart2: '00:00', lunchDur: 45,
        timeMode: ud.timeMode, resIz: '', coefK: '', orderName: '', itemName: '',
        postingDate: _todayStr, statusBefore: ud.statusBefore, workExtra: ud.workExtra, devRec: ud.devRec,
        sortMode: ud.sortMode
    };
}

// Обработчик ввода только цифр
function digitOnlyHandler(e) {
    e.target.value = e.target.value.replaceAll(/\D/g, '');
}

// Обработчики для десятичных полей (sanitizeDecimalInput)
function decimalInputHandler(e) {
    e.target.value = sanitizeDecimalInput(e.target.value);
}
function decimalBlurHandler(e) {
    let v = sanitizeDecimalInput(e.target.value);
    if (v === '') v = '0';
    // Обеспечиваем 2 знака после запятой
    const num = Number.parseFloat(v);
    if (!Number.isNaN(num)) {
        v = num.toFixed(2);
    }
    e.target.value = v;
}

// Кнопка удаления записи истории с подтверждением
function createHistoryDeleteButton(entryDiv) {
    const delBtn = createEl('button', { className: 'btn-sm btn-del-history' }, 'Удалить');
    delBtn.onclick = async () => {
        if (await confirmAction('Удалить эту запись из истории?')) {
            entryDiv.remove();
            await saveHistoryToStorage();
            updateFirstPauseVisibility();
        }
    };
    return delBtn;
}

// Разблокировка контролов формы (НЕ трогает workerCount — он сбрасывается только через "Сброс")
function unlockFormControls() {
    const ids = [
        { id: 'totalOps', cls: 'locked-input' },
        { id: 'techCardSelect', cls: 'locked-input' },
        { id: 'saveCardBtn', cls: 'locked-control' },
        { id: 'deleteCardBtn', cls: 'locked-control' },
        { id: 'analyzeCardBtn', cls: 'locked-control' }
    ];
    ids.forEach(({ id, cls }) => {
        try {
            const el = document.getElementById(id);
            if (el) { el.disabled = false; el.classList.remove(cls); el.title = ''; }
        } catch (e) {}
    });
    // Разблокировка поля поиска кастомного dropdown
    if (globalThis._tcDropdown) globalThis._tcDropdown.unlock();
    // Восстановление кнопок удаления и снятие состояния удаления при разблокировке
    try {
        document.querySelectorAll('.op-block').forEach(b => {
            b.classList.remove('deleted-op');
            // Кнопки удаления всегда видимы, сброс display не требуется
        });
    } catch (e) {}
}

// Полная разблокировка всех контролов (включая workerCount) — только для "Сброс"
function unlockAllFormControls() {
    unlockFormControls();
    try {
        const wcEl = document.getElementById('workerCount');
        if (wcEl) { wcEl.disabled = false; wcEl.classList.remove('locked-input'); wcEl.title = ''; }
    } catch (e) {}
}

// Построение Excel формулы сдвига обеда
function buildLunchShiftFormula(rawTimeExpr, lh, lm, lh2, lm2, ld) {
    const l1Val = `TIME(${lh},${lm},0)`;
    const l1End = `(TIME(${lh},${lm},0)+TIME(0,${ld},0))`;
    
    // Используем MOD для проверки времени (игнорируя дату/переполнение суток)
    const tp = `MOD(${rawTimeExpr}, 1)`;
    const cond1 = `AND(${tp}>=${l1Val}, ${tp}<${l1End})`;
    
    // Если попали в обед: берем целую часть (дни) + конец обеда
    const res1 = `(INT(${rawTimeExpr}) + ${l1End})`;
    const shifted1 = `IF(${cond1},${res1},${rawTimeExpr})`;
    
    const hasLunch2 = !(lh2 === 0 && lm2 === 0);
    if (hasLunch2) {
        const l2Val = `TIME(${lh2},${lm2},0)`;
        const l2End = `(TIME(${lh2},${lm2},0)+TIME(0,${ld},0))`;
        
        const tp2 = `MOD(${shifted1}, 1)`;
        const cond2 = `AND(${tp2}>=${l2Val}, ${tp2}<${l2End})`;
        const res2 = `(INT(${shifted1}) + ${l2End})`;
        
        return `IF(${cond2},${res2},${shifted1})`;
    }
    return shifted1;
}

function validateCardData(steps) {
    if (!Array.isArray(steps)) return false;
    return steps.every(s => 
        typeof s.name === 'string' && s.name.length <= 500 &&
        !Number.isNaN(Number.parseFloat(s.dur)) &&
        typeof s.unit === 'string' && ['min', 'hour'].includes(s.unit) &&
        typeof s.hasBreak === 'boolean' &&
        !Number.isNaN(Number.parseFloat(s.breakVal)) &&
        typeof s.breakUnit === 'string' && ['min', 'hour'].includes(s.breakUnit)
    );
}

// Безопасный парсинг JSON с защитой от prototype pollution
function safeJsonParse(jsonString) {
    try {
        const parsed = JSON.parse(jsonString);
        return sanitizeObject(parsed);
    } catch (e) {
        safeLogError('JSON parse error:', e);
        return null;
    }
}

// Очистка объекта от опасных свойств, чтобы предотвратить prototype pollution атаки
function sanitizeObject(obj) {
    if (obj === null || typeof obj !== 'object') {
        return obj;
    }
    
    if (Array.isArray(obj)) {
        return obj.map(sanitizeObject);
    }
    
    const clean = {};
    for (const key of Object.keys(obj)) {
        // Блокируем prototype pollution атаки
        if (key === '__proto__' || key === 'constructor' || key === 'prototype') {
            safeDebug('Blocked potentially dangerous key:', key);
            continue;
        }
        clean[key] = sanitizeObject(obj[key]);
    }
    return clean;
}

function validateImportData(obj) {
    if (typeof obj !== 'object' || obj === null) return false;
    return Object.entries(obj).every(([key, value]) => {
        if (!key.startsWith('z7_card_')) return false;
        // Дополнительная проверка на опасные ключи
        if (key.includes('__proto__') || key.includes('constructor')) return false;
        try {
            const parsed = safeJsonParse(value);
            return parsed && validateCardData(parsed);
        } catch (e) {
            return false;
        }
    });
}

function formatDurationToTime(val, unit) {
    let sec = 0;
    if (unit === 'min') sec = val * 60;
    else if (unit === 'hour') sec = val * 3600;
    else sec = val;
    
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = Math.floor(sec % 60);
    
    return [h, m, s].map(v => String(v).padStart(2, '0')).join(':');
}

// === ИНИЦИАЛИЗАЦИЯ ===
const startDateInput = document.getElementById('startDate');
const postingDateInput = document.getElementById('postingDate');
// Устанавливаем текущую локальную дату в формате YYYY-MM-DD для полей startDate и postingDate
const todayStr = formatDateISO(new Date());
startDateInput.value = todayStr;
if (postingDateInput) postingDateInput.value = todayStr;

const startTimeInput = document.getElementById('startTime');
const container = document.getElementById('fieldsContainer');
// Привязка обработчика к селектору timeMode для переключения UI исполнителей
try {
    const timeModeEl = document.getElementById('timeMode');
    if (timeModeEl) {
        timeModeEl.addEventListener('change', () => updateWorkerUIByTimeMode());
    }
} catch (e) { console.debug?.('attach timeMode listener failed:', e?.message); }
// Состояние модального окна операций
let operationFirstId = ''; // Первый 8-значный номер подтверждения
let lastOperationIndex = null; // Индекс операции, которая будет "последней"
let penultimateOperationIndex = null; // Индекс операции, которая будет "предпоследней"
let autoIncrementEnabled = false; // Состояние чекбокса "авто"
let workerIds = []; // Массив 8-значных номеров исполнителей

// === ПЕРСИСТЕНТНОСТЬ ИСПОЛНИТЕЛЕЙ ===
const WORKERS_SESSION_KEY = 'z7_workers_session';

async function saveWorkersSession() {
    try {
        const wcEl = document.getElementById('workerCount');
        const session = {
            count: wcEl ? Number.parseInt(wcEl.value, 10) || 1 : 1,
            ids: workerIds.slice(),
            locked: wcEl ? wcEl.disabled : false
        };
        await safeLocalStorageSet(WORKERS_SESSION_KEY, JSON.stringify(session));
    } catch (e) { console.debug?.('saveWorkersSession error:', e?.message); }
}

function loadWorkersSession() {
    try {
        const raw = localStorage.getItem(WORKERS_SESSION_KEY);
        if (!raw) return;
        const session = safeJsonParse(raw);
        if (!session || typeof session !== 'object') return;

        const wcEl = document.getElementById('workerCount');
        if (wcEl && session.count) {
            wcEl.value = Math.max(1, Math.min(10, session.count));
        }
        if (Array.isArray(session.ids)) {
            workerIds = session.ids.slice();
        }
        if (session.locked && wcEl) {
            wcEl.disabled = true;
            wcEl.classList.add('locked-input');
            wcEl.title = 'Нажмите "Очистить" (F5) или "Сброс" для разблокировки';
        }
    } catch (e) { console.debug?.('loadWorkersSession error:', e?.message); }
}

// Ограничение ввода в поля 'Заказ' и 'Rиз' — только цифры
try {
    ['orderName', 'resIz'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', digitOnlyHandler);
            el.setAttribute('inputmode', 'numeric');
            el.setAttribute('autocomplete', 'off');
        }
    });
} catch (e) {
    console.debug?.('Digit-only input listeners attach failed:', e?.message);
}

// Ограничение/санитизация ввода для длительности обеда (до 5 цифр + 2 дробных)
try {
    const lunchDurEl = document.getElementById('lunchDur');
    if (lunchDurEl) {
        lunchDurEl.addEventListener('input', decimalInputHandler);
        lunchDurEl.addEventListener('blur', decimalBlurHandler);
        lunchDurEl.setAttribute('inputmode', 'decimal');
        lunchDurEl.setAttribute('autocomplete', 'off');
    }
} catch (e) {
    console.debug?.('lunchDur listener attach failed:', e?.message);
}

// Счётчики символов в реальном времени для statusBefore, workExtra, devRec (макс 300)
try {
    const fields = ['statusBefore', 'workExtra', 'devRec'];
    fields.forEach(id => {
        const el = document.getElementById(id);
        const ctr = document.getElementById(id + '_counter');
        if (!el || !ctr) return;
        const update = () => {
            const max = Number.parseInt(el.getAttribute('maxlength') || '300', 10) || 300;
            const len = String(el.value || '').length;
            const remaining = Math.max(0, max - len);
            ctr.textContent = `осталось ${remaining} / ${max}`;
        };
        // Инициализация
        update();
        el.addEventListener('input', update);
    });
} catch (e) {
    console.debug?.('char counter attach failed:', e?.message);
}

// Ограничение ввода в поле 'Коэф. K' — числа с максимум 2 десятичными знаками
try {
    const kInputEl = document.getElementById('coefK');
    if (kInputEl) {
        kInputEl.addEventListener('input', (e) => {
            let v = String(e.target.value || '');
            // Разрешаем цифры, точку и запятую. Удаляем остальные символы.
            v = v.replaceAll(/[^0-9.,]/g, '');
            // Оставляем только первый разделитель (точку или запятую) и максимум 2 знака дробной части
            const sepMatch = /[.,]/.exec(v);
            if (sepMatch) {
                const sep = sepMatch[0];
                const idx = v.indexOf(sep);
                const intPart = v.slice(0, idx).replaceAll(/[.,]/g, '');
                const dec = v.slice(idx + 1).replaceAll(/[.,]/g, '').slice(0, 2);
                v = intPart + sep + dec;
            } else {
                // Нет разделителя — просто удалить все разделители
                v = v.replaceAll(/[.,]/g, '');
            }
            e.target.value = v;
        });
        kInputEl.setAttribute('inputmode', 'decimal');
        kInputEl.setAttribute('autocomplete', 'off');
    }
} catch (e) {
    console.debug?.('CoefK input listener attach failed:', e?.message);
}

// Строгая санитизация в реальном времени для текстовых полей: itemName, statusBefore, workExtra, devRec
try {
    const itemEl = document.getElementById('itemName');
    if (itemEl) {
        itemEl.addEventListener('input', (e) => {
            const v = sanitizeStrict(e.target.value || '', 70);
            e.target.value = v;
        });
    }

    const strictFields = ['statusBefore', 'workExtra', 'devRec'];
    strictFields.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('input', (e) => {
            const max = Number.parseInt(el.getAttribute('maxlength') || '300', 10) || 300;
            const v = sanitizeStrict(e.target.value || '', max);
            e.target.value = v;
            // обновляем счётчик символов, если существует
            try {
                const ctr = document.getElementById(id + '_counter');
                if (ctr) {
                    const len = String(v).length;
                    const remaining = Math.max(0, max - len);
                    ctr.textContent = `осталось ${remaining} / ${max}`;
                }
            } catch (error_) {
                console.debug?.('strict sanitizer counter update failed:', error_?.message);
            }
        });
    });
} catch (e) {
    console.debug?.('attach strict sanitizers failed:', e?.message);
}

// Синхронизация единиц времени: все операции используют единицу первой операции
function syncTimeUnits() {
    const firstUnitSelect = container.querySelector('.op-block:first-child .op-unit');
    if (!firstUnitSelect) return;
    
    const selectedUnit = firstUnitSelect.value;
    const allUnitSelects = container.querySelectorAll('.op-block .op-unit');
    
    allUnitSelects.forEach((select, idx) => {
        if (idx > 0) { // Пропускаем первую операцию
            select.value = selectedUnit;
        }
    });
}

function createEl(tag, props = {}, text = '') {
    const el = document.createElement(tag);
    for (const [key, value] of Object.entries(props)) {
        if (key.startsWith('on')) continue;
        if (key === 'className') el.className = value;
        else if (key === 'style') {
            value.split(';').forEach(part => {
                const idx = part.indexOf(':');
                if (idx > 0) {
                    el.style.setProperty(part.slice(0, idx).trim(), part.slice(idx + 1).trim());
                }
            });
        }
        else el.setAttribute(key, value);
    }
    if (text) el.textContent = text;
    return el;
}

// Вспомогательная функция для создания подтаблицы в разделённой разметке (для истории и основных результатов)
function createSplitTable(headers, flexGrow = 1) {
    const wrapper = createEl('div', {
        className: 'split-table-wrapper',
        style: `flex-grow:${flexGrow};`
    });
    const table = createEl('table');
    const thead = createEl('thead');
    const trHead = createEl('tr');
    headers.forEach(h => trHead.append(createEl('th', {}, h)));
    thead.append(trHead);
    const tbody = createEl('tbody');
    table.append(thead, tbody);
    wrapper.append(table);
    return { wrapper, tbody };
}


// Заполняет 5 подтаблиц строками данных (используется в generateTable, addToHistoryTable и restoreHistoryFromStorage)
function populateSplitTables(data, tblOps, tblDur, tblPostingDate, tblWorker, tblTime) {
    data.forEach((r, ri) => {
        const isNewOp = ri > 0 && r.originalOpIndex !== data[ri - 1].originalOpIndex;
        const separatorClass = isNewOp ? 'op-separator-row' : '';

        const trOps = createEl('tr');
        if (separatorClass) trOps.className = separatorClass;
        trOps.append(
            createEl('td', {}, r.originalOpIndex || (ri + 1)),
            createEl('td', {}, r.opIdx),
            createEl('td', { style: 'text-align:center; font-weight:600;' }, r.name),
            createEl('td', {}, r.crossedLunch ? '🍽️' : ''),
            createEl('td', { style: 'color: #555;' }, r.pauseText || '')
        );
        tblOps.tbody.append(trOps);

        const trDur = createEl('tr');
        if (separatorClass) trDur.className = separatorClass;
        trDur.append(createEl('td', {}, r.durText));
        tblDur.tbody.append(trDur);

        const trPostingDate = createEl('tr');
        if (separatorClass) trPostingDate.className = separatorClass;
        trPostingDate.append(createEl('td', {}, r.postingDate || ''));
        tblPostingDate.tbody.append(trPostingDate);

        const trWorker = createEl('tr');
        if (separatorClass) trWorker.className = separatorClass;
        trWorker.append(createEl('td', {}, r.worker));
        tblWorker.tbody.append(trWorker);

        const trTime = createEl('tr');
        if (separatorClass) trTime.className = separatorClass;
        trTime.append(
            createEl('td', {}, r.startDate),
            createEl('td', {}, r.startTime),
            createEl('td', {}, r.endDate),
            createEl('td', {}, r.endTime)
        );
        tblTime.tbody.append(trTime);
    });
}

// === КОНСТАНТЫ СЕССИЙ (используются в saveHistoryToStorage и далее) ===
const SESSIONS_META_KEY = 'z7_sessions_meta';
const SESSION_DATA_PREFIX = 'z7_session_data_';
let currentSessionId = null;
let sessionsMeta = []; // [{ id, name, created }]

// === ФУНКЦИИ ДЛЯ СОХРАНЕНИЯ И ЗАГРУЗКИ ИСТОРИИ ===
async function saveHistoryToStorage() {
    try {
        const historyList = document.getElementById('historyList');
        const entries = historyList.querySelectorAll('.history-entry');
        const historyData = Array.from(entries).map(entry => entry.dataset.jsonData);
        await safeLocalStorageSet('z7_history_session', JSON.stringify(historyData));
        // Синхронизируем слот текущей сессии
        if (currentSessionId) {
            await safeLocalStorageSet(SESSION_DATA_PREFIX + currentSessionId, JSON.stringify(historyData));
        }
    } catch (e) {
        console.error('Ошибка при сохранении истории:', e);
    }
}

function restoreHistoryFromStorage() {
    try {
        const historyJson = localStorage.getItem('z7_history_session');
        if (!historyJson) return;
        
        const historyList = document.getElementById('historyList');
        historyList.textContent = '';
        
        const historyData = safeJsonParse(historyJson);
        if (!Array.isArray(historyData)) return;
        
        historyData.forEach(jsonStr => {
            try {
                const data = safeJsonParse(jsonStr);
                if (!data) return;
                const entryDiv = createEl('div', { className: 'history-entry' });
                entryDiv.dataset.jsonData = jsonStr;

                const header = createEl('div', { className: 'history-header' });
                const leftSpan = createEl('span');
                const bName = createEl('b', {}, data.title);
                leftSpan.append(bName);

                const rightSpan = createEl('span', { style: 'display:flex; align-items:center;' });
                const infoText = createEl('span', { style: 'font-size:12px' }, ` Строк: ${data.rows.length}`);
                const delBtn = createHistoryDeleteButton(entryDiv);
                rightSpan.append(infoText, delBtn);
                header.append(leftSpan, rightSpan);
                
                // Определяем единицу измерения для заголовка
                const restoreHeaderUnit = getHeaderUnitSuffix(data.rows);

                // Разметка из 5 подтаблиц (повторяет основной вид расчёта)
                const splitContainer = createEl('div', { className: 'tables-container', style: 'display:flex; gap:10px; flex-wrap:wrap; width:100%; align-items:flex-start;' });
                const tblOps = createSplitTable(['№', 'ПДТВ', 'Операция', 'Обед?', 'Пауза'], 2);
                const tblDur = createSplitTable([`Работа${restoreHeaderUnit}`], 1);
                const tblPostingDate = createSplitTable(['Дата проводки'], 1);
                const tblWorker = createSplitTable(['Исполнитель'], 1);
                const tblTime = createSplitTable(['Дата Начала', 'Время Начала', 'Дата Конца', 'Время Конца'], 3);

                populateSplitTables(data.rows, tblOps, tblDur, tblPostingDate, tblWorker, tblTime);
                splitContainer.append(tblOps.wrapper, tblDur.wrapper, tblPostingDate.wrapper, tblWorker.wrapper, tblTime.wrapper);

                const z7Table = createZ7TableElement(data.z7);
                
                entryDiv.append(header, splitContainer, createEl('div', { style: 'height:10px' }), z7Table);
                historyList.append(entryDiv);
            } catch (e) {
                safeLogError('Ошибка при восстановлении записи:', e);
            }
        });
    } catch (e) {
        safeLogError('Ошибка при загрузке истории:', e);
    }
    updateStartTimeFromHistory();
}

async function clearHistoryData() {
    if (!await confirmAction('Вы уверены? Это удалит всю историю расчетов.')) return;
    
    try {
        const historyList = document.getElementById('historyList');
        historyList.textContent = '';
        await safeLocalStorageRemove('z7_history_session');
        // Очищаем слот текущей сессии
        if (currentSessionId) {
            await safeLocalStorageRemove(SESSION_DATA_PREFIX + currentSessionId);
        }
        try { await showMessage('История удалена'); } catch(e){}
        
        document.getElementById('startTime').value = "08:00:00";
        
        updateStartTimeFromHistory();
        updateFirstPauseVisibility();
    } catch (e) {
        console.error('Ошибка при очистке истории:', e);
        showMessage('Ошибка при очистке истории').catch(() => {});
    }
}

// === МЕНЕДЖЕР СЕССИЙ ===

function loadSessionsMeta() {
    try {
        const raw = localStorage.getItem(SESSIONS_META_KEY);
        if (raw) {
            const parsed = safeJsonParse(raw);
            if (Array.isArray(parsed) && parsed.length > 0) {
                sessionsMeta = parsed;
                return true;
            }
        }
    } catch (e) { safeLogError('loadSessionsMeta error:', e); }
    return false;
}

async function saveSessionsMeta() {
    try {
        await safeLocalStorageSet(SESSIONS_META_KEY, JSON.stringify(sessionsMeta));
    } catch (e) { safeLogError('saveSessionsMeta error:', e); }
}

function generateSessionId() {
    return String(Date.now());
}

// Сохраняет текущий активный буфер (z7_history_session) в слот сессии
async function saveCurrentSessionData() {
    if (!currentSessionId) return;
    try {
        const raw = localStorage.getItem('z7_history_session');
        const key = SESSION_DATA_PREFIX + currentSessionId;
        if (raw) {
            await safeLocalStorageSet(key, raw);
        } else {
            await safeLocalStorageRemove(key);
        }
    } catch (e) { safeLogError('saveCurrentSessionData error:', e); }
}

// Загружает данные сессии в активный буфер (z7_history_session) и восстанавливает DOM
async function loadSessionData(id) {
    try {
        const key = SESSION_DATA_PREFIX + id;
        const raw = localStorage.getItem(key);
        if (raw) {
            await safeLocalStorageSet('z7_history_session', raw);
        } else {
            await safeLocalStorageRemove('z7_history_session');
        }
        restoreHistoryFromStorage();
    } catch (e) { safeLogError('loadSessionData error:', e); }
}

function renderSessionDropdown() {
    const sel = document.getElementById('sessionSelect');
    if (!sel) return;
    sel.textContent = '';
    sessionsMeta.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s.id;
        opt.textContent = s.name;
        if (s.id === currentSessionId) opt.selected = true;
        sel.append(opt);
    });
}

async function initSessionManager() {
    const hasExisting = loadSessionsMeta();
    if (!hasExisting) {
        // Первый запуск — создаём сессию по умолчанию из текущей истории
        const id = generateSessionId();
        sessionsMeta = [{ id, name: 'Сессия по умолчанию', created: Date.now() }];
        currentSessionId = id;
        await saveSessionsMeta();
        // Текущие данные z7_history_session уже на месте — сохраняем в слот
        await saveCurrentSessionData();
    } else {
        // Загружаем последнюю активную сессию
        const lastActiveId = localStorage.getItem('z7_active_session');
        if (lastActiveId && sessionsMeta.some(s => s.id === lastActiveId)) {
            currentSessionId = lastActiveId;
        } else {
            currentSessionId = sessionsMeta[0].id;
        }
        // Загружаем данные сессии в активный буфер
        await loadSessionData(currentSessionId);
    }
    renderSessionDropdown();
    setupSessionControls();
}

async function createNewSession() {
    let name = null;
    try {
        name = globalThis.prompt('Название новой сессии:', `Сессия ${sessionsMeta.length + 1}`);
    } catch (e) { return; }
    if (!name) return;
    name = sanitizeStrict(name, 100).trim();
    if (!name) return;

    // Сохраняем текущую сессию
    await saveCurrentSessionData();

    const id = generateSessionId();
    sessionsMeta.push({ id, name, created: Date.now() });
    await saveSessionsMeta();

    // Очищаем активный буфер и DOM
    await safeLocalStorageRemove('z7_history_session');
    document.getElementById('historyList').textContent = '';

    currentSessionId = id;
    localStorage.setItem('z7_active_session', currentSessionId);
    renderSessionDropdown();
    updateChainCheckboxState();
    document.getElementById('startTime').value = "08:00:00";
}

async function switchSession(targetId) {
    if (targetId === currentSessionId) return;
    if (!sessionsMeta.some(s => s.id === targetId)) return;

    // Сохраняем текущую
    await saveCurrentSessionData();

    currentSessionId = targetId;
    localStorage.setItem('z7_active_session', currentSessionId);

    // Загружаем целевую
    await loadSessionData(targetId);
    updateChainCheckboxState();
}

async function deleteSession(id) {
    if (sessionsMeta.length <= 1) {
        try { await showMessage('Нельзя удалить единственную сессию.'); } catch(e){}
        return;
    }
    const session = sessionsMeta.find(s => s.id === id);
    if (!session) return;

    if (!await confirmAction(`Удалить сессию "${session.name}" и все её данные?`)) return;

    // Удаляем данные
    await safeLocalStorageRemove(SESSION_DATA_PREFIX + id);
    sessionsMeta = sessionsMeta.filter(s => s.id !== id);
    await saveSessionsMeta();

    if (id === currentSessionId) {
        // Переключаемся на последнюю из оставшихся
        const latest = sessionsMeta[sessionsMeta.length - 1];
        currentSessionId = latest.id;
        localStorage.setItem('z7_active_session', currentSessionId);
        await loadSessionData(currentSessionId);
        updateChainCheckboxState();
    }
    renderSessionDropdown();
}

async function renameSession(id) {
    const session = sessionsMeta.find(s => s.id === id);
    if (!session) return;

    let newName = null;
    try {
        newName = globalThis.prompt('Новое название сессии:', session.name);
    } catch (e) { return; }
    if (!newName) return;
    newName = sanitizeStrict(newName, 100).trim();
    if (!newName || newName === session.name) return;

    session.name = newName;
    await saveSessionsMeta();
    renderSessionDropdown();
}

function setupSessionControls() {
    const sel = document.getElementById('sessionSelect');
    const newBtn = document.getElementById('newSessionBtn');
    const delBtn = document.getElementById('deleteSessionBtn');
    const renBtn = document.getElementById('renameSessionBtn');

    if (sel) {
        sel.addEventListener('change', () => switchSession(sel.value));
    }
    if (newBtn) {
        newBtn.addEventListener('click', () => createNewSession());
    }
    if (delBtn) {
        delBtn.addEventListener('click', () => deleteSession(currentSessionId));
    }
    if (renBtn) {
        renBtn.addEventListener('click', () => renameSession(currentSessionId));
    }
}

// Функция для управления видимостью чекбокса паузы первого блока
function updateFirstPauseVisibility() {
    const firstOpBlock = document.querySelector('.op-block');
    if (!firstOpBlock) return;
    const historyList = document.getElementById('historyList');
    const isFirstCalculation = historyList.children.length === 0;

    // Если это самый первый расчёт (история пуста), скрываем поле паузы для первой операции
    try {
        const breakGroup = firstOpBlock.querySelector('.break-container');
        const breakInput = firstOpBlock.querySelector('.op-break-val');
        const breakUnit = firstOpBlock.querySelector('.op-break-unit');
        if (isFirstCalculation) {
            if (breakGroup) breakGroup.style.display = 'none';
            if (breakInput) {
                breakInput.value = '0';
                breakInput.dispatchEvent(new Event('input'));
            }
            if (breakUnit) breakUnit.value = 'min';
        } else {
            if (breakGroup) breakGroup.style.display = 'flex';
        }
    } catch (e) {
        console.debug?.('reset pause visibility error:', e?.message);
    }
}

function renderFields() {
    const targetCount = validateNumber(document.getElementById('totalOps').value, 1, 20);
    document.getElementById('totalOps').value = targetCount;

    // Валидация отрицательных значений для #workerCount
    let workerCount = Number.parseInt(document.getElementById('workerCount').value, 10);
    if (workerCount < 1) {
        document.getElementById('workerCount').value = 1;
    }

    const currentBlocks = Array.from(container.children);
    const currentCount = currentBlocks.length;

    if (targetCount > currentCount) {
        let maxIndex = 0;
        currentBlocks.forEach(b => {
            const idx = Number.parseInt(b.dataset.originalIndex, 10);
            if (!Number.isNaN(idx) && idx > maxIndex) maxIndex = idx;
        });
        for (let i = 0; i < (targetCount - currentCount); i++) {
            createOperationBlock(maxIndex + 1 + i);
        }
    } else if (targetCount < currentCount) {
        for (let i = currentCount - 1; i >= targetCount; i--) {
            currentBlocks[i].remove();
        }
    }
    // Если модальное окно операций открыто, перерисовываем его поля ввода и пересчитываем номера подтверждения
    try {
        const oModal = document.getElementById('opsModal');
        if (oModal?.classList.contains('active')) {
            renderOpsInputList();
            updateOpsCalculatedValues();
        }
    } catch (e) {
        console.debug?.('renderFields modal update error:', e?.message);
    }
    try {
        updateMainOperationLabels();
        updateOperationInputPrefixes();
        updateWorkerUIByTimeMode();
    } catch (error_) {
        console.debug?.('renderFields post-update error:', error_?.message);
    }
}

function createOperationBlock(index) {
    const block = createEl('div', { className: 'op-block' });
    block.dataset.originalIndex = index;
    // Метка номера операции (показывает номер подтверждения, если задан, иначе порядковый индекс)
    const totalOpsCurrent = Number.parseInt(document.getElementById('totalOps')?.value || '0', 10) || 0;
    const opNumText = (typeof getOperationLabel === 'function') ? getOperationLabel(index, totalOpsCurrent) : String(index);
    const numLabel = createEl('div', { className: 'op-num-label' }, opNumText);

    const prefix = `${index}) `;
    const nameInp = createEl('input', {
        className: 'op-header-input',
        name: `op_name_${index}`,
        value: `${prefix}Операция №${index}`,
        type: 'text',
        placeholder: 'Название операции',
        maxlength: '200',
        autocomplete: 'off'
    });
    // Делаем числовой префикс неизменяемым: оставляем в начале, санитизируем только суффикс
    try {
        const handleInput = (e) => {
            const el = e.target;
            let v = el.value || '';
            // удаляем ведущий числовой префикс, который пользователь может вставить/набрать
            v = v.replace(/^\s*\d+\)\s*/, '');
            // санитизируем только значимую часть
            v = sanitizeStrict(v, 200);
            el.value = prefix + v;
            // удерживаем курсор после префикса
            const pos = Math.max(prefix.length, (el.selectionStart || 0));
            try { el.setSelectionRange(pos, pos); } catch (ee) {}
        };

        nameInp.addEventListener('input', handleInput);
        nameInp.addEventListener('focus', (e) => {
            const el = e.target;
            if ((el.selectionStart || 0) < prefix.length) {
                try { el.setSelectionRange(prefix.length, prefix.length); } catch (ee) {}
            }
        });

        nameInp.addEventListener('keydown', (e) => {
            const el = e.target;
            const selStart = el.selectionStart || 0;
            const selEnd = el.selectionEnd || 0;
            // предотвращаем удаление или выделение префикса
            if ((e.key === 'Backspace' || e.key === 'Delete') && selEnd <= prefix.length) {
                e.preventDefault();
            }
            // предотвращаем выделение, включающее префикс, и замену его при наборе
            if (e.key.length === 1 && selStart < prefix.length && selEnd <= prefix.length) {
                // ставим курсор после префикса перед вставкой
                try { el.setSelectionRange(prefix.length, prefix.length); } catch (ee) {}
            }
        });

        nameInp.addEventListener('paste', (e) => {
            e.preventDefault();
            const paste = e.clipboardData.getData('text') || '';
            const sanitized = sanitizeStrict(paste, 200);
            const el = e.target;
            const cur = el.value || '';
            const insertPos = Math.max(prefix.length, el.selectionStart || prefix.length);
            const before = cur.slice(prefix.length, insertPos);
            const after = cur.slice(insertPos);
            const newBody = (before + sanitized + after).slice(0, 200);
            el.value = prefix + sanitizeStrict(newBody, 200);
            const pos = prefix.length + Math.min(newBody.length, 200);
            try { el.setSelectionRange(pos, pos); } catch (ee) {}
        });
    } catch (e) {
        console.debug?.('op name input attach failed:', e?.message);
    }
    
    
    const controls = createEl('div', { className: 'op-controls' });

    // Исполнители: чекбоксы под названием операции
    const workersWrapper = createEl('div', { className: 'op-workers-wrapper' });
    workersWrapper.append(createEl('label', { className: 'op-workers-label' }, 'Исполнители:'));
    const workersBox = createEl('div', { className: 'op-workers-box' });
    // заполняем в соответствии с текущим workerCount
    try {
        const curCount = Number.parseInt(document.getElementById('workerCount')?.value || '1', 10) || 1;
        for (let w = 1; w <= curCount; w++) {
            const id = `op_${index}_worker_${w}`;
            const cb = createEl('input', { type: 'checkbox', className: 'op-worker-checkbox', id, 'data-worker': String(w) });
            cb.checked = true;
            const lbl = createEl('label', { htmlFor: id, className: 'op-worker-label' }, String(w));
            const wrapper = createEl('span', { className: 'op-worker-item' });
            wrapper.append(cb, lbl);
            workersBox.append(wrapper);

            // распространяем логику цепочки
            cb.addEventListener('change', () => {
                updateWorkerChain();
            });
        }

    } catch (err) {
        console.debug?.('init op workers failed:', err?.message);
    }
    workersWrapper.append(workersBox);

    // Элемент, отображаемый когда задействованы все исполнители (заменяет чекбоксы в не-индивидуальном режиме)
    const workersAll = createEl('div', { className: 'op-workers-all', style: 'display:none;' }, 'ВСЕ');
    workersWrapper.append(workersAll);
    
    // Блок времени работы
    const workGroup = createEl('div', { className: 'time-group' });
    workGroup.append(createEl('label', { htmlFor: `op_duration_${index}` }, 'Время:'));
    const workInput = createEl('input', {
        type: 'text',
        className: 'op-duration',
        id: `op_duration_${index}`,
        name: `op_duration_${index}`,
        inputmode: 'decimal',
        pattern: String.raw`\d{0,5}([.,]\d{1,2})?`,
        maxlength: '8',
        size: '6',
        style: 'width:8ch',
        value: '10',
        autocomplete: 'off'
    });
    // Санитизация и ограничение ввода: до 5 цифр целой части и 2 дробных
    workInput.addEventListener('input', decimalInputHandler);
    workInput.addEventListener('blur', decimalBlurHandler);
    workGroup.append(workInput);
    // плейсхолдер, отображаемый в режиме individual
    const workAll = createEl('div', { className: 'op-dur-all', style: 'display:none;' }, 'В Excel');
    workGroup.append(workAll);
    const workUnit = createEl('select', {
        className: 'op-unit',
        name: `op_unit_${index}`,
        style: 'width:70px; background:transparent; border:none;'
    });
    workUnit.append(
        new Option('мин', 'min'),
        new Option('час', 'hour')
    );
    
    // Для всех операций кроме первой - disabled и синхронизация с первой
    if (index !== 1) {
        workUnit.disabled = true;
        // Синхронизируем с первой операцией
        const firstUnitSelect = container.querySelector('.op-block:first-child .op-unit');
        if (firstUnitSelect) {
            workUnit.value = firstUnitSelect.value;
        }
    } else {
        // Для первой операции - обработчик синхронизации
        workUnit.addEventListener('change', syncTimeUnits);
    }
    workGroup.append(workUnit);
    
    // Блок паузы между заказами (видим во всех карточках, кроме первой операции первой записи)
    const breakGroup = createEl('div', { className: 'time-group break-container' });
    // Видимая метка для поля паузы — тот же стиль, что и у метки «Время»
    breakGroup.append(createEl('label', { htmlFor: `op_break_${index}` }, 'Пауза:'));
    const breakInput = createEl('input', {
        type: 'text',
        className: 'op-break-val',
        id: `op_break_${index}`,
        name: `op_break_${index}`,
        inputmode: 'decimal',
        pattern: String.raw`\d{0,5}([.,]\d{1,2})?`,
        maxlength: '8',
        size: '6',
        style: 'width:8ch',
        value: '0',
        autocomplete: 'off'
    });
    // Санитизация и ограничение ввода: до 5 цифр целой части и 2 дробных
    breakInput.addEventListener('input', decimalInputHandler);
    breakInput.addEventListener('blur', decimalBlurHandler);
    breakGroup.append(breakInput);
    const breakUnit = createEl('select', {
        className: 'op-break-unit',
        name: `op_break_unit_${index}`,
        style: 'width:70px; background:transparent; border:none;'
    });
    breakUnit.append(
        new Option('мин', 'min'),
        new Option('час', 'hour')
    );
    breakGroup.append(breakUnit);
    const breakAll = createEl('div', { className: 'op-break-all', style: 'display:none;' }, 'В Excel');
    breakGroup.append(breakAll);
    
    // Добавляем блок паузы и элементы управления. Начальная видимость breakGroup
    // По умолчанию: показываем паузу для всех операций кроме первой (первая управляется updateFirstPauseVisibility)
    if (index !== 1) {
        breakGroup.style.display = 'flex';
    } else {
        breakGroup.style.display = 'none';
    }
    controls.append(breakGroup, workGroup);
    // поле ввода названия + UI исполнителей
    const nameCol = createEl('div', { className: 'op-name-col' });
    nameCol.append(nameInp, workersWrapper);

    // Кнопка удаления операции (X) — доступна до нажатия «Задать»
    const delOpBtn = createEl('button', {
        type: 'button',
        className: 'btn-del-op',
        title: 'Удалить операцию из расчета'
    });
    delOpBtn.textContent = '✕';
    // Всегда показываем кнопку удаления, но при нажатии она будет переключать состояние мягкого удаления (показано/скрыто)
    delOpBtn.addEventListener('click', () => {
        if (block.classList.contains('deleted-op')) {
            // Восстановить
            block.classList.remove('deleted-op');
            delOpBtn.title = 'Удалить операцию из расчета';
        } else {
            // Мягкое удаление
            block.classList.add('deleted-op');
            delOpBtn.title = 'Восстановить операцию';
        }
    });
    block.append(delOpBtn, numLabel, nameCol, controls);
    container.append(block);
    
    // Обновить видимость паузы первого блока после создания нового блока
    updateFirstPauseVisibility();
    // Обеспечиваем соответствие UI исполнителей текущему выбору timeMode
    try { updateWorkerUIByTimeMode(); } catch (e) {}
}

// Отключение опции «individual» в timeMode при активном режиме цепочки
function updateTimeModeByChain() {
    const chainCheckbox = document.getElementById('chainMode');
    const timeModeEl = document.getElementById('timeMode');
    if (!chainCheckbox || !timeModeEl) return;
    const isChain = chainCheckbox.checked;
    const individualOpt = timeModeEl.querySelector('option[value="individual"]');
    if (individualOpt) {
        individualOpt.disabled = isChain;
    }
    // Если цепочка только что включена и был выбран «individual», переключаем на «total»
    if (isChain && timeModeEl.value === 'individual') {
        timeModeEl.value = 'total';
        updateWorkerUIByTimeMode();
    }
}

// Переключение видимости чекбоксов исполнителей для каждой операции в зависимости от #timeMode
function updateWorkerUIByTimeMode() {
    const modeEl = document.getElementById('timeMode');
    if (!modeEl) return;
    const mode = modeEl.value;
    const blocks = Array.from(document.querySelectorAll('.op-block'));
    blocks.forEach(block => {
        const box = block.querySelector('.op-workers-box');
        const allEl = block.querySelector('.op-workers-all');
        const workInput = block.querySelector('.op-duration');
        const workAll = block.querySelector('.op-dur-all');
        const breakInput = block.querySelector('.op-break-val');
        const breakUnit = block.querySelector('.op-break-unit');
        const breakAll = block.querySelector('.op-break-all');
        if (!box || !allEl) return;
        if (mode === 'individual') {
            // Сохраняем значение, если ещё не сохранено, затем обнуляем
            if (workInput && workInput.dataset.savedVal === undefined) {
                workInput.dataset.savedVal = workInput.value;
                workInput.value = 0;
            }
            if (breakInput && breakInput.dataset.savedVal === undefined) {
                breakInput.dataset.savedVal = breakInput.value;
                breakInput.value = 0;
            }

            // Индивидуальный: показываем чекбоксы исполнителей для каждой операции, скрываем числовые поля и показываем плейсхолдеры «В Excel»
            box.style.display = 'grid';
            allEl.style.display = 'none';
            if (workInput) {
                workInput.style.display = 'none';
            }
            if (workAll) workAll.style.display = '';
            if (breakInput) {
                breakInput.style.display = 'none';
            }
            if (breakUnit) breakUnit.style.display = 'none';
            if (breakAll) breakAll.style.display = '';
        } else {
            // Восстанавливаем значения, если были сохранены
            if (workInput && workInput.dataset.savedVal !== undefined) {
                workInput.value = workInput.dataset.savedVal;
                delete workInput.dataset.savedVal;
            }
            if (breakInput && breakInput.dataset.savedVal !== undefined) {
                breakInput.value = breakInput.dataset.savedVal;
                delete breakInput.dataset.savedVal;
            }

            // total / per_worker: скрываем чекбоксы исполнителей для каждой операции и показываем общие поля ввода
            const cbs = Array.from(box.querySelectorAll('.op-worker-checkbox'));
            cbs.forEach(cb => { cb.checked = true; });
            box.style.display = 'none';
            allEl.style.display = 'inline-flex';
            if (workInput) workInput.style.display = '';
            if (workAll) workAll.style.display = 'none';
            if (breakInput) { breakInput.style.display = ''; }
            if (breakUnit) { breakUnit.style.display = ''; }
            if (breakAll) breakAll.style.display = 'none';
        }
    });
    // Повторная проверка логики цепочки после переключения режима
    if (mode === 'individual') {
        updateWorkerChain();
    }
}

// Правило: если исполнитель снят в операции N, он отключён и снят во всех операциях > N
function updateWorkerChain() {
    try {
        const workerCount = Number.parseInt(document.getElementById('workerCount')?.value || '1', 10) || 1;
        const blocks = Array.from(document.querySelectorAll('.op-block'));
        
        for (let w = 1; w <= workerCount; w++) {
            let chainActive = true;
            for (let i = 0; i < blocks.length; i++) {
                const block = blocks[i];
                const cb = block.querySelector(`.op-worker-checkbox[data-worker="${w}"]`);
                if (!cb) continue;
                if (!chainActive) {
                    // Предыдущая операция была снята -> отключаем и снимаем эту
                    cb.checked = false;
                    cb.disabled = true;
                    cb.parentElement.style.opacity = '0.5';
                } else {
                    // Предыдущие операции в порядке.
                    // Включаем эту
                    cb.disabled = false;
                    cb.parentElement.style.opacity = '1';
                    // Если снят, это разрывает цепочку для ПОСЛЕДУЮЩИХ
                    if (!cb.checked) {
                        chainActive = false;
                    }
                }
            }
        }
    } catch (e) {
        console.debug?.('updateWorkerChain error:', e?.message);
    }
}


let _generateInProgress = false;

async function generateTable() {
    if (_generateInProgress) return;
    _generateInProgress = true;
    const generateBtn = document.getElementById('generateBtn');
    if (generateBtn) generateBtn.disabled = true;
    try {
    const tableResult = document.getElementById('tableResult');
    const z7Result = document.getElementById('z7Result');
    tableResult.textContent = '';
    z7Result.textContent = '';

    const startD = document.getElementById('startDate').value;
    const startT = document.getElementById('startTime').value;
    const postingD = (document.getElementById('postingDate') && document.getElementById('postingDate').value) ? document.getElementById('postingDate').value : startD;
    const workerCount = validateNumber(document.getElementById('workerCount').value, 1, 10);
    const timeMode = document.getElementById('timeMode').value;
    const lunchStartInput = document.getElementById('lunchStart').value;
    const lunchStartInput2 = document.getElementById('lunchStart2').value;
    // Длительность обеда теперь может быть дробной; парсим как float и ограничиваем от 0 до 480
    let lunchDurMin = Number.parseFloat(String(document.getElementById('lunchDur').value).replaceAll(',', '.')) || 0;
    if (!Number.isFinite(lunchDurMin)) lunchDurMin = 0;
    lunchDurMin = Math.max(0, Math.min(480, lunchDurMin));
    const isChain = document.getElementById('chainMode').checked;
    // Валидация значений select по допустимым перечислениям
    if (timeMode !== 'per_worker' && timeMode !== 'total' && timeMode !== 'individual') {
        console.warn('Unexpected timeMode value, defaulting to "total"');
    }
    
    if (!startD || !startT) {
        showMessage("Пожалуйста, укажите дату и время начала.").catch(() => {});
        return;
    }

    // Проверяем, это первый расчет или нет
    const historyList = document.getElementById('historyList');
    const isFirstCalculation = historyList.children.length === 0;

    let [y, m, d] = startD.split('-').map(Number);
    let [th, tm, ts] = startT.split(':').map(Number);
    ts = ts || 0;
    let globalTime = new Date(y, m - 1, d, th, tm, ts);
 
    // --- Настройка обедов (JS) ---
    // Валидация формата времени обеда (HH:MM или HH:MM:SS)
    const timeRe = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;
    let lh = 0, lm = 0;
    try {
        const m = String(lunchStartInput || '').match(timeRe);
        if (m) { lh = Number(m[1]); lm = Number(m[2]); } else { throw new Error('invalid lunchStart'); }
    } catch (e) {
        lh = 12; lm = 0;
    }
    let lunchStartTime = new Date(y, m - 1, d, lh, lm, 0);
    let lunchEndTime = new Date(lunchStartTime.getTime() + lunchDurMin * 60000);

    // Второй обед
    let lh2 = 0, lm2 = 0;
    try {
        const m2 = String(lunchStartInput2 || '').match(timeRe);
        if (m2) { lh2 = Number(m2[1]); lm2 = Number(m2[2]); } else { throw new Error('invalid lunchStart2'); }
    } catch (e) {
        lh2 = 0; lm2 = 0; // fallback to midnight
    }
    let lunch2StartTime = new Date(y, m - 1, d, lh2, lm2, 0);
    // Если второй обед раньше старта (напр 00:00 vs 08:00), считаем что он на след. день
    if (lunch2StartTime < globalTime) {
        lunch2StartTime.setDate(lunch2StartTime.getDate() + 1);
    }
    let lunch2EndTime = new Date(lunch2StartTime.getTime() + lunchDurMin * 60000);

    const opsNodeList = document.querySelectorAll('.op-block');
    if (opsNodeList.length === 0) return;
    // Исключаем мягко удалённые операции из расчётов
    const ops = Array.from(opsNodeList).filter(b => !b.classList.contains('deleted-op'));
    if (ops.length === 0) { showMessage('Нет активных операций для расчёта (все удалены)', 'Ошибка', 'error'); return; }

    // Сортировка для расчета согласно выбранному режиму
    const sortMode = document.getElementById('opsSortMode')?.value || 'sequential';
    if (sortMode === 'confirmation') {
        ops.sort((a, b) => {
            const idA = Number(a.dataset.opId) || 0;
            const idB = Number(b.dataset.opId) || 0;
            return idA - idB;
        });
    } else {
        ops.sort((a, b) => {
            const idxA = Number(a.dataset.originalIndex) || 0;
            const idxB = Number(b.dataset.originalIndex) || 0;
            return idxA - idxB;
        });
    }

    const operationNames = [];
    const dataMain = [];
    const fmtTime = (date) => date.toLocaleTimeString('ru', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
    const fmtDate = (date) => date.toLocaleDateString('ru');

    ops.forEach((block, opIndex) => {
        // Удаляем возможный префикс вида "N) " перед санитаризацией
        const rawOpName = block.querySelector('.op-header-input').value || '';
        const name = sanitizeStrict(stripOrdinalPrefix(rawOpName), 200);
        operationNames.push(name);
        const originalOpIndex = block.dataset.originalIndex || (opIndex + 1);
        const dur = Math.max(0, Number.parseFloat(block.querySelector('.op-duration').value) || 0);
        let unit = block.querySelector('.op-unit').value;
        if (unit !== 'min' && unit !== 'hour') unit = 'min';

        // Исходная длительность в мс (из полей ввода карточки)
        let origDurationMs = 0;
        if (unit === 'hour') origDurationMs = dur * 3600 * 1000;
        else origDurationMs = dur * 60 * 1000;

        // Определяем, какая длительность используется для расчётов, а какая для отображения/экспорта
        let durationMsForCalc = origDurationMs;
        let displayDurVal = dur; // значение для отображения и экспорта (в мин или часах в зависимости от выбранной единицы)
        if (timeMode === 'total' && workerCount > 1) {
            durationMsForCalc = origDurationMs / workerCount;
            displayDurVal = displayDurVal / workerCount;
        }
        // В режиме «individual» сохраняем значения в UI/экспорте, но для расчёта временной шкалы длительности равны нулю
        if (timeMode === 'individual') {
            durationMsForCalc = 0;
            // displayDurVal остаётся dur, чтобы карточки сохраняли значения и ячейки Excel могли быть предзаполнены исходными длительностями
        }

        // Применяем пер-операционную паузу ПЕРЕД началом этой операции (даже если 0)
        const opBreakVal = Math.max(0, Number.parseFloat(block.querySelector('.op-break-val').value) || 0);
        let opBreakUnit = block.querySelector('.op-break-unit')?.value || 'min';
        if (opBreakUnit !== 'min' && opBreakUnit !== 'hour') opBreakUnit = 'min';
        const opBreakSec = (opBreakUnit === 'hour') ? (opBreakVal * 3600) : (opBreakVal * 60);
        const origOpBreakMs = Math.floor(opBreakSec * 1000);
        const opBreakMsForCalc = (timeMode === 'individual') ? 0 : origOpBreakMs;
        globalTime = new Date(globalTime.getTime() + opBreakMsForCalc);

        let opStart = new Date(globalTime);
        let opEnd = new Date(opStart.getTime() + durationMsForCalc);
        let crossedLunch = false;

        // Логика проверки двух обедов
        // Если второй обед установлен в 00:00, он не должен учитываться
        const hasLunch2 = !(lh2 === 0 && lm2 === 0);
        let lunches = [
            { s: lunchStartTime, e: lunchEndTime },
            ...(hasLunch2 ? [{ s: lunch2StartTime, e: lunch2EndTime }] : [])
        ].sort((a, b) => a.s - b.s);

        for (let l of lunches) {
            // 1. Если начало операции попадает внутрь обеда -> сдвигаем старт
            if (opStart >= l.s && opStart < l.e) {
                opStart = new Date(l.e);
                opEnd = new Date(opStart.getTime() + durationMsForCalc);
                crossedLunch = true;
            }

            // 2. Если операция накрывает начало обеда (началась до, заканчивается после)
            if (opStart < l.s && opEnd > l.s) {
                let lDur = l.e.getTime() - l.s.getTime();
                opEnd = new Date(opEnd.getTime() + lDur);
                crossedLunch = true;
            }
        }

        let displayDurText = new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(displayDurVal);

        // Метаданные ПДТВ для Excel-формул (вычисляются один раз на операцию, вне цикла по исполнителям)
        const _origIdxNum = Number(originalOpIndex);
        const _pdtvIsLast = lastOperationIndex !== null && _origIdxNum === lastOperationIndex;
        const _pdtvIsPenultimate = penultimateOperationIndex !== null && _origIdxNum === penultimateOperationIndex;
        let _pdtvOffset = 0;
        if (autoIncrementEnabled && operationFirstId && operationFirstId.trim() !== '') {
            const _totalOps = ops.length;
            if (_pdtvIsLast) {
                _pdtvOffset = _totalOps - 1;
            } else if (_pdtvIsPenultimate) {
                _pdtvOffset = _totalOps - 2;
            } else {
                let _pos = _origIdxNum;
                if (lastOperationIndex !== null && _origIdxNum > lastOperationIndex) _pos -= 1;
                if (penultimateOperationIndex !== null && _origIdxNum > penultimateOperationIndex) _pos -= 1;
                _pdtvOffset = _pos - 1;
            }
        }

        for (let w = 1; w <= workerCount; w++) {
            // Если чекбокс для этого исполнителя существует в данной операции и не отмечен, пропускаем создание строк
            try {
                const cb = block.querySelector(`.op-worker-checkbox[data-worker="${w}"]`);
                if (cb && !cb.checked) continue;
            } catch (err) { /* ignore selector errors */ }
            // Пауза операции (показываем текстовое значение, если >0)
            const rowPauseText = opBreakVal > 0 ? formatDurationToTime(opBreakVal, opBreakUnit) : "";
            const rowPauseExcel = opBreakSec / 86400.0;

            dataMain.push({
                opIdx: block.dataset.opId || getOperationLabel(opIndex + 1, ops.length), // Номер подтверждения или порядковый номер
                opNumeric: opIndex + 1, // Числовой индекс для Excel формул
                originalOpIndex: originalOpIndex,
                name: name,
                worker: getWorkerLabel(w),
                workerIndex: w, // сохраняем числовой индекс для Excel формул
                durVal: displayDurVal,
                durText: displayDurText,
                startObj: new Date(opStart),
                endObj: new Date(opEnd),
                startDate: fmtDate(opStart),
                startTime: fmtTime(opStart),
                endDate: fmtDate(opEnd),
                endTime: fmtTime(opEnd),
                crossedLunch: crossedLunch,
                pauseText: rowPauseText,
                pauseExcelVal: rowPauseExcel,
                postingDateIso: postingD,
                postingDate: fmtDate(new Date(postingD + 'T00:00:00')),
                unit: unit, // сохраняем единицу измерения
                pdtvAutoMode: autoIncrementEnabled,
                pdtvOffset: _pdtvOffset
            });
        }
        globalTime = opEnd;
    });

    if (isChain) {
        startDateInput.value = formatDateISO(globalTime);
        startTimeInput.value = formatTimeHMS(globalTime);
    }

    const tblOps = createSplitTable(['№', 'ПДТВ', 'Операция', 'Обед?', 'Пауза'], 2);
    
    // Определяем единицу измерения для заголовка Работа
    const headerUnit = getHeaderUnitSuffix(dataMain);
    
    const tblDur = createSplitTable([`Работа${headerUnit}`], 1);
    const tblPostingDate = createSplitTable(['Дата проводки'], 1);
    const tblWorker = createSplitTable(['Исполнитель'], 1);
    const tblTime = createSplitTable(['Дата Начала', 'Время Начала', 'Дата Конца', 'Время Конца'], 3);

    populateSplitTables(dataMain, tblOps, tblDur, tblPostingDate, tblWorker, tblTime);

    tableResult.append(tblOps.wrapper, tblDur.wrapper, tblPostingDate.wrapper, tblWorker.wrapper, tblTime.wrapper);

    const statusText = sanitizeStrict(document.getElementById('statusBefore').value, 300) || "замечаний нет";
    const extraWorks = sanitizeStrict(document.getElementById('workExtra').value, 300) || "нет";
    const devRec = sanitizeStrict(document.getElementById('devRec').value, 300) || "нет";
    const rizVal = sanitizeInput(document.getElementById('resIz').value, 6) || "";
    const kVal = sanitizeInput(document.getElementById('coefK').value, 5) || "";
        const kValForZ7 = kVal.replaceAll(',', '.');
    const worksText = operationNames.join(', ');
    const rizDisplay = rizVal ? `${rizVal} МОм` : "";

    const z7Lines = [
        `1. состояние объекта ремонта до начала работ: ${statusText}`,
        `2. выполненные работы в рамках планового объёма работ: ${worksText}`,
        `3. выполненные работы в рамках дополнительного объёма работ: ${extraWorks}`,
        `4. результаты испытаний, тестов, замеров, инспекций: Rиз= ${rizDisplay} K= ${kValForZ7}`,
        `5. отклонения от ТК и рекомендации по корректировке ТК: ${devRec}`
    ];
    
    const z7Div = createEl('div', { className: 'z7-report-wrapper' });
    const z7Table = createEl('table', { className: 'z7-table' });
    const z7Head = createEl('thead');
    const thZ7 = createEl('th', { className: 'z7-header-common', colspan: '12' }, 'Z7');
    const z7HeadTr = createEl('tr');
    z7HeadTr.append(thZ7);
    z7Head.append(z7HeadTr);

    const z7Body = createEl('tbody');
    const tr = createEl('tr', { className: 'z7-row' });
    const z7Td = createEl('td');
    z7Lines.forEach(line => z7Td.append(createEl('div', { className: 'z7-line-item' }, line)));
    tr.append(z7Td);
    z7Body.append(tr);
    z7Table.append(z7Head, z7Body);
    z7Div.append(z7Table);
    z7Result.append(z7Div);

    const select = document.getElementById('techCardSelect');
    const cardNameBase = select.value === 'manual' ? 'Ручной ввод' : select.options[select.selectedIndex].text;
    const orderInput = sanitizeInput(document.getElementById('orderName')?.value || '', 12);
    const nameInput = sanitizeStrict(document.getElementById('itemName')?.value || '', 70);
    const cardName = (orderInput ? (orderInput + ' | ') : '') + (nameInput ? nameInput : cardNameBase);
    
    const lunchConfig = { h: lh, m: lm, h2: lh2, m2: lm2, dur: lunchDurMin };
    const workersConfig = { count: Number.parseInt(document.getElementById('workerCount').value, 10) || 1, ids: workerIds.slice() };

    // orderPauseConfig сохраняем для совместимости, но данные уже в строках
    const orderPauseConfig = { dur: 0, unit: 'min', isApplied: !isFirstCalculation }; 
    await addToHistoryTable(dataMain, cardName, z7Lines, lunchConfig, isChain, orderPauseConfig, timeMode, workersConfig);
    } finally {
        _generateInProgress = false;
        if (generateBtn) generateBtn.disabled = false;
    }
}

async function addToHistoryTable(data, cardName, z7LinesArray, lunchConfig, isChain, orderPauseConfig, timeMode, workersConfig) {
    try {
        const historyList = document.getElementById('historyList');
        const now = new Date();
        const tsDate = now.toLocaleDateString('ru');
        const tsTime = now.toLocaleTimeString('ru');

        const entryDiv = createEl('div', { className: 'history-entry' });
        entryDiv.dataset.jsonData = JSON.stringify({
            title: `${cardName} | Сформировано: ${tsDate}; ${tsTime}`,
            rows: data,
            z7: z7LinesArray,
            lunch: lunchConfig,
            chain: isChain,
            orderPause: orderPauseConfig,
            timeMode: timeMode || 'total',
            workers: workersConfig
        });

        const header = createEl('div', { className: 'history-header' });
        const leftSpan = createEl('span');
        const bName = createEl('b', {}, cardName);
        leftSpan.append(bName, document.createTextNode(` | Сформировано: ${tsDate}; ${tsTime}`));

        const rightSpan = createEl('span', { style: 'display:flex; align-items:center;' });
        const infoText = createEl('span', { style: 'font-size:12px' }, ` Строк: ${data.length}`);
        const delBtn = createHistoryDeleteButton(entryDiv);
        rightSpan.append(infoText, delBtn);
        header.append(leftSpan, rightSpan);
        
        // Определяем единицу измерения для заголовка
        const histHeaderUnit = getHeaderUnitSuffix(data);

        // Разметка из 5 подтаблиц (повторяет основной вид расчёта)
        const splitContainer = createEl('div', { className: 'tables-container', style: 'display:flex; gap:10px; flex-wrap:wrap; width:100%; align-items:flex-start;' });
        const tblOps = createSplitTable(['№', 'ПДТВ', 'Операция', 'Обед?', 'Пауза'], 2);
        const tblDur = createSplitTable([`Работа${histHeaderUnit}`], 1);
        const tblPostingDate = createSplitTable(['Дата проводки'], 1);
        const tblWorker = createSplitTable(['Исполнитель'], 1);
        const tblTime = createSplitTable(['Дата Начала', 'Время Начала', 'Дата Конца', 'Время Конца'], 3);

        populateSplitTables(data, tblOps, tblDur, tblPostingDate, tblWorker, tblTime);
        splitContainer.append(tblOps.wrapper, tblDur.wrapper, tblPostingDate.wrapper, tblWorker.wrapper, tblTime.wrapper);

        const z7Table = createZ7TableElement(z7LinesArray);
        
        entryDiv.append(header, splitContainer, createEl('div', { style: 'height:10px' }), z7Table);
        historyList.prepend(entryDiv);
        
        // Сохраняем историю в localStorage
        await saveHistoryToStorage();
        updateStartTimeFromHistory();
        updateFirstPauseVisibility();
    } catch (e) {
        console.error(e);
        showMessage("Ошибка добавления в историю: " + e.message).catch(() => {});
    }
}

function updateStartTimeFromHistory() {
    const isChainMode = document.getElementById('chainMode').checked;
    const historyList = document.getElementById('historyList');
    const startTimeInput = document.getElementById('startTime');
    const startDateInput = document.getElementById('startDate');
    
    if (!isChainMode || historyList.children.length === 0) {
        // Если режим цепочки отключен или история пуста, поле активно
        startTimeInput.disabled = false; startTimeInput.title = '';
        startDateInput.disabled = false; startDateInput.title = '';
        return;
    }
    
    // Получаем последнюю запись из истории
    const lastEntry = historyList.firstElementChild;
    if (!lastEntry || !lastEntry.dataset.jsonData) {
        startTimeInput.disabled = false; startTimeInput.title = '';
        startDateInput.disabled = false; startDateInput.title = '';
        return;
    }

    try {
        const data = safeJsonParse(lastEntry.dataset.jsonData);
        if (!data || !data.rows || data.rows.length === 0) {
            startTimeInput.disabled = false; startTimeInput.title = '';
            startDateInput.disabled = false; startDateInput.title = '';
            return;
        }

        // Получаем время окончания последней операции
        const lastRow = data.rows[data.rows.length - 1];

        if (lastRow.endObj) {
            const dt = new Date(lastRow.endObj);
            startDateInput.value = formatDateISO(dt);
            startTimeInput.value = formatTimeHMS(dt);
        } else {
            startTimeInput.value = lastRow.endTime;
        }

        const timeTip = 'Автоматически из последней записи. Для разблокировки очистите историю или создайте новую.';
        startTimeInput.disabled = true; startTimeInput.title = timeTip;
        startDateInput.disabled = true; startDateInput.title = timeTip;
    } catch (e) {
        console.error('Ошибка при обновлении времени начала:', e);
        startTimeInput.disabled = false; startTimeInput.title = '';
        startDateInput.disabled = false; startDateInput.title = '';
    }
}

// Поддержка Отмены / Повтора (Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z)

// Стеки отмены/повтора в памяти. Храним ограниченную историю.
const _undoStack = [];
const _redoStack = [];
const _UNDO_LIMIT = 100;
let _snapshotTimer = null;
const _SNAPSHOT_DEBOUNCE = 500;

function captureAppState() {
    const getVal = (id) => {
        const el = document.getElementById(id);
        return el ? el.value : null;
    };

    const state = {
        totalOps: Number(getVal('totalOps') || 0),
        workerCount: Number(getVal('workerCount') || 1),
        timeMode: getVal('timeMode') || 'total',
        chainMode: !!document.getElementById('chainMode')?.checked,
        opsSortMode: getVal('opsSortMode') || 'sequential',
        techCardValue: getVal('techCardSelect') || 'manual',
        startDate: getVal('startDate') || '',
        startTime: getVal('startTime') || '',
        postingDate: getVal('postingDate') || '',
        lunchStart: getVal('lunchStart') || '',
        lunchStart2: getVal('lunchStart2') || '',
        lunchDur: getVal('lunchDur') || '',
        orderName: getVal('orderName') || '',
        itemName: getVal('itemName') || '',
        statusBefore: getVal('statusBefore') || '',
        workExtra: getVal('workExtra') || '',
        devRec: getVal('devRec') || '',
        resIz: getVal('resIz') || '',
        coefK: getVal('coefK') || '',
        ops: []
    };

    const blocks = document.querySelectorAll('.op-block');
    blocks.forEach(block => {
        const name = block.querySelector('.op-header-input')?.value || '';
        const dur = block.querySelector('.op-duration')?.value || '';
        const unit = block.querySelector('.op-unit')?.value || 'min';
        const breakVal = block.querySelector('.op-break-val')?.value || '';
        const breakUnit = block.querySelector('.op-break-unit')?.value || 'min';
        const workerCbs = [];
        const cbs = block.querySelectorAll('.op-worker-checkbox');
        cbs.forEach(cb => workerCbs.push({ w: cb.dataset.worker, checked: !!cb.checked }));
        state.ops.push({ name, dur, unit, breakVal, breakUnit, workers: workerCbs });
    });

    return state;
}

function restoreAppState(state) {
    if (!state || typeof state !== 'object') return;
    try {
        // Поля верхнего уровня
        if (document.getElementById('workerCount')) document.getElementById('workerCount').value = state.workerCount || 1;
        if (document.getElementById('totalOps')) document.getElementById('totalOps').value = state.totalOps || 1;
        if (document.getElementById('timeMode')) document.getElementById('timeMode').value = state.timeMode || 'total';
        if (document.getElementById('chainMode')) document.getElementById('chainMode').checked = !!state.chainMode;
        if (document.getElementById('opsSortMode')) document.getElementById('opsSortMode').value = state.opsSortMode || 'sequential';
        if (state.techCardValue && document.getElementById('techCardSelect')) {
            document.getElementById('techCardSelect').value = state.techCardValue;
            if (globalThis._tcDropdown) globalThis._tcDropdown.refresh();
        }
        if (document.getElementById('startDate')) document.getElementById('startDate').value = state.startDate || '';
        if (document.getElementById('startTime')) document.getElementById('startTime').value = state.startTime || '';
        if (document.getElementById('postingDate')) document.getElementById('postingDate').value = state.postingDate || state.startDate || '';
        if (document.getElementById('lunchStart')) document.getElementById('lunchStart').value = state.lunchStart || '';
        if (document.getElementById('lunchStart2')) document.getElementById('lunchStart2').value = state.lunchStart2 || '';
        if (document.getElementById('lunchDur')) document.getElementById('lunchDur').value = state.lunchDur || '';
        if (document.getElementById('orderName')) document.getElementById('orderName').value = state.orderName || '';
        if (document.getElementById('itemName')) document.getElementById('itemName').value = state.itemName || '';
        if (document.getElementById('statusBefore')) document.getElementById('statusBefore').value = state.statusBefore || '';
        if (document.getElementById('workExtra')) document.getElementById('workExtra').value = state.workExtra || '';
        if (document.getElementById('devRec')) document.getElementById('devRec').value = state.devRec || '';
        if (document.getElementById('resIz')) document.getElementById('resIz').value = state.resIz || '';
        if (document.getElementById('coefK')) document.getElementById('coefK').value = state.coefK || '';

        // Пересоздаём блоки операций до нужного количества и заполняем поля
        renderFields();
        const blocks = document.querySelectorAll('.op-block');
        blocks.forEach((block, idx) => {
            const row = state.ops[idx];
            if (!row) return;
            const nameInp = block.querySelector('.op-header-input');
            if (nameInp) nameInp.value = row.name;
            const durInp = block.querySelector('.op-duration');
            if (durInp) durInp.value = row.dur;
            const unitSel = block.querySelector('.op-unit');
            if (unitSel) unitSel.value = row.unit || 'min';
            const breakInp = block.querySelector('.op-break-val');
            if (breakInp) breakInp.value = row.breakVal;
            const breakUnit = block.querySelector('.op-break-unit');
            if (breakUnit) breakUnit.value = row.breakUnit || 'min';
            const cbs = block.querySelectorAll('.op-worker-checkbox');
            cbs.forEach(cb => {
                const w = cb.dataset.worker;
                const found = (row.workers || []).find(x => String(x.w) === String(w));
                if (found) cb.checked = !!found.checked;
            });
        });

        // Обновляем состояния UI
        try { updateWorkerUIByTimeMode(); syncTimeUnits(); updateFirstPauseVisibility(); } catch (e) {}
    } catch (e) {
        console.error('restoreAppState error:', e);
    }
}

function _pushUndoSnapshot() {
    try {
        const s = captureAppState();
        // Избегаем дублирования последовательных состояний
        const last = _undoStack[_undoStack.length - 1];
        if (JSON.stringify(last) === JSON.stringify(s)) return;
        _undoStack.push(s);
        if (_undoStack.length > _UNDO_LIMIT) _undoStack.shift();
        // Новое действие очищает стек повтора
        _redoStack.length = 0;
    } catch (e) { console.debug?.('pushUndo error', e?.message); }
}

function scheduleSnapshotDebounced() {
    if (_snapshotTimer) clearTimeout(_snapshotTimer);
    _snapshotTimer = setTimeout(() => { _pushUndoSnapshot(); _snapshotTimer = null; }, _SNAPSHOT_DEBOUNCE);
}

function undo() {
    if (_undoStack.length === 0) return;
    try {
        const current = captureAppState();
        _redoStack.push(current);
        const prev = _undoStack.pop();
        restoreAppState(prev);
    } catch (e) { console.error('undo error', e); }
}

function redo() {
    if (_redoStack.length === 0) return;
    try {
        const curr = captureAppState();
        _undoStack.push(curr);
        const next = _redoStack.pop();
        restoreAppState(next);
    } catch (e) { console.error('redo error', e); }
}

// Начальный снимок после загрузки
window.addEventListener('load', () => { try { _pushUndoSnapshot(); } catch (e) {} });

// Горячие клавиши: Ctrl/Cmd+Z = отмена, Ctrl/Cmd+Y или Ctrl+Shift+Z = повтор
document.addEventListener('keydown', (e) => {
    const key = (e.key || '').toLowerCase();
    const mod = (e.ctrlKey || e.metaKey);
    if (!mod) return;
    if (!e.shiftKey && key === 'z') {
        e.preventDefault();
        undo();
    } else if (key === 'y' || (e.shiftKey && key === 'z')) {
        e.preventDefault();
        redo();
    }
});

// Отложенные снимки для пользовательских правок: поля ввода в главном контейнере и некоторые элементы верхнего уровня
const _snapshotTargets = ['totalOps','workerCount','timeMode','startDate','startTime','postingDate','lunchStart','lunchStart2','lunchDur','orderName','itemName','statusBefore','workExtra','devRec','coefK','resIz','chainMode','opsSortMode'];
_snapshotTargets.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('input', scheduleSnapshotDebounced);
    el.addEventListener('change', scheduleSnapshotDebounced);
});
// делегируем события ввода внутри контейнера операций
const _opsContainer = document.getElementById('fieldsContainer');
if (_opsContainer) {
    _opsContainer.addEventListener('input', scheduleSnapshotDebounced, true);
    _opsContainer.addEventListener('change', scheduleSnapshotDebounced, true);
}

// === ЭКСПОРТ В EXCEL ===
let _excelExportInProgress = false;

function setupExcelExport() {
    document.getElementById('clearHistoryBtn').addEventListener('click', clearHistoryData);
    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
}

async function exportToExcel() {
    if (_excelExportInProgress) return;
    _excelExportInProgress = true;
    const btn = document.getElementById('exportExcelBtn');
    if (btn) { btn.disabled = true; }
    try {
    const historyList = document.getElementById('historyList');
    const entries = historyList.querySelectorAll('.history-entry');

    if (entries.length === 0) {
        await showMessage('История пуста!');
        return;
    }

    if (typeof ExcelJS === 'undefined') {
        await showMessage('Ошибка: библиотека ExcelJS не загружена', 'Ошибка', 'error');
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const sheetName = new Date().toLocaleDateString('ru-RU').replaceAll('.', '-');
    const ws = workbook.addWorksheet(sheetName);

    // --- Ширина колонок (от A до V) ---
    ws.columns = [
        { width: 3 },   // A spacer       
        { width: 4 },   // B №            
        { width: 65 },    // C Операция      
        { width: 7.3 },   // D Обед?        
        { width: 12.7 },  // E Пауза        
        { width: 12 },  // F Работа(alt)   
        { width: 14.5 },  // G ПДТВ         
        { width: 2.7 },   // H -            
        { width: 2.7 },   // I -           
        { width: 2.7 },   // J -           
        { width: 2.7 },   // K -            
        { width: 10.9 },  // L Работа(main)  
        { width: 2.7 },   // M -            
        { width: 2.7 },   // N -            
        { width: 18.2 },  // O Дата проводки 
        { width: 16.4 },  // P Исполнитель   
        { width: 2.7 },   // Q -            
        { width: 18.2 },  // R Дата Начала   
        { width: 18.2 },  // S Время Начала  
        { width: 18.2 },  // T Дата Конца    
        { width: 18.2 },  // U Время Конца   
        { width: 9.1 },   // V INDEX        
    ];

    // --- Строительные блоки стилей ExcelJS ---
    const THIN = { style: 'thin' };
    const MEDIUM = { style: 'medium' };
    function makeBorders(opts) {
        return {
            top: opts.thickTop ? MEDIUM : THIN,
            bottom: opts.thickBottom ? MEDIUM : THIN,
            left: opts.thickLeft ? MEDIUM : THIN,
            right: opts.thickRight ? MEDIUM : THIN,
        };
    }



    // === ЦВЕТОВАЯ ПАЛИТРА EXCEL — загружается из настроек (меняется в «Настройки По Умолчанию») ===
    const _ec = getUserDefaults().excelColors;
    const hexToArgb = (hex) => 'FF' + hex.replace('#', '').toUpperCase();
    const FILL_LOCKED    = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(_ec.locked) } };
    const FILL_EDITABLE  = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(_ec.editable) } };
    const FILL_HEADER    = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(_ec.header) } };
    const FILL_AUTHOR    = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(_ec.author) } };
    const FILL_PDTV      = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(_ec.pdtv ?? '#FFF9C4') } };
    // Цвет текста подбирается автоматически по WCAG для обеспечения контрастности
    const FONT_LOCKED   = { name: 'Arial', size: 12, color: { argb: hexToArgb(getContrastColor(_ec.locked)) } };
    const FONT_EDITABLE = { name: 'Arial', size: 12, color: { argb: hexToArgb(getContrastColor(_ec.editable)) } };
    const FONT_ICON     = { name: 'Arial', size: 14, color: { argb: hexToArgb(getContrastColor(_ec.locked)) } };
    const FONT_HEADER   = { name: 'Arial', size: 12, bold: true, color: { argb: hexToArgb(getContrastColor(_ec.header)) } };
    const FONT_AUTHOR   = { name: 'Arial', size: 30, bold: true, color: { argb: hexToArgb(getContrastColor(_ec.author)) } };
    const FONT_SETTINGS = { name: 'Arial', size: 14, bold: true, color: { argb: hexToArgb(getContrastColor(_ec.author)) } };
    const FONT_PDTV     = { name: 'Arial', size: 12, color: { argb: hexToArgb(getContrastColor(_ec.pdtv ?? '#FFF9C4')) } };

    const ALIGN_CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const ALIGN_LEFT = { horizontal: 'left', vertical: 'middle', wrapText: true };
    const ALIGN_CENTER_NOWRAP = { horizontal: 'center', vertical: 'middle' };

    function applyStyle(cell, style) {
        if (style.font) cell.font = style.font;
        if (style.fill) cell.fill = style.fill;
        if (style.alignment) cell.alignment = style.alignment;
        if (style.border) cell.border = style.border;
        if (style.numFmt) cell.numFmt = style.numFmt;
        cell.protection = { locked: style.locked !== false };
    }

    // Фабрика стилей — возвращает карту стилей в зависимости от того, является ли строка последней в группе операции
    function getStyleMap(isGroupEnd) {
        const bb = isGroupEnd;
        const brd = (extra) => makeBorders(Object.assign({ thickBottom: bb }, extra || {}));
        return {
            borderLocked:    { font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd(),                  locked: true },
            borderLeftLocked:{ font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd(),                  locked: true },
            iconLocked:      { font: FONT_ICON,     fill: FILL_LOCKED,   alignment: ALIGN_CENTER_NOWRAP,border: brd(),                  locked: true },
            timeLocked:      { font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd(),   numFmt: 'h:mm:ss',    locked: true },
            timeEditable:    { font: FONT_EDITABLE, fill: FILL_EDITABLE, alignment: ALIGN_CENTER,       border: brd(),   numFmt: 'h:mm:ss',    locked: false },
            durEditable:     { font: FONT_EDITABLE, fill: FILL_EDITABLE, alignment: ALIGN_CENTER,       border: brd(),   numFmt: '0.00',       locked: false },
            durLocked:       { font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd(),   numFmt: '0.00',       locked: true },
            dateLocked:      { font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd(),   numFmt: 'dd.mm.yyyy', locked: true },
            dateEditable:    { font: FONT_EDITABLE, fill: FILL_EDITABLE, alignment: ALIGN_CENTER,       border: brd(),   numFmt: 'dd.mm.yyyy', locked: false },
            borderEditable:  { font: FONT_EDITABLE, fill: FILL_EDITABLE, alignment: ALIGN_CENTER,       border: brd(),                  locked: false },
            pdtvLocked:      { font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd({ thickLeft: true }),locked: true },
            pdtvEditable:    { font: FONT_EDITABLE, fill: FILL_EDITABLE, alignment: ALIGN_CENTER,       border: brd({ thickLeft: true }),locked: false },
            pdtvFormula:     { font: FONT_PDTV,     fill: FILL_PDTV,     alignment: ALIGN_CENTER,       border: brd({ thickLeft: true }),locked: false },
            endTimeLocked:   { font: FONT_LOCKED,   fill: FILL_LOCKED,   alignment: ALIGN_CENTER,       border: brd({ thickRight: true }), numFmt: 'h:mm:ss', locked: true },
        };
    }

    // Вспомогательная: установить значение и стиль ячейки
    function setCell(row, col, value, style) {
        const c = row.getCell(col);
        if (value !== null && value !== undefined && typeof value === 'object' && 'formula' in value) {
            c.value = { formula: value.formula };
        } else {
            c.value = value;
        }
        applyStyle(c, style);
    }

    // Вспомогательная: установить объединённую строку (B:V) со стилем
    function setMergedRow(ws, rowNum, value, styleProps) {
        ws.mergeCells(`B${rowNum}:V${rowNum}`);
        const c = ws.getCell(`B${rowNum}`);
        c.value = value;
        if (styleProps.font) c.font = styleProps.font;
        if (styleProps.fill) c.fill = styleProps.fill;
        if (styleProps.alignment) c.alignment = styleProps.alignment;
        if (styleProps.border) c.border = styleProps.border;
        c.protection = { locked: styleProps.locked !== false };
    }

    // Константы букв колонок для формул в формате A1
    const CL = { PAUSE: 'E', DUR: 'L', START_DATE: 'R', START_TIME: 'S', END_DATE: 'T', END_TIME: 'U', KEY: 'V' };

    // --- Строка 1: Settings info (из сохранённых данных записей, не из полей ввода) ---
    let sheetRow = 0;
    const firstEntryData = safeJsonParse(entries[0]?.dataset?.jsonData);
    const lunchData = firstEntryData?.lunch;
    const isChainGlobal = firstEntryData?.chain ?? document.getElementById('chainMode')?.checked;
    const lunch1Val = lunchData ? `${String(lunchData.h ?? 0).padStart(2, '0')}:${String(lunchData.m ?? 0).padStart(2, '0')}` : '00:00';
    const lunch2H = lunchData?.h2 ?? 0, lunch2M = lunchData?.m2 ?? 0;
    const lunch2Val = lunchData ? `${String(lunch2H).padStart(2, '0')}:${String(lunch2M).padStart(2, '0')}` : '00:00';
    const lunchDurVal = lunchData?.dur ?? 0;
    const lunch2Text = (lunch2H === 0 && lunch2M === 0) ? 'НЕ Учитывается' : lunch2Val;
    const settingsText = `Режим Формирования: ${isChainGlobal ? 'Цепочка' : 'НЕ Цепочка'}  |  Обед 1: ${lunch1Val}  |  Обед 2: ${lunch2Text}  |  Обед(мин): ${lunchDurVal}  |  Записей: ${entries.length}`;

    const settingsRow = ws.addRow([]);
    sheetRow++;
    const settingsRowNum = sheetRow;
    setMergedRow(ws, sheetRow, settingsText,
        { font: FONT_SETTINGS, fill: FILL_AUTHOR, alignment: { horizontal: 'center', vertical: 'middle' }, locked: true });
    settingsRow.height = 30;

    // Закрепление строки 1 (настройки) при прокрутке
    ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 0 }];

    // --- Строка 2: Информация об авторе ---
    const authorRow = ws.addRow([]);
    sheetRow++;
    setMergedRow(ws, sheetRow, 'Создано при помощи калькулятора для ленивых v.0.6.9',
        { font: FONT_AUTHOR, fill: FILL_AUTHOR, alignment: { horizontal: 'center', vertical: 'middle' }, locked: true });
    authorRow.height = 50;

    // --- Строка 3: Разделитель ---
    ws.addRow([]);
    sheetRow++;
    setMergedRow(ws, sheetRow, '', { alignment: { vertical: 'middle', wrapText: true }, locked: true });

    let previousEntryData = null;
    const entriesArray = Array.from(entries).reverse();

    // Глобальное отслеживание логики «Установить один раз»
    let globalPostingRow = null;
    const globalWorkerRowMap = {};

    entriesArray.forEach((entry, entryIndex) => {
        const data = safeJsonParse(entry.dataset.jsonData);
        if (!data) return;
        const lh = data.lunch.h || 0;
        const lm = data.lunch.m || 0;
        const lh2 = (data.lunch.h2 !== undefined) ? data.lunch.h2 : 0;
        const lm2 = (data.lunch.m2 !== undefined) ? data.lunch.m2 : 0;
        const ld = data.lunch.dur || 60;
        const isChain = data.chain;

        // Без режима цепочки — привязка исполнителей и даты проводки в рамках одной записи
        if (!isChain) {
            globalPostingRow = null;
            for (const k of Object.keys(globalWorkerRowMap)) delete globalWorkerRowMap[k];
        }

        const headerUnit = getHeaderUnitSuffix(data.rows);
        let altHeaderUnit = '';
        if (headerUnit === ' (мин)') altHeaderUnit = ' (час)';
        else if (headerUnit === ' (час)') altHeaderUnit = ' (мин)';
        else altHeaderUnit = ' (алт.)';

        // --- Строка заголовка ---
        ws.addRow([]);
        sheetRow++;
        const modeStr = data.timeMode === 'per_worker' ? 'На Каждого' : data.timeMode === 'individual' ? 'Индивидуальный' : 'Общий';
        const pdtvModeStr = data.rows?.[0]?.pdtvAutoMode === true ? 'Авто' : 'НЕ Авто';
        const headerRowNum = sheetRow;

        // B:C — выпадающий список подтверждения
        ws.mergeCells(`B${headerRowNum}:C${headerRowNum}`);
        const confirmCell = ws.getCell(`B${headerRowNum}`);
        confirmCell.value = 'НЕ Подтверждено';
        confirmCell.font = FONT_HEADER;
        confirmCell.fill = FILL_HEADER;
        confirmCell.alignment = ALIGN_CENTER;
        confirmCell.border = makeBorders({});
        confirmCell.protection = { locked: false };
        confirmCell.dataValidation = {
            type: 'list',
            allowBlank: false,
            formulae: ['"Подтверждено,НЕ Подтверждено"']
        };

        // D:V — текст заголовка записи
        ws.mergeCells(`D${headerRowNum}:V${headerRowNum}`);
        const titleCell = ws.getCell(`D${headerRowNum}`);
        titleCell.value = excelSanitizeCell(data.title) + ' | Режим Времени: ' + modeStr + ' | Режим ПДТВ: ' + pdtvModeStr;
        titleCell.font = FONT_HEADER;
        titleCell.fill = FILL_HEADER;
        titleCell.alignment = ALIGN_CENTER;
        titleCell.border = makeBorders({});
        titleCell.protection = { locked: true };

        // --- Строка заголовков ---
        const headerLabels = [
            null, '№', 'Операция', 'Обед?', 'Пауза в начале операции',
            'Работа' + altHeaderUnit, 'ПДТВ', '-', '-', '-', '-',
            'Работа' + headerUnit, '-', '-', 'Дата проводки',
            'Исполнитель', '-', 'Дата Начала', 'Время Начала',
            'Дата Конца', 'Время Конца', 'INDEX'
        ];
        const hRow = ws.addRow(headerLabels);
        sheetRow++;
        const colHeaderRowNum = sheetRow;
        for (let ci = 2; ci <= 22; ci++) {
            const hCell = hRow.getCell(ci);
            hCell.font = FONT_HEADER;
            hCell.fill = FILL_HEADER;
            hCell.alignment = ALIGN_CENTER;
            hCell.protection = { locked: true };
            if (ci === 7) hCell.border = makeBorders({ thickLeft: true });
            else if (ci === 21) hCell.border = makeBorders({ thickRight: true });
            else hCell.border = makeBorders({});
        }

        // Сортировка строк по opNumeric, затем по workerIndex
        let rowsForExport = data.rows.slice().sort((a, b) => {
            const na = (a.opNumeric ?? a.opIdx) || 0;
            const nb = (b.opNumeric ?? b.opIdx) || 0;
            if (na !== nb) return na - nb;
            return (a.workerIndex || 1) - (b.workerIndex || 1);
        });

        const dataStartRow = sheetRow + 1;
        const dataEndRow = dataStartRow + rowsForExport.length - 1;
        const rowPosMap = {};

        // Якорная строка ПДТВ для текущей записи (используется для Excel-формул в авто-режиме)
        let entryBaseGRow = null;
        let entryBaseOffset = 0;
        // Строка первого исполнителя каждой операции (для формул ручного режима)
        const opFirstWorkerGRow = {};

        rowsForExport.forEach((r, idx) => {
            const pauseVal = typeof r.pauseExcelVal === 'number' ? r.pauseExcelVal : 0;
            const unitDiv = (r.unit === 'hour') ? 24.0 : 1440.0;
            const curOpNum = r.opNumeric ?? r.opIdx;
            const prevRowOpNum = (idx > 0) ? (rowsForExport[idx - 1].opNumeric ?? rowsForExport[idx - 1].opIdx) : -1;
            const nextRowOpNum = (idx < rowsForExport.length - 1) ? (rowsForExport[idx + 1].opNumeric ?? rowsForExport[idx + 1].opIdx) : -1;
            const isGroupEnd = (idx === rowsForExport.length - 1) || (curOpNum !== nextRowOpNum);
            const styles = getStyleMap(isGroupEnd);

            // Текущая абсолютная строка (нумерация с 1)
            const curRow = sheetRow + 1;
            // Вспомогательная функция для ссылки на ячейку той же строки в нотации A1
            const cr = (col) => `${col}${curRow}`;

            // === ДЛИТЕЛЬНОСТЬ ===
            let durValue, durStyle;
            if (data.timeMode === 'individual') {
                durValue = r.durVal;
                durStyle = styles.durEditable;
            } else {
                if (curOpNum === prevRowOpNum) {
                    durValue = { formula: `L${curRow - 1}` };
                    durStyle = styles.durLocked;
                } else {
                    durValue = r.durVal;
                    durStyle = styles.durEditable;
                }
            }

            // === АЛЬТ. ДЛИТЕЛЬНОСТЬ (формула) ===
            const altDurFormula = (r.unit === 'hour') ? `L${curRow}*60` : `L${curRow}/60`;

            // === ПАУЗА ===
            const isFirstEntryFirstOp = (entryIndex === 0 && curOpNum === 1);
            const isFirstOpOfEntry = (curOpNum === 1);
            const isFirstWorkerOfOp = (curOpNum !== prevRowOpNum);
            let pauseCellValue, pauseStyle;
            if (isFirstEntryFirstOp) {
                if (r.workerIndex === 1) {
                    pauseCellValue = pauseVal; pauseStyle = styles.timeLocked;
                } else {
                    pauseCellValue = { formula: `E${curRow - 1}` }; pauseStyle = styles.timeLocked;
                }
            } else if (isFirstOpOfEntry) {
                if (r.workerIndex === 1) {
                    pauseCellValue = pauseVal; pauseStyle = styles.timeEditable;
                } else {
                    pauseCellValue = { formula: `E${curRow - 1}` }; pauseStyle = styles.timeLocked;
                }
            } else {
                if (isFirstWorkerOfOp) {
                    pauseCellValue = pauseVal; pauseStyle = styles.timeEditable;
                } else {
                    pauseCellValue = { formula: `E${curRow - 1}` }; pauseStyle = styles.timeLocked;
                }
            }

            // === ВРЕМЯ НАЧАЛА (формулы A1) ===
            let startTimeValue, startTimeStyle;
            let fullStartFormula = '';
            if (idx === 0) {
                if (isChain && previousEntryData) {
                    const offset = 5 + (previousEntryData.z7.length * 1);
                    const prevRow = curRow - offset;
                    const rawTimeRef = `(T${prevRow} + U${prevRow} + ${cr(CL.PAUSE)})`;
                    fullStartFormula = buildLunchShiftFormula(rawTimeRef, lh, lm, lh2, lm2, ld);
                    startTimeValue = { formula: `MOD(${fullStartFormula}, 1)` };
                    startTimeStyle = styles.timeLocked;
                } else {
                    // Редактируемое время начала — сохраняем как дробную часть суток
                    const st = new Date(r.startObj);
                    startTimeValue = (st.getHours() * 3600 + st.getMinutes() * 60 + st.getSeconds()) / 86400;
                    startTimeStyle = styles.timeEditable;
                }
            } else {
                if (data.timeMode === 'individual') {
                    if ((curOpNum || 0) > 1) {
                        const prevKey = `${(curOpNum - 1)}_${r.workerIndex || 1}`;
                        const keyRange = `V${dataStartRow}:V${dataEndRow}`;
                        const timeRange = `U${dataStartRow}:U${dataEndRow}`;
                        const dateRange = `T${dataStartRow}:T${dataEndRow}`;
                        const lookupTime = `INDEX(${timeRange}, MATCH("${prevKey}", ${keyRange}, 0))`;
                        const lookupDate = `INDEX(${dateRange}, MATCH("${prevKey}", ${keyRange}, 0))`;
                        const rawTimeWithPause = `(${lookupDate}+${lookupTime}+${cr(CL.PAUSE)})`;
                        fullStartFormula = buildLunchShiftFormula(rawTimeWithPause, lh, lm, lh2, lm2, ld);
                        startTimeValue = { formula: `MOD(${fullStartFormula}, 1)` };
                        startTimeStyle = styles.timeLocked;
                    } else {
                        if (curOpNum === prevRowOpNum) {
                            startTimeValue = { formula: `S${curRow - 1}` };
                            startTimeStyle = styles.timeLocked;
                        } else {
                            const rawTimeWithPause = `(T${curRow - 1}+U${curRow - 1}+${cr(CL.PAUSE)})`;
                            fullStartFormula = buildLunchShiftFormula(rawTimeWithPause, lh, lm, lh2, lm2, ld);
                            startTimeValue = { formula: `MOD(${fullStartFormula}, 1)` };
                            startTimeStyle = styles.timeLocked;
                        }
                    }
                } else {
                    if (curOpNum === prevRowOpNum) {
                        startTimeValue = { formula: `S${curRow - 1}` };
                        startTimeStyle = styles.timeLocked;
                    } else {
                        const rawTimeWithPause = `(T${curRow - 1}+U${curRow - 1}+${cr(CL.PAUSE)})`;
                        fullStartFormula = buildLunchShiftFormula(rawTimeWithPause, lh, lm, lh2, lm2, ld);
                        startTimeValue = { formula: `MOD(${fullStartFormula}, 1)` };
                        startTimeStyle = styles.timeLocked;
                    }
                }
            }

            // === ФОРМУЛА ИКОНКИ (ОБЕД) ===
            const l1Val = `TIME(${lh},${lm},0)`;
            const l1End = `(TIME(${lh},${lm},0)+TIME(0,${ld},0))`;
            const lDurVal = `TIME(0,${ld},0)`;
            const hasLunch2 = !(lh2 === 0 && lm2 === 0);
            const l2Val = `TIME(${lh2},${lm2},0)`;
            const l2End = `(TIME(${lh2},${lm2},0)+TIME(0,${ld},0))`;

            const startTimeMod = `MOD(${cr(CL.START_TIME)}, 1)`;
            const endTimeRel = `(${startTimeMod}+(${cr(CL.DUR)}/${unitDiv}))`;
            const l1EndMod = `MOD(${l1End}, 1)`;
            const icWasShifted1 = `ABS(${startTimeMod}-${l1EndMod})<TIME(0,0,1)`;
            const icCovers1 = `OR(AND(${startTimeMod}<${l1Val}, ${endTimeRel}>(${l1Val}+TIME(0,0,1))), AND(${startTimeMod}<(${l1Val}+1), ${endTimeRel}>(${l1Val}+1+TIME(0,0,1))))`;
            const icC1 = `OR(${icWasShifted1}, ${icCovers1})`;
            const icShift1 = `IF(${icC1}, ${lDurVal}, 0)`;

            let formulaIcon;
            if (hasLunch2) {
                const shiftedStartMod = `MOD(${cr(CL.START_TIME)}+${icShift1}, 1)`;
                const shiftedEndRel = `(${shiftedStartMod}+(${cr(CL.DUR)}/${unitDiv}))`;
                const l2EndMod = `MOD(${l2End}, 1)`;
                const icWasShifted2 = `ABS(${shiftedStartMod}-${l2EndMod})<TIME(0,0,1)`;
                const icCovers2 = `OR(AND(${shiftedStartMod}<${l2Val}, ${shiftedEndRel}>(${l2Val}+TIME(0,0,1))), AND(${shiftedStartMod}<(${l2Val}+1), ${shiftedEndRel}>(${l2Val}+1+TIME(0,0,1))))`;
                const icC2 = `OR(${icWasShifted2}, ${icCovers2})`;
                formulaIcon = `IF(OR(${icC1}, ${icC2}), "🍽️", "")`;
            } else {
                formulaIcon = `IF(${icC1}, "🍽️", "")`;
            }

            // === ФОРМУЛЫ ВРЕМЕНИ ОКОНЧАНИЯ / ДАТЫ ОКОНЧАНИЯ ===
            const stMod = `MOD(${cr(CL.START_TIME)}, 1)`;
            const rawEndRel = `(${stMod}+(${cr(CL.DUR)}/${unitDiv}))`;
            const enC1 = `OR(AND(${stMod} < ${l1Val}, ${rawEndRel} > (${l1Val}+TIME(0,0,1))), AND(${stMod} < (${l1Val}+1), ${rawEndRel} > (${l1Val}+1+TIME(0,0,1))))`;
            const enShift1 = `IF(${enC1}, ${lDurVal}, 0)`;
            const mainMath = `${cr(CL.START_DATE)} + ${cr(CL.START_TIME)} + (${cr(CL.DUR)}/${unitDiv}) + ${enShift1}`;

            let formulaEnd, endDateFormula;
            if (hasLunch2) {
                const stMod2 = `MOD(${cr(CL.START_TIME)} + ${enShift1}, 1)`;
                const rawEndRel2 = `(${stMod2}+(${cr(CL.DUR)}/${unitDiv}))`;
                const enC2 = `OR(AND(${stMod2} < ${l2Val}, ${rawEndRel2} > (${l2Val}+TIME(0,0,1))), AND(${stMod2} < (${l2Val}+1), ${rawEndRel2} > (${l2Val}+1+TIME(0,0,1))))`;
                const enShift2 = `IF(${enC2}, ${lDurVal}, 0)`;
                formulaEnd = `MOD(${mainMath} + ${enShift2}, 1)`;
                endDateFormula = `INT(${mainMath} + ${enShift2})`;
            } else {
                formulaEnd = `MOD(${mainMath}, 1)`;
                endDateFormula = `INT(${mainMath})`;
            }

            // === ИНДЕКС ОПЕРАЦИИ (ПДТВ) ===
            let pdtvCellValue, pdtvCellStyle;
            const isAutoPdtv = r.pdtvAutoMode === true;
            if (!isAutoPdtv) {
                if (isFirstWorkerOfOp) {
                    // Первый исполнитель операции — якорная ячейка с числовым значением
                    const opIdxNum = Number(String(r.opIdx ?? '').replaceAll("'", ""));
                    pdtvCellValue = Number.isFinite(opIdxNum) ? opIdxNum : String(r.opIdx ?? '');
                    pdtvCellStyle = styles.pdtvEditable;
                    opFirstWorkerGRow[curOpNum] = curRow;
                } else {
                    // Последующие исполнители — ссылка на первую запись ПДТВ этой операции
                    pdtvCellValue = { formula: `G${opFirstWorkerGRow[curOpNum]}` };
                    pdtvCellStyle = styles.pdtvFormula;
                }
            } else {
                const currentOffset = typeof r.pdtvOffset === 'number' ? r.pdtvOffset : 0;
                if (entryBaseGRow === null) {
                    // Первая строка записи — якорная ячейка с числовым значением
                    entryBaseGRow = curRow;
                    entryBaseOffset = currentOffset;
                    const opIdxNum = Number(String(r.opIdx ?? '').replaceAll("'", ""));
                    pdtvCellValue = Number.isFinite(opIdxNum) ? opIdxNum : String(r.opIdx ?? '');
                    pdtvCellStyle = styles.pdtvEditable;
                } else {
                    // Остальные строки — формула относительно якорной ячейки
                    const diff = currentOffset - entryBaseOffset;
                    if (diff === 0) {
                        pdtvCellValue = { formula: `G${entryBaseGRow}` };
                    } else if (diff > 0) {
                        pdtvCellValue = { formula: `G${entryBaseGRow}+${diff}` };
                    } else {
                        pdtvCellValue = { formula: `G${entryBaseGRow}${diff}` };
                    }
                    pdtvCellStyle = styles.pdtvFormula;
                }
            }

            // === ИСПОЛНИТЕЛЬ (глобальная привязка) ===
            const workerRaw = String(r.worker || '');
            const workerNum = Number(workerRaw.replaceAll("'", ""));
            const wIdx = r.workerIndex || 1;
            if (globalWorkerRowMap[wIdx] === undefined) globalWorkerRowMap[wIdx] = curRow;
            const targetWorkerRow = globalWorkerRowMap[wIdx];
            let workerValue, workerStyle;
            if (targetWorkerRow === curRow) {
                workerValue = (workerRaw.trim() !== '' && Number.isFinite(workerNum)) ? workerNum : excelSanitizeCell(workerRaw);
                workerStyle = styles.borderEditable;
            } else {
                workerValue = { formula: `P${targetWorkerRow}` };
                workerStyle = styles.borderLocked;
            }

            // === ДАТА ПРОВОДКИ (глобальная привязка) ===
            let postingValue, postingStyle;
            const _pd = r.postingDateIso ? new Date(String(r.postingDateIso) + 'T00:00:00') : new Date(r.startObj);
            const postingDate = new Date(Date.UTC(_pd.getFullYear(), _pd.getMonth(), _pd.getDate()));
            if (globalPostingRow === null) {
                globalPostingRow = curRow;
                postingValue = postingDate;
                postingStyle = styles.dateEditable;
            } else {
                postingValue = { formula: `O${globalPostingRow}` };
                postingStyle = styles.dateLocked;
            }

            // === ДАТА НАЧАЛА ===
            let startDateValue, startDateStyle;
            const _sd = new Date(r.startObj);
            const startDate = new Date(Date.UTC(_sd.getFullYear(), _sd.getMonth(), _sd.getDate()));
            if (idx === 0 && !(isChain && previousEntryData)) {
                startDateValue = startDate;
                startDateStyle = styles.dateEditable;
            } else if (fullStartFormula) {
                startDateValue = { formula: `INT(${fullStartFormula})` };
                startDateStyle = styles.dateLocked;
            } else {
                startDateValue = { formula: `R${curRow - 1}` };
                startDateStyle = styles.dateLocked;
            }

            // === ФОРМИРОВАНИЕ СТРОКИ ===
            const dataRow = ws.addRow([]);
            sheetRow++;

            setCell(dataRow, 2, r.originalOpIndex || (idx + 1), styles.borderLocked);
            setCell(dataRow, 3, excelSanitizeCell(r.name), styles.borderLeftLocked);
            setCell(dataRow, 4, { formula: formulaIcon }, styles.iconLocked);
            setCell(dataRow, 5, pauseCellValue, pauseStyle);
            setCell(dataRow, 6, { formula: altDurFormula }, styles.durLocked);
            setCell(dataRow, 7, pdtvCellValue, pdtvCellStyle);
            for (let ci = 8; ci <= 11; ci++) setCell(dataRow, ci, '', styles.borderLocked);
            setCell(dataRow, 12, durValue, durStyle);
            setCell(dataRow, 13, '', styles.borderLocked);
            setCell(dataRow, 14, '', styles.borderLocked);
            setCell(dataRow, 15, postingValue, postingStyle);
            setCell(dataRow, 16, workerValue, workerStyle);
            setCell(dataRow, 17, '', styles.borderLocked);
            setCell(dataRow, 18, startDateValue, startDateStyle);
            setCell(dataRow, 19, startTimeValue, startTimeStyle);
            setCell(dataRow, 20, { formula: endDateFormula }, styles.dateLocked);
            setCell(dataRow, 21, { formula: formulaEnd }, styles.endTimeLocked);
            setCell(dataRow, 22, String(curOpNum) + '_' + String(r.workerIndex || 1), styles.borderLocked);

            rowPosMap[`${curOpNum}_${r.workerIndex || 1}`] = idx;
        });

        // --- Строка заголовка Z7 ---
        ws.addRow([]);
        sheetRow++;
        const z7HeaderRowNum = sheetRow;
        setMergedRow(ws, sheetRow, 'Z7',
            { font: FONT_HEADER, fill: FILL_HEADER, alignment: ALIGN_CENTER, border: makeBorders({}), locked: true });

        // Условное форматирование: при «Подтверждено» — строка заголовка записи + строка заголовков колонок + строка Z7
        const cfConfirmedRule = [{
            type: 'expression',
            formulae: [`$B$${headerRowNum}="Подтверждено"`],
            style: {
                fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: hexToArgb(_ec.confirmed) } },
                font: { color: { argb: hexToArgb(getContrastColor(_ec.confirmed)) } }
            }
        }];
        ws.addConditionalFormatting({ ref: `B${headerRowNum}:V${headerRowNum}`,   rules: cfConfirmedRule });
        ws.addConditionalFormatting({ ref: `B${colHeaderRowNum}:V${colHeaderRowNum}`, rules: cfConfirmedRule });
        ws.addConditionalFormatting({ ref: `B${z7HeaderRowNum}:V${z7HeaderRowNum}`,   rules: cfConfirmedRule });

        // --- Строки содержимого Z7 ---
        data.z7.forEach((line) => {
            const sanitizedZ7 = excelSanitizeCell(line);
            const charsPerLine = 80;
            const linesNeeded = Math.max(1, Math.ceil(String(sanitizedZ7).length / charsPerLine));
            const z7Row = ws.addRow([]);
            sheetRow++;
            setMergedRow(ws, sheetRow, sanitizedZ7,
                { font: FONT_LOCKED, fill: FILL_LOCKED, alignment: ALIGN_LEFT, border: makeBorders({}), locked: true });
            if (linesNeeded > 1) {
                z7Row.height = Math.min(400, 18 * linesNeeded);
            }
        });

        // --- Пустая строка-разделитель ---
        ws.addRow([]);
        sheetRow++;

        previousEntryData = data;
    });

    // Обновляем строку настроек: добавляем реактивный подсчёт статусов подтверждения
    const escapedSettings = settingsText.replaceAll('"', '""');
    const settingsCell = ws.getCell(`B${settingsRowNum}`);
    settingsCell.value = {
        formula: `"${escapedSettings}  |  Подтверждено: " & COUNTIF(B${settingsRowNum + 1}:B${sheetRow},"Подтверждено") & "  |  НЕ Подтверждено: " & COUNTIF(B${settingsRowNum + 1}:B${sheetRow},"НЕ Подтверждено")`
    };

    // Защита листа от случайного редактирования
    const sheetPassword = [71,71,82,45,49,51,48,49].map(c => String.fromCodePoint(c)).join('');
    await ws.protect(sheetPassword, {
        selectLockedCells: true,
        selectUnlockedCells: true,
    });

    // Генерация xlsx-буфера
    const buffer = await workbook.xlsx.writeBuffer();
    await downloadExcelFile(buffer);
    } catch (e) {
        console.error('exportToExcel error:', e);
        await showMessage('Ошибка при экспорте', 'Ошибка', 'error');
    } finally {
        _excelExportInProgress = false;
        const btn = document.getElementById('exportExcelBtn');
        if (btn) { btn.disabled = false; }
    }
}

async function downloadExcelFile(buffer) {
    const fileName = `История_Расчетов_${new Date().toLocaleDateString('ru-RU').replaceAll('.', '-')}.xlsx`;

    // Пробуем использовать Tauri API
    if (tauriDialog?.save && tauriInvoke) {
        try {
            const filePath = await tauriDialog.save({
                defaultPath: fileName,
                filters: [{ name: 'Excel', extensions: ['xlsx'] }]
            });

            if (filePath) {
                await tauriInvoke('save_file_binary', {
                    path: filePath,
                    content: Array.from(new Uint8Array(buffer))
                });
                await showMessage('Файл успешно сохранён!', 'Успех');
            }
            return;
        } catch (e) {
            console.error('Ошибка сохранения:', e);
            await showMessage(String(e), 'Ошибка', 'error');
            return;
        }
    }

    // Запасной вариант — браузерный метод
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = fileName;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// === УПРАВЛЕНИЕ ТЕХКАРТАМИ ===
function getCardData() {
    return Array.from(document.querySelectorAll('.op-block')).map(b => {
        const durInput = b.querySelector('.op-duration');
        const breakInput = b.querySelector('.op-break-val');
        
        // Если значения временно обнулены (напр. Индивидуальный режим), восстанавливаем сохранённые валидные значения
        const durRaw = (durInput && durInput.dataset.savedVal !== undefined) 
            ? durInput.dataset.savedVal 
            : (durInput ? durInput.value : 0);
            
        const breakRaw = (breakInput && breakInput.dataset.savedVal !== undefined)
            ? breakInput.dataset.savedVal
            : (breakInput ? breakInput.value : 0);

        return {
            // Сохраняем имя операции без порядкового префикса
            name: sanitizeStrict(stripOrdinalPrefix(b.querySelector('.op-header-input').value), 200),
            dur: Math.max(0, Number.parseFloat(durRaw) || 0),
            unit: b.querySelector('.op-unit').value,
            // hasBreak: определяем по значению перерыва (чекбокса больше нет)
            hasBreak: (Math.max(0, Number.parseFloat(breakRaw) || 0) > 0),
            breakVal: Math.max(0, Number.parseFloat(breakRaw) || 0),
            breakUnit: b.querySelector('.op-break-unit').value
        };
    });
}

function setCardData(steps) {
    if (!validateCardData(steps)) {
        showMessage('Ошибка: некорректные данные шаблона').catch(() => {});
        return;
    }
    if (document.getElementById('opsSortMode')) document.getElementById('opsSortMode').value = 'sequential';

    document.getElementById('totalOps').value = Math.min(steps.length, 20);
    container.textContent = '';
    renderFields();

    const blocks = document.querySelectorAll('.op-block');
    
    // Сначала устанавливаем единицу для первой операции
    if (steps[0] && blocks[0]) {
        blocks[0].querySelector('.op-unit').value = steps[0].unit;
    }
    
    steps.forEach((s, i) => {
        if (!blocks[i]) return;
        // Для отображения в UI добавляем порядковый префикс, но сохраняем внутри шаблона только имя
        blocks[i].querySelector('.op-header-input').value = `${i + 1}) ${sanitizeStrict(s.name, 200)}`;
        
        blocks[i].querySelector('.op-duration').value = Math.max(0, Number.parseFloat(s.dur) || 0);
        // Для всех операций кроме первой единица будет синхронизирована
        if (i === 0) {
            blocks[i].querySelector('.op-unit').value = s.unit;
        }

        if (s.hasBreak) {
            const breakGroup = blocks[i].querySelector('.break-container');
            try {
                if (breakGroup) breakGroup.style.display = 'flex';
            } catch (ee) {}
            blocks[i].querySelector('.op-break-val').value = Math.max(0, Number.parseFloat(s.breakVal) || 0);
            blocks[i].querySelector('.op-break-unit').value = s.breakUnit || 'min';
        }
    });
    
    // Синхронизируем единицы времени всех операций с первой
    syncTimeUnits();
}

function loadTechCards() {
    const userGroup = document.getElementById('userCards');
    userGroup.textContent = '';
    const keys = Object.keys(localStorage).filter(k => k.startsWith('z7_card_'));
    // Преобразуем в метки и сортируем с учётом числовых значений (чтобы '10' > '2')
    const mapped = keys.map(k => ({ key: k, label: k.replaceAll('z7_card_', '') }));
    mapped.sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true, sensitivity: 'base' }));
    mapped.forEach(({ key, label }) => {
        userGroup.append(createEl('option', { value: key }, label));
    });
    // Обновляем кастомный выпадающий список, если инициализирован
    if (globalThis._tcDropdown) globalThis._tcDropdown.refresh();
}

// === КАСТОМНЫЙ DROPDOWN ПОИСКА ТЕХКАРТ ===
(function initTcDropdown() {
    const input = document.getElementById('tcSearchInput');
    const list = document.getElementById('tcDropdownList');
    const hiddenSelect = document.getElementById('techCardSelect');
    if (!input || !list || !hiddenSelect) return;

    const DEFAULT_LABEL = '-- Ручной ввод --';
    let items = []; // {value, label, isDefault}
    let activeIndex = -1;
    let isOpen = false;

    function getCards() {
        const keys = Object.keys(localStorage).filter(k => k.startsWith('z7_card_'));
        const mapped = keys.map(k => ({ key: k, label: k.replaceAll('z7_card_', '') }));
        mapped.sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true, sensitivity: 'base' }));
        return mapped;
    }

    function buildItems(filter) {
        const cards = getCards();
        const q = (filter || '').toLowerCase().trim();
        const result = [{ value: 'manual', label: DEFAULT_LABEL, isDefault: true }];
        for (const c of cards) {
            if (!q || c.label.toLowerCase().includes(q)) {
                result.push({ value: c.key, label: c.label, isDefault: false });
            }
        }
        return result;
    }

    function renderList(filter) {
        items = buildItems(filter);
        list.textContent = '';
        activeIndex = -1;

        // Элемент по умолчанию
        const defItem = items[0];
        const defEl = document.createElement('div');
        defEl.className = 'tc-dropdown-item' + (hiddenSelect.value === 'manual' ? ' selected' : '');
        defEl.dataset.value = defItem.value;
        defEl.textContent = defItem.label;
        list.append(defEl);

        // Разделитель
        const userCards = items.filter(i => !i.isDefault);
        if (userCards.length > 0) {
            const sep = document.createElement('div');
            sep.className = 'tc-dropdown-separator';
            sep.textContent = 'Сохранённые';
            list.append(sep);
            for (const card of userCards) {
                const el = document.createElement('div');
                el.className = 'tc-dropdown-item' + (hiddenSelect.value === card.value ? ' selected' : '');
                el.dataset.value = card.value;
                el.textContent = card.label;
                list.append(el);
            }
        } else if (filter && filter.trim()) {
            const empty = document.createElement('div');
            empty.className = 'tc-dropdown-empty';
            empty.textContent = 'Ничего не найдено';
            list.append(empty);
        }
    }

    function openDropdown() {
        if (isOpen) return;
        isOpen = true;
        renderList(input.value === getDisplayLabel() ? '' : input.value);
        list.classList.add('open');
        // Выделяем весь текст для удобной замены
        input.select();
    }

    function closeDropdown() {
        if (!isOpen) return;
        isOpen = false;
        list.classList.remove('open');
        activeIndex = -1;
        // Восстанавливаем отображаемую метку
        input.value = getDisplayLabel();
    }

    function getDisplayLabel() {
        if (hiddenSelect.value === 'manual') return DEFAULT_LABEL;
        const opt = hiddenSelect.querySelector('option[value="' + CSS.escape(hiddenSelect.value) + '"]');
        return opt ? opt.textContent : DEFAULT_LABEL;
    }

    function selectItem(value) {
        hiddenSelect.value = value;
        // Генерируем событие change на скрытом select, чтобы другие части приложения могли отреагировать на изменение
        hiddenSelect.dispatchEvent(new Event('change'));
        input.value = getDisplayLabel();
        closeDropdown();
    }

    function scrollToActive() {
        const allItems = list.querySelectorAll('.tc-dropdown-item');
        if (activeIndex >= 0 && activeIndex < allItems.length) {
            allItems.forEach(el => el.classList.remove('active'));
            allItems[activeIndex].classList.add('active');
            allItems[activeIndex].scrollIntoView({ block: 'nearest' });
        }
    }

    // --- События ---
    input.addEventListener('focus', () => {
        openDropdown();
    });

    input.addEventListener('input', () => {
        if (!isOpen) openDropdown();
        renderList(input.value);
    });

    input.addEventListener('keydown', (e) => {
        if (!isOpen) {
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter') {
                openDropdown();
                e.preventDefault();
                return;
            }
            return;
        }
        const allItems = list.querySelectorAll('.tc-dropdown-item');
        const count = allItems.length;
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            activeIndex = (activeIndex + 1) % count;
            scrollToActive();
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            activeIndex = (activeIndex - 1 + count) % count;
            scrollToActive();
        } else if (e.key === 'Enter') {
            e.preventDefault();
            if (activeIndex >= 0 && activeIndex < count) {
                selectItem(allItems[activeIndex].dataset.value);
            }
        } else if (e.key === 'Escape') {
            e.preventDefault();
            closeDropdown();
            input.blur();
        }
    });

    list.addEventListener('mousedown', (e) => {
        // Предотвращаем срабатывание blur до click
        e.preventDefault();
        const item = e.target.closest('.tc-dropdown-item');
        if (item && item.dataset.value != null) {
            selectItem(item.dataset.value);
        }
    });

    input.addEventListener('blur', () => {
        // Небольшая задержка, чтобы mousedown на списке успел сработать
        setTimeout(() => {
            if (!list.matches(':hover')) {
                closeDropdown();
            }
        }, 150);
    });

    // Закрытие по клику вне области
    document.addEventListener('click', (e) => {
        const dropdown = document.getElementById('tcDropdown');
        if (isOpen && dropdown && !dropdown.contains(e.target)) {
            closeDropdown();
        }
    });

    // Публичный API
    const api = {
        refresh() {
            input.value = getDisplayLabel();
            if (isOpen) renderList(input.value === getDisplayLabel() ? '' : input.value);
        },
        setValue(val) {
            hiddenSelect.value = val;
            input.value = getDisplayLabel();
        },
        lock() {
            input.disabled = true;
            input.classList.add('locked-input');
            closeDropdown();
        },
        unlock() {
            input.disabled = false;
            input.classList.remove('locked-input');
        }
    };
    globalThis._tcDropdown = api;

    // Начальное отображение
    input.value = getDisplayLabel();
})();

// === ПРИВЯЗКА СОБЫТИЙ ===
document.getElementById('chainMode').addEventListener('change', async (e) => {
    const confirmed = await confirmAction(
        e.target.checked
            ? 'Включить режим "Цепочка"?\nВнимание!\nНевозможно будет выключить режим "Цепочка" при наличии ранее введенных данных.'
            : 'Выключить режим "Цепочка"?\nВнимание!\nНевозможно будет включить режим "Цепочка" при наличии ранее введенных данных.'
    );
    if (!confirmed) {
        e.target.checked = !e.target.checked;
        return;
    }
    updateStartTimeFromHistory();
    updateFirstPauseVisibility();
    updateTimeModeByChain();
});
const totalOpsEl = document.getElementById('totalOps');
if (totalOpsEl) {
    // При вводе: оставляем только цифры и ограничиваем максимумом сразу
    totalOpsEl.addEventListener('input', (e) => {
    let v = String(e.target.value).replaceAll(/[^0-9]/g, '');
        if (v !== '') {
            const n = Number.parseInt(v, 10);
            if (!Number.isNaN(n)) {
                const clamped = Math.max(1, Math.min(20, n));
                if (clamped !== n) v = String(clamped);
            }
        }
        e.target.value = v;
    });

    // Вставка: очистка и ограничение диапазона
    totalOpsEl.addEventListener('paste', (e) => {
        e.preventDefault();
        const text = e.clipboardData.getData('text') || '';
        const digits = text.replaceAll(/[^0-9]/g, '');
        const n = Number.parseInt(digits || '0', 10) || 0;
        const clamped = validateNumber(n, 1, 20);
        totalOpsEl.value = clamped;
        renderFields();
    });

    totalOpsEl.addEventListener('change', (e) => {
        // Ограничиваем допустимым диапазоном и перерисовываем
        const val = validateNumber(e.target.value, 1, 20);
        e.target.value = val;
        renderFields();
    });
    totalOpsEl.addEventListener('keyup', renderFields);
}
// Обработчик кнопки «ЗАДАТЬ» рядом с #totalOps: подтверждение и блокировка ввода
document.getElementById('generateBtn').addEventListener('click', generateTable);

document.getElementById('clearBtn').addEventListener('click', async () => {
    if (!await confirmAction('Очистить?')) return;

    // Сброс полей формы к значениям по умолчанию, но сохраняем историю И настройки исполнителей
    const defaults = getFormDefaults();

    if (document.getElementById('opsSortMode')) document.getElementById('opsSortMode').value = defaults.sortMode || 'sequential';
    try {
        // Если пользователь сохранил конфигурацию в localStorage, предпочитаем восстановить её для этих элементов
        let cfg = null;
        try { cfg = safeJsonParse(localStorage.getItem(CONFIG_KEY) || 'null'); } catch (ee) { cfg = null; }

        document.getElementById('totalOps').value = defaults.totalOps;
        // workerCount НЕ сбрасывается — сохраняется между очистками
        document.getElementById('startDate').value = defaults.startDate;
        try {
            // Сохраняем postingDate: восстанавливаем из сохранённой конфигурации если есть, иначе не перезаписываем текущее значение
            const pdEl = document.getElementById('postingDate');
            if (pdEl) {
                if (cfg && cfg.postingDate) {
                    pdEl.value = cfg.postingDate;
                }
            }
        } catch(e){}
        document.getElementById('startTime').value = defaults.startTime;
        if (!document.getElementById('chainMode').disabled) {
            document.getElementById('chainMode').checked = defaults.chainMode;
        }
        document.getElementById('lunchStart').value = (cfg && cfg.lunchStart) ? cfg.lunchStart : defaults.lunchStart;
        document.getElementById('lunchStart2').value = (cfg && cfg.lunchStart2) ? cfg.lunchStart2 : defaults.lunchStart2;
        document.getElementById('lunchDur').value = (cfg && cfg.lunchDur !== undefined) ? cfg.lunchDur : defaults.lunchDur;
        // Сброс timeMode к пользовательскому умолчанию (из Настроек)
        try { if (document.getElementById('timeMode')) document.getElementById('timeMode').value = defaults.timeMode; } catch(e) {}
        document.getElementById('resIz').value = defaults.resIz;
        document.getElementById('coefK').value = defaults.coefK;
        document.getElementById('orderName').value = defaults.orderName;
        document.getElementById('itemName').value = defaults.itemName;
        document.getElementById('statusBefore').value = defaults.statusBefore;
        document.getElementById('workExtra').value = defaults.workExtra;
        document.getElementById('devRec').value = defaults.devRec;
    } catch (e) {
        console.debug?.('clearBtn reset fields error:', e?.message);
    }

    // Разблокируем ввод totalOps и элементы техкарт, если они были заблокированы кнопкой «ЗАДАТЬ»
    unlockFormControls();
    // В режиме без цепочки — разблокировать workerCount через "Очистить"
    if (!document.getElementById('chainMode')?.checked) {
        const wcEl = document.getElementById('workerCount');
        if (wcEl) { wcEl.disabled = false; wcEl.classList.remove('locked-input'); wcEl.title = ''; }
    }

    // Очистка сгенерированных результатов и динамических полей
    container.textContent = '';
    const tableResult = document.getElementById('tableResult');
    const z7Result = document.getElementById('z7Result');
    if (tableResult) tableResult.textContent = '';
    if (z7Result) z7Result.textContent = '';

    // Сброс модальных окон и внутреннего состояния
    try {
        operationFirstId = '';
        lastOperationIndex = null;
        penultimateOperationIndex = null;
        // Перерисовываем списки модальных окон, если открыты
        const oModal = document.getElementById('opsModal');
        if (oModal && oModal.classList.contains('active')) renderOpsInputList();
        updateLunch2Label();
    } catch (e) {
        console.debug?.('clearBtn reset state error:', e?.message);
    }

    // Создаём заново один пустой блок операции
    renderFields();
});

// Обработчик деструктивной кнопки «Сброс»: очищает большую часть localStorage и сбрасывает поля к значениям по умолчанию
// Сохраняет: настройки (z7_defaults), техкарты (z7_card_*), историю (z7_history_session),
//            заметки исполнителей (z7_workers_cheat).
// Удаляет: временную конфигурацию обедов (z7_config) и другие данные.
document.getElementById('resetBtn').addEventListener('click', async () => {
    const msg = 'Сбросить все поля?\nИстория, техкарты, настройки и заметки исполнителей сохранятся.';
    if (!await confirmAction(msg)) return;

    const defaults = getFormDefaults();

    try {
        // Очищаем localStorage за исключением сохраняемых ключей
        const preservePrefixes = ['z7_card_', SESSION_DATA_PREFIX];
        const preserveKeys = new Set(['z7_history_session', 'z7_workers_cheat', DEFAULTS_KEY, SESSIONS_META_KEY, 'z7_active_session']);
        // Удаляем CONFIG_KEY и WORKERS_SESSION_KEY при сбросе
        const allKeys = Array.from(Object.keys(localStorage));
        for (const k of allKeys) {
            if (preserveKeys.has(k)) continue;
            if (preservePrefixes.some(p => k.startsWith(p))) continue;
            try { await safeLocalStorageRemove(k); } catch (e) { try { localStorage.removeItem(k); } catch (ee) {} }
        }

        document.getElementById('totalOps').value = defaults.totalOps;
        document.getElementById('workerCount').value = defaults.workerCount;
        document.getElementById('startDate').value = defaults.startDate;
        try { if (document.getElementById('postingDate')) document.getElementById('postingDate').value = defaults.postingDate; } catch(e){}
        document.getElementById('startTime').value = defaults.startTime;
        if (!document.getElementById('chainMode').disabled) {
            document.getElementById('chainMode').checked = defaults.chainMode;
        }
        document.getElementById('lunchStart').value = defaults.lunchStart;
        document.getElementById('lunchStart2').value = defaults.lunchStart2;
        document.getElementById('lunchDur').value = defaults.lunchDur;
        try { if (document.getElementById('timeMode')) document.getElementById('timeMode').value = defaults.timeMode; } catch(e){}
        try { if (document.getElementById('opsSortMode')) document.getElementById('opsSortMode').value = defaults.sortMode || 'sequential'; } catch(e){}
        document.getElementById('resIz').value = defaults.resIz;
        document.getElementById('coefK').value = defaults.coefK;
        document.getElementById('orderName').value = defaults.orderName;
        document.getElementById('itemName').value = defaults.itemName;
        document.getElementById('statusBefore').value = defaults.statusBefore;
        document.getElementById('workExtra').value = defaults.workExtra;
        document.getElementById('devRec').value = defaults.devRec;

        // Разблокируем ВСЕ элементы управления
        unlockAllFormControls();

        // Сброс данных исполнителей
        workerIds = [];

        container.textContent = '';
        const tableResult = document.getElementById('tableResult');
        const z7Result = document.getElementById('z7Result');
        if (tableResult) tableResult.textContent = '';
        if (z7Result) z7Result.textContent = '';

        // Перерисовываем пустые блоки операций
        try { renderFields(); } catch (e) {}

        lastOperationIndex = null;
        penultimateOperationIndex = null;
        updateLunch2Label();
        await showMessage('Сброс выполнен', 'Готово');
    } catch (e) {
        safeLogError('Reset error', e);
        await showMessage(String(e), 'Ошибка', 'error');
    }
});

document.getElementById('saveCardBtn').addEventListener('click', async () => {
    let name = null;
    
    // Tauri v2 не имеет встроенного prompt, используем fallback на globalThis.prompt
    // но оборачиваем в try-catch для безопасности
    try {
        name = globalThis.prompt("Название шаблона (техкарты):");
    } catch (e) {
        console.error('Prompt error:', e);
        return;
    }
    
    if (!name) return;

    // Строгая санитизация имени шаблона: ограничение по длине и очистка запрещённых символов
    name = sanitizeStrict(String(name), 100).trim();
    // Блокируем потенциально опасные имена ключей (prototype pollution и т.п.)
    if (name.length === 0 || name.includes('__proto__') || name.includes('constructor') || name.includes('prototype')) {
        await showMessage('Название не может быть пустым или содержать недопустимые последовательности', 'Ошибка', 'error');
        return;
    }

    // Проверяем, существует ли уже техкарта с таким именем
    const storageKey = 'z7_card_' + name;
    if (localStorage.getItem(storageKey) !== null) {
        const overwrite = await confirmAction(`Техкарта "${name}" уже существует.\nПерезаписать?`);
        if (!overwrite) return;
    }

    await safeLocalStorageSet(storageKey, JSON.stringify(getCardData()));
    loadTechCards();
    
    // Уведомление об успешном сохранении
    await showMessage(`Шаблон "${name}" сохранён`, 'Успешно');
});

document.getElementById('deleteCardBtn').addEventListener('click', async () => {
    const sel = document.getElementById('techCardSelect');
    if (sel.value === 'manual') return;

    if (await confirmAction('Удалить шаблон?')) {
        await safeLocalStorageRemove(sel.value);
        loadTechCards();
        sel.value = 'manual';
        if (globalThis._tcDropdown) globalThis._tcDropdown.setValue('manual');
    }
});

// === Модальное окно "Синтаксический Анализ" ===
(function initAnalyzeModal() {
    const modal = document.getElementById('analyzeModal');
    if (!modal) return;

    const closeBtn = document.getElementById('closeAnalyzeModal');
    const cancelBtn = document.getElementById('analyzeModalCancelBtn');
    const saveBtn = document.getElementById('analyzeModalSaveBtn');
    const nameInput = document.getElementById('analyzeCardName');
    const opsText = document.getElementById('analyzeOpsText');
    const unitSelect = document.getElementById('analyzeUnit');
    const multiplierSelect = document.getElementById('analyzeMultiplier');

    function openAnalyzeModal() {
        // Очищаем поля при открытии
        nameInput.value = '';
        opsText.value = '';
        unitSelect.value = 'min';
        multiplierSelect.value = '1';
        modal.classList.add('active');
        nameInput.focus();
    }

    function closeAnalyzeModal() {
        modal.classList.remove('active');
    }

    // Парсинг текста операций: каждая строка — "Название [промежуточное_число] длительность"
    // Автоматически удаляет порядковые префиксы ("1) ", "1. ", "1- " и т.п.)
    // Промежуточные числа между названием и длительностью отбрасываются
    // Длительность: до 5 цифр целой части, до 2 цифр дробной (разделитель . или ,)
    function parseOpsText(text) {
        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
        const ops = [];
        const errors = [];

        for (let i = 0; i < lines.length; i++) {
            let line = lines[i];

            // Удаляем порядковый префикс: "1) ", "2. ", "3- ", "10) " и т.п.
            line = line.replace(/^\s*\d+[).\-]\s*/, '');

            // Сворачиваем множественные пробелы в один
            line = line.replace(/\s{2,}/g, ' ').trim();

            if (line.length === 0) {
                errors.push(`Строка ${i + 1}: пустая строка после удаления префикса`);
                continue;
            }

            // Ищем последнее число в строке — это длительность
            // Длительность: до 5 цифр целой части, опционально разделитель (./,) и до 2 дробных
            // \S и \s взаимоисключающие — бэктрекинг невозможен
            const match = line.match(/^(\S+(?:\s+\S+)*?)\s+(\d{1,5}(?:[.,]\d{1,2})?)\s*$/);
            if (!match) {
                errors.push(`Строка ${i + 1}: не удалось распознать — "${lines[i]}"`);
                continue;
            }
            let opName = match[1].trim();
            const durStr = match[2].replaceAll(',', '.');
            const dur = Number.parseFloat(durStr);

            // Отбрасываем промежуточное число в конце названия (например, "Операция № 1 0" → "Операция № 1")
            // Промежуточное число: целое или дробное, отделённое пробелом от остальной части названия
            opName = opName.replace(/\s+\d+(?:[.,]\d+)?\s*$/, '').trim();

            if (!opName || opName.length === 0) {
                errors.push(`Строка ${i + 1}: пустое название операции`);
                continue;
            }
            if (opName.length > 200) {
                errors.push(`Строка ${i + 1}: название операции слишком длинное (макс. 200 символов)`);
                continue;
            }
            if (Number.isNaN(dur) || dur < 0) {
                errors.push(`Строка ${i + 1}: некорректная длительность — "${match[2]}"`);
                continue;
            }
            if (dur === 0) {
                errors.push(`Строка ${i + 1}: длительность не может быть 0`);
                continue;
            }
            ops.push({ name: sanitizeStrict(opName, 200), dur });
        }

        return { ops, errors };
    }

    // Кнопка "Сохранить"
    saveBtn.addEventListener('click', async () => {
        const cardName = sanitizeStrict(nameInput.value, 100).trim();
        if (!cardName || cardName.length === 0) {
            await showMessage('Введите название техкарты', 'Ошибка', 'error');
            nameInput.focus();
            return;
        }
        if (cardName.includes('__proto__') || cardName.includes('constructor') || cardName.includes('prototype')) {
            await showMessage('Название содержит недопустимые последовательности', 'Ошибка', 'error');
            return;
        }

        const rawOps = opsText.value.trim();
        if (!rawOps) {
            await showMessage('Введите операции и длительности', 'Ошибка', 'error');
            opsText.focus();
            return;
        }

        const { ops, errors } = parseOpsText(rawOps);

        if (errors.length > 0) {
            await showMessage('Ошибки парсинга:\n' + errors.join('\n'), 'Синтаксический Анализ', 'error');
            return;
        }

        if (ops.length === 0) {
            await showMessage('Не найдено ни одной операции', 'Ошибка', 'error');
            return;
        }

        if (ops.length > 20) {
            await showMessage('Максимум 20 операций', 'Ошибка', 'error');
            return;
        }

        const unit = unitSelect.value;
        const multiplier = Math.max(1, Math.min(10, Number.parseInt(multiplierSelect.value, 10) || 1));

        // Формируем массив шагов в формате техкарты (длительность × множитель)
        const steps = ops.map(op => ({
            name: op.name,
            dur: op.dur * multiplier,
            unit: unit,
            hasBreak: false,
            breakVal: 0,
            breakUnit: 'min'
        }));

        // Валидируем через существующую функцию
        if (!validateCardData(steps)) {
            await showMessage('Данные не прошли валидацию', 'Ошибка', 'error');
            return;
        }

        // Проверяем, существует ли уже техкарта с таким именем
        const storageKey = 'z7_card_' + cardName;
        if (localStorage.getItem(storageKey) !== null) {
            const overwrite = await confirmAction(`Техкарта "${cardName}" уже существует.\nПерезаписать?`);
            if (!overwrite) return;
        }

        // Сохраняем в localStorage
        await safeLocalStorageSet(storageKey, JSON.stringify(steps));
        loadTechCards();
        closeAnalyzeModal();
        await showMessage(`Техкарта "${cardName}" сохранена (${ops.length} операций)`, 'Успешно');
    });

    // Закрытие модалки
    closeBtn.addEventListener('click', closeAnalyzeModal);
    cancelBtn.addEventListener('click', closeAnalyzeModal);

    // Открытие по кнопке 🔍
    document.getElementById('analyzeCardBtn')?.addEventListener('click', openAnalyzeModal);
})();

// === Модальное окно "Настройки" ===
(function initSettingsModal() {
    const modal = document.getElementById('settingsModal');
    if (!modal) return;

    const closeBtn = document.getElementById('closeSettingsModal');
    const saveBtn = document.getElementById('settingsSaveBtn');
    const resetBtn = document.getElementById('settingsResetBtn');
    const cancelBtn = document.getElementById('settingsCancelBtn');

    // Элементы формы настроек
    const defTheme = document.getElementById('defTheme');
    const defChainMode = document.getElementById('defChainMode');
    const defTimeMode = document.getElementById('defTimeMode');
    const defStatusBefore = document.getElementById('defStatusBefore');
    const defWorkExtra = document.getElementById('defWorkExtra');
    const defDevRec = document.getElementById('defDevRec');
    const defSortMode = document.getElementById('defSortMode');
    const defColorLocked    = document.getElementById('defColorLocked');
    const defColorEditable  = document.getElementById('defColorEditable');
    const defColorHeader    = document.getElementById('defColorHeader');
    const defColorAuthor    = document.getElementById('defColorAuthor');
    const defColorConfirmed = document.getElementById('defColorConfirmed');
    const defColorPdtv      = document.getElementById('defColorPdtv');

    const previewLocked    = document.getElementById('previewLocked');
    const previewEditable  = document.getElementById('previewEditable');
    const previewHeader    = document.getElementById('previewHeader');
    const previewAuthor    = document.getElementById('previewAuthor');
    const previewConfirmed = document.getElementById('previewConfirmed');
    const previewPdtv      = document.getElementById('previewPdtv');

    function updateColorPreview(preview, hex) {
        const textColor = getContrastColor(hex);
        preview.style.backgroundColor = hex;
        preview.style.color = textColor;
        preview.textContent = 'Текст';
    }

    function populateFromStorage() {
        const d = getUserDefaults();
        defTheme.value = d.theme || 'light';
        defChainMode.checked = d.chainMode;
        defTimeMode.value = d.timeMode;
        defStatusBefore.value = d.statusBefore;
        defWorkExtra.value = d.workExtra;
        defDevRec.value = d.devRec;
        defSortMode.value = d.sortMode;
        const c = d.excelColors;
        defColorLocked.value    = c.locked;
        defColorEditable.value  = c.editable;
        defColorHeader.value    = c.header;
        defColorAuthor.value    = c.author;
        defColorConfirmed.value = c.confirmed;
        defColorPdtv.value      = c.pdtv ?? '#FFF9C4';
        updateColorPreview(previewLocked,    c.locked);
        updateColorPreview(previewEditable,  c.editable);
        updateColorPreview(previewHeader,    c.header);
        updateColorPreview(previewAuthor,    c.author);
        updateColorPreview(previewConfirmed, c.confirmed);
        updateColorPreview(previewPdtv,      c.pdtv ?? '#FFF9C4');
    }

    // Live-обновление превью при изменении пикеров
    defColorLocked.addEventListener('input',    () => updateColorPreview(previewLocked,    defColorLocked.value));
    defColorEditable.addEventListener('input',  () => updateColorPreview(previewEditable,  defColorEditable.value));
    defColorHeader.addEventListener('input',    () => updateColorPreview(previewHeader,    defColorHeader.value));
    defColorAuthor.addEventListener('input',    () => updateColorPreview(previewAuthor,    defColorAuthor.value));
    defColorConfirmed.addEventListener('input', () => updateColorPreview(previewConfirmed, defColorConfirmed.value));
    defColorPdtv.addEventListener('input',      () => updateColorPreview(previewPdtv,      defColorPdtv.value));

    // Мгновенный предпросмотр темы при переключении селектора
    defTheme.addEventListener('change', () => {
        applyTheme(defTheme.value);
    });

    function openSettingsModal() {
        populateFromStorage();
        modal.classList.add('active');
        defStatusBefore.focus();
    }

    function closeSettingsModal() {
        modal.classList.remove('active');
        // Откатываем предпросмотр темы к сохранённому значению
        const saved = getUserDefaults();
        applyTheme(saved.theme || 'light');
    }

    // Сохранить пользовательские умолчания
    saveBtn.addEventListener('click', async () => {
        const data = {
            theme: defTheme.value || 'light',
            chainMode: !!defChainMode.checked,
            timeMode: defTimeMode.value || 'total',
            statusBefore: sanitizeStrict(defStatusBefore.value || '', 300),
            workExtra: sanitizeStrict(defWorkExtra.value || '', 300),
            devRec: sanitizeStrict(defDevRec.value || '', 300),
            sortMode: defSortMode.value || 'sequential',
            excelColors: {
                locked:    defColorLocked.value,
                editable:  defColorEditable.value,
                header:    defColorHeader.value,
                author:    defColorAuthor.value,
                confirmed: defColorConfirmed.value,
                pdtv:      defColorPdtv.value
            }
        };
        // Применяем тему сразу
        applyTheme(data.theme);
        try {
            await safeLocalStorageSet(DEFAULTS_KEY, JSON.stringify(data));
            closeSettingsModal();
            await showMessage('Настройки сохранены', 'Готово');
        } catch (e) {
            console.error('Settings save error:', e);
            await showMessage('Ошибка сохранения настроек', 'Ошибка', 'error');
        }
    });

    // Сбросить к встроенным и удалить из localStorage
    resetBtn.addEventListener('click', async () => {
        if (!await confirmAction('Вернуть встроенные значения по умолчанию?')) return;
        try {
            await safeLocalStorageRemove(DEFAULTS_KEY);
            populateFromStorage(); // перечитает встроенные
            applyTheme('light'); // встроенная тема — светлая
            await showMessage('Настройки сброшены к встроенным', 'Готово');
        } catch (e) {
            console.error('Settings reset error:', e);
        }
    });

    closeBtn.addEventListener('click', closeSettingsModal);
    cancelBtn.addEventListener('click', closeSettingsModal);

    // Закрытие по Escape
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && modal.classList.contains('active')) {
            closeSettingsModal();
        }
    });

    // Открытие по кнопке "Настройки"
    document.getElementById('settingsBtn')?.addEventListener('click', openSettingsModal);
})();

document.getElementById('techCardSelect').addEventListener('change', (e) => {
    if (e.target.value !== 'manual') {
        try {
            const data = safeJsonParse(localStorage.getItem(e.target.value));
            if (data) {
                setCardData(data);
            }
            // Заполняем поле "Наименование" названием техкарты
            const cardName = e.target.value.replace(/^z7_card_/, '');
            const itemNameEl = document.getElementById('itemName');
            if (itemNameEl && cardName) {
                itemNameEl.value = cardName;
            }
        } catch (err) {
            console.error('Ошибка загрузки шаблона:', err);
        }
    }
});

document.getElementById('exportBtn').addEventListener('click', async () => {
    const obj = {};
    Object.keys(localStorage)
        .filter(k => k.startsWith('z7_card_'))
        .forEach(k => {
            obj[k] = localStorage.getItem(k);
        });
    // Примечание: `z7_workers_cheat` намеренно исключён из JSON-экспорта (локальные заметки остаются приватными)

    const jsonContent = JSON.stringify(obj, null, 2);
    const fileName = `z7_backup_${new Date().toISOString().slice(0, 10)}.json`;
    
    // Пробуем использовать Tauri API
    if (tauriDialog?.save && tauriInvoke) {
        try {
            const filePath = await tauriDialog.save({
                defaultPath: fileName,
                filters: [{ name: 'JSON', extensions: ['json'] }]
            });
            
            if (filePath) {
                await saveFileSecure(filePath, jsonContent);
                await showMessage('Файл успешно сохранён!', 'Успех');
            }
            return;
        } catch (e) {
            console.error('Ошибка сохранения:', e);
            await showMessage(String(e), 'Ошибка', 'error');
            return;
        }
    }
    
    // Запасной вариант — браузерный метод
    const a = document.createElement('a');
    const url = URL.createObjectURL(new Blob([jsonContent], { type: "application/json" }));
    a.href = url;
    a.download = fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});

document.getElementById('importBtn').addEventListener('click', () => {
    document.getElementById('fileInput').click();
});

document.getElementById('opsSortMode').addEventListener('change', (e) => {
    const container = document.getElementById('fieldsContainer');
    const blocks = Array.from(container.children);
    const sortMode = e.target.value;
    
    if (sortMode === 'confirmation') {
        blocks.sort((a, b) => (Number(a.dataset.opId) || 0) - (Number(b.dataset.opId) || 0));
    } else {
        blocks.sort((a, b) => {
            const idxA = Number(a.dataset.originalIndex) || 0;
            const idxB = Number(b.dataset.originalIndex) || 0;
            return idxA - idxB;
        });
    }
    
    blocks.forEach(b => container.appendChild(b));
    try { updateMainOperationLabels(); updateOperationInputPrefixes(); } catch (e) { /* ignore */ }
});

document.getElementById('fileInput').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    // Проверка размера файла (макс 1 МБ)
    const MAX_FILE_SIZE = 1024 * 1024;
    if (file.size > MAX_FILE_SIZE) {
        showMessage('Ошибка: файл слишком большой (макс. 1 МБ)').catch(() => {});
        e.target.value = '';
        return;
    }
    
    const reader = new FileReader();
    reader.onload = async (ev) => {
        try {
            const d = safeJsonParse(ev.target.result);
            if (!d || !validateImportData(d)) {
                showMessage('Ошибка: файл содержит некорректные данные').catch(() => {});
                return;
            }

            const importKeys = Object.keys(d).filter(k => k.startsWith('z7_card_'));
            if (importKeys.length === 0) {
                showMessage('Ошибка: файл не содержит техкарт').catch(() => {});
                return;
            }

            // Определяем конфликтующие техкарты (уже есть в localStorage)
            const existingKeys = Object.keys(localStorage).filter(k => k.startsWith('z7_card_'));
            const conflicts = importKeys.filter(k => existingKeys.includes(k));
            const newKeys = importKeys.filter(k => !existingKeys.includes(k));

            // mode: 'all' — перезаписать всё, 'new' — только новые, 'cancel' — отмена
            let mode = 'all';

            if (conflicts.length > 0) {
                mode = await showImportConflictDialog(conflicts);
                if (mode === 'cancel') return;
            }

            const keysToSave = mode === 'new' ? newKeys : importKeys;
            for (const k of keysToSave) {
                await safeLocalStorageSet(k, d[k]);
            }

            loadTechCards();
            const added = newKeys.length;
            const overwritten = mode === 'new' ? 0 : conflicts.length;
            const skipped = mode === 'new' ? conflicts.length : 0;
            const parts = [];
            if (added > 0) parts.push(`добавлено: ${added}`);
            if (overwritten > 0) parts.push(`перезаписано: ${overwritten}`);
            if (skipped > 0) parts.push(`пропущено: ${skipped}`);
            if (parts.length === 0) parts.push('без изменений');
            await showMessage(`Импорт завершён (${parts.join(', ')}).`, 'Готово');
        } catch (e) {
            showMessage("Ошибка при импорте: " + e.message).catch(() => {});
        }
    };
    reader.readAsText(file);
    e.target.value = ''; // Сброс input для повторного выбора того же файла
});

// === ИНИЦИАЛИЗАЦИЯ ===
loadTechCards();
loadWorkersSession(); // Восстановление кол-ва исполнителей, ID и статуса блокировки (ДО renderFields!)
renderFields();
setupExcelExport();
initSessionManager(); // Инициализирует сессии и вызывает restoreHistoryFromStorage() внутри
updateFirstPauseVisibility();

// === Режим Цепочка: принудительно включён и заблокирован при наличии истории ===
function updateChainCheckboxState() {
    try {
        const historyList = document.getElementById('historyList');
        const chainCheckbox = document.getElementById('chainMode');
        if (!historyList || !chainCheckbox) return;
        const hasHistory = historyList.children.length > 0;
        if (hasHistory) {
            // Восстанавливаем состояние чекбокса из первой записи истории
            const firstEntry = historyList.querySelector('.history-entry');
            if (firstEntry && firstEntry.dataset.jsonData) {
                try {
                    const data = safeJsonParse(firstEntry.dataset.jsonData);
                    if (data && typeof data.chain === 'boolean') {
                        chainCheckbox.checked = data.chain;
                    }
                } catch (_) { /* ignore parse errors */ }
            }
        }
        chainCheckbox.disabled = hasHistory;
        chainCheckbox.title = hasHistory ? 'Для разблокировки очистите историю или создайте новую.' : '';

        // Парсим данные первой записи один раз для обеда и workers
        let firstEntryData = null;
        {
            const fe = historyList.querySelector('.history-entry');
            if (fe && fe.dataset.jsonData) {
                try { firstEntryData = safeJsonParse(fe.dataset.jsonData); } catch (_) {}
            }
        }
        const isChainFromHistory = firstEntryData?.chain ?? false;

        // Блокировка/разблокировка полей обеда
        const lunchStartEl = document.getElementById('lunchStart');
        const lunchStart2El = document.getElementById('lunchStart2');
        const lunchDurEl = document.getElementById('lunchDur');
        if (hasHistory) {
            if (firstEntryData?.lunch) {
                const l = firstEntryData.lunch;
                if (lunchStartEl) lunchStartEl.value = String(l.h ?? 0).padStart(2, '0') + ':' + String(l.m ?? 0).padStart(2, '0');
                if (lunchStart2El) lunchStart2El.value = String(l.h2 ?? 0).padStart(2, '0') + ':' + String(l.m2 ?? 0).padStart(2, '0');
                if (lunchDurEl) lunchDurEl.value = l.dur ?? 0;
            }
            const lunchTip = 'Для разблокировки очистите историю или создайте новую.';
            if (lunchStartEl) { lunchStartEl.disabled = true; lunchStartEl.classList.add('locked-input'); lunchStartEl.title = lunchTip; }
            if (lunchStart2El) { lunchStart2El.disabled = true; lunchStart2El.classList.add('locked-input'); lunchStart2El.title = lunchTip; }
            if (lunchDurEl) { lunchDurEl.disabled = true; lunchDurEl.classList.add('locked-input'); lunchDurEl.title = lunchTip; }
        } else {
            if (lunchStartEl) { lunchStartEl.disabled = false; lunchStartEl.classList.remove('locked-input'); lunchStartEl.title = ''; }
            if (lunchStart2El) { lunchStart2El.disabled = false; lunchStart2El.classList.remove('locked-input'); lunchStart2El.title = ''; }
            if (lunchDurEl) { lunchDurEl.disabled = false; lunchDurEl.classList.remove('locked-input'); lunchDurEl.title = ''; }
        }

        // Блокировка/разблокировка workerCount
        const wcEl = document.getElementById('workerCount');
        if (hasHistory && isChainFromHistory) {
            // Цепочка + есть записи: восстанавливаем и блокируем
            if (firstEntryData?.workers) {
                if (wcEl) wcEl.value = Math.max(1, Math.min(10, firstEntryData.workers.count || 1));
                if (Array.isArray(firstEntryData.workers.ids)) workerIds = firstEntryData.workers.ids.slice();
            }
            if (wcEl) { wcEl.disabled = true; wcEl.classList.add('locked-input'); wcEl.title = 'Для разблокировки очистите историю или создайте новую.'; }
        } else if (!hasHistory) {
            // Нет записей — разблокируем в любом режиме
            if (wcEl) { wcEl.disabled = false; wcEl.classList.remove('locked-input'); wcEl.title = ''; }
        }
        // Без цепочки + есть записи — не трогаем текущее состояние workerCount

        updateStartTimeFromHistory();
        updateFirstPauseVisibility();
        updateTimeModeByChain();
    } catch (e) { console.debug?.('updateChainCheckboxState error:', e?.message); }
}

// MutationObserver: автообновление состояния чекбокса при изменении списка истории
try {
    const _histListEl = document.getElementById('historyList');
    if (_histListEl) {
        const _histObserver = new MutationObserver(() => updateChainCheckboxState());
        _histObserver.observe(_histListEl, { childList: true });
    }
} catch (e) { console.debug?.('historyList MutationObserver setup error:', e?.message); }

// Начальное состояние чекбокса при загрузке
updateChainCheckboxState();

// === СОХРАНЯЕМАЯ КОНФИГУРАЦИЯ (timeMode + настройки обеда) ===
const CONFIG_KEY = 'z7_config';

function updateLunch2Label() {
    const input = document.getElementById('lunchStart2');
    if (!input) return;
    const label = document.querySelector('label[for="lunchStart2"]');
    if (!label) return;
    
    if (input.value === '00:00') {
        label.style.textDecoration = 'line-through';
        label.style.opacity = '0.6';
    } else {
        label.style.textDecoration = 'none';
        label.style.opacity = '1';
    }
}

async function saveConfig() {
    try {
        const cfg = {
                lunchStart: document.getElementById('lunchStart')?.value || '12:00',
                lunchStart2: document.getElementById('lunchStart2')?.value || '00:00',
                lunchDur: document.getElementById('lunchDur')?.value || '45',
                postingDate: document.getElementById('postingDate')?.value || null
            };
        await safeLocalStorageSet(CONFIG_KEY, JSON.stringify(cfg));
    } catch (e) { console.debug?.('saveConfig error', e?.message); }
}

function loadConfig() {
    try {
        const raw = localStorage.getItem(CONFIG_KEY);
        if (!raw) return null;
        const cfg = safeJsonParse(raw);
        if (!cfg) return null;
        if (cfg.lunchStart && document.getElementById('lunchStart')) document.getElementById('lunchStart').value = cfg.lunchStart;
        if (cfg.lunchStart2 && document.getElementById('lunchStart2')) document.getElementById('lunchStart2').value = cfg.lunchStart2;
        if (cfg.lunchDur !== undefined && document.getElementById('lunchDur')) document.getElementById('lunchDur').value = cfg.lunchDur;
        if (cfg.postingDate && document.getElementById('postingDate')) document.getElementById('postingDate').value = cfg.postingDate;
        try { updateWorkerUIByTimeMode(); } catch (e) {}
        updateLunch2Label();
        return cfg;
    } catch (e) { console.debug?.('loadConfig error', e?.message); return null; }
}

// Подключаем автосохранение для этих элементов управления
try {
    const ids = ['lunchStart','lunchStart2','lunchDur','postingDate'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('change', saveConfig);
        el.addEventListener('input', saveConfig);
        if (id === 'lunchStart2') {
            el.addEventListener('input', updateLunch2Label);
            el.addEventListener('change', updateLunch2Label);
        }
    });
} catch (e) { console.debug?.('attach saveConfig listeners failed', e?.message); }

// Загружаем сохранённую конфигурацию (чтобы «Очистить»/перезагрузка восстанавливали эти значения)
loadConfig();
// Применяем пользовательские умолчания для полей, управляемых настройками, при каждой загрузке/обновлении
try {
    const _ud = getUserDefaults();
    const tEl = document.getElementById('timeMode');
    if (tEl) { tEl.value = _ud.timeMode || 'total'; updateWorkerUIByTimeMode(); }
    const cEl = document.getElementById('chainMode');
    if (cEl) cEl.checked = _ud.chainMode;
    const sbEl = document.getElementById('statusBefore');
    if (sbEl) sbEl.value = _ud.statusBefore;
    const weEl = document.getElementById('workExtra');
    if (weEl) weEl.value = _ud.workExtra;
    const drEl = document.getElementById('devRec');
    if (drEl) drEl.value = _ud.devRec;
    const smEl = document.getElementById('opsSortMode');
    if (smEl) smEl.value = _ud.sortMode || 'sequential';
} catch(e) { console.debug?.('apply user defaults error', e?.message); }
updateLunch2Label();


// === МОДАЛЬНОЕ ОКНО "О ПРОГРАММЕ" ===

const aboutTabConfig = [
    { tab: 'about',      btnId: 'aboutTabAbout',      bodyId: 'aboutModalBody',    file: 'about.md',        errMsg: 'Ошибка загрузки информации о программе.' },
    { tab: 'help',       btnId: 'aboutTabHelp',       bodyId: 'aboutHelpBody',     file: 'instruction.md',  errMsg: 'Ошибка загрузки инструкции.' },
    { tab: 'license',    btnId: 'aboutTabLicense',    bodyId: 'aboutLicenseBody',  file: 'license.md',       errMsg: 'Ошибка загрузки лицензии.' },
    { tab: 'licenseRu',  btnId: 'aboutTabLicenseRu',  bodyId: 'aboutLicenseRuBody', file: 'license(ru).md',  errMsg: 'Ошибка загрузки лицензии (рус).' },
];
const aboutTabCache = {};

async function loadAboutTabText(cfg) {
    if (aboutTabCache[cfg.tab]) return aboutTabCache[cfg.tab];
    try {
        const response = await fetch(cfg.file);
        if (!response.ok) throw new Error(`Failed to load ${cfg.file}`);
        aboutTabCache[cfg.tab] = await response.text();
        return aboutTabCache[cfg.tab];
    } catch (e) {
        console.error(`Error loading ${cfg.file}:`, e);
        return cfg.errMsg;
    }
}

// Backward-compat helpers used elsewhere in the file
async function loadAboutText() { return loadAboutTabText(aboutTabConfig[0]); }
async function loadInstructionText() { return loadAboutTabText(aboutTabConfig[1]); }

async function switchAboutTab(tab) {
    const cfg = aboutTabConfig.find(c => c.tab === tab);
    if (!cfg) return;
    for (const c of aboutTabConfig) {
        document.getElementById(c.bodyId).style.display = 'none';
        document.getElementById(c.btnId).classList.remove('about-tab--active');
    }
    const body = document.getElementById(cfg.bodyId);
    body.style.display = '';
    document.getElementById(cfg.btnId).classList.add('about-tab--active');
    const text = await loadAboutTabText(cfg);
    body.textContent = text;
}

for (const c of aboutTabConfig) {
    document.getElementById(c.btnId).addEventListener('click', () => switchAboutTab(c.tab));
}

document.getElementById('aboutBtn').addEventListener('click', async () => {
    const modal = document.getElementById('aboutModal');
    for (const c of aboutTabConfig) {
        const body = document.getElementById(c.bodyId);
        body.style.display = 'none';
        body.textContent = 'Загрузка...';
        document.getElementById(c.btnId).classList.remove('about-tab--active');
    }
    const firstCfg = aboutTabConfig[0];
    document.getElementById(firstCfg.btnId).classList.add('about-tab--active');
    document.getElementById(firstCfg.bodyId).style.display = '';
    modal.classList.add('active');
    const text = await loadAboutTabText(firstCfg);
    document.getElementById(firstCfg.bodyId).textContent = text;
});

document.getElementById('closeAboutModal').addEventListener('click', () => {
    document.getElementById('aboutModal').classList.remove('active');
});

// Закрытие по Escape
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        const modal = document.getElementById('aboutModal');
        if (modal && modal.classList.contains('active')) {
            modal.classList.remove('active');
        }
        const workersModal = document.getElementById('workersModal');
        if (workersModal && workersModal.classList.contains('active')) {
            workersModal.classList.remove('active');
        }
    }
});

// === МОДАЛЬНОЕ ОКНО НОМЕРОВ ПОДТВЕРЖДЕНИЯ ОПЕРАЦИЙ ===

// Возвращает номер подтверждения для операции (index начинается с 1)
// Логика: если операция отмечена как "последняя" - она пропускается в нумерации,
// остальные нумеруются последовательно, а пропущенная получает последний номер
function getOperationLabel(index, totalOps) {
    if (!operationFirstId || operationFirstId.trim() === '') {
        return String(index); // По умолчанию порядковый номер
    }

    const firstNum = Number.parseInt(operationFirstId, 10);
    if (Number.isNaN(firstNum)) return String(index);

    // Если эта операция отмечена как "последняя" - присваиваем ей последний номер
    if (lastOperationIndex !== null && index === lastOperationIndex) {
        const lastNum = firstNum + (totalOps - 1);
        return String(lastNum).padStart(10, '0');
    }

    // Если эта операция отмечена как "предпоследняя" - присваиваем ей предпоследний номер
    if (penultimateOperationIndex !== null && index === penultimateOperationIndex) {
        const penNum = firstNum + (totalOps - 2);
        return String(penNum).padStart(10, '0');
    }

    // Для остальных операций: считаем позицию без учёта "последней"
    let position = index;
    // Вычитаем 1 за каждую спец. операцию, которая находится перед текущей (по индексу)
    // Но проще считать последовательно, пропуская спец. индексы
    if (lastOperationIndex !== null && index > lastOperationIndex) {
        position = index - 1;
    }
    if (penultimateOperationIndex !== null && index > penultimateOperationIndex) {
        position = position - 1;
    }

    const opNum = firstNum + (position - 1);
    return String(opNum).padStart(10, '0');
}

// Обновляет текстовые метки с номерами операций в основной части (справа/слева)
function updateMainOperationLabels() {
    const blocks = document.querySelectorAll('.op-block');
    if (!blocks || blocks.length === 0) return;
    const total = Number.parseInt(document.getElementById('totalOps').value, 10) || blocks.length;
    blocks.forEach((blk, i) => {
        const lbl = blk.querySelector('.op-num-label');
        if (lbl) {
            try {
                if (blk.dataset.opId) {
                    lbl.textContent = blk.dataset.opId;
                } else {
                    lbl.textContent = getOperationLabel(i + 1, total);
                }
            } catch (e) {
                // Безопасность: не ломаем UI, если getOperationLabel завершится с ошибкой
                lbl.textContent = String(i + 1);
            }
        }
    });
}

// Обновляет префиксы в полях ввода операций (например после изменения количества)
function updateOperationInputPrefixes() {
    const blocks = document.querySelectorAll('.op-block');
    if (!blocks) return;
    blocks.forEach((blk, i) => {
        const inp = blk.querySelector('.op-header-input');
        if (!inp) return;
        const idx = blk.dataset.originalIndex ? blk.dataset.originalIndex : (i + 1);
        const prefix = `${idx}) `;
        const body = stripOrdinalPrefix(inp.value || '');
        inp.value = prefix + sanitizeStrict(body, 200);
    });
}

function renderOpsInputList() {
    const container = document.getElementById('opsInputList');
    const count = Number.parseInt(document.getElementById('totalOps').value, 10) || 1;
    container.replaceChildren();

    // Сброс индексов, если они выходят за пределы количества операций
    if (lastOperationIndex !== null && lastOperationIndex > count) {
        lastOperationIndex = null;
        penultimateOperationIndex = null;
    }
    if (penultimateOperationIndex !== null && penultimateOperationIndex > count) penultimateOperationIndex = null;

    // Получаем названия операций из полей ввода
    // Создаём map по originalIndex для корректного порядка независимо от сортировки в основном окне
    const opBlocks = document.querySelectorAll('.op-block');
    const opBlocksByIndex = new Map();
    opBlocks.forEach(b => {
        const idx = Number.parseInt(b.dataset.originalIndex, 10);
        if (!Number.isNaN(idx)) opBlocksByIndex.set(idx, b);
    });

    for (let i = 1; i <= count; i++) {
        const row = createEl('div', { className: 'op-input-row' });

        // Берём название операции из блока с соответствующим порядковым номером
        let opName = `Операция ${i}`;
        const block = opBlocksByIndex.get(i);
        if (block) {
            const nameInput = block.querySelector('.op-header-input');
            if (nameInput && nameInput.value.trim()) {
                // Показываем полное значение поля, включая префикс (например "1) ...")
                opName = nameInput.value.trim();
            }
        }

        const label = createEl('label', { className: 'op-label', htmlFor: `op_id_${i}` }, `${opName}:`);

        // Для первой операции - редактируемый input, для остальных - disabled или enabled в зависимости от состояния "авто"
        const isFirst = (i === 1);
        const input = createEl('input', {
            type: 'text',
            id: `op_id_${i}`,
            name: `op_id_${i}`,
            maxLength: '10',
            placeholder: isFirst ? '0000000000' : '№ ПДТВ',
            autocomplete: 'off'
        });

        if (isFirst) {
            if (block && block.dataset.opId) {
                input.value = block.dataset.opId;
                operationFirstId = block.dataset.opId;
            } else {
                input.value = operationFirstId || '';
            }
            // Разрешаем только цифры
            input.addEventListener('input', (e) => {
                e.target.value = e.target.value.replaceAll(/[^0-9]/g, '').substring(0, 10);
                updateOpsCalculatedValues();
            });

            // Создаем контейнер для чекбокса "авто", чтобы выровнять с чекбоксами "последняя"
            const autoCheckboxWrapper = createEl('div', { className: 'op-checkbox-wrapper' });
            const autoCheckbox = createEl('input', {
                type: 'checkbox',
                id: 'op_auto_checkbox',
                name: 'op_auto'
            });
            autoCheckbox.checked = autoIncrementEnabled;  // Используем сохраненное состояние

            autoCheckbox.addEventListener('change', (e) => {
                const isChecked = e.target.checked;

                // Запрет снятия "авто" пока есть "последняя" или "предпоследняя"
                if (!isChecked && (lastOperationIndex !== null || penultimateOperationIndex !== null)) {
                    e.target.checked = true;
                    showMessage('Сначала снимите галочки «последняя» и «предпоследняя»', 'Внимание', 'warning');
                    return;
                }

                // Обновляем состояние переменной
                autoIncrementEnabled = isChecked;

                // Обновляем состояния полей ввода и чекбоксов "последняя"
                for (let j = 2; j <= count; j++) {
                    const inputField = document.getElementById(`op_id_${j}`);
                    const lastCheckbox = document.getElementById(`op_special_${j}`);

                    if (inputField) {
                        inputField.disabled = isChecked;
                    }
                    if (lastCheckbox) {
                        lastCheckbox.disabled = !isChecked;
                        // Снимаем галочку при отключении "авто"
                        if (!isChecked) {
                            lastCheckbox.checked = false;
                        }
                    }
                }

                // Обновляем значения полей ввода
                updateOpsCalculatedValues();
            });

            const autoCheckboxLabel = createEl('label', { htmlFor: 'op_auto_checkbox' }, 'авто');
            autoCheckboxWrapper.append(autoCheckbox, autoCheckboxLabel);

            // Добавляем элемент в строку после input, чтобы чекбокс "авто" был правее поля ввода
            row.append(label, input, autoCheckboxWrapper);
        } else {
            // Для остальных операций определяем, нужно ли разблокировать поле ввода
            if (autoIncrementEnabled) {
                // Если автоподсчет включен, поле заблокировано
                input.disabled = true;
            } else {
                // Если автоподсчет выключен, поле разблокировано
                input.disabled = false;
            }

            // Рассчитываем значение с учётом галочки "последняя"
            if (block && block.dataset.opId) {
                input.value = block.dataset.opId;
            } else if (operationFirstId && operationFirstId.trim()) {
                const firstNum = Number.parseInt(operationFirstId, 10);
                if (!Number.isNaN(firstNum)) {
                    // Используем общую логику расчета
                    input.value = getOperationLabel(i, count);
                }
            }

            // Логика чекбоксов "последняя" / "предпоследняя"
            const checkboxWrapper = createEl('div', {
                className: 'op-checkbox-wrapper'
            });
            
            const checkbox = createEl('input', {
                type: 'checkbox',
                id: `op_special_${i}`,
                name: 'op_special'
            });
            checkbox.disabled = !autoIncrementEnabled; // Чекбокс "последняя" доступен только при включенном "авто"
            
            let labelText = 'последняя ';

            if (lastOperationIndex !== null) {
                if (i === lastOperationIndex) {
                    // Это выбранная последняя операция
                    checkbox.checked = true;
                    labelText = 'последняя ';
                    checkbox.addEventListener('change', (e) => {
                        // Запрет снятия "последняя" пока установлена "предпоследняя"
                        if (!e.target.checked && penultimateOperationIndex !== null) {
                            e.target.checked = true;
                            showMessage('Сначала снимите галочку «предпоследняя»', 'Внимание', 'warning');
                            return;
                        }
                        // Снятие галочки "последняя"
                        lastOperationIndex = null;
                        penultimateOperationIndex = null;
                        renderOpsInputList();
                        updateOpsCalculatedValues();
                    });
                } else {
                    // Остальные становятся "предпоследняя"
                    checkbox.checked = (i === penultimateOperationIndex);
                    labelText = 'предпоследняя';
                    checkbox.addEventListener('change', (e) => {
                        if (e.target.checked) {
                            penultimateOperationIndex = i;
                        } else {
                            penultimateOperationIndex = null;
                        }
                        renderOpsInputList();
                        updateOpsCalculatedValues();
                    });
                }
            } else {
                // Ни одна операция не выбрана как последняя
                checkbox.checked = false;
                labelText = 'последняя';
                checkbox.addEventListener('change', (e) => {
                    if (e.target.checked) {
                        lastOperationIndex = i;
                        penultimateOperationIndex = null;
                    }
                    renderOpsInputList();
                    updateOpsCalculatedValues();
                });
            }

            const checkboxLabel = createEl('label', { htmlFor: `op_special_${i}` }, labelText);
            checkboxWrapper.append(checkbox, checkboxLabel);

            row.append(label, input, checkboxWrapper);
        }
        
        container.append(row);
    }
}

function updateOpsCalculatedValues() {
    const firstInput = document.getElementById('op_id_1');
    if (!firstInput) return;

    const firstVal = firstInput.value.trim();
    operationFirstId = firstVal;
    const count = Number.parseInt(document.getElementById('totalOps').value, 10) || 1;

    // Проверяем, включена ли функция "авто"
    const isAutoEnabled = autoIncrementEnabled;

    for (let i = 2; i <= count; i++) {
        const input = document.getElementById(`op_id_${i}`);
        if (input) {
            // Если "авто" включено, автоматически рассчитываем номера
            if (isAutoEnabled) {
                if (firstVal && firstVal.length > 0) {
                    const firstNum = Number.parseInt(firstVal, 10);
                    if (!Number.isNaN(firstNum)) {
                        input.value = getOperationLabel(i, count);
                    } else {
                        input.value = '';
                    }
                } else {
                    input.value = '';
                }
                
                // Если "авто" включено, поле должно быть заблокировано
                input.disabled = true;
            } else {
                // Если "авто" выключено, оставляем значение как есть (пользователь может ввести вручную)
                // Но если поле было заблокировано ранее, разблокируем его
                input.disabled = false;
            }
        }
    }
    // Обновляем метки в основной части, чтобы изменения в модальном окне были видны сразу
    try { updateMainOperationLabels(); } catch (e) { /* ignore */ }
}

function saveOperationIds() {
    // Сохраняем ID для каждого блока по originalIndex
    const blocks = Array.from(document.querySelectorAll('.op-block'));
    const blocksByIndex = new Map();
    blocks.forEach(b => {
        const idx = Number.parseInt(b.dataset.originalIndex, 10);
        if (!Number.isNaN(idx)) blocksByIndex.set(idx, b);
    });
    const count = blocks.length;
    for (let i = 1; i <= count; i++) {
        const input = document.getElementById(`op_id_${i}`);
        const block = blocksByIndex.get(i);
        if (input && block) {
            let val = input.value.trim();
            if (val && val.length > 0 && val.length < 10) val = val.padStart(10, '0');
            block.dataset.opId = val;
        }
    }

    // Сохраняем метки "последняя" и "предпоследняя" по originalIndex
    blocks.forEach(b => {
        const origIdx = Number.parseInt(b.dataset.originalIndex, 10);
        if (origIdx === lastOperationIndex) {
            b.dataset.isLast = "true";
        } else {
            delete b.dataset.isLast;
        }
        if (origIdx === penultimateOperationIndex) {
            b.dataset.isPenultimate = "true";
        } else {
            delete b.dataset.isPenultimate;
        }
    });

    // Сохраняем состояние чекбокса "авто"
    const autoCheckbox = document.getElementById('op_auto_checkbox');
    if (autoCheckbox) {
        autoIncrementEnabled = autoCheckbox.checked;
    }

    // Сортировка блоков
    const sortMode = document.getElementById('opsSortMode').value;
    if (sortMode === 'confirmation') {
        blocks.sort((a, b) => {
            const idA = Number(a.dataset.opId) || 0;
            const idB = Number(b.dataset.opId) || 0;
            return idA - idB;
        });
    } else {
        // Последовательный
        blocks.sort((a, b) => {
            const idxA = Number(a.dataset.originalIndex) || 0;
            const idxB = Number(b.dataset.originalIndex) || 0;
            return idxA - idxB;
        });
    }
    const container = document.getElementById('fieldsContainer');
    blocks.forEach(b => container.appendChild(b));

    // Восстанавливаем индексы по originalIndex (а не по позиции в DOM)
    lastOperationIndex = null;
    penultimateOperationIndex = null;
    blocks.forEach(b => {
        const origIdx = Number.parseInt(b.dataset.originalIndex, 10);
        if (b.dataset.isLast === "true") {
            lastOperationIndex = origIdx;
        }
        if (b.dataset.isPenultimate === "true") {
            penultimateOperationIndex = origIdx;
        }
    });

    // Обновляем operationFirstId на основе блока с originalIndex=1
    const firstBlock = blocksByIndex.get(1);
    if (firstBlock && firstBlock.dataset.opId) {
        operationFirstId = firstBlock.dataset.opId;
    }

    document.getElementById('opsModal').classList.remove('active');

    // Перерисовываем поля операций с учётом "последней операции"
    renderFields();
    // Обновляем метки операций в основной части
    try { updateMainOperationLabels(); } catch (e) { /* ignore */ }
}

function resetOperationIds() {
    operationFirstId = '';
    lastOperationIndex = null;
    penultimateOperationIndex = null;
    autoIncrementEnabled = false; // Сбрасываем состояние чекбокса "авто"
    if (document.getElementById('opsSortMode')) document.getElementById('opsSortMode').value = 'sequential';
    
    // Сброс сортировки и ID
    const blocks = Array.from(document.querySelectorAll('.op-block'));
    blocks.sort((a, b) => (Number(a.dataset.originalIndex) || 0) - (Number(b.dataset.originalIndex) || 0));
    const container = document.getElementById('fieldsContainer');
    blocks.forEach(b => { delete b.dataset.opId; delete b.dataset.isLast; delete b.dataset.isPenultimate; container.appendChild(b); });

    renderOpsInputList();
    try { updateMainOperationLabels(); } catch (e) { /* ignore */ }
}

document.getElementById('setOpsBtn').addEventListener('click', async () => {
    const totalEl = document.getElementById('totalOps');
    if (!totalEl) return;

    // Если totalOps ещё не заблокирован, сначала запрашиваем подтверждение
    if (!totalEl.disabled) {
        const msg = 'Вы уверены? Количество операций нельзя будет изменить.\nРазблокировка кнопкой "Очистить" или F5.';
        if (!await confirmAction(msg)) return;

        // Блокируем ввод и визуально помечаем
        const lockTip = 'Нажмите «Очистить» (F5) или «Сброс» для разблокировки';
        totalEl.disabled = true;
        totalEl.classList.add('locked-input');
        totalEl.title = lockTip;
        try { renderFields(); } catch (e) { console.debug?.('renderFields after setOps lock failed', e?.message); }

        // Также отключаем выбор техкарты и кнопки Сохранить/Удалить
        try {
            const sel = document.getElementById('techCardSelect');
            if (sel) { sel.disabled = true; sel.classList.add('locked-input'); sel.title = lockTip; }
            // Блокируем поле поиска кастомного выпадающего списка
            if (globalThis._tcDropdown) globalThis._tcDropdown.lock();
            const saveBtn = document.getElementById('saveCardBtn');
            if (saveBtn) { saveBtn.disabled = true; saveBtn.classList.add('locked-control'); saveBtn.title = lockTip; }
            const delBtn = document.getElementById('deleteCardBtn');
            if (delBtn) { delBtn.disabled = true; delBtn.classList.add('locked-control'); delBtn.title = lockTip; }
            const analyzeBtn = document.getElementById('analyzeCardBtn');
            if (analyzeBtn) { analyzeBtn.disabled = true; analyzeBtn.classList.add('locked-control'); analyzeBtn.title = lockTip; }
        } catch (e) { console.debug?.('lock tech card controls failed', e?.message); }
    }

    renderOpsInputList();
    document.getElementById('opsModal').classList.add('active');
});

document.getElementById('closeOpsModal').addEventListener('click', () => {
    document.getElementById('opsModal').classList.remove('active');
});

document.getElementById('saveOpsBtn').addEventListener('click', saveOperationIds);
document.getElementById('resetOpsBtn').addEventListener('click', resetOperationIds);

// При изменении количества операций обновляем модальное окно (если открыто)
document.getElementById('totalOps').addEventListener('change', () => {
    const modal = document.getElementById('opsModal');
    if (modal && modal.classList.contains('active')) {
        renderOpsInputList();
    }
});

// === МОДАЛЬНОЕ ОКНО НОМЕРОВ ИСПОЛНИТЕЛЕЙ ===
// workerIds объявлен выше для корректной работы loadWorkersSession()

function getWorkerLabel(index) {
    // index начинается с 1
    if (workerIds[index - 1] && workerIds[index - 1].trim()) {
        // Извлекаем только цифры из строки, если они есть, и используем их как номер исполнителя
        const raw = String(workerIds[index - 1]).trim();
        const digits = raw.replaceAll(/[^0-9]/g, '');
        if (digits.length === 0) return String(index);
        return digits.length >= 8 ? digits : digits.padStart(8, '0');
    }
    return String(index); // По умолчанию порядковый номер
}

function renderWorkersInputList() {
    const container = document.getElementById('workersInputList');
    const count = Number.parseInt(document.getElementById('workerCount').value, 10) || 1;
    container.replaceChildren();
    
    for (let i = 1; i <= count; i++) {
        const row = createEl('div', { className: 'worker-input-row' });
        const label = createEl('label', { htmlFor: `worker_id_${i}` }, `Исполнитель ${i}:`);
        const input = createEl('input', {
            type: 'text',
            id: `worker_id_${i}`,
            name: `worker_id_${i}`,
            maxLength: '8',
            placeholder: '00000000',
            pattern: '[0-9]{8}',
            autocomplete: 'off'
        });
        input.value = workerIds[i - 1] || '';
        input.dataset.workerIndex = i - 1;
        
        // Разрешаем только цифры
        input.addEventListener('input', (e) => {
            // заменить все не-цифры (используем replaceAll с глобальным regex)
            e.target.value = e.target.value.replaceAll(/[^0-9]/g, '').substring(0, 8);
        });
        
        row.append(label, input);
        container.append(row);
    }
    // При каждом рендере обновляем состояние поля для шпаргалки (если оно есть), чтобы отразить сохранённое значение
    try {
        const cheatEl = document.getElementById('workersCheat');
        const editBtn = document.getElementById('editWorkersBtn');
        if (cheatEl) {
            const saved = localStorage.getItem('z7_workers_cheat') || '';
            cheatEl.value = saved;
            // При открытии модального окна по умолчанию блокируем редактирование шпаргалки, чтобы избежать случайных изменений
            cheatEl.disabled = true;
        }
        if (editBtn) editBtn.textContent = 'Изменить';
    } catch (e) {
        console.debug?.('renderWorkersInputList cheat load error:', e?.message);
    }
}

async function saveWorkerIds() {
    const inputs = document.querySelectorAll('#workersInputList input');
    workerIds = [];
    inputs.forEach((input, idx) => {
        const val = input.value.trim();
        // Если номер введён, проверяем что он 8-значный
        if (val && val.length === 8) {
            workerIds[idx] = val;
        } else if (val && val.length > 0 && val.length < 8) {
            // Дополняем нулями слева до 8 цифр
            workerIds[idx] = val.padStart(8, '0');
        } else {
            workerIds[idx] = '';
        }
    });
    // Сохраняем сессию исполнителей в localStorage
    saveWorkersSession();
    // Обновляем отображение в основной части
    document.getElementById('workersModal').classList.remove('active');
}

function resetWorkerIds() {
    workerIds = [];
    renderWorkersInputList();
}

document.getElementById('setWorkersBtn').addEventListener('click', async () => {
    const wcEl = document.getElementById('workerCount');
    if (!wcEl) return;

    // Если workerCount ещё не заблокирован, сначала запрашиваем подтверждение
    if (!wcEl.disabled) {
        const isChain = document.getElementById('chainMode')?.checked;
        const msg = isChain
            ? 'Вы уверены? Количество исполнителей нельзя будет изменить, пока в истории есть записи.'
            : 'Вы уверены? Количество исполнителей нельзя будет изменить.\nРазблокировка кнопкой "Очистить" (F5) или "Сброс".';
        if (!await confirmAction(msg)) return;

        // Блокируем ввод и визуально помечаем
        wcEl.disabled = true;
        wcEl.classList.add('locked-input');
        wcEl.title = isChain ? 'Для разблокировки очистите историю или создайте новую.' : 'Нажмите "Очистить" (F5) или "Сброс" для разблокировки';
        try { renderFields(); } catch (e) { console.debug?.('renderFields after setWorkers lock failed', e?.message); }
        // Сохраняем состояние блокировки
        saveWorkersSession();
    }

    // При открытии модального окна рендерим список полей для ввода номеров исполнителей в соответствии с текущим количеством
    renderWorkersInputList();
    document.getElementById('workersModal').classList.add('active');
});

document.getElementById('closeWorkersModal').addEventListener('click', () => {
    document.getElementById('workersModal').classList.remove('active');
});

document.getElementById('saveWorkersBtn').addEventListener('click', saveWorkerIds);
document.getElementById('resetWorkersBtn').addEventListener('click', resetWorkerIds);
// При изменении количества исполнителей обновляем модальное окно (если открыто)

// Кнопка для разблокировки поля редактирования шпаргалки и сохранения её значения при повторном нажатии
document.getElementById('editWorkersBtn').addEventListener('click', async (e) => {
    try {
        const cheatEl = document.getElementById('workersCheat');
        const btn = e.target;
        if (!cheatEl || !btn) return;
        if (cheatEl.disabled) {
            // разблокируем редактирование, пользователь может внести изменения в шпаргалку, а при повторном нажатии она сохранится автоматически
            cheatEl.disabled = false;
            cheatEl.focus();
            btn.textContent = 'Готово';
        } else {
            // сохраняем значение шпаргалки и блокируем редактирование
            cheatEl.disabled = true;
            btn.textContent = 'Изменить';
            try {
                const safeText = sanitizeInput(cheatEl.value || '', 5000);
                await safeLocalStorageSet('z7_workers_cheat', safeText);
                showMessage('Шпаргалка сохранена', 'Инфо');
            } catch (saveErr) {
                console.error('Auto-save workersCheat error:', saveErr);
            }
        }
    } catch (err) {
        console.error('editWorkersBtn toggle error:', err);
    }
});

// При изменении количества исполнителей обновляем модальное окно (если открыто)
const workerCountEl = document.getElementById('workerCount');
if (workerCountEl) {
    // Input: разрешаем ввод только цифр и ограничиваем длину, а также сразу же применяем эти ограничения при вставке текста
    workerCountEl.addEventListener('input', (e) => {
        let v = String(e.target.value).replaceAll(/[^0-9]/g, '');
        if (v !== '') {
            const n = Number.parseInt(v, 10);
            if (!Number.isNaN(n)) {
                const clamped = Math.max(1, Math.min(10, n));
                if (clamped !== n) v = String(clamped);
            }
        }
        e.target.value = v;
    });

    // При вставке текста в поле количества исполнителей извлекаем из вставляемого текста только цифры, а также применяем ограничения на количество исполнителей
    workerCountEl.addEventListener('paste', (e) => {
        e.preventDefault();
        const text = e.clipboardData.getData('text') || '';
        const digits = text.replaceAll(/[^0-9]/g, '');
        const n = Number.parseInt(digits || '0', 10) || 0;
        const clamped = validateNumber(n, 1, 10);
        workerCountEl.value = clamped;
        const modal = document.getElementById('workersModal');
        if (modal && modal.classList.contains('active')) {
            renderWorkersInputList();
        }
    });

    workerCountEl.addEventListener('change', (e) => {
        const val = validateNumber(e.target.value, 1, 10);
        e.target.value = val;
        const modal = document.getElementById('workersModal');
        if (modal && modal.classList.contains('active')) {
            renderWorkersInputList();
        }
        // Синхронизируем количество чекбоксов исполнителей в блоках операций с новым количеством исполнителей
        try { syncOpWorkersToCount(); } catch (ee) {}
    });
}

function syncOpWorkersToCount() {
    const count = Number.parseInt(document.getElementById('workerCount')?.value || '1', 10) || 1;
    const blocks = Array.from(document.querySelectorAll('.op-block'));
    blocks.forEach((block, idx) => {
        const box = block.querySelector('.op-workers-box');
        if (!box) return;
        const existing = Array.from(box.querySelectorAll('.op-worker-item'));
        const cur = existing.length;
        if (cur < count) {
            for (let w = cur + 1; w <= count; w++) {
                const id = `op_${idx+1}_worker_${w}`;
                const cb = createEl('input', { type: 'checkbox', className: 'op-worker-checkbox', id, 'data-worker': String(w) });
                cb.checked = true;
                const lbl = createEl('label', { htmlFor: id, className: 'op-worker-label' }, String(w));
                const wrapper = createEl('span', { className: 'op-worker-item' });
                wrapper.append(cb, lbl);
                box.append(wrapper);
                cb.addEventListener('change', () => {
                   updateWorkerChain();
                });
            }
        } else if (cur > count) {
            // Удаляем лишние чекбоксы исполнителей, если их стало больше, чем нужно
            for (let i = cur; i > count; i--) {
                const item = existing[i-1];
                if (item) box.removeChild(item);
            }
        }
    });
    // После синхронизации количества чекбоксов исполнителей в блоках операций обновляем цепочку исполнителей для всех операций, чтобы отразить изменения в количестве исполнителей
    updateWorkerChain();
}

