// MenuWebApp.gs

function debugAlwaysByDay() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const menus = MenuManager.findMenuSheets();
  const data  = menus[0].sheet.getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);
  const days  = _buildWeekDays(data, _getTodayDayIndex(), today);
  Logger.log('Дней: ' + days.length);
  days.forEach(d => Logger.log(d.dayName + ': alwaysOn=' + d.alwaysOn.length + ' → ' + d.alwaysOn.join(' | ')));
}

function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'app';

  if (mode === 'json') {
    const today = getTodayMenuData();
    const week  = getWeekMenuData();
    return ContentService
      .createTextOutput(JSON.stringify({ today: today, week: week }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return htmlTemplate_('MenuDisplay')
    .evaluate()
    .setTitle(L('Меню сегодня'))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveMenuDataToStorage(sheet, nextWeek) {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const officeName = getOfficeName();
    const data       = sheet.getDataRange().getValues();
    const today      = new Date(); today.setHours(0,0,0,0);
    const todayIdx   = _getTodayDayIndex();

    let weekLabel = '';
    if (data[0] && data[0][0]) {
      const h = data[0][0].toString();
      const m = h.match(/(?:с|from)\s+(\d+\.\d+)\s+(?:по|to)\s+(\d+\.\d+)/i);
      if (m) weekLabel = L('с ') + m[1] + L(' по ') + m[2];
    }
    if (!weekLabel) {
      const nm = sheet.getName().match(/(\d{2}\.\d{2})-(\d{2}\.\d{2})/);
      if (nm) weekLabel = nm[1] + ' - ' + nm[2];
    }

    const weekDays = _buildWeekDays(data, todayIdx, today);
    const toStore  = { days: weekDays, officeName, weekLabel };
    const json     = JSON.stringify(toStore);

    if (nextWeek) {
      PropertiesService.getScriptProperties().setProperty('NEXT_MENU_DATA', json);
      Logger.log('✅ NEXT_MENU_DATA сохранён (' + json.length + ' байт)');
    } else {
      PropertiesService.getScriptProperties().setProperty('CURRENT_MENU_DATA', json);
      PropertiesService.getScriptProperties().deleteProperty('NEXT_MENU_DATA');
      Logger.log('✅ CURRENT_MENU_DATA сохранён (' + json.length + ' байт)');
    }

  } catch(e) {
    Logger.log('❌ saveMenuDataToStorage: ' + e.toString());
  }
}

function forceSaveMenuDataFromSheets() {
  const menus = MenuManager.findMenuSheets();
  if (menus.length === 0) {
    SpreadsheetApp.getUi().alert(L('❌ Листы меню не найдены'));
    return;
  }

  const today = new Date(); today.setHours(0,0,0,0);

  menus.forEach(function(menu) {
    const mon    = new Date(menu.monday); mon.setHours(0,0,0,0);
    const isNext = mon > today;
    saveMenuDataToStorage(menu.sheet, isNext);
    Logger.log('✅ Сохранён ' + (isNext ? 'NEXT' : 'CURRENT') + ': ' + menu.name);
  });

  SpreadsheetApp.getUi().alert(L('✅ Данные меню сохранены!\n\nПланшеты обновятся при следующем открытии.'));
}

function getWeekPeriodFromStorage() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty('CURRENT_MENU_DATA');
    if (stored) {
      const parsed = JSON.parse(stored);
      if (parsed.weekLabel) return parsed.weekLabel;
    }
  } catch(e) {}

  try {
    const menus = MenuManager.findMenuSheets();
    if (menus.length > 0) {
      const nm = menus[0].name.match(/(\d{2}\.\d{2})-(\d{2}\.\d{2})/);
      if (nm) return nm[1] + ' - ' + nm[2];
    }
  } catch(e) {}

  return '';
}

function getTodayMenuData() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const officeName = getOfficeName();
    const todayName  = _getTodayName();
    const todayDate  = _getTodayDateFull();
    const todayIdx   = _getTodayDayIndex();

    const MONTHS = [L('января'),L('февраля'),L('марта'),L('апреля'),L('мая'),L('июня'),
                    L('июля'),L('августа'),L('сентября'),L('октября'),L('ноября'),L('декабря')];
    const now = new Date();
    const todayFormatted = now.getDate() + ' ' + MONTHS[now.getMonth()];

    const stored = PropertiesService.getScriptProperties().getProperty('CURRENT_MENU_DATA');
    if (stored) {
      try {
        const parsed  = JSON.parse(stored);
        const allDays = parsed.days || [];


        const todayDay = allDays.find(function(d) {
          if (d.dayDate && d.dayDate !== '') {
            return d.dayDate === todayFormatted;
          }
          return d.dayIdx === todayIdx;
        });

        Logger.log('getTodayMenuData: из PropertiesService, сегодня="' + todayFormatted +
          '", день: ' + (todayDay ? todayDay.dayName + ' ' + todayDay.dayDate : 'не найден'));

        return {
          todayName,
          todayDate,
          weekLabel:  parsed.weekLabel  || '',
          officeName: parsed.officeName || officeName,
          timings:    todayDay ? todayDay.timings  : [],
          alwaysOn:   todayDay ? todayDay.alwaysOn : []
        };
      } catch(parseErr) {
        Logger.log('⚠️ Ошибка парсинга CURRENT_MENU_DATA: ' + parseErr.toString());
      }
    }

    Logger.log('getTodayMenuData: PropertiesService пуст, читаем из листа');
    const todayMenu = MenuManager.findTodaySheet();
    if (!todayMenu) {
      return { todayName, todayDate, weekLabel: '', officeName,
               timings: [], alwaysOn: [], error: L('Меню не найдено') };
    }

    Logger.log('getTodayMenuData: используем лист "' + todayMenu.name + '"');
    const data = todayMenu.sheet.getDataRange().getValues();
    return _parse(data, todayIdx, officeName, todayName, todayDate);

  } catch(e) {
    Logger.log('❌ getTodayMenuData: ' + e.toString());
    return {
      todayName:  _getTodayName(),
      todayDate:  _getTodayDateFull(),
      weekLabel:  '',
      officeName: '',
      timings:    [],
      alwaysOn:   [],
      error:      e.toString()
    };
  }
}

function getWeekMenuData() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const officeName = getOfficeName();
    const todayIdx   = _getTodayDayIndex();
    const allDays    = [];

    const currentStored = PropertiesService.getScriptProperties().getProperty('CURRENT_MENU_DATA');
    if (currentStored) {
      try {
        const parsed = JSON.parse(currentStored);
        allDays.push(...(parsed.days || []));
      } catch(e) { Logger.log('⚠️ CURRENT_MENU_DATA parse: ' + e); }
    }

    const nextStored = PropertiesService.getScriptProperties().getProperty('NEXT_MENU_DATA');
    if (nextStored) {
      try {
        const parsed = JSON.parse(nextStored);
        allDays.push(...(parsed.days || []));
      } catch(e) { Logger.log('⚠️ NEXT_MENU_DATA parse: ' + e); }
    }

    if (allDays.length > 0) {

      const MONTHS = [L('января'),L('февраля'),L('марта'),L('апреля'),L('мая'),L('июня'),
                      L('июля'),L('августа'),L('сентября'),L('октября'),L('ноября'),L('декабря')];
      const now = new Date();
      const todayFormatted = now.getDate() + ' ' + MONTHS[now.getMonth()];

      allDays.forEach(function(d) {
        if (d.dayDate && d.dayDate !== '') {
          d.isToday = (d.dayDate === todayFormatted);
        } else {
          d.isToday = (d.dayIdx === todayIdx);
        }
      });

      Logger.log('getWeekMenuData: из PropertiesService, дней: ' + allDays.length +
        ', сегодня="' + todayFormatted + '"');
      return { days: allDays, officeName };
    }

    Logger.log('getWeekMenuData: PropertiesService пуст, читаем из листов');
    const menus  = MenuManager.findMenuSheets();
    const today  = new Date(); today.setHours(0,0,0,0);
    const days   = [];

    menus.forEach(function(menu) {
      const data    = menu.sheet.getDataRange().getValues();
      const dayData = _buildWeekDays(data, todayIdx, today);
      days.push(...dayData);
    });

    const MONTHS2 = [L('января'),L('февраля'),L('марта'),L('апреля'),L('мая'),L('июня'),
                     L('июля'),L('августа'),L('сентября'),L('октября'),L('ноября'),L('декабря')];
    const now2 = new Date();
    const todayFormatted2 = now2.getDate() + ' ' + MONTHS2[now2.getMonth()];

    days.forEach(function(d) {
      if (d.dayDate && d.dayDate !== '') {
        d.isToday = (d.dayDate === todayFormatted2);
      } else {
        d.isToday = (d.dayIdx === todayIdx);
      }
    });

    return { days, officeName };

  } catch(e) {
    Logger.log('❌ getWeekMenuData: ' + e.toString());
    return { days: [], error: e.toString() };
  }
}

function _buildWeekDays(data, todayIdx, today) {
  const DAY_NAMES = [L('Понедельник'),L('Вторник'),L('Среда'),L('Четверг'),L('Пятница'),L('Суббота'),L('Воскресенье')];
  const MONTHS    = [L('января'),L('февраля'),L('марта'),L('апреля'),L('мая'),L('июня'),
                     L('июля'),L('августа'),L('сентября'),L('октября'),L('ноября'),L('декабря')];

  const alwaysByDay = { 0:[], 1:[], 2:[], 3:[], 4:[], 5:[], 6:[] };
  let inAlways = false;

  for (let i = 1; i < data.length; i++) {
    const c0    = _str(data[i][0]);
    const c0lo  = c0.toLowerCase();
    if (isMenuAlwaysLabel_(c0)) { inAlways = true; continue; }
    if (isMenuTimeLabel_(c0)) { inAlways = false; break; }
    if (!inAlways) continue;

    const vals = [];
    for (let c = 1; c < data[i].length; c++) {
      const v = _cleanVal(data[i][c]);
      if (v) vals.push(v);
    }
    if (vals.length === 0) continue;

    if (!c0) {
      for (let d = 0; d < 7; d++) {
        vals.forEach(v => { if (!alwaysByDay[d].includes(v)) alwaysByDay[d].push(v); });
      }
    } else {
      const DAY_ABBR = { 'пн':0,'вт':1,'ср':2,'чт':3,'пт':4,'сб':5,'вс':6,
                       'mon':0,'tue':1,'wed':2,'thu':3,'fri':4,'sat':5,'sun':6 };
      const c0norm = c0lo.replace(/\s+/g, ' ').trim();
      const m = c0norm.match(/^(?:с|from)\s+([а-яёa-z]+)\s+(?:по|to)\s+([а-яёa-z]+)/i);
      if (m && DAY_ABBR[m[1]] !== undefined && DAY_ABBR[m[2]] !== undefined) {
        const from = DAY_ABBR[m[1]], to = DAY_ABBR[m[2]];
        for (let d = from; d <= to; d++) {
          vals.forEach(v => { if (!alwaysByDay[d].includes(v)) alwaysByDay[d].push(v); });
        }
      } else {
        for (let d = 0; d < 7; d++) {
          vals.forEach(v => { if (!alwaysByDay[d].includes(v)) alwaysByDay[d].push(v); });
        }
      }
    }
  }

  const days        = [];
  let dayCounter    = -1;
  let currentDay    = null;
  let currentTiming = null;
  let inSchedule    = false;

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const c0   = _str(row[0]);
    const c0lo = c0.toLowerCase();

    if (isMenuTimeLabel_(c0)) { inSchedule = true; continue; }
    if (!inSchedule) continue;

    const isDayHeader = (row[0] instanceof Date) || _isDayHeaderText(c0);

    if (isDayHeader) {
      if (currentTiming && currentTiming.items.length) {
        currentDay.timings.push(currentTiming); currentTiming = null;
      }
      if (currentDay) days.push(currentDay);
      dayCounter++;

      let name, dateStr, isToday;

      if (row[0] instanceof Date) {
        const d = row[0];
        name    = DAY_NAMES[dayCounter] || (L('День ') + (dayCounter + 1));
        dateStr = d.getDate() + ' ' + MONTHS[d.getMonth()];
        const rowDate   = new Date(d); rowDate.setHours(0,0,0,0);
        const todayCopy = new Date(today); todayCopy.setHours(0,0,0,0);
        isToday = rowDate.getTime() === todayCopy.getTime();
      } else {
        const parsed = _parseDayHeaderText(c0);
        name    = DAY_NAMES[dayCounter] || parsed.dayName;
        dateStr = parsed.dateStr;
        const todayShort = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
        isToday = parsed.dateRaw === todayShort;
      }

      currentDay = {
        dayName:  name,
        dayDate:  dateStr,
        dayIdx:   dayCounter,
        isToday:  isToday,
        timings:  [],
        alwaysOn: []
      };
      continue;
    }

    if (!currentDay) continue;

    const timeM = c0.match(/^(\d{1,2}:\d{2})\s*[-–—]\s*(.+)/);
    if (timeM) {
      if (currentTiming && currentTiming.items.length) currentDay.timings.push(currentTiming);
      currentTiming = { time: c0, items: [] };
      for (let c = 1; c < row.length; c++) {
        const v = _cleanVal(row[c]); if (v) currentTiming.items.push(v);
      }
      continue;
    }
    if (currentTiming && !c0) {
      for (let c = 1; c < row.length; c++) {
        const v = _cleanVal(row[c]); if (v) currentTiming.items.push(v);
      }
    }
  }

  if (currentTiming && currentTiming.items.length && currentDay) currentDay.timings.push(currentTiming);
  if (currentDay) days.push(currentDay);

  days.forEach(function(d) { d.alwaysOn = (alwaysByDay[d.dayIdx] || []).slice(); });
  return days;
}

function _parse(data, todayIdx, officeName, todayName, todayDate) {
  const result = { todayName, todayDate, weekLabel: '', officeName, timings: [], alwaysOn: [] };

  if (data[0] && data[0][0]) {
    const h = data[0][0].toString();
    const m = h.match(/(?:с|from)\s+(\d+\.\d+)\s+(?:по|to)\s+(\d+\.\d+)/i);
    if (m) result.weekLabel = L('с ') + m[1] + L(' по ') + m[2];
  }

  let mode = 'skip', dayCounter = -1, inToday = false, currentTiming = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i], c0 = _str(row[0]), c0lo = c0.toLowerCase();

    if (isMenuAlwaysLabel_(c0)) {
      mode = 'always';
      if (currentTiming) { _pushTiming(result, currentTiming); currentTiming = null; }
      continue;
    }
    if (isMenuTimeLabel_(c0)) {
      mode = 'schedule'; dayCounter = -1; inToday = false;
      if (currentTiming) { _pushTiming(result, currentTiming); currentTiming = null; }
      continue;
    }

    if (mode === 'always') {
      const period     = c0;
      const applicable = _periodApplies(period, todayIdx);
      for (let c = 1; c < row.length; c++) {
        const val = _cleanVal(row[c]); if (!val) continue;
        if (period && !applicable) continue;
        if (!result.alwaysOn.includes(val)) result.alwaysOn.push(val);
      }
      if (!period) {
        for (let c = 0; c < row.length; c++) {
          const val = _cleanVal(row[c]);
          if (val && !result.alwaysOn.includes(val)) result.alwaysOn.push(val);
        }
      }
      continue;
    }

    if (mode === 'schedule') {
      if (row[0] instanceof Date || _isDayHeaderText(c0)) {
        if (currentTiming) { _pushTiming(result, currentTiming); currentTiming = null; }
        dayCounter++;
        if (row[0] instanceof Date) {
          const rowDate    = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd.MM');
          const todayShort = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM');
          inToday = (rowDate === todayShort) || (dayCounter === todayIdx);
        } else {
          const parsed     = _parseDayHeaderText(c0);
          const todayShort = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
          inToday = (parsed.dateRaw === todayShort) || (dayCounter === todayIdx);
        }
        continue;
      }
      if (!inToday) continue;
      const timeM = c0.match(/^(\d{1,2}:\d{2})\s*[-–—]\s*(.+)/);
      if (timeM) {
        if (currentTiming) _pushTiming(result, currentTiming);
        currentTiming = { time: c0, items: [] };
        for (let c = 1; c < row.length; c++) { const v = _cleanVal(row[c]); if (v) currentTiming.items.push(v); }
        continue;
      }
      if (currentTiming && !c0) {
        for (let c = 1; c < row.length; c++) { const v = _cleanVal(row[c]); if (v) currentTiming.items.push(v); }
      }
    }
  }

  if (currentTiming) _pushTiming(result, currentTiming);
  result.alwaysOn = result.alwaysOn.filter((v, i, a) => a.indexOf(v) === i);
  return result;
}

function _periodApplies(period, todayIdx) {
  if (!period) return true;
  const DAY_ABBR = { 'пн':0,'вт':1,'ср':2,'чт':3,'пт':4,'сб':5,'вс':6,
                       'mon':0,'tue':1,'wed':2,'thu':3,'fri':4,'sat':5,'sun':6 };
  const m = period.toLowerCase().match(/(?:с|from)\s+([а-яёa-z]+)\s+(?:по|to)\s+([а-яёa-z]+)/);
  if (!m) return true;
  const from = DAY_ABBR[m[1]], to = DAY_ABBR[m[2]];
  if (from === undefined || to === undefined) return true;
  return todayIdx >= from && todayIdx <= to;
}

function _pushTiming(result, timing) {
  if (timing && timing.items.length > 0) result.timings.push(timing);
}

function _str(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) return '';
  const s = val.toString().trim();
  if (s.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s/)) return '';
  if (s.match(/GMT[+-]\d{4}/)) return '';
  return s;
}

function _cleanVal(val) {
  const s = _str(val);
  if (!s || isMenuFiller_(s)) return '';
  return s;
}

function _getTodayDayIndex() {
  const d = new Date().getDay();
  return d === 0 ? 6 : d - 1;
}

function _getTodayName() {
  return [L('Воскресенье'),L('Понедельник'),L('Вторник'),L('Среда'),L('Четверг'),L('Пятница'),L('Суббота')][new Date().getDay()];
}

function _getTodayDateFull() {
  const now    = new Date();
  const months = [L('января'),L('февраля'),L('марта'),L('апреля'),L('мая'),L('июня'),
                  L('июля'),L('августа'),L('сентября'),L('октября'),L('ноября'),L('декабря')];
  return now.getDate() + ' ' + months[now.getMonth()];
}

function _isDayHeaderText(str) {
  if (!str) return false;
  const lo = str.toLowerCase();
  const DAY_WORDS = ['понедельник','вторник','среда','четверг','пятница','суббота','воскресенье',
                     'monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
  return DAY_WORDS.some(d => lo.startsWith(d)) && /\d{2}\.\d{2}\.\d{4}/.test(str);
}

function _parseDayHeaderText(str) {
  const MONTHS = [L('января'),L('февраля'),L('марта'),L('апреля'),L('мая'),L('июня'),
                  L('июля'),L('августа'),L('сентября'),L('октября'),L('ноября'),L('декабря')];
  const m = str.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (!m) return { dayName: str, dateStr: '', dateRaw: '' };
  const day      = parseInt(m[1]);
  const month    = parseInt(m[2]) - 1;
  const dateStr  = day + ' ' + MONTHS[month];
  return { dayName: str, dateStr, dateRaw: m[1] + '.' + m[2] + '.' + m[3] };
}

function activateNextWeekMenu() {
  try {
    const next = PropertiesService.getScriptProperties().getProperty('NEXT_MENU_DATA');
    if (!next) {
      Logger.log('activateNextWeekMenu: NEXT_MENU_DATA пуст — нечего активировать');
      return;
    }
    PropertiesService.getScriptProperties().setProperty('CURRENT_MENU_DATA', next);
    PropertiesService.getScriptProperties().deleteProperty('NEXT_MENU_DATA');
    Logger.log('✅ activateNextWeekMenu: меню следующей недели стало текущим');
  } catch(e) {
    Logger.log('❌ activateNextWeekMenu: ' + e.toString());
  }
}
