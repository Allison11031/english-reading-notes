// 英文閱讀筆記：第三版
// 流程：先存輸入內容 → 查字典與翻譯 → 把結果寫回同一筆紀錄。
// AI 只有在按下「AI 解析」時才會使用。

var TIME_ZONE = 'Asia/Taipei';

// 欄位順序就是試算表的欄位順序。之後要加欄位，請加在該列最後面。
var SHEET_COLUMNS = {
  '單字': ['ID', '新增時間', '書名', '篇名', '單字', '音標', '解釋', '例句', '原文句子', '這篇文章中的意思', '我的註解', '來源', '簡短釋義', '語境依據句子', 'AI解析'],
  '片語': ['ID', '新增時間', '書名', '篇名', '片語', '解釋', '例句', '原文句子', '這篇文章中的意思', '我的註解', '來源', '簡短釋義', '語境依據句子', 'AI解析'],
  '閱讀理解': ['ID', '新增時間', '書名', '篇名', '段落或主題', '我的理解', '原文句子', '疑問']
};

var SETTINGS_SHEET = '設定';
var AI_LOG_SHEET = 'AI使用紀錄';
var AI_LOG_COLUMNS = ['時間', '用途', '內容', '結果'];

var DEFAULT_SETTINGS = [
  ['最後書名', ''],
  ['最後篇名', ''],
  ['AI每日上限', '30'],
  ['AI模型', 'gemini-3.5-flash-lite'],
  ['AI使用日期', ''],
  ['AI今日次數', '0']
];

var PART_OF_SPEECH = {
  'noun': 'n.',
  'verb': 'v.',
  'adjective': 'adj.',
  'adverb': 'adv.',
  'pronoun': 'pron.',
  'preposition': 'prep.',
  'conjunction': 'conj.',
  'interjection': 'int.',
  'exclamation': 'int.'
};


// ---------- 開啟網頁與建立工作表 ----------

function doGet() {
  prepareSheets();

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('英文閱讀筆記')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function prepareSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEET_COLUMNS).forEach(function (sheetName) {
    var columns = SHEET_COLUMNS[sheetName];
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1, sheet.getMaxRows(), columns.length).setNumberFormat('@');
      sheet.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      return;
    }

    // 第一版的工作表：只在最後補上新欄位，舊資料不動
    var currentColumnCount = sheet.getLastColumn();
    if (currentColumnCount < columns.length) {
      if (sheet.getMaxColumns() < columns.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), columns.length - sheet.getMaxColumns());
      }
      var newColumns = columns.slice(currentColumnCount);
      sheet.getRange(1, currentColumnCount + 1, sheet.getMaxRows(), newColumns.length).setNumberFormat('@');
      sheet.getRange(1, currentColumnCount + 1, 1, newColumns.length).setValues([newColumns]).setFontWeight('bold');
    }
  });

  var settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(SETTINGS_SHEET);
    settingsSheet.getRange(1, 1, 1, 2).setNumberFormat('@').setValues([['項目', '值']]).setFontWeight('bold');
    settingsSheet.setFrozenRows(1);
  }

  var existingKeys = settingsSheet.getRange(1, 1, Math.max(settingsSheet.getLastRow(), 1), 1).getDisplayValues().map(function (row) {
    return row[0];
  });
  DEFAULT_SETTINGS.forEach(function (setting) {
    if (existingKeys.indexOf(setting[0]) === -1) {
      var nextRow = settingsSheet.getLastRow() + 1;
      settingsSheet.getRange(nextRow, 1, 1, 2).setNumberFormat('@').setValues([setting]);
    }
  });

  if (!spreadsheet.getSheetByName(AI_LOG_SHEET)) {
    var logSheet = spreadsheet.insertSheet(AI_LOG_SHEET);
    logSheet.getRange(1, 1, logSheet.getMaxRows(), AI_LOG_COLUMNS.length).setNumberFormat('@');
    logSheet.getRange(1, 1, 1, AI_LOG_COLUMNS.length).setValues([AI_LOG_COLUMNS]).setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
}


// ---------- 共用的小工具 ----------

function readSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  var settings = {};
  sheet.getRange(1, 1, sheet.getLastRow(), 2).getDisplayValues().forEach(function (row) {
    settings[row[0]] = row[1];
  });
  return settings;
}


function writeSetting(key, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  var keys = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(String(value));
      return;
    }
  }
}


function now() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
}


function today() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
}


function aiUsage() {
  var settings = readSettings();
  return {
    used: settings['AI使用日期'] === today() ? Number(settings['AI今日次數']) || 0 : 0,
    limit: Number(settings['AI每日上限']) || 0
  };
}


function addRow(sheet, columns, record) {
  var nextRow = sheet.getLastRow() + 1;
  if (nextRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 100);
  }
  var row = columns.map(function (columnName) {
    return record[columnName] === undefined ? '' : String(record[columnName]);
  });
  // 先設成純文字，避免 -ing、=word 這類內容被當成公式
  sheet.getRange(nextRow, 1, 1, columns.length).setNumberFormat('@').setValues([row]);
}


function findRecord(id) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var types = ['單字', '片語'];

  for (var i = 0; i < types.length; i++) {
    var type = types[i];
    var sheet = spreadsheet.getSheetByName(type);
    var cell = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (cell) {
      return rowToRecord(sheet, type, cell.getRow());
    }
  }
  return null;
}


function rowToRecord(sheet, type, rowNumber) {
  var columns = SHEET_COLUMNS[type];
  var values = sheet.getRange(rowNumber, 1, 1, columns.length).getDisplayValues()[0];
  var record = { '類型': type, rowNumber: rowNumber, '內容': values[4] };
  columns.forEach(function (columnName, index) {
    record[columnName] = values[index];
  });
  return record;
}


function writeFields(id, fields) {
  var record = findRecord(id);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(record['類型']);
  var columns = SHEET_COLUMNS[record['類型']];

  Object.keys(fields).forEach(function (fieldName) {
    var columnIndex = columns.indexOf(fieldName);
    if (columnIndex !== -1) {
      sheet.getRange(record.rowNumber, columnIndex + 1).setValue(String(fields[fieldName]));
    }
  });
}


function writeFieldsLocked(id, fields) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    writeFields(id, fields);
  } finally {
    lock.releaseLock();
  }
}


function findSameText(text) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var target = text.toLowerCase();
  var types = ['單字', '片語'];
  var latest = null;

  types.forEach(function (type) {
    var sheet = spreadsheet.getSheetByName(type);
    if (sheet.getLastRow() < 2) {
      return;
    }
    var values = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (var i = values.length - 1; i >= 0; i--) {
      if (values[i][0].trim().toLowerCase() === target) {
        var record = rowToRecord(sheet, type, i + 2);
        if (!latest || record['新增時間'] > latest['新增時間']) {
          latest = record;
        }
        return;
      }
    }
  });

  return latest;
}


// ---------- 網頁載入與設定 ----------

function getInitialData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var articlesByBook = {};

  Object.keys(SHEET_COLUMNS).forEach(function (sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet.getLastRow() < 2) {
      return;
    }

    var rows = sheet.getRange(2, 3, sheet.getLastRow() - 1, 2).getDisplayValues();
    rows.forEach(function (row) {
      var book = row[0];
      var article = row[1];
      if (!book) {
        return;
      }
      if (!articlesByBook[book]) {
        articlesByBook[book] = [];
      }
      if (article && articlesByBook[book].indexOf(article) === -1) {
        articlesByBook[book].push(article);
      }
    });
  });

  var settings = readSettings();

  return {
    lastBook: settings['最後書名'],
    lastArticle: settings['最後篇名'],
    aiUsage: aiUsage(),
    aiReady: Boolean(PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')),
    articlesByBook: articlesByBook
  };
}


function saveSource(request) {
  var book = String(request.book || '').trim();
  var article = String(request.article || '').trim();
  if (!book || !article) {
    throw new Error('請填寫書名和篇名。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    writeSetting('最後書名', book);
    writeSetting('最後篇名', article);
  } finally {
    lock.releaseLock();
  }
  return true;
}


// ---------- 字典、翻譯、AI ----------

function lookupDictionary(text) {
  var result = { found: false, phonetic: '', meanings: [] };

  try {
    var response = UrlFetchApp.fetch(
      'https://api.dictionaryapi.dev/api/v2/entries/en/' + encodeURIComponent(text.toLowerCase()),
      { muteHttpExceptions: true }
    );
    if (response.getResponseCode() !== 200) {
      return result;
    }

    var entries = JSON.parse(response.getContentText());

    entries.forEach(function (entry) {
      if (!result.phonetic && entry.phonetic) {
        result.phonetic = entry.phonetic;
      }
      (entry.phonetics || []).forEach(function (phonetic) {
        if (!result.phonetic && phonetic.text) {
          result.phonetic = phonetic.text;
        }
      });
      (entry.meanings || []).forEach(function (meaning) {
        var definitions = (meaning.definitions || []).filter(function (item) {
          return item.definition;
        }).map(function (item) {
          return { zh: '', en: item.definition, example: item.example || '' };
        });
        if (definitions.length > 0) {
          result.meanings.push({
            partOfSpeech: PART_OF_SPEECH[meaning.partOfSpeech] || meaning.partOfSpeech || '',
            definitions: definitions
          });
        }
      });
    });

    result.found = result.meanings.length > 0;
  } catch (error) {
    result.found = false;
  }

  return result;
}


function translateMeanings(text, meanings) {
  var result = { complete: false, shortMeaning: '' };

  try {
    var shortMeaning = LanguageApp.translate(text, 'en', 'zh-TW').trim();
    if (shortMeaning && shortMeaning.toLowerCase() !== text.toLowerCase()) {
      result.shortMeaning = shortMeaning;
    }

    var englishLines = [];
    meanings.forEach(function (meaning) {
      meaning.definitions.forEach(function (definition) {
        englishLines.push(definition.en);
      });
    });

    var chineseLines = LanguageApp.translate(englishLines.join('\n'), 'en', 'zh-TW').split('\n').map(function (line) {
      return line.trim();
    }).filter(function (line) {
      return line;
    });

    // 行數對不上時，代表翻譯和解釋可能錯位，寧可當成翻譯失敗
    if (chineseLines.length === englishLines.length) {
      var lineIndex = 0;
      meanings.forEach(function (meaning) {
        meaning.definitions.forEach(function (definition) {
          definition.zh = chineseLines[lineIndex];
          lineIndex = lineIndex + 1;
        });
      });
      result.complete = Boolean(result.shortMeaning);
    }
  } catch (error) {
    result.complete = false;
  }

  return result;
}


function callAi(purpose, text, prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { ok: false, message: '還沒有設定 AI 金鑰，AI 功能暫時不能用。' };
  }

  var modelName = '';
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var usage = aiUsage();
    if (usage.used >= usage.limit) {
      return { ok: false, message: '今天的 AI 次數已用完（' + usage.used + '/' + usage.limit + '）。' };
    }
    writeSetting('AI使用日期', today());
    writeSetting('AI今日次數', usage.used + 1);
    modelName = readSettings()['AI模型'];
  } finally {
    lock.releaseLock();
  }

  var outcome = { ok: false, message: 'AI 暫時無法使用，請稍後再試。' };

  try {
    var response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelName) + ':generateContent',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': apiKey },
        muteHttpExceptions: true,
        payload: JSON.stringify({
          contents: [{ role: 'user', parts: [{ text: prompt }] }],
          generationConfig: { responseMimeType: 'application/json' }
        })
      }
    );

    if (response.getResponseCode() === 200) {
      var body = JSON.parse(response.getContentText());
      var answer = body.candidates[0].content.parts.map(function (part) {
        return part.text || '';
      }).join('').replace(/^```(?:json)?\s*|\s*```$/g, '');
      outcome = { ok: true, data: JSON.parse(answer) };
    } else {
      outcome.logDetail = 'HTTP ' + response.getResponseCode();
    }
  } catch (error) {
    outcome.logDetail = '回應無法解析';
  }

  var logLock = LockService.getScriptLock();
  logLock.waitLock(10000);
  try {
    addRow(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AI_LOG_SHEET), AI_LOG_COLUMNS, {
      '時間': now(),
      '用途': purpose,
      '內容': text,
      '結果': outcome.ok ? '成功' : '失敗：' + (outcome.logDetail || '')
    });
  } finally {
    logLock.releaseLock();
  }

  return outcome;
}


function buildAiPrompt(text, sentence) {
  return [
    '你是給繁體中文母語者使用的英文老師。請解析下面的英文單字或片語。',
    '請用繁體中文說明常見意思與用法，並提供一到兩個英文例句與中文翻譯。',
    sentence ? '另外提供了原文句子，請再說明它在這個句子中的意思。' : '',
    '只輸出 JSON，格式：{"short_meaning":"最常見的簡短繁體中文釋義","analysis":"完整解析，可換行"}',
    '單字或片語：' + JSON.stringify(text),
    sentence ? '原文句子：' + JSON.stringify(sentence) : ''
  ].join('\n');
}


function formatMeanings(meanings) {
  var definitionLines = [];
  var exampleLines = [];

  meanings.forEach(function (meaning) {
    definitionLines.push(meaning.partOfSpeech);
    meaning.definitions.forEach(function (definition, index) {
      var number = index + 1;
      if (definition.zh) {
        definitionLines.push(number + '. ' + definition.zh);
        if (definition.en) {
          definitionLines.push('   ' + definition.en);
        }
      } else {
        definitionLines.push(number + '. ' + definition.en);
      }
      if (definition.example) {
        definitionLines.push('   例句：' + definition.example);
        exampleLines.push(meaning.partOfSpeech + ' ' + number + '. ' + definition.example);
      }
    });
    definitionLines.push('');
  });

  return {
    // 試算表單格上限約 50,000 字
    meaningText: definitionLines.join('\n').trim().slice(0, 49000),
    exampleText: exampleLines.join('\n').slice(0, 49000)
  };
}


function publicRecord(record) {
  var result = {};
  Object.keys(record).forEach(function (key) {
    if (key !== 'rowNumber') {
      result[key] = record[key];
    }
  });
  return result;
}


// ---------- 查詢並儲存（只用字典與翻譯，不使用 AI） ----------

function saveAndLookup(request) {
  var id = String(request.id || '').trim();
  var text = String(request.text || '').trim().replace(/\s+/g, ' ');
  var sentence = String(request.sentence || '').trim();
  var book = String(request.book || '').trim();
  var article = String(request.article || '').trim();

  if (!book || !article) {
    throw new Error('請先設定書名和篇名。');
  }
  if (!text) {
    throw new Error('請填寫「單字或片語」。');
  }
  if (!id) {
    throw new Error('缺少紀錄編號，請重新整理頁面後再試一次。');
  }

  var type = request.type;
  if (type !== '單字' && type !== '片語') {
    type = /\s/.test(text) ? '片語' : '單字';
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    prepareSheets();

    var existing = findRecord(id);

    // 網路中斷後重送時，不會多存一筆
    if (existing && existing['來源'] !== '查詢中') {
      return { record: publicRecord(existing) };
    }

    if (!existing) {
      var sameText = findSameText(text);

      if (sameText && !request.saveAnother) {
        return { duplicate: true, record: publicRecord(sameText) };
      }

      var newRecord = {
        'ID': id,
        '新增時間': now(),
        '書名': book,
        '篇名': article,
        '原文句子': sentence,
        '來源': '查詢中'
      };

      if (sameText) {
        type = sameText['類型'];
        newRecord['音標'] = sameText['音標'] || '';
        newRecord['解釋'] = sameText['解釋'];
        newRecord['例句'] = sameText['例句'];
        newRecord['簡短釋義'] = sameText['簡短釋義'];
        newRecord['來源'] = sameText['來源'];
      }

      newRecord[type] = text;
      addRow(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(type), SHEET_COLUMNS[type], newRecord);
      writeSetting('最後書名', book);
      writeSetting('最後篇名', article);

      if (sameText) {
        return { record: publicRecord(findRecord(id)) };
      }
    }
  } finally {
    lock.releaseLock();
  }

  var dictionary = lookupDictionary(text);
  var translation = translateMeanings(text, dictionary.meanings);
  var formatted = formatMeanings(dictionary.meanings);

  writeFieldsLocked(id, {
    '音標': dictionary.phonetic,
    '解釋': formatted.meaningText,
    '例句': formatted.exampleText,
    '簡短釋義': translation.shortMeaning,
    '來源': dictionary.found ? '字典' : '查無結果'
  });

  return { record: publicRecord(findRecord(id)) };
}


// 只有按下「AI 解析」才會執行。結果只寫入 AI解析 欄位，不改動字典結果。
function analyzeWithAi(request) {
  var id = String(request.id || '').trim();
  var record = findRecord(id);

  if (!record) {
    throw new Error('請確認這筆紀錄仍在試算表中。');
  }

  var sentence = String(request.sentence || '').trim() || record['原文句子'];
  var text = record['內容'];

  // 同一筆紀錄、同一個原文句子已解析過，直接沿用，不使用額度
  if (record['AI解析'] && record['語境依據句子'] === sentence) {
    return { record: publicRecord(record), aiUsage: aiUsage(), reused: true };
  }

  var answer = callAi('AI解析', text, buildAiPrompt(text, sentence));

  if (!answer.ok) {
    return { record: publicRecord(record), aiUsage: aiUsage(), message: answer.message };
  }

  var analysis = String(answer.data.analysis || '').trim().slice(0, 49000);
  if (!analysis) {
    return { record: publicRecord(record), aiUsage: aiUsage(), message: 'AI 沒有回傳內容，請再試一次。' };
  }

  var fields = {
    'AI解析': analysis,
    '語境依據句子': sentence,
    '原文句子': sentence
  };

  if (!record['簡短釋義'] && answer.data.short_meaning) {
    fields['簡短釋義'] = answer.data.short_meaning;
  }

  writeFieldsLocked(id, fields);

  return { record: publicRecord(findRecord(id)), aiUsage: aiUsage() };
}


// 紀錄頁：一次讀取全部紀錄，搜尋與排序在網頁上進行
function getRecords() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var records = [];

  Object.keys(SHEET_COLUMNS).forEach(function (type) {
    var sheet = spreadsheet.getSheetByName(type);
    if (sheet.getLastRow() < 2) {
      return;
    }

    var columns = SHEET_COLUMNS[type];
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, columns.length).getDisplayValues();

    rows.forEach(function (values) {
      if (!values[0]) {
        return;
      }

      var record = { '類型': type, '內容': values[4] };
      columns.forEach(function (columnName, index) {
        record[columnName] = values[index];
      });
      records.push(record);
    });
  });

  return records;
}


// 「我的註解」離開欄位時自動儲存
function saveNote(request) {
  var id = String(request.id || '').trim();
  var note = String(request.note || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!findRecord(id)) {
      throw new Error('請確認這筆紀錄仍在試算表中。');
    }
    writeFields(id, { '我的註解': note });
    return { saved: true };
  } finally {
    lock.releaseLock();
  }
}


// 閱讀理解不需要查詢，直接儲存
function saveReading(request) {
  var id = String(request.id || '').trim();
  var values = request.values || {};
  var book = String(request.book || '').trim();
  var article = String(request.article || '').trim();
  var understanding = String(values['我的理解'] || '').trim();

  if (!book || !article) {
    throw new Error('請先設定書名和篇名。');
  }
  if (!understanding) {
    throw new Error('請填寫「我的理解」。');
  }
  if (!id) {
    throw new Error('缺少紀錄編號，請重新整理頁面後再試一次。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    prepareSheets();

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('閱讀理解');
    var alreadySaved = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (alreadySaved) {
      return { saved: true };
    }

    addRow(sheet, SHEET_COLUMNS['閱讀理解'], {
      'ID': id,
      '新增時間': now(),
      '書名': book,
      '篇名': article,
      '段落或主題': String(values['段落或主題'] || '').trim(),
      '我的理解': understanding,
      '原文句子': String(values['原文句子'] || '').trim(),
      '疑問': String(values['疑問'] || '').trim()
    });
    writeSetting('最後書名', book);
    writeSetting('最後篇名', article);

    return { saved: true };
  } finally {
    lock.releaseLock();
  }
}
