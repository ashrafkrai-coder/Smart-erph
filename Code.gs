const ERPH = Object.freeze({
  DAYS: ['ISNIN', 'SELASA', 'RABU', 'KHAMIS', 'JUMAAT'],
  DAY_BY_INDEX: ['Ahad', 'Isnin', 'Selasa', 'Rabu', 'Khamis', 'Jumaat', 'Sabtu'],
  DAY_JAWI_BY_INDEX: ['احد', 'اثنين', 'سلاسا', 'رابو', 'خميس', 'جمعة', 'سبت'],
  DEFAULT_OUTPUT_SCRIPT: 'jawi',
  SUBJECTS: {
    PAI: 'Pendidikan Agama Islam (PAI)',
    KKQ: 'Kelas Kemahiran al-Quran (KKQ)'
  },
  SUBJECTS_JAWI: {
    PAI: 'ڤنديديقن اݢام إسلام (ڤاي)',
    KKQ: 'کلس کماهيرن القرءان (ککق)'
  },
  SUBJECT_ALIASES_JAWI: {
    PAI: ['ڤنديديقن اݢام إسلام (ڤاي)', 'ڤنديديقن اݢام إسلام', 'ڤنديديقن إسلام', 'ڤنديديقن اسلام'],
    KKQ: ['کلس کماهيرن القرءان (ککق)', 'کلس کماهيرن القرءان', 'کلس کماهيرن القرآن (KKQ)', 'کلس کماهيرن القرآن']
  },
  ACTIVE_STATUS_ALIASES: ['AKTIF', 'اکتيف'],
  STUDENT_CENTRED_STRATEGY: 'Pembelajaran Berpusatkan Murid dan Kolaboratif',
  STUDENT_CENTRED_STRATEGY_JAWI: 'ڤمبلاجرن برڤوستکن موريد دان کولابوراتيف',
  ASSESSMENT_ROWS: {
    'Amali / Eksperimen': 53,
    Projek: 54,
    Pembentangan: 55,
    Ujian: 56,
    Peperiksaan: 57,
    'Latihan / Kerja Rumah': 58,
    'Lembaran Kerja': 59,
    Pemerhatian: 60,
    Kuiz: 61,
    Lisan: 62,
    Tugasan: 63
  }
});

/** Web App endpoint for the Smart eRPH PWA. Deploy as a Web App after saving. */
function doGet() {
  return jsonResponse_({ ok: true, service: 'Smart eRPH AI', status: 'ready' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData?.contents || '{}');
    verifyPwaToken_(payload.apiToken);
    if (payload.action === 'generate_week') return jsonResponse_(generateWeekFromPwa_(payload));
    if (payload.action === 'prepare_classroom') return jsonResponse_(prepareClassroomFromPwa_(payload));
    throw new Error('Tindakan API tidak sah.');
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function onOpen() {
  const ss = SpreadsheetApp.getActive();
  if (ss) {
    // Simpan ID fail Jawi secara automatik setiap kali fail dibuka.
    PropertiesService.getScriptProperties().setProperty('ERPH_JAWI_SPREADSHEET_ID', ss.getId());
  }

  SpreadsheetApp.getUi()
    .createMenu('Smart eRPH AI — Jawi')
    .addItem('Jana semua slot tab aktif (Jawi)', 'generateAllActiveSlots')
    .addItem('Baiki strategi murid & kolaboratif yang kosong', 'fillMissingStudentCentredStrategies')
    .addSeparator()
    .addItem('Cipta salinan Jawi seluruh fail', 'createJawiWorkbookCopy')
    .addItem('Tukar DSKP/Sukatan fail Jawi kepada Jawi', 'convertJawiCurriculumTabs')
    .addSeparator()
    .addItem('Tetapkan API Gemini', 'setGeminiApiKey')
    .addItem('Uji sambungan Gemini', 'testGeminiApi')
    .addItem('Sambungkan fail Jawi ini untuk PWA', 'setErphSpreadsheet')
    .addToUi();
}

/**
 * Cipta salinan penuh workbook dan tukar teks literal Rumi kepada Jawi.
 * Fail asal tidak disentuh. Formula, nombor, tarikh, checkbox dan URL dikekalkan.
 * Nama tab juga dikekalkan supaya struktur Smart eRPH tidak rosak.
 */
function createJawiWorkbookCopy() {
  const ui = SpreadsheetApp.getUi();
  const source = SpreadsheetApp.getActive();
  const answer = ui.alert(
    'Cipta salinan Jawi',
    'Sistem akan mencipta salinan fail ini dan menukar teks Rumi kepada Jawi. Fail asal, formula, nombor, tarikh, checkbox dan nama tab tidak akan diubah.',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer !== ui.Button.OK) return;

  source.toast('Mencipta salinan eRPH versi Jawi...', 'Smart eRPH AI', 5);
  const sourceFile = DriveApp.getFileById(source.getId());
  const copyFile = sourceFile.makeCopy(source.getName() + ' — Jawi');
  const jawiWorkbook = SpreadsheetApp.openById(copyFile.getId());

  try {
    const stats = convertWorkbookToJawi_(jawiWorkbook, source);
    SpreadsheetApp.flush();
    ui.alert(
      'Penukaran Jawi selesai',
      'Fail asal tidak diubah.\n\n' +
      'Tab diproses: ' + stats.sheets + '\n' +
      'Sel teks ditukar: ' + stats.cells + '\n' +
      'Teks unik dihantar ke AI: ' + stats.uniqueTexts + '\n\n' +
      'Salinan Jawi:\n' + copyFile.getUrl(),
      ui.ButtonSet.OK
    );
  } catch (error) {
    copyFile.setName(source.getName() + ' — Jawi (TIDAK LENGKAP)');
    throw error;
  }
}

function convertWorkbookToJawi_(workbook, progressWorkbook) {
  const sourceUi = progressWorkbook || SpreadsheetApp.getActive();
  const sheets = workbook.getSheets();
  const unique = new Map();
  let candidateCells = 0;

  sheets.forEach((sheet, sheetIndex) => {
    sourceUi.toast(
      'Membaca tab ' + (sheetIndex + 1) + '/' + sheets.length + ': ' + sheet.getName(),
      'Smart eRPH AI — Jawi',
      5
    );

    const range = sheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();

    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        if (formulas[row][col]) continue;
        const value = values[row][col];
        if (typeof value !== 'string' || !shouldTransliterateToJawi_(value)) continue;

        candidateCells++;
        if (!unique.has(value)) unique.set(value, []);
        unique.get(value).push({
          sheet: sheet,
          row: range.getRow() + row,
          col: range.getColumn() + col
        });
      }
    }
  });

  const entries = Array.from(unique.keys()).map((text, index) => ({ id: index + 1, text: text }));
  if (!entries.length) return { sheets: sheets.length, cells: 0, uniqueTexts: 0 };

  sourceUi.toast('Menukar ' + entries.length + ' teks unik kepada Jawi...', 'Smart eRPH AI — Jawi', 5);
  const translatedById = transliterateJawiEntries_(entries, sourceUi);
  const idByOriginal = new Map(entries.map(item => [item.text, item.id]));
  let changedCells = 0;

  unique.forEach((locations, originalText) => {
    const id = idByOriginal.get(originalText);
    const jawiText = translatedById.get(id);
    if (!jawiText || jawiText === originalText) return;
    locations.forEach(location => {
      location.sheet.getRange(location.row, location.col).setValue(jawiText);
      changedCells++;
    });
  });

  return { sheets: sheets.length, cells: changedCells, uniqueTexts: entries.length, candidates: candidateCells };
}

function hasArabicScript_(value) {
  return /[؀-ۿ]/.test(String(value == null ? '' : value));
}

// Lajur teks kurikulum (indeks 0) yang ditukar. Kod SK, minggu RPT, pelajaran dan status AKTIF sengaja dikekalkan
// kerana kod padankan minggu (digit) dan AKTIF berdasarkan nilai Rumi.
const JAWI_CURRICULUM_TABS = ['DSKP_PAI_T4', 'DSKP_PAI_T5', 'SUKATAN_KKQ_T1', 'SUKATAN_KKQ_T2', 'SUKATAN_KKQ_T3'];
const JAWI_CURRICULUM_COLUMNS = [4, 5, 6, 8, 9, 10, 11, 12];

/**
 * Tukar teks DSKP/Sukatan dalam fail Jawi kepada Jawi (sekali jalan, boleh dijalankan berulang untuk menyambung).
 * Hanya sel yang masih Rumi dihantar ke AI; sel yang sudah ada tulisan Arab dilangkau.
 */
function convertJawiCurriculumTabs() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Tukar DSKP/Sukatan kepada Jawi',
    'Sistem akan menukar teks Bidang, Tajuk, Alias, Standard Kandungan, Standard Pembelajaran, Objektif, Kata Kunci dan Sumber ' +
    'dalam 5 tab DSKP/Sukatan pada fail e-RPH Jawi. Kod SK, Minggu RPT, Status AKTIF dan tab lain tidak diubah. ' +
    'Jika terhenti (kuota/masa), jalankan semula untuk menyambung.',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer !== ui.Button.OK) return;

  const deadline = Date.now() + 4.5 * 60 * 1000;
  const workbook = getWorkbook_('jawi');
  const missingTabs = [];
  const tabs = [];
  const unique = new Map();

  JAWI_CURRICULUM_TABS.forEach(name => {
    const sheet = getWorksheet_(name, workbook);
    if (!sheet) { missingTabs.push(name); return; }
    const range = sheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();
    tabs.push({ sheet, range, values, formulas });
    for (let r = 1; r < values.length; r++) {
      if (!values[r][0]) continue;
      JAWI_CURRICULUM_COLUMNS.forEach(c => {
        const value = values[r][c];
        if (formulas[r][c] || typeof value !== 'string' || !needsJawi_(value)) return;
        if (!unique.has(value)) unique.set(value, []);
        unique.get(value).push({ tab: tabs.length - 1, r, c });
      });
    }
  });

  const entries = Array.from(unique.keys()).map((text, index) => ({ id: index + 1, text }));
  if (!entries.length) {
    ui.alert('Tiada teks Rumi tinggal dalam tab DSKP/Sukatan Jawi.' + (missingTabs.length ? '\nTab tidak ditemui: ' + missingTabs.join(', ') : ''));
    return;
  }

  workbook.toast('Menukar ' + entries.length + ' teks unik kepada Jawi...', 'Smart eRPH AI — Jawi', 5);
  const converted = transliterateJawiWithRetry_(entries, deadline, workbook);

  let changedCells = 0;
  const changedTabs = new Set();
  entries.forEach(entry => {
    const jawi = converted.map.get(entry.id);
    if (!jawi) return;
    unique.get(entry.text).forEach(loc => {
      tabs[loc.tab].values[loc.r][loc.c] = jawi;
      changedTabs.add(loc.tab);
      changedCells++;
    });
  });

  changedTabs.forEach(tabIndex => {
    const tab = tabs[tabIndex];
    const rows = tab.values.length - 1;
    JAWI_CURRICULUM_COLUMNS.forEach(c => {
      const column = [];
      for (let r = 1; r < tab.values.length; r++) column.push([tab.formulas[r][c] || tab.values[r][c]]);
      tab.sheet.getRange(tab.range.getRow() + 1, tab.range.getColumn() + c, rows, 1).setValues(column);
    });
  });
  SpreadsheetApp.flush();

  const remaining = entries.length - converted.map.size;
  ui.alert(
    remaining ? 'Penukaran separa — jalankan semula untuk menyambung' : 'Penukaran DSKP/Sukatan selesai',
    'Sel ditukar: ' + changedCells + '\n' +
    'Teks unik ditukar: ' + converted.map.size + ' / ' + entries.length + '\n' +
    (remaining ? 'Baki belum ditukar: ' + remaining + '\n' : '') +
    (converted.stopReason ? 'Sebab berhenti: ' + converted.stopReason + '\n' : '') +
    (missingTabs.length ? 'Tab tidak ditemui: ' + missingTabs.join(', ') + '\n' : ''),
    ui.ButtonSet.OK
  );
}

function needsJawi_(value) {
  const text = String(value).trim();
  if (!text || !/[A-Za-z]/.test(text) || hasArabicScript_(text)) return false;
  if (/^(https?:\/\/|www\.)/i.test(text)) return false;
  return !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text);
}

function transliterateJawiWithRetry_(entries, deadline, progressWorkbook) {
  const CHUNK_SIZE = 25;
  const PARALLEL_CHUNKS = 2;
  const MAX_ATTEMPTS = 3;
  const chunks = [];
  for (let i = 0; i < entries.length; i += CHUNK_SIZE) chunks.push(entries.slice(i, i + CHUNK_SIZE));
  const map = new Map();
  let stopReason = '';

  for (let start = 0; start < chunks.length && !stopReason; start += PARALLEL_CHUNKS) {
    if (Date.now() > deadline) { stopReason = 'had masa tercapai'; break; }
    const group = chunks.slice(start, start + PARALLEL_CHUNKS);
    let results = null;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS && !results; attempt++) {
      try {
        results = callGeminiBatch_(group.map(buildJawiPrompt_));
      } catch (error) {
        const message = String(error && error.message || error);
        const quota = /quota|429|rate.?limit|RESOURCE_EXHAUSTED/i.test(message);
        if (!quota || attempt === MAX_ATTEMPTS || Date.now() + 45000 > deadline) {
          stopReason = quota ? 'kuota AI habis (tunggu seminit, kemudian jalankan semula)' : message;
          break;
        }
        Utilities.sleep(45000);
      }
    }
    if (!results) break;

    results.forEach((result, index) => {
      const items = result && Array.isArray(result.items) ? result.items : [];
      const returned = new Map(items
        .filter(item => item && Number.isFinite(Number(item.id)))
        .map(item => [Number(item.id), String(item.text == null ? '' : item.text)]));
      group[index].forEach(entry => {
        const jawi = returned.get(entry.id);
        if (jawi && hasArabicScript_(jawi)) map.set(entry.id, jawi);
      });
    });
    if (progressWorkbook) {
      progressWorkbook.toast('Kemajuan: ' + map.size + ' / ' + entries.length + ' teks', 'Smart eRPH AI — Jawi', 5);
    }
  }
  return { map, stopReason };
}

function shouldTransliterateToJawi_(value) {
  const trimmed = String(value).trim();
  if (!trimmed) return false;
  if (!/[A-Za-z]/.test(trimmed)) return false;
  if (/^(https?:\/\/|www\.)/i.test(trimmed)) return false;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) return false;
  return true;
}

function transliterateJawiEntries_(entries, progressWorkbook) {
  const CHUNK_SIZE = 35;
  const PARALLEL_CHUNKS = 4;
  const chunks = [];
  for (let i = 0; i < entries.length; i += CHUNK_SIZE) chunks.push(entries.slice(i, i + CHUNK_SIZE));
  const resultMap = new Map();

  for (let start = 0; start < chunks.length; start += PARALLEL_CHUNKS) {
    const group = chunks.slice(start, start + PARALLEL_CHUNKS);
    const prompts = group.map(buildJawiPrompt_);
    const results = callGeminiBatch_(prompts);

    results.forEach((result, resultIndex) => {
      const expectedChunk = group[resultIndex];
      const items = result && Array.isArray(result.items) ? result.items : [];
      const returned = new Map(
        items
          .filter(item => item && Number.isFinite(Number(item.id)))
          .map(item => [Number(item.id), String(item.text == null ? '' : item.text)])
      );
      expectedChunk.forEach(item => {
        const converted = returned.get(item.id);
        if (!converted) throw new Error('AI tidak memulangkan teks Jawi lengkap bagi ID ' + item.id + '. Sila jalankan semula.');
        resultMap.set(item.id, converted);
      });
    });

    if (progressWorkbook) {
      const done = Math.min(start + PARALLEL_CHUNKS, chunks.length);
      progressWorkbook.toast('Kemajuan Jawi: ' + done + '/' + chunks.length + ' kelompok selesai', 'Smart eRPH AI — Jawi', 5);
    }
  }
  return resultMap;
}

function buildJawiPrompt_(items) {
  return [
    'Anda ialah pakar transliterasi tulisan Jawi Bahasa Melayu Malaysia.',
    '',
    'TUGAS',
    'Tukar teks Rumi dalam setiap item kepada tulisan Jawi berdasarkan Pedoman Ejaan Jawi Yang Disempurnakan. Ini ialah transliterasi, BUKAN terjemahan.',
    '',
    'PERATURAN WAJIB',
    '1. Kekalkan maksud dan susunan ayat asal.',
    '2. Kekalkan nombor, tarikh, masa, tanda baca, simbol, baris baharu dan susunan senarai.',
    '3. Jika sebahagian teks sudah Arab/Jawi, kekalkan bahagian itu dan tukar hanya bahagian Rumi yang sesuai.',
    '4. Jangan tambah penerangan, nota atau markdown.',
    '5. Nama orang/tempat dan istilah pinjaman ditulis dalam Jawi mengikut sebutan yang munasabah.',
    '6. Akronim/kod teknikal seperti DSKP, KSSM, PBD, PAK21, KBAT, KBKK, EMK, AI, KKQ, eRPH, URL dan kod standard boleh dikekalkan dalam Rumi jika penukaran akan menghilangkan identiti kod.',
    '7. Pulangkan SEMUA ID sekali dan jangan ubah ID.',
    '',
    'PULANGKAN JSON SAHAJA dengan struktur tepat:',
    '{"items":[{"id":1,"text":"..."},{"id":2,"text":"..."}]}',
    '',
    'DATA:',
    JSON.stringify(items)
  ].join('\n');
}

function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt(
    'Tetapkan API Gemini',
    'Tampalkan API key daripada projek Google AI Studio yang menggunakan Tier 1. Key disimpan dalam Script Properties projek ini.',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = answer.getResponseText().trim();
  if (!apiKey) {
    ui.alert('API key tidak dimasukkan.');
    return;
  }

  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('GEMINI_API_KEY', apiKey);

  // Bersihkan tetapan provider lama supaya projek ini menggunakan Gemini sahaja.
  properties.deleteProperty('GEMINI_API_KEY_BACKUP');
  properties.deleteProperty('OPENROUTER_API_KEY');
  properties.deleteProperty('OPENROUTER_MODEL');
  properties.deleteProperty('AI_PROVIDER');

  ui.alert('API Gemini Tier 1 telah disimpan. Smart eRPH kini menggunakan Gemini sahaja.');
}

function testGeminiApi() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = callGeminiBatch_([
      'Pulangkan JSON sahaja: {"status":"ok","mesej":"Gemini aktif"}'
    ])[0];
    ui.alert(
      'Sambungan Gemini berjaya',
      'Model Gemini boleh dicapai menggunakan GEMINI_API_KEY semasa.\n\n' +
      JSON.stringify(result),
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Sambungan Gemini gagal',
      error.message || String(error),
      ui.ButtonSet.OK
    );
  }
}

function setErphSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Buka fail eRPH Jawi terlebih dahulu.');
  PropertiesService.getScriptProperties().setProperty('ERPH_JAWI_SPREADSHEET_ID', ss.getId());
  SpreadsheetApp.getUi().alert('Fail eRPH Jawi ini telah disambungkan kepada PWA.');
}

// Pembetulan sekali jalan untuk rekod lama. Dalam template, M(slot+1) ialah label
// "STRATEGI P&P" dan M(slot+2) ialah nilai strategi yang sebenar.
function fillMissingStudentCentredStrategies() {
  let updated = 0;
  const missingTabs = [];
  ERPH.DAYS.forEach(day => {
    const sheet = getWorksheet_(day);
    if (!sheet) {
      missingTabs.push(day);
      return;
    }
    getSlotsFromSheet_(sheet).forEach(slot => {
      const slotStartRow = Number(slot.value);
      sheet.getRange(slotStartRow + 1, 13).setValue('ستراتيݢي ڤڠاجرن دان ڤمبلاجرن');
      const strategyCell = sheet.getRange(slotStartRow + 2, 13);
      if (!String(strategyCell.getDisplayValue()).trim()) {
        strategyCell.setValue(ERPH.STUDENT_CENTRED_STRATEGY_JAWI);
        updated++;
      }
    });
  });
  const note = missingTabs.length ? ` Tab tidak ditemui: ${missingTabs.join(', ')}.` : '';
  SpreadsheetApp.getUi().alert(`Selesai: ${updated} slot kosong telah diisi dengan strategi pembelajaran berpusatkan murid dan kolaboratif.${note}`);
}

function getWorkbook_() {
  const properties = PropertiesService.getScriptProperties();
  const jawiId = properties.getProperty('ERPH_JAWI_SPREADSHEET_ID');
  if (jawiId) return SpreadsheetApp.openById(jawiId);

  // Fallback hanya ketika kod dijalankan dari fail Jawi yang terikat.
  const active = SpreadsheetApp.getActive();
  if (active) return active;

  throw new Error('Fail eRPH Jawi belum disambungkan. Buka Sheet Jawi dan pilih Smart eRPH AI — Jawi > Sambungkan fail Jawi ini untuk PWA.');
}

function getWorksheet_(expectedName, workbook) {
  workbook = workbook || getWorkbook_();
  const exact = workbook.getSheetByName(expectedName);
  if (exact) return exact;
  const normalise = name => String(name).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const expected = normalise(expectedName);
  return workbook.getSheets().find(sheet => normalise(sheet.getName()) === expected) || null;
}

function getSlotsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const labels = sheet.getRange(1, 2, lastRow, 1).getDisplayValues().flat();
  const formulas = sheet.getRange(1, 4, lastRow, 1).getFormulas().flat();
  const starts = labels
    .map((label, index) => ({ label: label.trim().toUpperCase(), row: index + 1 }))
    .filter(item => ['TARIKH', 'تاريخ'].includes(item.label) && formulas[item.row - 1]);
  if (!starts.length) throw new Error(`Tiada slot eRPH ditemui dalam tab ${sheet.getName()}.`);
  return starts.map((item, index) => ({ value: item.row, label: `Slot ${index + 1} — rekod kelas ${index + 1}` }));
}

function generateAllActiveSlots() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  if (!ERPH.DAYS.includes(sheet.getName())) throw new Error('Buka salah satu tab ISNIN hingga JUMAAT dahulu.');
  const date = coerceSheetDate_(sheet.getRange('D11').getValue());
  const expectedSheetName = ERPH.DAY_BY_INDEX[date.getDay()].toUpperCase();
  if (expectedSheetName !== sheet.getName()) {
    throw new Error(`Tarikh di D11 ialah ${Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy')}, tetapi tab aktif ialah ${sheet.getName()}. Sila betulkan tarikh dahulu.`);
  }
  const outputScript = 'jawi';
  const week = Number(sheet.getRange('D7').getValue());
  if (!week) throw new Error('Masukkan nombor minggu pada sel D7 dahulu.');

  const jobs = getSlotsFromSheet_(sheet).map(slot => buildAutoJob_(sheet, Number(slot.value), week, date, outputScript));
  const results = callGeminiBatch_(jobs.map(job => buildPrompt_(job.form, job.curriculum, outputScript)));
  results.forEach(result => validateGeneratedOutput_(result, outputScript));
  jobs.forEach((job, index) => writeErph_(sheet, job.slotStartRow, job.form, job.curriculum, results[index], outputScript));
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(`Selesai: ${jobs.length} slot pada tab ${sheet.getName()} telah dijana.`);
}

function generateWeekFromPwa_(payload) {
  const week = Number(payload.week);
  if (!Number.isInteger(week) || week < 1 || week > 45) throw new Error('Minggu mestilah antara 1 hingga 45.');
  const monday = parseIsoDate_(payload.monday);
  if (monday.getDay() !== 1) throw new Error('Tarikh yang dipilih mestilah hari Isnin.');
  const outputScript = 'jawi';
  const outputWorkbook = getWorkbook_(outputScript);
  const selectedDays = payload.days || {};
  const dayJobs = [];

  ERPH.DAYS.forEach((sheetName, offset) => {
    if (!selectedDays[sheetName]) return;
    const date = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + offset, 12, 0, 0);
    const sheet = getWorksheet_(sheetName, outputWorkbook);
    if (!sheet) throw new Error(`Tab ${sheetName} tidak ditemui dalam fail eRPH ${outputScript === 'jawi' ? 'Jawi' : 'Rumi'}.`);
    sheet.getRange('D7').setValue(week);
    sheet.getRange('D9').setValue(outputDay_(date.getDay(), outputScript));
    sheet.getRange('D11').setValue(date);
    getSlotsFromSheet_(sheet).forEach(slot => {
      const job = buildAutoJob_(sheet, Number(slot.value), week, date, outputScript);
      dayJobs.push({ ...job, sheet, sheetName });
    });
  });
  if (!dayJobs.length) throw new Error('Tandakan sekurang-kurangnya satu hari untuk dijana.');

  const results = callGeminiBatch_(dayJobs.map(job => buildPrompt_(job.form, job.curriculum, outputScript)));
  results.forEach(result => validateGeneratedOutput_(result, outputScript));
  dayJobs.forEach((job, index) => writeErph_(job.sheet, job.slotStartRow, job.form, job.curriculum, results[index], outputScript));
  SpreadsheetApp.flush();
  const summary = ERPH.DAYS.reduce((all, day) => {
    const count = dayJobs.filter(job => job.sheetName === day).length;
    if (count) all[day] = count;
    return all;
  }, {});
  return {
    ok: true,
    week,
    monday: payload.monday,
    outputScript,
    outputLabel: outputScript === 'jawi' ? 'Jawi' : 'Rumi',
    generatedSlots: dayJobs.length,
    summary,
    workbookUrl: outputWorkbook.getUrl()
  };
}

function prepareClassroomFromPwa_(payload) {
  const week = Number(payload.week);
  if (!Number.isInteger(week) || week < 1 || week > 45) throw new Error('Minggu mestilah antara 1 hingga 45.');
  const monday = parseIsoDate_(payload.monday);
  if (monday.getDay() !== 1) throw new Error('Tarikh yang dipilih mestilah hari Isnin.');
  const outputScript = 'jawi';

  const workbook = getWorkbook_(outputScript);
  const mondayLabel = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  const fileName = outputScript === 'jawi'
    ? `eRPH Jawi Minggu ${week} (${mondayLabel})`
    : `eRPH Minggu ${week} (${mondayLabel})`;
  const copy = DriveApp.getFileById(workbook.getId()).makeCopy(fileName);
  return { ok: true, outputScript, fileName: copy.getName(), fileUrl: copy.getUrl() };
}

function buildAutoJob_(sheet, slotStartRow, week, date, outputScript) {
  const rawSubject = String(sheet.getRange(slotStartRow + 3, 4).getDisplayValue()).trim();
  const rawForm = sheet.getRange(slotStartRow + 2, 4).getDisplayValue();
  const form = {
    week,
    date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    slotStartRow,
    form: parseFormNumber_(rawForm),
    className: String(sheet.getRange(slotStartRow + 2, 5).getDisplayValue()).trim(),
    subject: canonicalSubject_(rawSubject),
    startTime: String(sheet.getRange(slotStartRow + 1, 5).getDisplayValue()).trim(),
    endTime: String(sheet.getRange(slotStartRow + 1, 9).getDisplayValue()).trim()
  };
  const missing = [];
  if (!form.startTime) missing.push(`masa mula (E${slotStartRow + 1})`);
  if (!form.endTime) missing.push(`masa tamat (I${slotStartRow + 1})`);
  if (!form.form) missing.push(`Tingkatan (D${slotStartRow + 2}${rawForm ? `: "${rawForm}" bukan nombor` : ''})`);
  if (!form.className) missing.push(`kelas (E${slotStartRow + 2})`);
  if (!form.subject) missing.push(`mata pelajaran (D${slotStartRow + 3}${rawSubject ? `: "${rawSubject}" tidak dikenali sebagai PAI/KKQ` : ''})`);
  if (missing.length) {
    throw new Error(`Slot bermula baris ${slotStartRow} di tab ${sheet.getName()} belum lengkap: ${missing.join('; ')}.`);
  }
  const existingTitle = String(sheet.getRange(slotStartRow + 5, 4).getDisplayValue()).trim();
  const curriculum = findCurriculumByExistingTitle_(form.subject, form.form, existingTitle, outputScript) ||
    findCurriculumByWeek_(form.subject, form.form, week, outputScript);
  return { slotStartRow, form, curriculum };
}

function canonicalSubject_(value) {
  const subject = String(value || '').trim();
  const key = normaliseForMatch_(subject);

  const paiAliases = [
    ERPH.SUBJECTS.PAI, 'Pendidikan Islam', 'Pendidikan Agama Islam', 'PAI',
    ERPH.SUBJECTS_JAWI.PAI, ...ERPH.SUBJECT_ALIASES_JAWI.PAI
  ].map(normaliseForMatch_);

  const kkqAliases = [
    ERPH.SUBJECTS.KKQ, 'Kelas Kemahiran al-Quran', 'KKQ',
    ERPH.SUBJECTS_JAWI.KKQ, ...ERPH.SUBJECT_ALIASES_JAWI.KKQ
  ].map(normaliseForMatch_);

  if (!key) return '';
  if (paiAliases.includes(key)) return ERPH.SUBJECTS.PAI;
  if (kkqAliases.includes(key)) return ERPH.SUBJECTS.KKQ;

  // Padanan longgar untuk ejaan Jawi/Rumi yang sedikit berbeza.
  if (/kkq|ككق|قران|قرءان|quran/.test(key)) return ERPH.SUBJECTS.KKQ;
  if (/pai|ڤاي|islam|اسلام/.test(key)) return ERPH.SUBJECTS.PAI;
  return '';
}

function normaliseForMatch_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[ً-ٰٟـ​-‏﻿]/g, '')
    .replace(/[أإآٱ]/g, 'ا')
    .replace(/[ىی]/g, 'ي')
    .replace(/[کڪ]/g, 'ك')
    .replace(/[ەۀة]/g, 'ه')
    .replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 0x0660))
    .replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 0x06F0))
    .replace(/[^a-z0-9؀-ۿ]+/g, '');
}

function parseFormNumber_(value) {
  const digits = normaliseForMatch_(value).match(/\d+/);
  return digits ? Number(digits[0]) : 0;
}

function findCurriculumByWeek_(subject, form, week, outputScript) {
  const options = getCurriculumOptions(subject, form, outputScript).options;
  const match = options.find(item => {
    const numbers = String(item.week).match(/\d+/g);
    if (!numbers) return false;
    const range = numbers.map(Number);
    return week >= Math.min(...range) && week <= Math.max(...range);
  });
  if (!match) throw new Error(`Tiada tajuk ${subject} Tingkatan ${form} untuk Minggu ${week} dalam tab data.`);
  return match;
}

function findCurriculumByExistingTitle_(subject, form, existingTitle, outputScript) {
  if (!existingTitle) return null;
  const normalise = text => String(text || '')
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\u064B-\u065F\u0670\u200B-\u200D\uFEFF]/g, '')
    .replace(/[^a-z0-9\u0600-\u06FF]+/g, '');
  const target = normalise(existingTitle);
  if (!target) return null;
  return getCurriculumOptions(subject, form, outputScript).options.find(item =>
    [item.title, item.alias].some(value => normalise(value) === target)
  ) || null;
}

function getCurriculumOptions(subject, form, outputScript) {
  const sourceName = getSourceSheetName_(subject, Number(form));
  const sheet = getWorksheet_(sourceName, getWorkbook_(outputScript));
  if (!sheet) throw new Error(`Tab data ${sourceName} tidak ditemui dalam fail eRPH yang disambungkan.`);

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  const activeKeys = ERPH.ACTIVE_STATUS_ALIASES.map(normaliseForMatch_);
  const rows = values.filter(row =>
    row[0] && activeKeys.includes(normaliseForMatch_(row[13]))
  );
  return {
    sourceName,
    headers,
    options: rows.map(row => ({
      lesson: row[2], week: row[3], field: row[4], title: row[5], alias: row[6],
      code: row[7], standardContent: row[8], standardLearning: row[9],
      suggestedObjectives: row[10], keywords: row[11], source: row[12]
    }))
  };
}

function getSourceSheetName_(subject, form) {
  if (subject === ERPH.SUBJECTS.PAI && [4, 5].includes(form)) return `DSKP_PAI_T${form}`;
  if (subject === ERPH.SUBJECTS.KKQ && [1, 2, 3].includes(form)) return `SUKATAN_KKQ_T${form}`;
  throw new Error('Padanan tidak dibenarkan: PAI hanya Tingkatan 4–5, manakala KKQ hanya Tingkatan 1–3.');
}

function getTargetSheet_(dateText) {
  const date = parseIsoDate_(dateText);
  const day = ERPH.DAY_BY_INDEX[date.getDay()].toUpperCase();
  if (!ERPH.DAYS.includes(day)) throw new Error('Tarikh yang dipilih ialah hujung minggu. Sila pilih tarikh Isnin hingga Jumaat.');
  const sheet = getWorksheet_(day);
  if (!sheet) throw new Error(`Tab ${day} tidak ditemui dalam fail eRPH yang disambungkan.`);
  return sheet;
}

function buildPrompt_(form, c, outputScript) {
  const mode = normaliseOutputScript_(outputScript);
  const outputInstruction = mode === 'jawi'
    ? `OUTPUT JAWI — WAJIB
Semua nilai teks yang anda jana hendaklah dalam tulisan Jawi Bahasa Melayu Malaysia berdasarkan Pedoman Ejaan Jawi Yang Disempurnakan.
Jangan terjemah maksud kurikulum. Rujukan kurikulum di bawah kebiasaannya sudah dalam Jawi: salin tepat ke medan theme, title, standardContent dan standardLearning tanpa mengubah ejaan. Jika sesuatu rujukan masih Rumi, transliterasikannya dengan setepat mungkin tanpa menambah atau membuang fakta.
Kekalkan nombor, kod standard, PAI, KKQ, DSKP, KSSM, PBD, PA21, KBAT, KBKK, EMK, i-THINK dan nama aktiviti antarabangsa jika perlu.
PENTING: medan assessmentTypes ialah kod sistem. Nilainya MESTI kekal dalam Rumi dan hanya boleh menggunakan: "Amali / Eksperimen", "Projek", "Pembentangan", "Ujian", "Peperiksaan", "Latihan / Kerja Rumah", "Lembaran Kerja", "Pemerhatian", "Kuiz", "Lisan", "Tugasan".`
    : 'OUTPUT RUMI — Gunakan Bahasa Melayu Rumi yang kemas dan standard.';

  return `Anda ialah guru Pendidikan Islam KSSM Malaysia. Hasilkan eRPH Bahasa Melayu yang praktikal, tepat dan selaras dengan DSKP/sukatan yang diberi. Jangan cipta Standard Kandungan atau Standard Pembelajaran baharu. Gunakan 3 item ringkas bagi objektif, kriteria kejayaan dan setiap fasa aktiviti. Aktiviti mestilah sesuai untuk tempoh masa ${form.startTime} hingga ${form.endTime}, berpusatkan murid, beradab dan boleh dilaksanakan.

${outputInstruction}

MAKLUMAT KELAS
Minggu: ${form.week}
Tarikh: ${form.date}
Tingkatan: ${form.form}
Kelas: ${form.className}
Mata pelajaran: ${form.subject}
Masa: ${form.startTime} hingga ${form.endTime}

RUJUKAN KURIKULUM — WAJIB KEKAL TEPAT
Bidang/Tema: ${c.field}
Tajuk: ${c.title}
Kod SK: ${c.code}
Standard Kandungan: ${c.standardContent}
Standard Pembelajaran: ${c.standardLearning}
Cadangan objektif: ${c.suggestedObjectives}
Kata kunci: ${c.keywords}
Sumber: ${c.source}

Untuk medan "strategy", pilih SATU strategi pengajaran berpusatkan murid dan kolaboratif yang PALING SESUAI dengan tajuk, kemahiran dan aktiviti di atas (contoh: Pembelajaran Berasaskan Inkuiri, Pembelajaran Berasaskan Projek, Pembelajaran Koperatif, Pembelajaran Berasaskan Masalah, Think-Pair-Share, Round Robin, Perbincangan Berkumpulan, Simulasi/Main Peranan). Jangan ulang frasa generik yang sama setiap kali — nyatakan strategi khusus yang relevan dengan kandungan pelajaran ini.

PULANGKAN JSON SAHAJA, tanpa markdown, dengan skema tepat ini:
{
 "theme":"", "title":"", "skill":"", "standardContent":"", "standardLearning":"",
 "objectives":["","",""], "successCriteria":["","",""],
 "starter":["","",""], "activity":["","",""], "explanation":["","",""], "closure":["","",""], "assessmentDetails":["","",""],
 "references":"", "reflection":"", "followUp":"", "strategy":"", "method":"", "teachingAids":"", "pa21":"", "kbkk":"", "iThink":"", "values":"", "thinkingSkill":"", "multipleIntelligences":"", "kbatCurriculum":"", "kbatCoCurriculum":"", "emk":"", "assessmentTypes":["Pemerhatian","Lisan","Kuiz"]
}`;
}

function callGeminiBatch_(prompts) {
  // Versi ini menggunakan Gemini sahaja.
  return callGeminiApiBatch_(prompts);
}

function callGeminiApiBatch_(prompts) {
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty('GEMINI_API_KEY');
  const model = properties.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';

  if (!apiKey) {
    throw new Error('GEMINI_API_KEY belum disetkan dalam Script Properties.');
  }

  // Elakkan lonjakan permintaan apabila satu minggu mempunyai banyak slot.
  // Maksimum 4 request serentak, kemudian jeda sebelum kelompok seterusnya.
  const MAX_PARALLEL = 4;
  const GROUP_PAUSE_MS = 1200;
  const MAX_RETRIES = 2;
  const results = new Array(prompts.length);

  const buildRequest_ = prompt => ({
    url: `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        responseMimeType: 'application/json',
        temperature: 0.35
      }
    }),
    muteHttpExceptions: true
  });

  for (let start = 0; start < prompts.length; start += MAX_PARALLEL) {
    let pending = prompts
      .slice(start, start + MAX_PARALLEL)
      .map((prompt, offset) => ({
        index: start + offset,
        prompt: prompt
      }));

    let attempt = 0;

    while (pending.length) {
      const responses = UrlFetchApp.fetchAll(
        pending.map(item => buildRequest_(item.prompt))
      );

      const retryItems = [];
      const retryResponses = [];

      responses.forEach((response, i) => {
        const item = pending[i];
        const status = response.getResponseCode();

        if ((status === 429 || status === 503) && attempt < MAX_RETRIES) {
          retryItems.push(item);
          retryResponses.push(response);
          return;
        }

        results[item.index] = parseGeminiResponse_(response, model);
      });

      if (!retryItems.length) break;

      const waitMs = getGeminiRetryDelayMs_(retryResponses, attempt);
      Utilities.sleep(waitMs);
      pending = retryItems;
      attempt++;
    }

    if (start + MAX_PARALLEL < prompts.length) {
      Utilities.sleep(GROUP_PAUSE_MS);
    }
  }

  return results;
}

function getGeminiRetryDelayMs_(responses, attempt) {
  let delayMs = Math.min(30000, 2000 * Math.pow(2, attempt));

  responses.forEach(response => {
    try {
      const bodyText = response.getContentText();
      const match =
        bodyText.match(/retry in\s+([0-9.]+)s/i) ||
        bodyText.match(/"retryDelay"\s*:\s*"([0-9.]+)s"/i);

      if (match) {
        delayMs = Math.max(delayMs, Math.ceil(Number(match[1]) * 1000) + 500);
      }
    } catch (error) {
      // Gunakan exponential fallback.
    }
  });

  return Math.min(delayMs, 35000);
}

function parseGeminiResponse_(response, model) {
  const status = response.getResponseCode();
  let body;

  try {
    body = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error(`Gemini (${model}) memulangkan respons bukan JSON. HTTP ${status}.`);
  }

  if (status >= 300) {
    const message = body.error?.message || `Gemini gagal menjana eRPH. HTTP ${status}.`;

    if (status === 429) {
      throw new Error(
        `Gemini (${model}) masih terkena had quota/rate limit selepas retry. ` +
        `Pastikan GEMINI_API_KEY benar-benar milik projek Tier 1. Butiran: ${message}`
      );
    }

    throw new Error(message);
  }

  const text = body.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error(`Gemini (${model}) tidak memulangkan kandungan eRPH.`);

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Respons Gemini (${model}) bukan JSON yang sah. Sila jana semula.`);
  }
}

function writeErph_(sheet, slotStartRow, form, c, g, outputScript) {
  const mode = normaliseOutputScript_(outputScript);
  const isJawi = mode === 'jawi';
  const date = parseIsoDate_(form.date);
  const day = outputDay_(date.getDay(), mode);
  const subject = outputSubject_(form.subject, mode);
  const put = (cell, value) => sheet.getRange(cell).setValue(value || '');
  const putSlot = (rowOffset, column, value) => sheet.getRange(slotStartRow + rowOffset, column).setValue(value || '');
  const putSlotList = (rowOffsets, values) => rowOffsets.forEach((offset, i) => putSlot(offset, 4, (values || [])[i] || ''));

  // Identiti dan jadual — hanya tab hari yang sepadan dengan tarikh akan disentuh.
  put('D7', Number(form.week)); put('D9', day); put('D11', date);
  putSlot(1, 5, form.startTime); putSlot(1, 9, form.endTime); putSlot(2, 4, Number(form.form)); putSlot(2, 5, form.className); putSlot(3, 4, subject);

  // Mod Jawi: guna teks kurikulum terus jika tab DSKP/Sukatan Jawi sudah ditukar; jika masih Rumi, guna transliterasi AI.
  const curriculumText = (source, generated) => (isJawi && !hasArabicScript_(source) ? (generated || source) : source);
  putSlot(4, 4, isJawi ? (g.theme || curriculumText(c.field, g.theme)) : c.field);
  putSlot(5, 4, isJawi ? (g.title || curriculumText(c.title, g.title)) : c.title);
  putSlot(6, 4, g.skill);
  putSlot(7, 4, curriculumText(c.standardContent, g.standardContent));
  putSlot(8, 4, curriculumText(c.standardLearning, g.standardLearning));
  putSlotList([11, 12, 13], g.objectives); putSlotList([14, 15, 16], g.successCriteria);
  putSlotList([18, 19, 20], g.starter); putSlotList([22, 23, 24], g.activity);
  putSlotList([26, 27, 28], g.explanation); putSlotList([30, 31, 32], g.closure); putSlotList([34, 35, 36], g.assessmentDetails);
  putSlot(37, 4, isJawi && hasArabicScript_(c.source) ? c.source : (g.references || c.source)); putSlot(42, 4, g.reflection); putSlot(47, 4, g.followUp);

  // Lajur sokongan sebelah kanan template.
  putSlot(1, 13, isJawi ? 'ستراتيݢي ڤڠاجرن دان ڤمبلاجرن' : 'STRATEGI P&P');
  putSlot(2, 13, String(g.strategy || '').trim() || (isJawi ? ERPH.STUDENT_CENTRED_STRATEGY_JAWI : ERPH.STUDENT_CENTRED_STRATEGY));
  putSlot(5, 13, g.method); putSlot(8, 13, g.teachingAids); putSlot(11, 13, g.pa21);
  putSlot(15, 13, g.kbkk); putSlot(17, 13, g.iThink); putSlot(20, 13, g.values); putSlot(23, 13, g.thinkingSkill);
  putSlot(26, 13, g.multipleIntelligences); putSlot(31, 13, g.kbatCurriculum); putSlot(33, 13, g.kbatCoCurriculum); putSlot(36, 13, g.emk);

  Object.values(ERPH.ASSESSMENT_ROWS).forEach(row => sheet.getRange(slotStartRow + row - 14, 16).setValue(false));
  (g.assessmentTypes || []).forEach(type => {
    const row = ERPH.ASSESSMENT_ROWS[type];
    if (row) sheet.getRange(slotStartRow + row - 14, 16).setValue(true);
  });
}

function normaliseOutputScript_(value) {
  // Versi ini khusus untuk workbook Jawi penuh.
  return 'jawi';
}

function detectOutputScript_(sheet) {
  return 'jawi';
}

function outputDay_(dayIndex, outputScript) {
  return normaliseOutputScript_(outputScript) === 'jawi'
    ? ERPH.DAY_JAWI_BY_INDEX[dayIndex]
    : ERPH.DAY_BY_INDEX[dayIndex];
}

function outputSubject_(subject, outputScript) {
  if (subject === ERPH.SUBJECTS.PAI) return ERPH.SUBJECTS_JAWI.PAI;
  if (subject === ERPH.SUBJECTS.KKQ) return ERPH.SUBJECTS_JAWI.KKQ;
  return subject;
}

function validateGeneratedOutput_(generated, outputScript) {
  if (normaliseOutputScript_(outputScript) !== 'jawi') return;
  const sample = [
    generated.theme, generated.title, generated.skill,
    generated.standardContent, generated.standardLearning,
    ...(generated.objectives || []), ...(generated.activity || []),
    generated.reflection, generated.followUp, generated.strategy
  ].filter(Boolean).join(' ');
  if (!/[\u0600-\u06FF]/.test(sample)) {
    throw new Error('AI tidak menghasilkan tulisan Jawi. Tiada perubahan ditulis; sila jana semula.');
  }
}

function parseIsoDate_(value) {
  const parts = String(value).split('-').map(Number);
  if (parts.length !== 3 || parts.some(Number.isNaN)) throw new Error('Format tarikh tidak sah.');
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
}

function coerceSheetDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  throw new Error('Sel D11 perlu mengandungi tarikh yang sah.');
}

function verifyPwaToken_(apiToken) {
  const expected = PropertiesService.getScriptProperties().getProperty('PWA_API_TOKEN');
  if (!expected) throw new Error('PWA_API_TOKEN belum ditetapkan dalam Script Properties.');
  if (!apiToken || apiToken !== expected) throw new Error('Akses PWA tidak dibenarkan.');
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}