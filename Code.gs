const ERPH = Object.freeze({
  DAYS: ['ISNIN', 'SELASA', 'RABU', 'KHAMIS', 'JUMAAT'],
  DAY_BY_INDEX: ['Ahad', 'Isnin', 'Selasa', 'Rabu', 'Khamis', 'Jumaat', 'Sabtu'],
  SUBJECTS: {
    PAI: 'Pendidikan Agama Islam (PAI)',
    KKQ: 'Kelas Kemahiran al-Quran (KKQ)'
  },
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
    if (payload.action !== 'generate_week') throw new Error('Tindakan API tidak sah.');
    return jsonResponse_(generateWeekFromPwa_(payload));
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Smart eRPH AI')
    .addItem('Buka penjana eRPH', 'showErphSidebar')
    .addItem('Jana semua slot tab aktif', 'generateAllActiveSlots')
    .addItem('Tetapkan API Gemini', 'setGeminiApiKey')
    .addToUi();
}

function showErphSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Smart eRPH AI');
  SpreadsheetApp.getUi().showSidebar(html);
}

function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt(
    'Tetapkan API Gemini',
    'Tampalkan API key daripada Google AI Studio. Ia disimpan dalam Script Properties projek ini.',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const apiKey = answer.getResponseText().trim();
  if (!apiKey) {
    ui.alert('API key tidak dimasukkan.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
  ui.alert('API Gemini telah disimpan. Anda kini boleh membuka Smart eRPH AI.');
}

function getSidebarBootstrap() {
  return {
    subjects: ERPH.SUBJECTS,
    activeSheet: SpreadsheetApp.getActiveSheet().getName(),
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

function getSlotsForDate(dateText) {
  const sheet = getTargetSheet_(dateText);
  return { targetSheet: sheet.getName(), slots: getSlotsFromSheet_(sheet) };
}

function getSlotsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const labels = sheet.getRange(1, 2, lastRow, 1).getDisplayValues().flat();
  const formulas = sheet.getRange(1, 4, lastRow, 1).getFormulas().flat();
  const starts = labels
    .map((label, index) => ({ label: label.trim().toUpperCase(), row: index + 1 }))
    .filter(item => item.label === 'TARIKH' && formulas[item.row - 1]);
  if (!starts.length) throw new Error(`Tiada slot eRPH ditemui dalam tab ${sheet.getName()}.`);
  return starts.map((item, index) => ({ value: item.row, label: `Slot ${index + 1} — rekod kelas ${index + 1}` }));
}

function generateAllActiveSlots() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  if (!ERPH.DAYS.includes(sheet.getName())) throw new Error('Buka salah satu tab ISNIN hingga JUMAAT dahulu.');
  const date = coerceSheetDate_(sheet.getRange('D11').getValue());
  const expectedSheet = getTargetSheet_(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  if (expectedSheet.getName() !== sheet.getName()) {
    throw new Error(`Tarikh di D11 ialah ${Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy')}, tetapi tab aktif ialah ${sheet.getName()}. Sila betulkan tarikh dahulu.`);
  }
  const week = Number(sheet.getRange('D7').getValue());
  if (!week) throw new Error('Masukkan nombor minggu pada sel D7 dahulu.');

  const jobs = getSlotsFromSheet_(sheet).map(slot => buildAutoJob_(sheet, Number(slot.value), week, date));
  const results = callGeminiBatch_(jobs.map(job => buildPrompt_(job.form, job.curriculum)));
  jobs.forEach((job, index) => writeErph_(sheet, job.slotStartRow, job.form, job.curriculum, results[index]));
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(`Selesai: ${jobs.length} slot pada tab ${sheet.getName()} telah dijana.`);
}

function generateWeekFromPwa_(payload) {
  const week = Number(payload.week);
  if (!Number.isInteger(week) || week < 1 || week > 45) throw new Error('Minggu mestilah antara 1 hingga 45.');
  const monday = parseIsoDate_(payload.monday);
  if (monday.getDay() !== 1) throw new Error('Tarikh yang dipilih mestilah hari Isnin.');
  const selectedDays = payload.days || {};
  const dayJobs = [];

  ERPH.DAYS.forEach((sheetName, offset) => {
    if (!selectedDays[sheetName]) return;
    const date = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + offset, 12, 0, 0);
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) throw new Error(`Tab ${sheetName} tidak ditemui.`);
    sheet.getRange('D7').setValue(week);
    sheet.getRange('D9').setValue(ERPH.DAY_BY_INDEX[date.getDay()]);
    sheet.getRange('D11').setValue(date);
    getSlotsFromSheet_(sheet).forEach(slot => {
      const job = buildAutoJob_(sheet, Number(slot.value), week, date);
      dayJobs.push({ ...job, sheet, sheetName });
    });
  });
  if (!dayJobs.length) throw new Error('Tandakan sekurang-kurangnya satu hari untuk dijana.');

  const results = callGeminiBatch_(dayJobs.map(job => buildPrompt_(job.form, job.curriculum)));
  dayJobs.forEach((job, index) => writeErph_(job.sheet, job.slotStartRow, job.form, job.curriculum, results[index]));
  SpreadsheetApp.flush();
  const summary = ERPH.DAYS.reduce((all, day) => {
    const count = dayJobs.filter(job => job.sheetName === day).length;
    if (count) all[day] = count;
    return all;
  }, {});
  return { ok: true, week, monday: payload.monday, generatedSlots: dayJobs.length, summary };
}

function buildAutoJob_(sheet, slotStartRow, week, date) {
  const form = {
    week,
    date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    slotStartRow,
    form: Number(sheet.getRange(slotStartRow + 2, 4).getValue()),
    className: String(sheet.getRange(slotStartRow + 2, 5).getDisplayValue()).trim(),
    subject: canonicalSubject_(sheet.getRange(slotStartRow + 3, 4).getDisplayValue()),
    startTime: String(sheet.getRange(slotStartRow + 1, 5).getDisplayValue()).trim(),
    endTime: String(sheet.getRange(slotStartRow + 1, 9).getDisplayValue()).trim()
  };
  if (!form.form || !form.className || !form.subject || !form.startTime || !form.endTime) {
    throw new Error(`Slot bermula baris ${slotStartRow} belum lengkap. Pastikan masa, Tingkatan, kelas dan mata pelajaran telah diisi.`);
  }
  const existingTitle = String(sheet.getRange(slotStartRow + 5, 4).getDisplayValue()).trim();
  const curriculum = findCurriculumByExistingTitle_(form.subject, form.form, existingTitle) ||
    findCurriculumByWeek_(form.subject, form.form, week);
  return { slotStartRow, form, curriculum };
}

function canonicalSubject_(value) {
  const subject = String(value).trim();
  if ([ERPH.SUBJECTS.PAI, 'Pendidikan Islam'].includes(subject)) return ERPH.SUBJECTS.PAI;
  if ([ERPH.SUBJECTS.KKQ, 'KKQ'].includes(subject)) return ERPH.SUBJECTS.KKQ;
  return '';
}

function findCurriculumByWeek_(subject, form, week) {
  const options = getCurriculumOptions(subject, form).options;
  const match = options.find(item => {
    const numbers = String(item.week).match(/\d+/g);
    if (!numbers) return false;
    const range = numbers.map(Number);
    return week >= Math.min(...range) && week <= Math.max(...range);
  });
  if (!match) throw new Error(`Tiada tajuk ${subject} Tingkatan ${form} untuk Minggu ${week} dalam tab data.`);
  return match;
}

function findCurriculumByExistingTitle_(subject, form, existingTitle) {
  if (!existingTitle) return null;
  const normalise = text => String(text).toLowerCase().replace(/[^a-z0-9]+/g, '');
  const target = normalise(existingTitle);
  return getCurriculumOptions(subject, form).options.find(item =>
    [item.title, item.alias].some(value => normalise(value) === target)
  ) || null;
}

function getCurriculumOptions(subject, form) {
  const sourceName = getSourceSheetName_(subject, Number(form));
  const sheet = SpreadsheetApp.getActive().getSheetByName(sourceName);
  if (!sheet) throw new Error(`Tab data ${sourceName} tidak ditemui.`);

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  const rows = values.filter(row => row[0] && row[13].toUpperCase() === 'AKTIF');
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

function generateAndWriteErph(form) {
  validateRequest_(form);
  const target = getTargetSheet_(form.date);
  const slotStartRow = validateSlot_(target, Number(form.slotStartRow));
  const selected = resolveCurriculum_(form.subject, Number(form.form), form.curriculum);
  const generated = callGemini_(buildPrompt_(form, selected));
  writeErph_(target, slotStartRow, form, selected, generated);
  SpreadsheetApp.flush();
  return { ok: true, targetSheet: target.getName(), message: `eRPH berjaya dimasukkan ke tab ${target.getName()}.` };
}

function validateSlot_(sheet, slotStartRow) {
  // Semakan dibuat terus pada tab sasaran supaya slot daripada hari lain tidak boleh digunakan.
  const isSlot = sheet.getRange(slotStartRow, 2).getDisplayValue().trim().toUpperCase() === 'TARIKH' &&
    Boolean(sheet.getRange(slotStartRow, 4).getFormula());
  if (!isSlot) throw new Error('Slot eRPH tidak sah bagi tab hari ini. Sila muat semula sidebar dan pilih slot semula.');
  return slotStartRow;
}

function resolveCurriculum_(subject, form, selectedFromSidebar) {
  const source = getCurriculumOptions(subject, form).options;
  const selected = source.find(item =>
    item.code === selectedFromSidebar.code && item.title === selectedFromSidebar.title
  );
  if (!selected) throw new Error('Tajuk tidak sepadan dengan tab DSKP/sukatan bagi mata pelajaran dan tingkatan yang dipilih.');
  return selected;
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
  const sheet = SpreadsheetApp.getActive().getSheetByName(day);
  if (!sheet) throw new Error(`Tab ${day} tidak ditemui.`);
  return sheet;
}

function validateRequest_(form) {
  if (!form || !form.date || !form.week || !form.form || !form.className || !form.subject || !form.curriculum || !form.slotStartRow) {
    throw new Error('Lengkapkan semua maklumat wajib dahulu.');
  }
  getSourceSheetName_(form.subject, Number(form.form));
  if (![ERPH.SUBJECTS.PAI, ERPH.SUBJECTS.KKQ].includes(form.subject)) {
    throw new Error('Mata pelajaran tidak sah.');
  }
}

function buildPrompt_(form, c) {
  return `Anda ialah guru Pendidikan Islam KSSM Malaysia. Hasilkan eRPH Bahasa Melayu yang praktikal, tepat dan selaras dengan DSKP/sukatan yang diberi. Jangan cipta Standard Kandungan atau Standard Pembelajaran baharu. Gunakan 3 item ringkas bagi objektif, kriteria kejayaan dan setiap fasa aktiviti. Aktiviti mestilah sesuai untuk tempoh masa ${form.startTime} hingga ${form.endTime}, berpusatkan murid, beradab dan boleh dilaksanakan.

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

PULANGKAN JSON SAHAJA, tanpa markdown, dengan skema tepat ini:
{
 "theme":"", "title":"", "skill":"", "standardContent":"", "standardLearning":"",
 "objectives":["","",""], "successCriteria":["","",""],
 "starter":["","",""], "activity":["","",""], "explanation":["","",""], "closure":["","",""], "assessmentDetails":["","",""],
 "references":"", "reflection":"", "followUp":"", "strategy":"", "method":"", "teachingAids":"", "pa21":"", "kbkk":"", "iThink":"", "values":"", "thinkingSkill":"", "multipleIntelligences":"", "kbatCurriculum":"", "kbatCoCurriculum":"", "emk":"", "assessmentTypes":["Pemerhatian","Lisan","Kuiz"]
}`;
}

function callGemini_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY belum disetkan dalam Script Properties.');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${encodeURIComponent(key)}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0.35 } };
  const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  return parseGeminiResponse_(response);
}

function callGeminiBatch_(prompts) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY belum disetkan dalam Script Properties.');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${encodeURIComponent(key)}`;
  const requests = prompts.map(prompt => ({
    url,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0.35 } }),
    muteHttpExceptions: true
  }));
  return UrlFetchApp.fetchAll(requests).map(parseGeminiResponse_);
}

function parseGeminiResponse_(response) {
  const body = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 300) throw new Error(body.error?.message || 'Gemini gagal menjana eRPH.');
  const text = body.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('Gemini tidak memulangkan kandungan eRPH.');
  try { return JSON.parse(text); } catch (e) { throw new Error('Respons Gemini bukan JSON yang sah. Sila jana semula.'); }
}

function writeErph_(sheet, slotStartRow, form, c, g) {
  const date = parseIsoDate_(form.date);
  const day = ERPH.DAY_BY_INDEX[date.getDay()];
  const put = (cell, value) => sheet.getRange(cell).setValue(value || '');
  const putSlot = (rowOffset, column, value) => sheet.getRange(slotStartRow + rowOffset, column).setValue(value || '');
  const putSlotList = (rowOffsets, values) => rowOffsets.forEach((offset, i) => putSlot(offset, 4, (values || [])[i] || ''));

  // Identiti dan jadual — hanya tab hari yang sepadan dengan tarikh akan disentuh.
  put('D7', Number(form.week)); put('D9', day); put('D11', date);
  putSlot(1, 5, form.startTime); putSlot(1, 9, form.endTime); putSlot(2, 4, Number(form.form)); putSlot(2, 5, form.className); putSlot(3, 4, form.subject);

  // DSKP/Sukatan: data asal dikunci daripada pilihan kurikulum, bukan direka Gemini.
  putSlot(4, 4, c.field); putSlot(5, 4, c.title); putSlot(6, 4, g.skill);
  putSlot(7, 4, c.standardContent); putSlot(8, 4, c.standardLearning);
  putSlotList([11, 12, 13], g.objectives); putSlotList([14, 15, 16], g.successCriteria);
  putSlotList([18, 19, 20], g.starter); putSlotList([22, 23, 24], g.activity);
  putSlotList([26, 27, 28], g.explanation); putSlotList([30, 31, 32], g.closure); putSlotList([34, 35, 36], g.assessmentDetails);
  putSlot(37, 4, g.references || c.source); putSlot(42, 4, g.reflection); putSlot(47, 4, g.followUp);

  // Lajur sokongan sebelah kanan template.
  putSlot(1, 13, g.strategy); putSlot(5, 13, g.method); putSlot(8, 13, g.teachingAids); putSlot(11, 13, g.pa21);
  putSlot(15, 13, g.kbkk); putSlot(17, 13, g.iThink); putSlot(20, 13, g.values); putSlot(23, 13, g.thinkingSkill);
  putSlot(26, 13, g.multipleIntelligences); putSlot(31, 13, g.kbatCurriculum); putSlot(33, 13, g.kbatCoCurriculum); putSlot(36, 13, g.emk);

  Object.values(ERPH.ASSESSMENT_ROWS).forEach(row => sheet.getRange(slotStartRow + row - 14, 16).setValue(false));
  (g.assessmentTypes || []).forEach(type => {
    const row = ERPH.ASSESSMENT_ROWS[type];
    if (row) sheet.getRange(slotStartRow + row - 14, 16).setValue(true);
  });
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
