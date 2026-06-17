const SHEET_IDS = {
  jump:    '1D45RvGaOw_4YXY8AXVwyvcpZnjno2U2jeL-EgVaJ-lw',
  gym:     '1sr1Fj4FXcx7Ba_xVVafFNs2Yey5TXUlZ1ONX_nAXY2I',
  fitness: '1vZwafXh0s9fEnOz6IgxgKNLhwhmPckRpC9vdsA1htO4',
  tt:      '1p5sf87h7UVMtSV3izTmeZsEwOmF7JwEf7j0psNDzZw0',
  ath:     '1RAQjdoFAjWtwH35Ulh5o5UqMptFsdCxNCjccHyG8PfQ',
  bad:     '1Vr8AgVv6e9swNVOsr2eYPNwfQff4CMn9UCWPQI7zs3o',
  pkl:     '1yFtkhjKQPKPYWI-8pNo1kGDodXuX9egbsX5bzSH0QgQ',
  bball:   '1vRGtCn1x3Poexdw-cM7RPu7AI2lVDDhSk5lAiapI8FY',
};

const TIMESLOT_MAP = {
  jump: [
    { match: '星期一', id: 'quota-jump-mon' },
    { match: '星期二', id: 'quota-jump-tue' },
    { match: '星期三', id: 'quota-jump-wed' },
    { match: '星期四', id: 'quota-jump-thu' },
  ],
  gym: [
    { match: '星期一', id: 'quota-gym-mon' },
    { match: '星期四', id: 'quota-gym-thu' },
  ],
  fitness: [
    { match: '星期一', id: 'quota-fit-mon' },
    { match: '星期五', id: 'quota-fit-fri' },
  ],
  tt: [
    { match: '星期三', id: 'quota-tt-wed' },
    { match: '星期五', id: 'quota-tt-fri' },
  ],
  ath: [
    { match: '星期二', id: 'quota-ath-tue' },
    { match: '星期四', id: 'quota-ath-thu' },
    { match: '星期六', id: 'quota-ath-sat' },
  ],
  bad: [
    { match: '星期二', id: 'quota-bad-tue' },
    { match: '星期四', id: 'quota-bad-thu' },
  ],
  pkl: [
    { match: '低小', id: 'quota-pkl-tue1' },
    { match: '高小', id: 'quota-pkl-tue2' },
    { match: '星期五', id: 'quota-pkl-fri' },
  ],
  bball: [
    { match: '星期三', id: 'quota-bball-wed' },
    { match: '星期五', id: 'quota-bball-fri' },
  ],
};

function doGet(e) {
  const result = {};
  for (const [courseKey, sheetId] of Object.entries(SHEET_IDS)) {
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      const sheet = ss.getSheets()[0];
      const data = sheet.getDataRange().getValues();
      const slots = TIMESLOT_MAP[courseKey] || [];
      slots.forEach(s => result[s.id] = 0);
      if (data.length < 2) continue;
      const headers = data[0];
      let timeCol = -1;
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]);
        if (h.includes('上課時間') || h.includes('時段') || h.includes('班別')) {
          timeCol = i;
          break;
        }
      }
      if (timeCol === -1) continue;
      for (let row = 1; row < data.length; row++) {
        const answer = String(data[row][timeCol]).trim();
        if (!answer) continue;
        for (const slot of slots) {
          if (answer.includes(slot.match)) {
            result[slot.id]++;
            break;
          }
        }
      }
    } catch(e) {
      Logger.log('Error: ' + courseKey + ' - ' + e);
    }
  }
  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function setupEmailNotifications() {
  let count = 0;
  for (const [courseKey, sheetId] of Object.entries(SHEET_IDS)) {
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      const formUrl = ss.getFormUrl();
      if (!formUrl) { Logger.log('No form: ' + courseKey); continue; }
      const form = FormApp.openByUrl(formUrl);
      const triggers = ScriptApp.getUserTriggers(form);
      let has = false;
      for (const t of triggers) {
        if (t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) { has = true; break; }
      }
      if (!has) {
        ScriptApp.newTrigger('sendEmailOnSubmit').forForm(form).onFormSubmit().create();
        Logger.log('Set: ' + form.getTitle());
        count++;
      } else {
        Logger.log('Already set: ' + form.getTitle());
      }
    } catch(e) { Logger.log('Error: ' + courseKey + ' - ' + e); }
  }
  Logger.log('Done. Added ' + count + ' triggers.');
}

function sendEmailOnSubmit(e) {
  const formTitle = e.source.getTitle();
  const responses = e.response.getItemResponses();
  let body = '【有新報名！】\n\n表格：' + formTitle + '\n報名時間：' + new Date().toLocaleString('zh-HK') + '\n\n回覆內容：\n';
  for (const r of responses) {
    body += r.getItem().getTitle() + '：' + r.getResponse() + '\n';
  }
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: '【INITIATE SPORTS】新報名 - ' + formTitle,
    body: body
  });
}