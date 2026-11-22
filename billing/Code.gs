const PLANS = {
  sponsorship: ['تجهيز مشروع للرعاية', 400, 1, 'ملف رعاية ونص صفحة تعريفية وقائمة جهات مستهدفة', 'دفعة واحدة'],
  marketing: ['باقة التسويق السنوي لمشاريع الجمعيات', 700, 12, 'تسويق أربعة مشاريع جاهزة ومعتمدة وتقرير متابعة شهري', '١٢ شهرًا · السداد شهري'],
  consulting: ['باقة الاستشارات وتنمية الموارد المالية', 1500, 12, 'تطوير وتجهيز وتسويق عشرة مشاريع للاستدامة المالية وخطة أول ٩٠ يومًا وتقرير شهري', '١٢ شهرًا · السداد شهري'],
  social: ['إدارة المحتوى والسوشيال ميديا', 899, 1, '١٠ منشورات و٤ قصص وكتابة وتصميم وتقويم محتوى وإدارة النشر', 'شهر واحد']
};
const SHEET = 'الاشتراكات والدفعات';
const PENDING = 'بانتظار التحويل';
const PAID = 'تم السداد';
const ZONE = 'Asia/Riyadh';
const COLUMNS = ['رقم الفاتورة','تاريخ الإصدار','الجمعية','العنوان','البريد','الجوال','الباقة','المبلغ المستحق','إجمالي الباقة','الدفعة','تاريخ الاستحقاق','الحالة','تاريخ التأكيد','مرجع التحويل','إرسال البريد','معرّف الطلب','المبلغ المستلم'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('الاشتراكات')
    .addItem('تأكيد سداد الدفعة المحددة', 'confirmSelectedPayment').addToUi();
}

function setupBilling_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('افتح المشروع من جدول إدارة الفوترة.');
  PropertiesService.getScriptProperties().setProperty('BILLING_SHEET_ID', book.getId());
  book.setSpreadsheetTimeZone(ZONE);
  const sheet = book.getSheetByName(SHEET) || book.insertSheet(SHEET);
  if (!sheet.getLastRow()) sheet.appendRow(COLUMNS);
  sheet.setRightToLeft(true).setFrozenRows(1);
  sheet.getRange(1,1,1,COLUMNS.length).setBackground('#0A1931').setFontColor('#ffffff').setFontWeight('bold');
  for (const column of ['B','F','K','M','N']) {
    sheet.getRange(column + ':' + column).setNumberFormat('@');
  }
  sheet.hideColumns(16);
}

function doGet(e) {
  const planId = e.parameter.plan;
  if (!Object.hasOwn(PLANS, planId)) return HtmlService.createHtmlOutput('اختر باقة من صفحة الباقات.');
  const form = HtmlService.createTemplateFromFile('Order');
  form.plan = PLANS[planId];
  form.planId = planId;
  return form.evaluate().setTitle('طلب الباقة وإصدار الفاتورة');
}

function invoiceSettings_() {
  const props = PropertiesService.getScriptProperties();
  const values = ['ISSUER', 'PERMIT', 'BANK', 'BENEFICIARY', 'IBAN', 'LOGO_FILE_ID'].map(k => props.getProperty(k));
  if (values.some(v => !v) || !/^SA\d{22}$/.test(values[4])) throw new Error('الفوترة غير متاحة حاليًا.');
  return values;
}

function issueInvoice(input) {
  const { planId, requestId } = input;
  if (!Object.hasOwn(PLANS, planId) || !/^[0-9a-f-]{36}$/i.test(requestId)) throw new Error('بيانات الطلب غير صالحة.');
  const fields = ['association','address','email','phone'].map(k => String(input[k] || '').trim());
  const [association,address,email,phone] = fields;
  if (association.length < 2 || association.length > 160 || address.length < 3 || address.length > 250 ||
      email.length > 150 || !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(email) ||
      !/^\+?[0-9٠-٩][0-9٠-٩ ()-]{6,19}$/.test(phone)) {
    throw new Error('تحقق من بيانات الجمعية والبريد والجوال.');
  }
  const settings = invoiceSettings_();
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const sheet = orders_();
    const saved = sheet.getRange('P:P').createTextFinder(requestId).matchEntireCell(true).findAll()
      .find(cell => sheet.getRange(cell.getRow(),10).getValue() === 1);
    if (saved) {
      const record = sheet.getRange(saved.getRow(),1,1,COLUMNS.length).getDisplayValues()[0];
      if (!fields.every((value,index) => value === record[index+2]) || record[6] !== PLANS[planId][0]) throw new Error('رقم الطلب مستخدم.');
      return deliver_(sheet, saved.getRow(), row_(sheet, saved.getRow()), PLANS[planId], settings);
    }
    const cache = CacheService.getScriptCache();
    const emailKey = 'mail:' + email.toLowerCase();
    const total = Number(cache.get('orders-hour') || 0);
    if (total >= 30 || cache.get(emailKey)) throw new Error('تعذر إصدار طلب جديد الآن. حاول مرة أخرى لاحقًا.');
    cache.put(emailKey,'1',300);
    cache.put('orders-hour',String(total+1),3600);
    const plan = PLANS[planId];
    const date = date_(new Date());
    const invoiceId = 'INV-' + Utilities.getUuid().toUpperCase();
    const data = [invoiceId,date, ...fields,plan[0],plan[1],plan[1]*plan[2],1,date,PENDING,'','','قيد الإرسال',requestId,''];
    const index = sheet.getLastRow()+1;
    sheet.getRange(index,1,1,COLUMNS.length).setValues([data.map(cell_)]);
    SpreadsheetApp.flush();
    return deliver_(sheet, index, data, plan, settings);
  } finally {
    lock.releaseLock();
  }
}

function row_(sheet, index) {
  return sheet.getRange(index,1,1,COLUMNS.length).getValues()[0];
}

function orders_() {
  const id = PropertiesService.getScriptProperties().getProperty('BILLING_SHEET_ID');
  if (!id) throw new Error('الفوترة غير متاحة حاليًا.');
  return SpreadsheetApp.openById(id).getSheetByName(SHEET);
}

function date_(value) {
  return Utilities.formatDate(value, ZONE, 'yyyy-MM-dd');
}

function dueDate_(starting, offset) {
  const [year,month,day] = starting.split('-').map(Number);
  const lastDay = new Date(Date.UTC(year,month+offset,0)).getUTCDate();
  return new Date(Date.UTC(year,month-1+offset,Math.min(day,lastDay))).toISOString().slice(0,10);
}

function cell_(value) {
  return typeof value === 'string' && /^[=+\-@']/ .test(value) ? "'" + value : value;
}

function deliver_(sheet, index, data, plan, settings) {
  const invoice = HtmlService.createTemplateFromFile('Invoice');
  invoice.data = data;
  invoice.plan = plan;
  invoice.settings = settings;
  const logo = DriveApp.getFileById(settings[5]).getBlob();
  invoice.logo = 'data:' + logo.getContentType() + ';base64,' + Utilities.base64Encode(logo.getBytes());
  const pdf = Utilities.newBlob(invoice.evaluate().getContent(), MimeType.HTML, 'invoice.html')
    .getAs(MimeType.PDF).setName(data[0]+'.pdf');
  let emailSent = data[14] === 'تم الإرسال';
  if (!emailSent) {
    try {
      MailApp.sendEmail(data[4], 'فاتورة الخدمة ' + data[0],
        'مرحبًا، تجدون فاتورة خدمتكم وتعليمات التحويل في المرفق. يبدأ العمل بعد تأكيد السداد واستكمال الاتفاق للباقات السنوية.',
        { attachments: [pdf], name: 'المعتزبالله عقيل للاستشارات' });
      emailSent = true;
    } catch (error) {
      sheet.getRange(index,15).setValue('تعذر الإرسال');
    }
    if (emailSent) sheet.getRange(index,15).setValue('تم الإرسال');
    SpreadsheetApp.flush();
  }
  return { id: data[0], emailSent, pdf: Utilities.base64Encode(pdf.getBytes()) };
}

function confirmSelectedPayment() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) throw new Error('تأكيد السداد متاح من لوحة الإدارة فقط.');
  const index = sheet.getActiveCell().getRow();
  if (sheet.getName() !== SHEET || index < 2) { ui.alert('حدد صف الدفعة في لوحة الاشتراكات.'); return; }
  const selected = row_(sheet,index);
  if (selected[11] !== PENDING) { ui.alert('هذه الدفعة ليست بانتظار التحويل.'); return; }
  const result = ui.prompt('تأكيد سداد ' + selected[0], 'بعد مطابقة مبلغ ' + selected[7] + ' ريال مع كشف الحساب، أدخل مرجع التحويل البنكي.', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const reference = result.getResponseText().trim();
  if (!reference || reference.length > 120) { ui.alert('أدخل مرجع تحويل صالح.'); return; }
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const data = row_(sheet,index);
    if (data[0] !== selected[0] || data[11] !== PENDING) return;
    if (sheet.getRange('N:N').createTextFinder(reference).matchEntireCell(true).findNext()) {
      throw new Error('سبق تسجيل هذا المرجع البنكي.');
    }
    const confirmed = date_(new Date());
    const plan = Object.values(PLANS).find(entry => entry[0] === data[6]);
    if (plan && plan[2] > 1 && data[9] === 1) {
      const subscription = data[15];
      const existing = new Set(sheet.getRange(2,10,sheet.getLastRow()-1,7).getValues()
        .filter(row => row[6] === subscription).map(row => row[0]));
      const installments = [];
      for (let number = 2; number <= plan[2]; number++) {
        if (existing.has(number)) continue;
        const next = data.slice();
        next[0] = data[0] + '-' + String(number).padStart(2,'0');
        next[1] = '';
        next[9] = number;
        next[10] = dueDate_(confirmed,number-1);
        next[14] = '';
        installments.push(next.map(cell_));
      }
      if (installments.length) sheet.getRange(sheet.getLastRow()+1,1,installments.length,COLUMNS.length).setValues(installments);
    }
    sheet.getRange(index,12,1,3).setValues([[PAID,confirmed,cell_(reference)]]);
    sheet.getRange(index,17).setValue(data[7]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}
