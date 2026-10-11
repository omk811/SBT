/** نظام رصد السلوكيات — الأستاذ خليل البلوشي */
const SBT_HEADERS = ['المعرف','التاريخ','الطالب','الشعبة','نوع التقرير','السلوكيات','الملاحظات','الإجراء','وقت الحفظ','آخر تحديث'];
function doGet() { return output_({success:true,message:'نظام رصد السلوكيات — الأستاذ خليل البلوشي. العمليات تتطلب رمز الدخول.'}); }
function doPost(e) {
  let lock;
  try {
    const p = JSON.parse(e && e.postData && e.postData.contents || '{}');
    const configured = PropertiesService.getScriptProperties().getProperty('ACCESS_CODE');
    if(!configured || String(p.accessCode || '') !== configured) throw new Error('رمز الدخول غير صحيح أو لم يتم إعداد الخادم');
    if(!['list','create','update','delete'].includes(p.action)) throw new Error('عملية غير مدعومة');
    lock = LockService.getScriptLock(); lock.waitLock(20000);
    const sheet = sheet_();
    const rows = sheet.getDataRange().getValues();
    if(rows[0].join('|') !== SBT_HEADERS.join('|')) throw new Error('شغّل migrateLegacySheet قبل استخدام النظام');
    if(p.action === 'list') return output_({success:true,records:rows.slice(1).filter(r=>r[0]).map(record_)});
    if(p.action === 'create') {
      const record = validate_(p);
      const now = new Date().toISOString(); record.id=Utilities.getUuid(); record.timestamp=now;record.updatedAt=now;
      sheet.appendRow(values_(record)); SpreadsheetApp.flush();
      return output_({success:true,id:record.id});
    }
    const index = rows.findIndex((r,i)=>i>0 && String(r[0])===String(p.id));
    if(index<1) throw new Error('لم يتم العثور على التقرير');
    if(p.action==='delete') sheet.deleteRow(index+1);
    else {
      const old=record_(rows[index]);
      const next=validate_({...old,...p});next.id=old.id;next.timestamp=old.timestamp;next.updatedAt=new Date().toISOString();
      sheet.getRange(index+1,1,1,10).setValues([values_(next)]);
    }
    SpreadsheetApp.flush(); return output_({success:true});
  } catch(error) {return output_({success:false,error:String(error.message || error)});}
  finally {if(lock && lock.hasLock()) lock.releaseLock();}
}
function output_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function book_(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(id)return SpreadsheetApp.openById(id);const b=SpreadsheetApp.getActiveSpreadsheet();if(!b)throw new Error('أضف SPREADSHEET_ID في خصائص النص البرمجي');return b;}
function sheet_(){const b=book_(),name=PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || 'SBT_Reports';const s=b.getSheetByName(name);if(!s)throw new Error('شغّل setupSystem أولًا');return s;}
function setupSystem(){const b=book_(),name=PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || 'SBT_Reports';let s=b.getSheetByName(name);if(!s)s=b.insertSheet(name);if(s.getLastRow()===0){s.appendRow(SBT_HEADERS);s.setFrozenRows(1);}return 'تم إعداد ورقة التقارير';}
function date_(v){return v instanceof Date?Utilities.formatDate(v,'Asia/Muscat','yyyy-MM-dd'):String(v || '').slice(0,10);}
function record_(r){let behaviors;try{behaviors=JSON.parse(r[5]);}catch(_){behaviors=String(r[5]||'').split(/\n|،/).filter(Boolean);}return {id:String(r[0]),date:date_(r[1]),studentName:String(r[2]),className:String(r[3]),reportType:r[4]==='تحصيلي'?'academic':r[4]==='إيجابي'?'positive':'negative',isPositive:r[4]===true || r[4]==='true' || r[4]==='إيجابي',behaviors,notes:String(r[6]||''),actionTaken:String(r[7]||''),timestamp:r[8] instanceof Date?r[8].toISOString():String(r[8]||''),updatedAt:String(r[9]||'')};}
function validate_(p){const text=(v,max)=>{const s=String(v || '').trim();if(s.length>max)throw new Error('تجاوز طول الحقل المسموح');return s;};const r={...p,studentName:text(p.studentName,200),className:text(p.className,60),date:text(p.date,10),notes:text(p.notes,4000),actionTaken:text(p.actionTaken,4000),isPositive:p.isPositive===true};if(!r.studentName || !r.className || !/^\d{4}-\d{2}-\d{2}$/.test(r.date) || !Array.isArray(p.behaviors) || !p.behaviors.length || p.behaviors.length>60)throw new Error('بيانات التقرير غير مكتملة');r.reportType=p.reportType==='academic'?'academic':r.isPositive?'positive':'negative';r.behaviors=p.behaviors.map(b=>text(b,400));if(r.behaviors.some(b=>!b))throw new Error('سلوك فارغ');return r;}
function safe_(v){const s=String(v ?? '');return /^[=+@-]/.test(s)?"'"+s:s;}
function values_(r){return [r.id,r.date,r.studentName,r.className,r.reportType==='academic'?'تحصيلي':r.isPositive?'إيجابي':'سلبي',JSON.stringify(r.behaviors),r.notes,r.actionTaken,r.timestamp,r.updatedAt].map(safe_);}
/** ترحيل ورقة قديمة محددة صراحةً دون تعديلها. لا تُستخدم الورقة النشطة. */
function migrateLegacySheet(){
 const props=PropertiesService.getScriptProperties(),name=props.getProperty('LEGACY_SHEET_NAME');
 if(!name)throw new Error('أضف LEGACY_SHEET_NAME باسم ورقة البيانات القديمة');
 const b=book_(),source=b.getSheetByName(name);if(!source)throw new Error('الورقة القديمة غير موجودة');
 setupSystem();const target=sheet_();if(source.getSheetId()===target.getSheetId())throw new Error('يجب اختلاف ورقة المصدر عن الوجهة');if(target.getLastRow()>1)throw new Error('الوجهة تحتوي سجلات؛ أوقف الترحيل لمنع التكرار');
 const rows=source.getDataRange().getValues(),headers=rows.shift().map(String);
 const aliases={id:['id','المعرف'],date:['date','التاريخ'],studentName:['studentName','اسم الطالب','الطالب'],className:['className','الشعبة','الصف'],isPositive:['isPositive','نوع التقرير','نوع السلوك'],behaviors:['behaviors','السلوكيات'],notes:['notes','الملاحظات'],actionTaken:['actionTaken','action','الإجراء','الإجراء المتخذ'],timestamp:['timestamp','وقت الحفظ','الطابع الزمني']};
 const positions={};Object.keys(aliases).forEach(k=>positions[k]=headers.findIndex(h=>aliases[k].includes(h.trim())));
 if(['date','studentName','className','behaviors'].some(k=>positions[k]<0))throw new Error('عناوين الورقة القديمة غير معروفة؛ لا توجد تغييرات على البيانات');
 const records=rows.filter(r=>r.some(v=>v!=='')).map(row=>{const p={};Object.keys(positions).forEach(k=>p[k]=positions[k]<0?'':row[positions[k]]);p.date=date_(p.date);p.isPositive=p.isPositive===true || p.isPositive==='true' || p.isPositive==='إيجابي';try{p.behaviors=JSON.parse(p.behaviors);}catch(_){p.behaviors=String(p.behaviors).split(/\n|،/).filter(Boolean);}const r=validate_(p);r.id=Utilities.getUuid();r.timestamp=p.timestamp instanceof Date?p.timestamp.toISOString():String(p.timestamp||new Date().toISOString());r.updatedAt=new Date().toISOString();return values_(r);});
 if(records.length)target.getRange(2,1,records.length,10).setValues(records);return 'تم نسخ '+records.length+' تقريرًا مع الحفاظ على الورقة القديمة';
}
