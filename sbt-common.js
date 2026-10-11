/* الأستاذ خليل البلوشي */
window.SBT = {
  level(value) {
    const text = String(value);
    if (['الشجار','النميمة','المشاكل','التنمر','الاستيلاء','ألفاظ نابية'].some(k=>text.includes(k))) return 3;
    if (['الحديث','الضحك','التحرك','اللعب','التشويش','عدم الإنصات','المقاطعة','تناول المأكولات','شرب الماء','الخروج للحمام','الإهمال'].some(k=>text.includes(k))) return 1;
    return 2;
  },
  async request(url, action, payload = {}, code = sessionStorage.getItem('accessCode') || '') {
    const controller = new AbortController();
    const timer = setTimeout(()=>controller.abort(), 25000);
    try {
      const response = await fetch(url, {method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},
        body:JSON.stringify({...payload,action,accessCode:code}),signal:controller.signal});
      if(!response.ok) throw new Error('تعذر الاتصال بالخادم');
      const result = await response.json();
      if(!result.success) throw new Error(result.error || 'تعذر تنفيذ العملية');
      if(action === 'list') {
        if(!Array.isArray(result.records)) throw new Error('استجابة غير متوافقة؛ حدّث ملف Code.gs');
        result.records=result.records.map(r=>({...r,id:String(r.id || r.timestamp || ''),isPositive:r.isPositive===true || r.isPositive==='true',behaviors:Array.isArray(r.behaviors)?r.behaviors:[],actionTaken:r.actionTaken ?? r.action ?? ''}));
      }
      return result;
    } catch(e) {if(e.name==='AbortError') throw new Error('انتهت مهلة الاتصال؛ حدّث التقارير للتحقق قبل إعادة الحفظ'); throw e;}
    finally {clearTimeout(timer);}
  },
  logout(){sessionStorage.removeItem('accessCode');sessionStorage.removeItem('isLoggedIn');location.replace('index.html');}
};
