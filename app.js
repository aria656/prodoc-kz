/* ProDoc KZ. Форма, проверка, предпросмотр и создание файла. Всё работает в браузере, данные никуда не отправляются. */
(function(){
'use strict';

/* ---------- DOCX без внешних библиотек ---------- */
function crc32(b){var crc=-1;for(var i=0;i<b.length;i++){crc^=b[i];for(var k=0;k<8;k++)crc=(crc>>>1)^(0xEDB88320&-(crc&1));}return(crc^-1)>>>0;}
function zipStore(files){
  var enc=new TextEncoder(),parts=[],central=[],off=0;
  files.forEach(function(f){
    var name=enc.encode(f.name),data=enc.encode(f.data),crc=crc32(data);
    var lh=new DataView(new ArrayBuffer(30));
    lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0x0800,true);
    lh.setUint16(12,0x21,true);lh.setUint32(14,crc,true);lh.setUint32(18,data.length,true);
    lh.setUint32(22,data.length,true);lh.setUint16(26,name.length,true);
    parts.push(new Uint8Array(lh.buffer),name,data);
    var ch=new DataView(new ArrayBuffer(46));
    ch.setUint32(0,0x02014b50,true);ch.setUint16(4,20,true);ch.setUint16(6,20,true);ch.setUint16(8,0x0800,true);
    ch.setUint16(14,0x21,true);ch.setUint32(16,crc,true);ch.setUint32(20,data.length,true);
    ch.setUint32(24,data.length,true);ch.setUint16(28,name.length,true);ch.setUint32(42,off,true);
    central.push(new Uint8Array(ch.buffer),name);
    off+=30+name.length+data.length;
  });
  var csize=0;central.forEach(function(x){csize+=x.length;});
  var end=new DataView(new ArrayBuffer(22));
  end.setUint32(0,0x06054b50,true);end.setUint16(8,files.length,true);end.setUint16(10,files.length,true);
  end.setUint32(12,csize,true);end.setUint32(16,off,true);
  return new Blob(parts.concat(central,[new Uint8Array(end.buffer)]),{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
function xmlEsc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function buildDocx(blocks){
  var W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  var body=blocks.map(function(b){
    var jc={r:'left',c:'center',l:'left',j:'both'}[b.a]||'left';
    var ind=b.a==='r'?'<w:ind w:left="4536"/>':'';
    return '<w:p><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/>'+ind+'<w:jc w:val="'+jc+'"/></w:pPr>'+
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'+(b.b?'<w:b/>':'')+
      '<w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">'+xmlEsc(b.t)+'</w:t></w:r></w:p>';
  }).join('');
  var doc='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document '+W+'><w:body>'+body+
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';
  var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  var rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  return zipStore([{name:'[Content_Types].xml',data:ct},{name:'_rels/.rels',data:rels},{name:'word/document.xml',data:doc}]);
}

/* ---------- общие помощники ---------- */
var $=function(s,e){return (e||document).querySelector(s);};
var $$=function(s,e){return Array.prototype.slice.call((e||document).querySelectorAll(s));};
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function lines(s){return String(s||'').split(/\n+/).map(function(x){return x.trim();}).filter(Boolean);}
function fmtDate(iso){var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(iso||'');return m?m[3]+'.'+m[2]+'.'+m[1]:'[дата]';}
function today(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);}
function toast(t){var e=$('#toast');if(!e)return;e.textContent=t;e.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(function(){e.hidden=true;},3200);}
function trimDot(s){return String(s).replace(/[.\s]+$/,'');}

/* ---------- главная: поиск и категории ---------- */
function initHome(){
  var q=$('#q'),cat='all';
  function apply(){
    var t=(q.value||'').trim().toLowerCase(),shown=0;
    $$('#cards .card').forEach(function(c){
      var ok=(cat==='all'||c.dataset.cat===cat)&&(!t||(c.dataset.words||'').toLowerCase().indexOf(t)>-1);
      c.hidden=!ok;if(ok)shown++;
    });
    $('#none').hidden=shown>0;
  }
  q.addEventListener('input',apply);
  $$('.chip[data-cat]').forEach(function(b){
    b.addEventListener('click',function(){
      cat=b.dataset.cat;$$('.chip[data-cat]').forEach(function(x){x.setAttribute('aria-pressed',String(x===b));});apply();
    });
  });
}

/* ---------- шаблоны документов ---------- */
var STATUSES=['Подозреваемый','Обвиняемый','Подсудимый','Потерпевший','Гражданский истец','Гражданский ответчик','Защитник','Представитель','Иной участник процесса'];
function common(addrLabel,addrHint,addrPh){
  return [
    {k:'addressee',label:addrLabel,req:1,ph:addrPh,hint:addrHint},
    {k:'fio',label:'Ваши Ф. И. О.',req:1,ph:'Введите Ф. И. О.'},
    {k:'status',label:'Процессуальный статус',type:'select',req:1,options:STATUSES},
    {k:'caseNo',label:'Номер уголовного дела',req:1,ph:'Например: 123456'},
    {k:'address',label:'Адрес',ph:'Адрес проживания или места нахождения'},
    {k:'phone',label:'Телефон',ph:'+7 ...'}
  ];
}
function artField(def){return {k:'art',label:'Ссылка на закон (по желанию)',ph:def,hint:'Если оставить пустым, в текст попадёт: '+def+'. Перед подачей сверьте с действующей редакцией УПК РК.'};}
var DATEF={k:'date',label:'Дата документа',type:'date',req:1};

function head(h,t1,t2){
  var b=[{a:'r',t:h.v('addressee')},{a:'r',t:'Заявитель: '+h.v('fio')+', '+h.v('status').toLowerCase()}];
  if(h.o('address'))b.push({a:'r',t:'адрес: '+h.o('address')});
  if(h.o('phone'))b.push({a:'r',t:'тел.: '+h.o('phone')});
  b.push({a:'r',t:'по уголовному делу № '+h.v('caseNo')});
  b.push({a:'c',b:1,t:t1});
  b.push({a:'c',b:1,t:t2});
  return b;
}
function sign(h){return [{a:'l',t:'Дата: '+fmtDate(h.raw('date'))+'          Подпись: ____________ /'+h.v('fio')+'/'}];}
function intro(h){return {a:'j',t:'Я, '+h.v('fio')+', являюсь участником уголовного процесса по делу № '+h.v('caseNo')+' (процессуальный статус: '+h.v('status').toLowerCase()+').'};}

var DOCS={
evidence:{
  title:'Ходатайство о приобщении доказательств',file:'Ходатайство_о_приобщении_доказательств',
  step1:common('Кому адресовано','Лицо, ведущее досудебное расследование, прокурор или суд','Например: Следователю [орган, Ф. И. О.]'),
  step2:[
    {k:'evidence',label:'Что нужно приобщить (каждое с новой строки)',type:'textarea',req:1,ph:'Например:\nДоговор от 01.01.2025\nФотографии с места события (3 шт.)'},
    {k:'relevance',label:'Для установления каких обстоятельств это нужно',type:'textarea',req:1,ph:'Например: времени и места события, размера ущерба',hint:'По ст. 99 УПК РК в ходатайстве нужно указать, для установления каких обстоятельств оно заявлено.'},
    artField('ст. 99 УПК РК'),DATEF],
  demo:{addressee:'Следователю [орган, Ф. И. О.]',fio:'Иванов Иван Иванович',status:'Потерпевший',caseNo:'123456',address:'г. Астана, ул. Примерная, д. 1, кв. 1',phone:'+7 700 000 00 00',
    evidence:'Договор от 01.01.2025\nФотографии с места события (3 шт.)',relevance:'времени и обстоятельств события, о котором даны показания'},
  build:function(h){
    var items=lines(h.v('evidence')),b=head(h,'ХОДАТАЙСТВО','о приобщении доказательств');
    b.push(intro(h));
    b.push({a:'j',t:'В соответствии со '+h.art('ст. 99 УПК РК')+' прошу приобщить к материалам уголовного дела следующие доказательства:'});
    items.forEach(function(x,i){b.push({a:'l',t:(i+1)+'. '+x});});
    b.push({a:'j',t:'Приобщение необходимо для установления следующих обстоятельств: '+trimDot(h.v('relevance'))+'.'});
    b.push({a:'j',t:'На основании изложенного прошу:'});
    b.push({a:'l',t:'1. Приобщить указанные доказательства к материалам уголовного дела.'});
    b.push({a:'l',t:'2. О результатах рассмотрения ходатайства уведомить меня в установленном порядке.'});
    b.push({a:'l',t:'Приложение:'});
    items.forEach(function(x,i){b.push({a:'l',t:(i+1)+'. '+x});});
    return b.concat(sign(h));
  }},
witness:{
  title:'Ходатайство о допросе свидетеля',file:'Ходатайство_о_допросе_свидетеля',
  step1:common('Кому адресовано','Лицо, ведущее досудебное расследование, прокурор или суд','Например: Следователю [орган, Ф. И. О.]'),
  step2:[
    {k:'wName',label:'Ф. И. О. свидетеля',req:1,ph:'Введите Ф. И. О. свидетеля'},
    {k:'wAddr',label:'Адрес или телефон свидетеля',ph:'Если известны'},
    {k:'circ',label:'Какие обстоятельства известны свидетелю',type:'textarea',req:1,ph:'Что именно свидетель видел или знает. Например: свидетелю известно, что ...'},
    {k:'reason',label:'Для установления каких обстоятельств нужен допрос',type:'textarea',req:1,ph:'Например: порядка событий, участия конкретных лиц',hint:'По ст. 99 УПК РК в ходатайстве нужно указать, для установления каких обстоятельств оно заявлено.'},
    {k:'appear',label:'Явка свидетеля',type:'select',options:['Прошу вызвать свидетеля повесткой','Явку свидетеля обеспечу самостоятельно'],hint:'Если явку обеспечивает сторона, орган, ведущий процесс, не вправе отказать в допросе (ст. 99 УПК РК). Если не выбрать, в тексте будет просьба о вызове повесткой.'},
    artField('ст. 99 УПК РК'),DATEF],
  demo:{addressee:'Следователю [орган, Ф. И. О.]',fio:'Петров Пётр Петрович',status:'Подозреваемый',caseNo:'123456',address:'г. Алматы, ул. Примерная, д. 2',phone:'+7 700 000 00 01',
    wName:'Сидоров Сергей Сергеевич',wAddr:'г. Алматы, ул. Образцовая, д. 5',circ:'свидетель находился на месте события и видел, что произошло',reason:'порядка событий и участия в них конкретных лиц',appear:'Прошу вызвать свидетеля повесткой'},
  build:function(h){
    var b=head(h,'ХОДАТАЙСТВО','о допросе свидетеля');
    b.push(intro(h));
    var self=h.o('appear').indexOf('самостоятельно')>-1;
    b.push({a:'j',t:'В соответствии со '+h.art('ст. 99 УПК РК')+' прошу '+(self?'допросить':'вызвать и допросить')+' в качестве свидетеля следующее лицо: '+h.v('wName')+(h.o('wAddr')?' ('+h.o('wAddr')+')':'')+'.'});
    if(self)b.push({a:'j',t:'Явку свидетеля в назначенное время обязуюсь обеспечить самостоятельно.'});
    b.push({a:'j',t:'Свидетелю известны следующие обстоятельства: '+trimDot(h.v('circ'))+'.'});
    b.push({a:'j',t:'Допрос необходим для установления следующих обстоятельств: '+trimDot(h.v('reason'))+'.'});
    b.push({a:'j',t:'На основании изложенного прошу:'});
    b.push({a:'l',t:'1. '+(self?'Допросить':'Вызвать и допросить')+' указанного свидетеля.'});
    b.push({a:'l',t:'2. О результатах рассмотрения ходатайства уведомить меня в установленном порядке.'});
    return b.concat(sign(h));
  }},
complaint:{
  title:'Жалоба на процессуальное решение или действие',file:'Жалоба_на_решение_или_действие',
  step1:common('Кому адресована жалоба','Должность и Ф. И. О. того, кто будет рассматривать жалобу. Выберите порядок подачи на следующем шаге.','Должность, орган, Ф. И. О.'),
  step2:[
    {k:'channel',label:'Куда подаёте жалобу',type:'select',req:1,options:['Начальнику органа или прокурору (досудебный порядок)','В суд (судебный порядок)'],hint:'В суд можно обратиться по вопросам, перечисленным в ст. 106 УПК РК, в течение 15 суток. От выбора зависит ссылка на статью в тексте.'},
    {k:'kind',label:'Что вы обжалуете',type:'select',req:1,options:['решение','действие','бездействие']},
    {k:'official',label:'Чьё это решение, действие или бездействие',req:1,ph:'Должность, орган, Ф. И. О.'},
    {k:'actDate',label:'Дата решения или действия',type:'date',req:1},
    {k:'knowDate',label:'Когда вы узнали о нём (по желанию)',type:'date',hint:'Для жалобы в суд 15 суток по ст. 106 УПК РК считаются, в частности, с момента, когда вы ознакомились с решением.'},
    {k:'what',label:'Что произошло',type:'textarea',req:1,ph:'Опишите, какое решение принято или что сделано (не сделано)'},
    {k:'viol',label:'Какие ваши права нарушены',type:'textarea',req:1,ph:'Почему вы считаете это незаконным или необоснованным'},
    {k:'demand',label:'Что вы просите (каждое требование с новой строки)',type:'textarea',req:1,ph:'Например: признать бездействие незаконным\nОбязать устранить нарушение'},
    artField('ст. 100, 105 УПК РК (досудебный порядок) или ст. 100, 106 УПК РК (судебный порядок)'),DATEF],
  demo:{addressee:'Прокурору [наименование, Ф. И. О.]',fio:'Смирнов Алексей Алексеевич',status:'Обвиняемый',caseNo:'123456',address:'г. Астана, ул. Примерная, д. 3',phone:'+7 700 000 00 02',
    channel:'Начальнику органа или прокурору (досудебный порядок)',kind:'бездействие',official:'Следователь [орган, Ф. И. О.]',actDate:'2026-10-01',knowDate:'2026-10-05',
    what:'ходатайство, поданное мной, не было рассмотрено, ответа я не получил',viol:'я лишён возможности защищать свои интересы и представить доказательства',demand:'Признать бездействие следователя незаконным\nОбязать рассмотреть ходатайство'},
  build:function(h){
    var kind=h.v('kind'),court=h.o('channel').indexOf('суд')===0||h.o('channel').indexOf('В суд')===0;
    var def=court?'ст. 100, 106 УПК РК':'ст. 100, 105 УПК РК';
    var b=head(h,'ЖАЛОБА','на '+kind+' (орган или должностное лицо: '+h.v('official')+')');
    b.push(h.o('status').indexOf('Иное лицо')===0?{a:'j',t:'Я, '+h.v('fio')+', обращаюсь по уголовному делу № '+h.v('caseNo')+' как лицо, чьи права и законные интересы затронуты (ст. 100 УПК РК).'}:intro(h));
    b.push({a:'j',t:fmtDate(h.raw('actDate'))+' '+h.v('official')+': '+trimDot(h.v('what'))+'.'});
    if(h.raw('knowDate'))b.push({a:'j',t:'Об обжалуемом решении, действии или бездействии мне стало известно '+fmtDate(h.raw('knowDate'))+'.'});
    b.push({a:'j',t:'Считаю, что данное '+kind+' нарушает мои права и законные интересы, поскольку '+trimDot(h.v('viol'))+'.'});
    b.push({a:'j',t:'На основании изложенного и руководствуясь '+h.art(def)+', прошу:'});
    var d=lines(h.v('demand'));
    d.forEach(function(x,i){b.push({a:'l',t:(d.length>1?(i+1)+'. ':'')+x});});
    return b.concat(sign(h));
  }}
};

DOCS.complaint.step1[2].options=STATUSES.concat(['Иное лицо, чьи права затронуты']);

function helpers(V){
  return {
    raw:function(k){return (V[k]||'').trim();},
    v:function(k){var x=(V[k]||'').trim();return x||'[__]';},
    o:function(k){return (V[k]||'').trim();},
    art:function(def){var x=(V.art||'').trim();return x||def;}
  };
}

/* ---------- форма ---------- */
var D,S={step:1,V:{},dl:true};
function fieldHtml(f){
  var val=S.V[f.k]||'',id='f-'+f.k,inp;
  if(f.type==='textarea')inp='<textarea id="'+id+'" data-k="'+f.k+'" placeholder="'+esc(f.ph||'')+'">'+esc(val)+'</textarea>';
  else if(f.type==='select')inp='<select id="'+id+'" data-k="'+f.k+'"><option value="">Выберите</option>'+f.options.map(function(o){return '<option'+(o===val?' selected':'')+'>'+esc(o)+'</option>';}).join('')+'</select>';
  else inp='<input id="'+id+'" data-k="'+f.k+'" type="'+(f.type||'text')+'" value="'+esc(val)+'" placeholder="'+esc(f.ph||'')+'" autocomplete="off">';
  return '<div class="field" data-f="'+f.k+'"><label for="'+id+'">'+esc(f.label)+(f.req?' <em aria-hidden="true">*</em>':'')+'</label>'+inp+(f.hint?'<span class="hint">'+esc(f.hint)+'</span>':'')+'<span class="err" role="alert"></span></div>';
}
function stepper(){
  var n=['Общие сведения','Содержание','Просмотр и файл'];
  return '<ol class="steps">'+n.map(function(t,i){var s=i+1;return '<li'+(s===S.step?' aria-current="step"':'')+(s<S.step?' class="done"':'')+'><i>'+s+'</i><span>'+t+'</span></li>';}).join('')+'</ol>';
}
function draw(){
  var root=$('#tool'),html;
  if(S.step<3){
    var fs=S.step===1?D.step1:D.step2;
    html=stepper()+'<form class="panel" id="form" novalidate><div class="row" style="justify-content:space-between"><h2>'+(S.step===1?'Общие сведения':'Содержание')+'</h2><button type="button" class="demo" id="demo">Подставить пример</button></div>'+
      fs.map(fieldHtml).join('')+
      '<div class="actions">'+(S.step>1?'<button type="button" class="btn ghost" id="back">Назад</button>':'<span></span>')+'<button type="submit" class="btn grow">'+(S.step===1?'Далее':'Проверить и показать документ')+'</button></div></form>';
  }else{
    var h=helpers(S.V),blocks=D.build(h),req=D.step1.concat(D.step2).filter(function(f){return f.req;});
    var hasPh=blocks.some(function(b){return /\[[^\]]*\]/.test(b.t);});
    html=stepper()+'<div class="panel"><h2>Проверка и просмотр</h2><ul class="checks">'+
      '<li><span class="ok">✓</span><span>Обязательные поля заполнены: '+req.length+' из '+req.length+'.</span></li>'+
      '<li><span class="ok">✓</span><span>Вид документа: '+esc(D.title)+'.</span></li>'+
      '<li><span class="'+(hasPh?'wn':'ok')+'">'+(hasPh?'!':'✓')+'</span><span>'+(hasPh?'В тексте остались места в квадратных скобках, они выделены жёлтым. Замените их своими данными в Word.':'Незаполненных мест в тексте нет.')+'</span></li></ul>'+
      '<div class="paperwrap"><div class="paper" id="paper">'+blocks.map(function(b){
        var cls=(b.a==='r'?'r':b.a==='c'?'c':b.a==='j'?'j':'')+(b.b?' b':'');
        return '<p class="'+cls+'">'+esc(b.t).replace(/(\[[^\]]*\])/g,'<span class="ph">$1</span>')+'</p>';
      }).join('')+'</div></div>'+
      '<div class="actions"><button class="btn ghost" id="back">Изменить данные</button><div class="row grow"><button class="btn" id="dl-docx">Скачать DOCX</button><button class="btn ghost" id="dl-txt">Скачать TXT</button><button class="btn ghost" id="copy">Копировать текст</button></div></div></div>'+
      '<p class="notice" style="margin-top:16px"><strong>Проверьте документ перед подачей.</strong> Это учебный проект студентов, авторы не являются юристами. Покажите документ юристу или в юридической консультации.</p>';
  }
  root.innerHTML=html;bind();
}
function validate(fields){
  var first=null;
  fields.forEach(function(f){
    var box=$('[data-f="'+f.k+'"]');if(!box)return;
    var bad=f.req&&!(S.V[f.k]||'').trim();
    box.classList.toggle('bad',bad);
    $('.err',box).textContent=bad?'Заполните это поле.':'';
    var inp=$('[data-k]',box);if(inp)inp.setAttribute('aria-invalid',bad?'true':'false');
    if(bad&&!first)first=inp;
  });
  if(first){first.focus();toast('Заполните обязательные поля, они отмечены звёздочкой.');}
  return !first;
}
function go(step){S.step=step;draw();var t=$('#tool');if(t)t.scrollIntoView();}
function bind(){
  var form=$('#form');
  if(form){
    if(!S.V.date)S.V.date=today();
    var de=$('#f-date');if(de&&!de.value)de.value=S.V.date;
    form.addEventListener('input',function(e){var k=e.target.dataset.k;if(k)S.V[k]=e.target.value;});
    form.addEventListener('change',function(e){var k=e.target.dataset.k;if(k)S.V[k]=e.target.value;});
    form.addEventListener('submit',function(e){e.preventDefault();if(validate(S.step===1?D.step1:D.step2))go(S.step+1);});
    $('#demo').addEventListener('click',function(){Object.keys(D.demo).forEach(function(k){S.V[k]=D.demo[k];});draw();toast('Подставлены вымышленные данные.');});
  }
  var b=$('#back');if(b)b.addEventListener('click',function(){go(S.step-1);});
  var c=$('#copy');if(c)c.addEventListener('click',copyText);
  var x=$('#dl-docx');if(x)x.addEventListener('click',function(){save('docx');});
  var t=$('#dl-txt');if(t)t.addEventListener('click',function(){save('txt');});
}
function docBlocks(){return D.build(helpers(S.V));}
function docText(){return docBlocks().map(function(b){return b.t;}).join('\n\n');}
function copyText(){
  function fallback(){var r=document.createRange();r.selectNodeContents($('#paper'));var s=getSelection();s.removeAllRanges();s.addRange(r);toast('Текст выделен. Скопируйте его через меню телефона.');}
  try{navigator.clipboard.writeText(docText()).then(function(){toast('Текст скопирован.');},fallback);}catch(e){fallback();}
}
function save(kind){
  try{
    var blob=kind==='docx'?buildDocx(docBlocks()):new Blob([docText()],{type:'text/plain;charset=utf-8'});
    var url=URL.createObjectURL(blob),a=document.createElement('a');
    a.href=url;a.download=D.file+'.'+kind;document.body.appendChild(a);a.click();
    setTimeout(function(){URL.revokeObjectURL(url);a.remove();},4000);
    toast('Файл создан. Проверьте папку загрузок.');
  }catch(e){toast('Не удалось создать файл. Нажмите «Копировать текст».');}
}

/* ---------- запуск ---------- */
document.addEventListener('DOMContentLoaded',function(){
  if($('#q'))initHome();
  var id=document.body.getAttribute('data-doc');
  if(id&&DOCS[id]&&$('#tool')){D=DOCS[id];draw();}
});
window.__ProDoc={buildDocx:buildDocx,DOCS:DOCS,helpers:helpers};
})();
