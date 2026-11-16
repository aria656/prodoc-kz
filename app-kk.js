/* ProDoc KZ. Қазақша нұсқа. Бәрі браузерде жұмыс істейді, деректер ешқайда жіберілмейді. */
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
function fmtDate(iso){var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(iso||'');return m?m[3]+'.'+m[2]+'.'+m[1]:'[күні]';}
function today(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);}
function toast(t){var e=$('#toast');if(!e)return;e.textContent=t;e.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(function(){e.hidden=true;},3200);}
function trimDot(s){return String(s).replace(/[.\s]+$/,'');}

/* ---------- басты бет: іздеу және санаттар ---------- */
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

/* ---------- құжат үлгілері ---------- */
var STATUSES=['Күдікті','Айыпталушы','Сотталушы','Жәбірленуші','Азаматтық талапкер','Азаматтық жауапкер','Қорғаушы','Өкіл','Іс жүргізудің өзге қатысушысы'];
var OTHER='Құқықтары қозғалған өзге адам';
function common(addrLabel,addrHint,addrPh){
  return [
    {k:'addressee',label:addrLabel,req:1,ph:addrPh,hint:addrHint},
    {k:'fio',label:'Сіздің Т. А. Ә.',req:1,ph:'Т. А. Ә. енгізіңіз'},
    {k:'status',label:'Іс жүргізу мәртебесі',type:'select',req:1,options:STATUSES},
    {k:'caseNo',label:'Қылмыстық істің нөмірі',req:1,ph:'Мысалы: 123456'},
    {k:'address',label:'Мекенжай',ph:'Тұрғылықты жері немесе орналасқан жері'},
    {k:'phone',label:'Телефон',ph:'+7 ...'}
  ];
}
function artField(def){return {k:'art',label:'Заңға сілтеме (қалауыңыз бойынша)',ph:def,hint:'Бос қалдырсаңыз, мәтінге мынау қойылады: '+def+'. Тапсырар алдында ҚР ҚПК-нің қолданыстағы редакциясымен салыстырыңыз. Өз сілтемеңізді жазсаңыз, соңына «сәйкес» сөзі қосылады, сондықтан септік жалғауын өзіңіз жазыңыз.'};}
var DATEF={k:'date',label:'Құжат күні',type:'date',req:1};
var ADDR1='Мысалы: Тергеушіге [орган, Т. А. Ә.]';
var ADDR_HINT='Сотқа дейінгі тергеп-тексеруді жүргізетін адам, прокурор немесе сот';

function head(h,t1,t2){
  var b=[{a:'r',t:h.v('addressee')},{a:'r',t:'Өтініш беруші: '+h.v('fio')+', '+h.v('status').toLowerCase()}];
  if(h.o('address'))b.push({a:'r',t:'мекенжайы: '+h.o('address')});
  if(h.o('phone'))b.push({a:'r',t:'тел.: '+h.o('phone')});
  b.push({a:'r',t:'№ '+h.v('caseNo')+' қылмыстық іс бойынша'});
  b.push({a:'c',b:1,t:t1});
  b.push({a:'c',b:1,t:t2});
  return b;
}
function sign(h){return [{a:'l',t:'Күні: '+fmtDate(h.raw('date'))+'          Қолы: ____________ /'+h.v('fio')+'/'}];}
function intro(h){return {a:'j',t:'Мен, '+h.v('fio')+', № '+h.v('caseNo')+' іс бойынша қылмыстық процеске қатысушымын (іс жүргізу мәртебесі: '+h.v('status').toLowerCase()+').'};}
var NOTIFY='Өтінішхатты қарау нәтижелері туралы белгіленген тәртіппен маған хабарлауды.';
var BASIS='Жоғарыда айтылғанның негізінде сұраймын:';
var A99='ҚР ҚПК-нің 99-бабына';
var DEMO_ADDR='Тергеушіге [орган, Т. А. Ә.]';
var DEMO_PROS='Прокурорға [атауы, Т. А. Ә.]';
var CH1='Орган басшысына немесе прокурорға (сотқа дейінгі тәртіп)';
var CH2='Сотқа (сот тәртібі)';

var DOCS={
evidence:{
  title:'Дәлелдемелерді қосып тігу туралы өтінішхат',file:'Dalelderdi_kosyp_tigu_turaly_otinishhat',
  step1:common('Кімге жолданады',ADDR_HINT,ADDR1),
  step2:[
    {k:'evidence',label:'Нені қосып тігу керек (әрқайсысы жаңа жолдан)',type:'textarea',req:1,ph:'Мысалы:\n01.01.2025 жылғы шарт\nОқиға орнынан түсірілген фотосуреттер (3 дана)'},
    {k:'relevance',label:'Қандай мән-жайларды анықтау үшін қажет',type:'textarea',req:1,ph:'Мысалы: оқиғаның уақыты мен орны, залалдың мөлшері',hint:'ҚР ҚПК-нің 99-бабы бойынша өтінішхатта қандай мән-жайларды анықтау үшін берілгені көрсетілуі керек.'},
    artField(A99),DATEF],
  demo:{addressee:DEMO_ADDR,fio:'Иванов Иван Иванович',status:'Жәбірленуші',caseNo:'123456',address:'Астана қ., Үлгі к-сі, 1-үй, 1-пәтер',phone:'+7 700 000 00 00',
    evidence:'01.01.2025 жылғы шарт\nОқиға орнынан түсірілген фотосуреттер (3 дана)',relevance:'жауап берілген оқиғаның уақыты мен мән-жайлары'},
  build:function(h){
    var items=lines(h.v('evidence')),b=head(h,'ӨТІНІШХАТ','дәлелдемелерді қылмыстық іс материалдарына қосып тігу туралы');
    b.push(intro(h));
    b.push({a:'j',t:h.art(A99)+' сәйкес қылмыстық іс материалдарына мына дәлелдемелерді қосып тігуді сұраймын:'});
    items.forEach(function(x,i){b.push({a:'l',t:(i+1)+'. '+x});});
    b.push({a:'j',t:'Қосып тігу мына мән-жайларды анықтау үшін қажет: '+trimDot(h.v('relevance'))+'.'});
    b.push({a:'j',t:BASIS});
    b.push({a:'l',t:'1. Көрсетілген дәлелдемелерді қылмыстық іс материалдарына қосып тігуді.'});
    b.push({a:'l',t:'2. '+NOTIFY});
    b.push({a:'l',t:'Қосымша:'});
    items.forEach(function(x,i){b.push({a:'l',t:(i+1)+'. '+x});});
    return b.concat(sign(h));
  }},
witness:{
  title:'Куәден жауап алу туралы өтінішхат',file:'Kuaden_zhauap_alu_turaly_otinishhat',
  step1:common('Кімге жолданады',ADDR_HINT,ADDR1),
  step2:[
    {k:'wName',label:'Куәнің Т. А. Ә.',req:1,ph:'Куәнің Т. А. Ә. енгізіңіз'},
    {k:'wAddr',label:'Куәнің мекенжайы немесе телефоны',ph:'Белгілі болса'},
    {k:'circ',label:'Куәге қандай мән-жайлар белгілі',type:'textarea',req:1,ph:'Куә нені көрді немесе нені біледі. Мысалы: куәге ... белгілі'},
    {k:'reason',label:'Қандай мән-жайларды анықтау үшін жауап алу қажет',type:'textarea',req:1,ph:'Мысалы: оқиғалардың тәртібі, нақты адамдардың қатысуы',hint:'ҚР ҚПК-нің 99-бабы бойынша өтінішхатта қандай мән-жайларды анықтау үшін берілгені көрсетілуі керек.'},
    {k:'appear',label:'Куәнің келуі',type:'select',options:['Куәні хабарламамен шақыруды сұраймын','Куәнің келуін өзім қамтамасыз етемін'],hint:'Егер келуін тарап қамтамасыз етсе, іс жүргізуші орган жауап алудан бас тартуға құқылы емес (ҚР ҚПК-нің 99-бабы). Таңдамасаңыз, мәтінде шақыру туралы өтініш болады.'},
    artField(A99),DATEF],
  demo:{addressee:DEMO_ADDR,fio:'Петров Петр Петрович',status:'Күдікті',caseNo:'123456',address:'Алматы қ., Үлгі к-сі, 2-үй',phone:'+7 700 000 00 01',
    wName:'Сидоров Сергей Сергеевич',wAddr:'Алматы қ., Үлгілі к-сі, 5-үй',circ:'куә оқиға орнында болған және не болғанын көрген',reason:'оқиғалардың тәртібі мен оларға нақты адамдардың қатысуы',appear:'Куәні хабарламамен шақыруды сұраймын'},
  build:function(h){
    var b=head(h,'ӨТІНІШХАТ','куәден жауап алу туралы');
    b.push(intro(h));
    var self=h.o('appear').indexOf('өзім')>-1;
    b.push({a:'j',t:h.art(A99)+' сәйкес куә ретінде мына адамнан '+(self?'':'шақырып, ')+'жауап алуды сұраймын: '+h.v('wName')+(h.o('wAddr')?' ('+h.o('wAddr')+')':'')+'.'});
    if(self)b.push({a:'j',t:'Куәнің тағайындалған уақытта келуін өзім қамтамасыз етуге міндеттенемін.'});
    b.push({a:'j',t:'Куәге мына мән-жайлар белгілі: '+trimDot(h.v('circ'))+'.'});
    b.push({a:'j',t:'Жауап алу мына мән-жайларды анықтау үшін қажет: '+trimDot(h.v('reason'))+'.'});
    b.push({a:'j',t:BASIS});
    b.push({a:'l',t:'1. Көрсетілген куәден '+(self?'':'шақырып, ')+'жауап алуды.'});
    b.push({a:'l',t:'2. '+NOTIFY});
    return b.concat(sign(h));
  }},
complaint:{
  title:'Іс жүргізу шешіміне немесе әрекетіне шағым',file:'Shagym_sheshimge_nemese_areketke',
  step1:common('Шағым кімге жолданады','Шағымды қарайтын адамның лауазымы және Т. А. Ә. Берілу тәртібін келесі қадамда таңдайсыз.','Лауазымы, органы, Т. А. Ә.'),
  step2:[
    {k:'channel',label:'Шағымды қайда бересіз',type:'select',req:1,options:[CH1,CH2],hint:'Сотқа ҚР ҚПК-нің 106-бабында көрсетілген мәселелер бойынша 15 тәулік ішінде жүгінуге болады. Таңдауыңызға мәтіндегі баптың сілтемесі байланысты.'},
    {k:'kind',label:'Нені дауласасыз',type:'select',req:1,options:['шешім','әрекет','әрекетсіздік']},
    {k:'official',label:'Кімнің шешімі, әрекеті немесе әрекетсіздігі',req:1,ph:'Лауазымы, органы, Т. А. Ә.'},
    {k:'actDate',label:'Шешімнің немесе әрекеттің күні',type:'date',req:1},
    {k:'knowDate',label:'Ол туралы қашан білдіңіз (қалауыңыз бойынша)',type:'date',hint:'Сотқа шағымдану үшін ҚР ҚПК-нің 106-бабы бойынша 15 тәулік, атап айтқанда, шешіммен танысқан сәттен есептеледі.'},
    {k:'what',label:'Не болды',type:'textarea',req:1,ph:'Қандай шешім қабылданғанын немесе не істелгенін (істелмегенін) сипаттаңыз'},
    {k:'viol',label:'Сіздің қандай құқықтарыңыз бұзылды',type:'textarea',req:1,ph:'Мұны неліктен заңсыз немесе негізсіз деп санайсыз'},
    {k:'demand',label:'Не сұрайсыз (әр талап жаңа жолдан)',type:'textarea',req:1,ph:'Мысалы: әрекетсіздікті заңсыз деп тану\nБұзушылықты жоюды міндеттеу'},
    artField('ҚР ҚПК-нің 100, 105-баптарына (сотқа дейінгі тәртіп) немесе ҚР ҚПК-нің 100, 106-баптарына (сот тәртібі)'),DATEF],
  demo:{addressee:DEMO_PROS,fio:'Смирнов Алексей Алексеевич',status:'Айыпталушы',caseNo:'123456',address:'Астана қ., Үлгі к-сі, 3-үй',phone:'+7 700 000 00 02',
    channel:CH1,kind:'әрекетсіздік',official:'Тергеуші [орган, Т. А. Ә.]',actDate:'2026-10-01',knowDate:'2026-10-05',
    what:'мен берген өтінішхат қаралмады, жауап алмадым',viol:'мен өз мүдделерімді қорғау және дәлелдемелер ұсыну мүмкіндігінен айрылдым',demand:'Тергеушінің әрекетсіздігін заңсыз деп тануды\nӨтінішхатты қарауды міндеттеуді'},
  build:function(h){
    var kind=h.v('kind'),court=h.o('channel').indexOf('Сотқа')===0;
    var def=court?'ҚР ҚПК-нің 100, 106-баптарына':'ҚР ҚПК-нің 100, 105-баптарына';
    var dat={'шешім':'шешіміне','әрекет':'әрекетіне','әрекетсіздік':'әрекетсіздігіне'}[kind]||kind;
    var b=head(h,'ШАҒЫМ',h.v('official')+' (орган немесе лауазымды адам) '+dat);
    b.push(h.o('status')===OTHER?{a:'j',t:'Мен, '+h.v('fio')+', № '+h.v('caseNo')+' қылмыстық іс бойынша құқықтары мен заңды мүдделері қозғалған адам ретінде жүгінемін (ҚР ҚПК-нің 100-бабы).'}:intro(h));
    b.push({a:'j',t:fmtDate(h.raw('actDate'))+' '+h.v('official')+': '+trimDot(h.v('what'))+'.'});
    if(h.raw('knowDate'))b.push({a:'j',t:'Шағымдалып отырған шешім, әрекет немесе әрекетсіздік туралы маған '+fmtDate(h.raw('knowDate'))+' белгілі болды.'});
    b.push({a:'j',t:'Аталған '+kind+' менің құқықтарым мен заңды мүдделерімді бұзады деп санаймын, өйткені '+trimDot(h.v('viol'))+'.'});
    b.push({a:'j',t:'Жоғарыда айтылғанның негізінде және '+h.art(def)+' сәйкес сұраймын:'});
    var d=lines(h.v('demand'));
    d.forEach(function(x,i){b.push({a:'l',t:(d.length>1?(i+1)+'. ':'')+x});});
    return b.concat(sign(h));
  }}
};
DOCS.complaint.step1[2].options=STATUSES.concat([OTHER]);

DOCS.specialist={
  title:'Маманнан жауап алу туралы өтінішхат',file:'Mamannan_zhauap_alu_turaly_otinishhat',
  step1:common('Кімге жолданады',ADDR_HINT,ADDR1),
  step2:[
    {k:'sName',label:'Маманның Т. А. Ә.',req:1,ph:'Т. А. Ә. енгізіңіз'},
    {k:'sSpec',label:'Мамандығы немесе білім саласы',req:1,ph:'Мысалы: сарапшы-дәрігер, бағалаушы, құрылыс инженері'},
    {k:'circ',label:'Маманның қандай мәселелер бойынша түсініктеме бергені қажет',type:'textarea',req:1,ph:'Мысалы: зақымдану себептері мен мөлшері туралы'},
    {k:'reason',label:'Қандай мән-жайларды анықтау үшін жауап алу қажет',type:'textarea',req:1,ph:'Мысалы: мүліктің зақымдану сипаты мен құны',hint:'ҚР ҚПК-нің 99-бабы бойынша өтінішхатта қандай мән-жайларды анықтау үшін берілгені көрсетілуі керек.'},
    {k:'appear',label:'Маманның келуі',type:'select',options:['Маманды хабарламамен шақыруды сұраймын','Маманның келуін өзім қамтамасыз етемін'],hint:'Егер келуін тарап қамтамасыз етсе, іс жүргізуші орган маманнан жауап алудан бас тартуға құқылы емес (ҚР ҚПК-нің 99-бабы). Таңдамасаңыз, мәтінде шақыру туралы өтініш болады.'},
    artField(A99),DATEF],
  demo:{addressee:DEMO_ADDR,fio:'Иванов Иван Иванович',status:'Жәбірленуші',caseNo:'123456',address:'Астана қ., Үлгі к-сі, 1-үй, 1-пәтер',phone:'+7 700 000 00 00',
    sName:'Кузнецов Кирилл Константинович',sSpec:'бағалаушы',circ:'зақымданған мүліктің құны туралы',reason:'келтірілген залалдың мөлшері',appear:'Маманның келуін өзім қамтамасыз етемін'},
  build:function(h){
    var b=head(h,'ӨТІНІШХАТ','маманнан жауап алу туралы'),self=h.o('appear').indexOf('өзім')>-1;
    b.push(intro(h));
    b.push({a:'j',t:h.art(A99)+' сәйкес маман ретінде мына адамнан '+(self?'':'шақырып, ')+'жауап алуды сұраймын: '+h.v('sName')+' ('+h.v('sSpec')+').'});
    if(self)b.push({a:'j',t:'Маманның тағайындалған уақытта келуін өзім қамтамасыз етуге міндеттенемін.'});
    b.push({a:'j',t:'Маманның түсініктемесі мына мәселелер бойынша қажет: '+trimDot(h.v('circ'))+'.'});
    b.push({a:'j',t:'Жауап алу мына мән-жайларды анықтау үшін қажет: '+trimDot(h.v('reason'))+'.'});
    b.push({a:'j',t:BASIS});
    b.push({a:'l',t:'1. Көрсетілген маманнан '+(self?'':'шақырып, ')+'жауап алуды.'});
    b.push({a:'l',t:'2. '+NOTIFY});
    return b.concat(sign(h));
  }};

DOCS.refusal={
  title:'Өтінішхатты қанағаттандырудан бас тартуға шағым',file:'Shagym_otinishhattan_bas_tartuga',
  step1:common('Шағым кімге жолданады','Шағымды қарайтын адамның лауазымы және Т. А. Ә. Берілу тәртібін келесі қадамда таңдайсыз.','Лауазымы, органы, Т. А. Ә.'),
  step2:[
    {k:'channel',label:'Шағымды қайда бересіз',type:'select',req:1,options:[CH1,CH2],hint:'Сотқа ҚР ҚПК-нің 106-бабында көрсетілген мәселелер бойынша 15 тәулік ішінде жүгінуге болады.'},
    {k:'official',label:'Өтінішхатты кім қанағаттандырмады',req:1,ph:'Лауазымы, органы, Т. А. Ә.'},
    {k:'about',label:'Өтінішхат не туралы болды',type:'textarea',req:1,ph:'Мысалы: куәден жауап алу туралы, шартты қосып тігу туралы'},
    {k:'hodDate',label:'Өтінішхат күні (қалауыңыз бойынша)',type:'date'},
    {k:'decDate',label:'Бас тарту туралы қаулы күні',type:'date',req:1},
    {k:'knowDate',label:'Бас тарту туралы қашан білдіңіз (қалауыңыз бойынша)',type:'date',hint:'Сотқа шағымдану үшін ҚР ҚПК-нің 106-бабы бойынша 15 тәулік, атап айтқанда, шешіммен танысқан сәттен есептеледі.'},
    {k:'viol',label:'Бас тарту неліктен негізсіз',type:'textarea',req:1,ph:'Сіздің ойыңызша, өтінішхатты неліктен қанағаттандыру керек еді',hint:'ҚР ҚПК-нің 99-бабы бойынша өтінішхат істің мән-жайларын жан-жақты, толық және объективті зерттеуге ықпал етсе, қанағаттандырылады.'},
    {k:'demand',label:'Не сұрайсыз (әр талап жаңа жолдан)',type:'textarea',req:1,ph:'Мысалы: бас тарту туралы қаулыны жою\nӨтінішхатты қанағаттандыру'},
    artField('ҚР ҚПК-нің 99, 100, 105-баптарына (сотқа дейінгі тәртіп) немесе ҚР ҚПК-нің 99, 100, 106-баптарына (сот тәртібі)'),DATEF],
  demo:{addressee:DEMO_PROS,fio:'Смирнов Алексей Алексеевич',status:'Айыпталушы',caseNo:'123456',address:'Астана қ., Үлгі к-сі, 3-үй',phone:'+7 700 000 00 02',
    channel:CH1,official:'Тергеуші [орган, Т. А. Ә.]',about:'куәден жауап алу туралы',hodDate:'2026-09-20',decDate:'2026-10-01',knowDate:'2026-10-05',
    viol:'куәнің жауабы істің мән-жайларын анықтау үшін маңызды, ал куәнің келуін мен қамтамасыз етемін',demand:'Бас тарту туралы қаулыны жоюды\nӨтінішхатты қанағаттандыруды'},
  build:function(h){
    var court=h.o('channel').indexOf('Сотқа')===0,def=court?'ҚР ҚПК-нің 99, 100, 106-баптарына':'ҚР ҚПК-нің 99, 100, 105-баптарына';
    var b=head(h,'ШАҒЫМ',h.v('official')+' (орган немесе лауазымды адам) өтінішхатты қанағаттандырудан бас тарту туралы қаулыға');
    b.push(h.o('status')===OTHER?{a:'j',t:'Мен, '+h.v('fio')+', № '+h.v('caseNo')+' қылмыстық іс бойынша құқықтары мен заңды мүдделері қозғалған адам ретінде жүгінемін (ҚР ҚПК-нің 100-бабы).'}:intro(h));
    b.push({a:'j',t:'Мен өтінішхат берген едім'+(h.raw('hodDate')?' ('+fmtDate(h.raw('hodDate'))+')':'')+': '+trimDot(h.v('about'))+'.'});
    b.push({a:'j',t:fmtDate(h.raw('decDate'))+' '+h.v('official')+' осы өтінішхатты қанағаттандырудан бас тарту туралы қаулы шығарды.'});
    if(h.raw('knowDate'))b.push({a:'j',t:'Бас тарту туралы маған '+fmtDate(h.raw('knowDate'))+' белгілі болды.'});
    b.push({a:'j',t:'Бас тартуды негізсіз деп санаймын, өйткені '+trimDot(h.v('viol'))+'.'});
    b.push({a:'j',t:'Жоғарыда айтылғанның негізінде және '+h.art(def)+' сәйкес сұраймын:'});
    var d=lines(h.v('demand'));
    d.forEach(function(x,i){b.push({a:'l',t:(d.length>1?(i+1)+'. ':'')+x});});
    return b.concat(sign(h));
  }};
DOCS.refusal.step1[2].options=STATUSES.concat([OTHER]);

DOCS.withdraw={
  title:'Шағымды кері қайтарып алу туралы арыз',file:'Shagymdy_keri_kaitaryp_alu_turaly_aryz',
  step1:common('Арыз кімге жолданады','Шағымыңызды қарайтын адамға','Лауазымы, органы, Т. А. Ә.'),
  step2:[
    {k:'filedDate',label:'Шағым берілген күн',type:'date',req:1},
    {k:'subject',label:'Шағым нені дауласу туралы болды',type:'textarea',req:1,ph:'Мысалы: тергеушінің менің өтінішхатым бойынша әрекетсіздігіне'},
    {k:'why',label:'Қайтарып алу себебі (қалауыңыз бойынша)',type:'textarea',ph:'Мысалы: мәселе шешілді'},
    artField('ҚР ҚПК-нің 100-бабына'),DATEF],
  demo:{addressee:DEMO_PROS,fio:'Смирнов Алексей Алексеевич',status:'Айыпталушы',caseNo:'123456',address:'Астана қ., Үлгі к-сі, 3-үй',phone:'+7 700 000 00 02',
    filedDate:'2026-10-10',subject:'тергеушінің менің өтінішхатым бойынша әрекетсіздігі',why:'өтінішхат қаралды'},
  build:function(h){
    var b=head(h,'АРЫЗ','шағымды кері қайтарып алу туралы');
    b.push(intro(h));
    b.push({a:'j',t:fmtDate(h.raw('filedDate'))+' мен шағым бердім: '+trimDot(h.v('subject'))+'.'});
    if(h.o('why'))b.push({a:'j',t:'Қайтарып алу себебі: '+trimDot(h.o('why'))+'.'});
    b.push({a:'j',t:h.art('ҚР ҚПК-нің 100-бабына')+' сәйкес сұраймын:'});
    b.push({a:'l',t:'Көрсетілген шағымды кері қайтарып алуды қабылдауды.'});
    return b.concat(sign(h));
  }};

DOCS.restore={
  title:'Өткізіп алған мерзімді қалпына келтіру туралы өтінішхат',file:'Merzimdi_kalpyna_keltiru_turaly_otinishhat',
  step1:common('Кімге жолданады','Іс өндірісінде жүрген анықтаушыға, тергеушіге, прокурорға немесе судьяға','Лауазымы, органы немесе соты, Т. А. Ә.'),
  step2:[
    {k:'missed',label:'Қандай мерзім өткізіп алынды',type:'textarea',req:1,ph:'Мысалы: 01.10.2026 жылғы қаулыға шағым беру мерзімі'},
    {k:'endDate',label:'Мерзім аяқталған күн',type:'date',req:1},
    {k:'reason',label:'Мерзімді өткізіп алудың дәлелді себебі',type:'textarea',req:1,ph:'Мысалы: стационарда емделдім',hint:'ҚР ҚПК-нің 50-бабы бойынша мерзім дәлелді себеппен өткізіп алынса, қалпына келтірілуі мүмкін. Оны растайтын құжаттарды қоса беріңіз.'},
    {k:'proofs',label:'Себепті растайтын құжаттар (әрқайсысы жаңа жолдан)',type:'textarea',ph:'Мысалы: 10.10.2026 жылғы аурухана анықтамасы'},
    {k:'action',label:'Мерзім қалпына келтірілгеннен кейін не істейсіз',type:'textarea',req:1,ph:'Мысалы: 01.10.2026 жылғы қаулыға шағым беремін'},
    {k:'suspend',label:'Шешімнің орындалуын тоқтата тұру',type:'select',options:['Шағымданылған шешімнің орындалуын мерзімді қалпына келтіру мәселесі шешілгенге дейін тоқтата тұруды сұраймын','Сұрамаймын'],hint:'ҚР ҚПК-нің 50-бабы бойынша мерзім өткізіліп шағымданылған шешімнің орындалуы мерзімді қалпына келтіру мәселесі шешілгенге дейін тоқтатылуы мүмкін.'},
    artField('ҚР ҚПК-нің 50-бабына'),DATEF],
  demo:{addressee:DEMO_ADDR,fio:'Смирнов Алексей Алексеевич',status:'Айыпталушы',caseNo:'123456',address:'Астана қ., Үлгі к-сі, 3-үй',phone:'+7 700 000 00 02',
    missed:'01.10.2026 жылғы қаулыға шағым беру мерзімі',endDate:'2026-10-16',reason:'осы мерзім ішінде мен стационарда емделдім',proofs:'20.10.2026 жылғы аурухана анықтамасы',action:'01.10.2026 жылғы қаулыға шағым беремін',suspend:'Сұрамаймын'},
  build:function(h){
    var b=head(h,'ӨТІНІШХАТ','өткізіп алған мерзімді қалпына келтіру туралы');
    b.push(intro(h));
    b.push({a:'j',t:'Мен мерзімді өткізіп алдым: '+trimDot(h.v('missed'))+'. Мерзім '+fmtDate(h.raw('endDate'))+' аяқталды.'});
    b.push({a:'j',t:'Мерзім дәлелді себеппен өткізіп алынды: '+trimDot(h.v('reason'))+'.'});
    var pr=lines(h.o('proofs'));
    b.push({a:'j',t:'Мерзім қалпына келтірілгеннен кейін: '+trimDot(h.v('action'))+'.'});
    b.push({a:'j',t:'Жоғарыда айтылғанның негізінде және '+h.art('ҚР ҚПК-нің 50-бабына')+' сәйкес сұраймын:'});
    var n=1;
    b.push({a:'l',t:(n++)+'. Өткізіп алынған мерзімді қалпына келтіруді.'});
    if(h.o('suspend').indexOf('Шағымданылған')===0)b.push({a:'l',t:(n++)+'. Шағымданылған шешімнің орындалуын мерзімді қалпына келтіру мәселесі шешілгенге дейін тоқтата тұруды.'});
    b.push({a:'l',t:(n++)+'. '+NOTIFY});
    if(pr.length){b.push({a:'l',t:'Қосымша:'});pr.forEach(function(x,i){b.push({a:'l',t:(i+1)+'. '+x});});}
    return b.concat(sign(h));
  }};

DOCS.crime={
  title:'Қылмыстық құқық бұзушылық туралы арыз',file:'Kylmystyk_kukyk_buzushylyk_turaly_aryz',
  step1:[
    {k:'addressee',label:'Арыз кімге жолданады',req:1,ph:'Жүгінетін қылмыстық қудалау органы',hint:'Арыз беретін органды көрсетіңіз. Анонимді хабар тергеп-тексеруге себеп бола алмайды (ҚР ҚПК-нің 181-бабы), сондықтан өз деректеріңізді жазыңыз.'},
    {k:'fio',label:'Сіздің Т. А. Ә.',req:1,ph:'Т. А. Ә. енгізіңіз'},
    {k:'idDoc',label:'Жеке басты куәландыратын құжат (қалауыңыз бойынша)',ph:'Мысалы: жеке куәлік № ...'},
    {k:'address',label:'Тұрғылықты немесе жұмыс орнының мекенжайы',req:1,ph:'Қайда тұрасыз немесе жұмыс істейсіз'},
    {k:'phone',label:'Телефон',ph:'+7 ...'}
  ],
  step2:[
    {k:'when',label:'Қашан болды',req:1,ph:'Белгілі болса, күні мен уақыты'},
    {k:'where',label:'Қайда болды',req:1,ph:'Мекенжайы немесе орны'},
    {k:'what',label:'Не болды',type:'textarea',req:1,ph:'Ретімен сипаттаңыз: не болды, не істелді, салдары қандай'},
    {k:'who',label:'Сіздің ойыңызша, кім қатысты (белгілі болса)',type:'textarea',ph:'Т. А. Ә., сипаттамасы немесе «белгісіз адам»'},
    {k:'harm',label:'Қандай зиян келтірілді',type:'textarea',ph:'Моральдық, физикалық немесе мүліктік: нақты қандай және қандай сомада'},
    {k:'evid',label:'Сіздің сөздеріңізді не растайды (әрқайсысы жаңа жолдан)',type:'textarea',ph:'Құжаттар, фото, куәлардың аты-жөні'},
    artField('ҚР ҚПК-нің 180, 181-баптарына'),DATEF],
  demo:{addressee:'Басшыға [орган, Т. А. Ә.]',fio:'Иванов Иван Иванович',idDoc:'жеке куәлік № 000000000',address:'Астана қ., Үлгі к-сі, 1-үй, 1-пәтер',phone:'+7 700 000 00 00',
    when:'01.10.2026, сағат 15:00 шамасында',where:'Астана қ., Үлгілі к-сі, 5-үй',what:'белгісіз адам менің сөмкемді ашық тартып алып, жасырынып кетті',who:'белгісіз адам',harm:'шамамен 50 000 теңге сомасындағы мүлік жоғалды',evid:'Куәнің арызы\nБақылау камерасының жазбасы'},
  build:function(h){
    var b=[{a:'r',t:h.v('addressee')},{a:'r',t:'Өтініш беруші: '+h.v('fio')}];
    if(h.o('idDoc'))b.push({a:'r',t:'жеке басты куәландыратын құжат: '+h.o('idDoc')});
    b.push({a:'r',t:'тұрғылықты немесе жұмыс орнының мекенжайы: '+h.v('address')});
    if(h.o('phone'))b.push({a:'r',t:'тел.: '+h.o('phone')});
    b.push({a:'c',b:1,t:'АРЫЗ'});b.push({a:'c',b:1,t:'қылмыстық құқық бұзушылық туралы'});
    b.push({a:'j',t:'Хабарлаймын: '+trimDot(h.v('when'))+', мына мекенжай бойынша: '+trimDot(h.v('where'))+' мынадай жағдай болды: '+trimDot(h.v('what'))+'.'});
    if(h.o('who'))b.push({a:'j',t:'Менің ойымша, оған мына адам қатысты: '+trimDot(h.o('who'))+'.'});
    if(h.o('harm'))b.push({a:'j',t:'Нәтижесінде маған зиян келтірілді: '+trimDot(h.o('harm'))+'.'});
    var ev=lines(h.o('evid'));
    if(ev.length){b.push({a:'j',t:'Менің сөздерімді мыналар растайды:'});ev.forEach(function(x,i){b.push({a:'l',t:(i+1)+'. '+x});});}
    b.push({a:'j',t:'Жоғарыда айтылғанның негізінде және '+h.art('ҚР ҚПК-нің 180, 181-баптарына')+' сәйкес осы арызды қабылдап, оны Қазақстан Республикасының Қылмыстық-процестік кодексінде белгіленген тәртіппен қарауды, ал қабылданған шешім туралы маған хабарлауды сұраймын.'});
    b.push({a:'j',t:'Көрінеу жалған сөз жеткізгені үшін қылмыстық жауаптылық туралы ескертілдім.'});
    return b.concat(sign(h));
  }};

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
  else if(f.type==='select')inp='<select id="'+id+'" data-k="'+f.k+'"><option value="">Таңдаңыз</option>'+f.options.map(function(o){return '<option'+(o===val?' selected':'')+'>'+esc(o)+'</option>';}).join('')+'</select>';
  else inp='<input id="'+id+'" data-k="'+f.k+'" type="'+(f.type||'text')+'" value="'+esc(val)+'" placeholder="'+esc(f.ph||'')+'" autocomplete="off">';
  return '<div class="field" data-f="'+f.k+'"><label for="'+id+'">'+esc(f.label)+(f.req?' <em aria-hidden="true">*</em>':'')+'</label>'+inp+(f.hint?'<span class="hint">'+esc(f.hint)+'</span>':'')+'<span class="err" role="alert"></span></div>';
}
function stepper(){
  var n=['Жалпы мәліметтер','Мазмұны','Қарау және файл'];
  return '<ol class="steps">'+n.map(function(t,i){var s=i+1;return '<li'+(s===S.step?' aria-current="step"':'')+(s<S.step?' class="done"':'')+'><i>'+s+'</i><span>'+t+'</span></li>';}).join('')+'</ol>';
}
function draw(){
  var root=$('#tool'),html;
  if(S.step<3){
    var fs=S.step===1?D.step1:D.step2;
    html=stepper()+'<form class="panel" id="form" novalidate><div class="row" style="justify-content:space-between"><h2>'+(S.step===1?'Жалпы мәліметтер':'Мазмұны')+'</h2><button type="button" class="demo" id="demo">Мысалды қою</button></div>'+
      fs.map(fieldHtml).join('')+
      '<div class="actions">'+(S.step>1?'<button type="button" class="btn ghost" id="back">Артқа</button>':'<span></span>')+'<button type="submit" class="btn grow">'+(S.step===1?'Әрі қарай':'Тексеріп, құжатты көрсету')+'</button></div></form>';
  }else{
    var h=helpers(S.V),blocks=D.build(h),req=D.step1.concat(D.step2).filter(function(f){return f.req;});
    var hasPh=blocks.some(function(b){return /\[[^\]]*\]/.test(b.t);});
    html=stepper()+'<div class="panel"><h2>Тексеру және қарау</h2><ul class="checks">'+
      '<li><span class="ok">✓</span><span>Міндетті өрістер толтырылған: '+req.length+'/'+req.length+'.</span></li>'+
      '<li><span class="ok">✓</span><span>Құжат түрі: '+esc(D.title)+'.</span></li>'+
      '<li><span class="'+(hasPh?'wn':'ok')+'">'+(hasPh?'!':'✓')+'</span><span>'+(hasPh?'Мәтінде төртбұрышты жақшадағы орындар қалды, олар сары түспен белгіленген. Оларды Word-та өз деректеріңізбен ауыстырыңыз.':'Мәтінде толтырылмаған орындар жоқ.')+'</span></li></ul>'+
      '<div class="paperwrap"><div class="paper" id="paper">'+blocks.map(function(b){
        var cls=(b.a==='r'?'r':b.a==='c'?'c':b.a==='j'?'j':'')+(b.b?' b':'');
        return '<p class="'+cls+'">'+esc(b.t).replace(/(\[[^\]]*\])/g,'<span class="ph">$1</span>')+'</p>';
      }).join('')+'</div></div>'+
      '<section class="checklist noprint"><h3>Тапсырар алдында</h3><ul><li>Адресатты, Т. А. Ә., іс нөмірін және күнді тексеріңіз.</li><li>Төртбұрышты жақшадағы орындарды өз деректеріңізбен ауыстырыңыз.</li><li>Құжатқа өз қолыңызбен қол қойыңыз.</li><li>Көшірмесін жасап, қабылданғаны туралы белгі қойдырыңыз: құжатты жеке тапсырсаңыз, күні кеңсенің немесе лауазымды адамның белгісі бойынша анықталады (ҚР ҚПК-нің 49-бабы).</li><li>Поштамен жіберсеңіз, мерзім аяқталғанға дейін тапсырып, түбіртекті сақтаңыз: күні пошта штемпелі бойынша анықталады (ҚР ҚПК-нің 49-бабы).</li><li>Құжатты заңгерге немесе заң консультациясына көрсетіңіз.</li></ul><p class="small"><a href="kk-sroki.html">Мерзімді калькуляторда тексеру</a></p></section>'+'<div class="actions noprint"><button class="btn ghost" id="back">Деректерді өзгерту</button><div class="row grow"><button class="btn" id="dl-docx">DOCX жүктеу</button><button class="btn ghost" id="dl-txt">TXT жүктеу</button><button class="btn ghost" id="copy">Мәтінді көшіру</button><button class="btn ghost" id="print">Басып шығару / PDF</button></div></div></div>'+
      '<p class="notice" style="margin-top:16px"><strong>Құжатты тапсырар алдында тексеріңіз.</strong> Бұл студенттердің оқу жобасы, авторлар заңгер емес. Құжатты заңгерге немесе заң консультациясына көрсетіңіз. Қазақ тіліндегі мәтінді тіл маманы тексермеген.</p>';
  }
  root.innerHTML=html;bind();
}
function validate(fields){
  var first=null;
  fields.forEach(function(f){
    var box=$('[data-f="'+f.k+'"]');if(!box)return;
    var bad=f.req&&!(S.V[f.k]||'').trim();
    box.classList.toggle('bad',bad);
    $('.err',box).textContent=bad?'Бұл өрісті толтырыңыз.':'';
    var inp=$('[data-k]',box);if(inp)inp.setAttribute('aria-invalid',bad?'true':'false');
    if(bad&&!first)first=inp;
  });
  if(first){first.focus();toast('Міндетті өрістерді толтырыңыз, олар жұлдызшамен белгіленген.');}
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
    $('#demo').addEventListener('click',function(){Object.keys(D.demo).forEach(function(k){S.V[k]=D.demo[k];});draw();toast('Ойдан шығарылған деректер қойылды.');});
  }
  var b=$('#back');if(b)b.addEventListener('click',function(){go(S.step-1);});
  var c=$('#copy');if(c)c.addEventListener('click',copyText);
  var pr=$('#print');if(pr)pr.addEventListener('click',function(){try{window.print();}catch(e){toast('Басып шығаруды ашу мүмкін болмады. DOCX жүктеңіз.');}});
  var x=$('#dl-docx');if(x)x.addEventListener('click',function(){save('docx');});
  var t=$('#dl-txt');if(t)t.addEventListener('click',function(){save('txt');});
}
function docBlocks(){return D.build(helpers(S.V));}
function docText(){return docBlocks().map(function(b){return b.t;}).join('\n\n');}
function copyText(){
  function fallback(){var r=document.createRange();r.selectNodeContents($('#paper'));var s=getSelection();s.removeAllRanges();s.addRange(r);toast('Мәтін таңдалды. Телефон мәзірі арқылы көшіріңіз.');}
  try{navigator.clipboard.writeText(docText()).then(function(){toast('Мәтін көшірілді.');},fallback);}catch(e){fallback();}
}
function save(kind){
  try{
    var blob=kind==='docx'?buildDocx(docBlocks()):new Blob([docText()],{type:'text/plain;charset=utf-8'});
    var url=URL.createObjectURL(blob),a=document.createElement('a');
    a.href=url;a.download=D.file+'.'+kind;document.body.appendChild(a);a.click();
    setTimeout(function(){URL.revokeObjectURL(url);a.remove();},4000);
    toast('Файл жасалды. Жүктеулер қалтасын тексеріңіз.');
  }catch(e){toast('Файл жасау мүмкін болмады. «Мәтінді көшіру» түймесін басыңыз.');}
}

/* ---------- құжатты таңдау ---------- */
var PICK={
  q1:{t:'Не істегіңіз келеді?',o:[['Органнан немесе соттан бір нәрсе сұрау (өтінішхат)','q2h'],['Шешімге, әрекетке немесе әрекетсіздікке шағымдану','q2z'],['Бұрын берген шағымымды кері қайтарып алу','withdraw'],['Мерзімді өткізіп алдым, оны қалпына келтіргім келеді','restore'],['Қылмыс туралы хабарлау (арыз беру)','crime']]},
  q2h:{t:'Нені сұрайсыз?',o:[['Іске құжаттарды немесе заттарды қосып тігу','evidence'],['Іс туралы білетін адамнан (куәден) жауап алу','witness'],['Маманнан (бағалаушыдан, дәрігерден, инженерден және т. б.) жауап алу','specialist']]},
  q2z:{t:'Нені дауласасыз?',o:[['Менің өтінішхатымды қанағаттандырудан бас тартуды','refusal'],['Өзге шешімді, әрекетті немесе әрекетсіздікті','complaint']]}
};
var PICKFILE={
  evidence:['kk-hodatajstvo-o-priobshchenii-dokazatelstv.html','Дәлелдемелерді қосып тігу туралы өтінішхат'],
  witness:['kk-hodatajstvo-o-doprose-svidetelya.html','Куәден жауап алу туралы өтінішхат'],
  specialist:['kk-hodatajstvo-o-doprose-specialista.html','Маманнан жауап алу туралы өтінішхат'],
  complaint:['kk-zhaloba-na-reshenie-ili-dejstvie.html','Іс жүргізу шешіміне немесе әрекетіне шағым'],
  refusal:['kk-zhaloba-na-otkaz-v-hodatajstve.html','Өтінішхатты қанағаттандырудан бас тартуға шағым'],
  withdraw:['kk-zayavlenie-ob-otzyve-zhaloby.html','Шағымды кері қайтарып алу туралы арыз'],
  restore:['kk-hodatajstvo-o-vosstanovlenii-sroka.html','Өткізіп алған мерзімді қалпына келтіру туралы өтінішхат'],
  crime:['kk-zayavlenie-o-pravonarushenii.html','Қылмыстық құқық бұзушылық туралы арыз']
};
function initPick(){
  var root=$('#pick'),trail=[];
  function show(id){
    if(PICKFILE[id]){
      var f=PICKFILE[id];
      root.innerHTML='<div class="panel"><h2>Сізге мыналар сәйкес келуі мүмкін</h2><p class="pickres">'+esc(f[1])+'</p><p class="small">Бұл сіздің жауаптарыңыз бойынша берілген кеңес, заңгерлік консультация емес. Күмәніңіз болса, жағдайды заңгерге көрсетіңіз.</p><div class="actions"><button class="btn ghost" id="pk-back">Артқа</button><a class="btn grow" href="'+f[0]+'">Құжат жасау</a></div></div>';
      $('#pk-back').addEventListener('click',function(){show(trail.pop()||'q1');});
      return;
    }
    var q=PICK[id];
    root.innerHTML='<div class="panel"><h2>'+esc(q.t)+'</h2><div class="opts">'+q.o.map(function(o,i){return '<button class="opt" data-i="'+i+'">'+esc(o[0])+'</button>';}).join('')+'</div>'+(trail.length?'<div class="actions"><button class="btn ghost" id="pk-back">Артқа</button><span></span></div>':'')+'</div>';
    $$('.opt',root).forEach(function(b){b.addEventListener('click',function(){trail.push(id);show(q.o[+b.dataset.i][1]);});});
    var bk=$('#pk-back');if(bk)bk.addEventListener('click',function(){show(trail.pop());});
  }
  show('q1');
}

/* ---------- мерзім калькуляторы (ҚР ҚПК-нің 48-бабы) ---------- */
var WD=['жексенбі','дүйсенбі','сейсенбі','сәрсенбі','бейсенбі','жұма','сенбі'];
function isoToDate(iso){var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(iso||'');return m?new Date(Date.UTC(+m[1],+m[2]-1,+m[3])):null;}
function dateToIso(d){return d.getUTCFullYear()+'-'+('0'+(d.getUTCMonth()+1)).slice(-2)+'-'+('0'+d.getUTCDate()).slice(-2);}
function deadline(iso,n){
  var d=isoToDate(iso);if(!d)return null;
  d.setUTCDate(d.getUTCDate()+n);var shifted=false;
  while(d.getUTCDay()===0||d.getUTCDay()===6){d.setUTCDate(d.getUTCDate()+1);shifted=true;}
  return {iso:dateToIso(d),wd:WD[d.getUTCDay()],shifted:shifted};
}
function fmtDeadline(r){return fmtDate(r.iso)+' ('+r.wd+'), 24:00-ге дейін'+(r.shifted?'. Соңғы тәулік демалыс күніне сәйкес келді, сондықтан мерзім келесі жұмыс күніне ауыстырылды':'');}
var CALC=[
  {k:'hod',t:'Өтінішхатты қарау (99-бап)',lab:'Өтінішхатты берген күн',n:3,out:'Өтінішхат дереу қаралады, ал бұл мүмкін болмаса, ол мәлімделген күннен бастап үш тәуліктен кешіктірілмей қаралады. Соңғы мерзім:'},
  {k:'zh',t:'Шағымды басшының немесе прокурордың қарауы (105-бап)',lab:'Шағымды берген күн',n:7,out:'Қарау мерзімі жалпы жағдайда 7 тәулік; бірқатар мәселе бойынша 3 тәулік; ерекше жағдайларда 15 тәулікке дейін. Жалпы мерзім бойынша соңғы күн:'},
  {k:'sud',t:'Сотқа шағым беру: шешіммен танысқан күннен бастап (106-бап)',lab:'Шешіммен танысқан күніңіз',n:15,out:'Сотқа шағымды 15 тәулік ішінде беруге болады. Соңғы мерзім:'},
  {k:'sud2',t:'Сотқа шағым беру: прокурорға берілген шағымға жауап болмаса (106-бап)',lab:'Прокурорға шағым берген күніңіз',n:15,twice:1,out:'Алдымен прокурорға қарауға 15 тәулік беріледі, содан кейін сотқа жүгінуге тағы 15 тәулік.'},
  {k:'own',t:'Өз мерзіміңіз (тәулікпен)',lab:'Есептеуді бастайтын күн',n:0,own:1,out:'Соңғы мерзім:'}
];
function initCalc(){
  var root=$('#calc');
  root.innerHTML='<div class="panel"><div class="field"><label for="c-k">Қандай мерзімді есептеу керек</label><select id="c-k">'+CALC.map(function(c,i){return '<option value="'+i+'">'+esc(c.t)+'</option>';}).join('')+'</select></div>'+
    '<div class="field"><label for="c-d" id="c-l"></label><input id="c-d" type="date"></div>'+
    '<div class="field" id="c-nbox" hidden><label for="c-n">Неше тәулік</label><input id="c-n" type="number" min="1" max="365" value="10" inputmode="numeric"></div>'+
    '<div id="c-r" class="calcres" role="status" aria-live="polite"></div></div>';
  function upd(){
    var c=CALC[+$('#c-k').value],d=$('#c-d').value;$('#c-l').textContent=c.lab;$('#c-nbox').hidden=!c.own;
    if(!d){$('#c-r').innerHTML='<p class="small">Күнді таңдаңыз.</p>';return;}
    var n=c.own?Math.max(1,Math.min(365,parseInt($('#c-n').value,10)||1)):c.n,h;
    if(c.twice){var r1=deadline(d,15),r2=deadline(r1.iso,15);h='<p>'+esc(c.out)+'</p><p>Прокурор шағымды мына күнге дейін қарауы керек:</p><p class="pickres">'+esc(fmtDeadline(r1))+'</p><p>Жауап болмаса, сотқа шағымды мына күнге дейін беруге болады:</p><p class="pickres">'+esc(fmtDeadline(r2))+'</p>';}
    else{var r=deadline(d,n);h='<p>'+esc(c.out)+'</p><p class="pickres">'+esc(fmtDeadline(r))+'</p>';}
    $('#c-r').innerHTML=h+'<p class="small">Есеп ҚР ҚПК-нің 48-бабы бойынша жүргізілді: мерзім басталған тәулік есептелмейді; тәулікпен есептелетін мерзім соңғы тәуліктің жиырма төртінші сағатында аяқталады; егер ол жұмыс істемейтін күн болса, мерзім келесі жұмыс күні аяқталады. Қазақстанның мереке күндері мұнда ескерілмеген: соңғы күні мереке жоқ екенін күнтізбеден тексеріңіз. 49-бап бойынша шағым немесе құжат мерзім өтпей тұрып поштаға тапсырылса, мерзім өткізілген болып саналмайды (күні штемпель бойынша).</p>';
  }
  $('#c-k').addEventListener('change',upd);$('#c-d').addEventListener('input',upd);$('#c-n').addEventListener('input',upd);$('#c-d').value=today();upd();
}


/* ---------- запуск ---------- */
document.addEventListener('DOMContentLoaded',function(){
  if($('#q'))initHome();
  if($('#pick'))initPick();
  if($('#calc'))initCalc();
  var id=document.body.getAttribute('data-doc');
  if(id&&DOCS[id]&&$('#tool')){D=DOCS[id];draw();}
});
window.__ProDoc={buildDocx:buildDocx,DOCS:DOCS,helpers:helpers};
})();
