/** 友成幸福團隊｜Google Apps Script V8。只有帶有效門店 session 的呼叫能讀寫資料。 */
function props_() { return PropertiesService.getScriptProperties(); }
function dbId_() { const id=props_().getProperty('SPREADSHEET_ID'); if(!id) throw new Error('請管理者設定 SPREADSHEET_ID。'); return id; }

const DATA_SHEET = '業績圖卡_v1';
const PERSONNEL_SHEET = '人事資料_v1';
const PLATFORM_SHEET = '廣告平台_v1';
const PLATFORMS = ['591','YES319','Houseweb房屋網','樂屋網','好房網','我家網','YOUTUBE','Facebook臉書','痞客邦','其他網路平台'];
const HEADERS = ['識別碼','門店','月份','人員識別碼','姓名快照','庫存數量','年度業績目標_萬元','本月業績目標_萬元','目前業績_萬元','廣告明細_JSON','版本','更新時間','最後請求識別碼'];
const PERSONNEL_HEADERS = ['人員識別碼','門店','姓名','職務','狀態','建立時間','離職時間','版本','更新時間','最後請求識別碼'];
const NANPING_PEOPLE = [
  ['龔泰禧','主管'],['黃國興','業務'],['廖茂宏','業務'],['陳洋弘','業務'],['魏汝殷','業務'],['陳彥華','業務'],
  ['杜杰陽','業務'],['范瑞洪','業務'],['林湘縈','業務'],['詹明軒','業務'],['陳芮慈','業務'],['曾鈺傑','業務'],
  ['葉志傑','業務'],['林煥為','業務'],['卓正得','業務'],['董凱元','業務'],['古郁潔','業務'],['何百玲','業務'],
  ['劉淑燕','業務'],['李明憲(阿達)','業務'],['林佩蓉','業務'],['湯雅玲','業務'],['劉健安','業務']
];

function doGet() {
  const t = HtmlService.createTemplateFromFile('App');
  t.trustedOrigin = props_().getProperty('GITHUB_ORIGIN') || 'https://rakutennis923.github.io';
  return t.evaluate().setTitle('友成幸福團隊｜業績行動看板')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 僅在 Apps Script 編輯器執行。可重複執行，不刪除既有資料。 */
function setup() {
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss=SpreadsheetApp.openById(dbId_());
    const data=ensureSheet_(ss,DATA_SHEET,HEADERS);
    const people=ensureSheet_(ss,PERSONNEL_SHEET,PERSONNEL_HEADERS);
    let platforms=ss.getSheetByName(PLATFORM_SHEET);
    if(!platforms){platforms=ss.insertSheet(PLATFORM_SHEET);platforms.getRange(1,1,PLATFORMS.length+1,1).setValues([['廣告平台']].concat(PLATFORMS.map(p=>[p])));platforms.setFrozenRows(1);}
    if(people.getLastRow()===1) seedNanping_(people);
    data.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#123348').setFontColor('#ffffff');
    people.getRange(1,1,1,PERSONNEL_HEADERS.length).setFontWeight('bold').setBackground('#123348').setFontColor('#ffffff');
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
function ensureSheet_(ss,name,headers){let s=ss.getSheetByName(name);if(!s){s=ss.insertSheet(name);s.appendRow(headers);s.setFrozenRows(1);}checkHeaders_(s,headers);return s;}
function checkHeaders_(s,headers){if(JSON.stringify(s.getRange(1,1,1,headers.length).getValues()[0])!==JSON.stringify(headers))throw new Error('資料欄位已變更，請管理者還原欄位順序。');}
function seedNanping_(sheet){const now=new Date().toISOString();const rows=NANPING_PEOPLE.map(([name,role])=>[digest_(JSON.stringify(['友成南平店',name])),'友成南平店',name,role,'在職',now,'',1,now,'setup']);sheet.getRange(2,1,rows.length,PERSONNEL_HEADERS.length).setValues(rows);}

function codes_(){const codes=JSON.parse(props_().getProperty('STORE_CODES_JSON')||'{}');if(!codes||typeof codes!=='object'||Array.isArray(codes)||!Object.keys(codes).length)throw new Error('請管理者先設定各店通行碼。');Object.entries(codes).forEach(([k,v])=>{if(typeof v!=='string'||v.length<12||!k.trim())throw new Error('各店通行碼至少 12 字元。');});return codes;}
function getPublicConfig(){return{stores:Object.keys(codes_())};}
function login(store,code){const codes=codes_();if(typeof code!=='string'||code.length>200||!Object.prototype.hasOwnProperty.call(codes,store)||codes[store]!==code)throw new Error('門店或通行碼不正確。');const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put('session:'+token,JSON.stringify({store,codeHash:digest_(codes[store])}),21600);return{token,store};}
function auth_(token){if(typeof token!=='string'||token.length>100)throw new Error('登入已過期，請重新輸入通行碼。');const raw=CacheService.getScriptCache().get('session:'+token);if(!raw)throw new Error('登入已過期，請重新輸入通行碼。');const s=JSON.parse(raw),codes=codes_();if(!Object.prototype.hasOwnProperty.call(codes,s.store)||digest_(codes[s.store])!==s.codeHash)throw new Error('通行碼已更新，請重新登入。');return s.store;}
function logout(token){if(typeof token==='string'&&token.length<=100)CacheService.getScriptCache().remove('session:'+token);}
function digest_(s){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s)));}
function month_(m){if(typeof m!=='string'||!/^20\d{2}-(0[1-9]|1[0-2])$/.test(m))throw new Error('月份格式錯誤。');return m;}
function db_(){const s=SpreadsheetApp.openById(dbId_()).getSheetByName(DATA_SHEET);if(!s)throw new Error('尚未初始化資料庫，請管理者執行 setup_。');checkHeaders_(s,HEADERS);return s;}
function personnelDb_(){const s=SpreadsheetApp.openById(dbId_()).getSheetByName(PERSONNEL_SHEET);if(!s)throw new Error('尚未初始化人事資料庫，請管理者執行 setup_。');checkHeaders_(s,PERSONNEL_HEADERS);return s;}
function rows_(sheet,headers){return sheet.getLastRow()<2?[]:sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();}
function platforms_(){const s=SpreadsheetApp.openById(dbId_()).getSheetByName(PLATFORM_SHEET);if(!s||s.getLastRow()<2)throw new Error('廣告平台清單尚未設定。');return[...new Set(s.getRange(2,1,s.getLastRow()-1,1).getDisplayValues().flat().map(x=>x.trim()).filter(Boolean))];}

function cleanText_(v,label,max){if(typeof v!=='string')throw new Error(label+'格式不正確。');const s=v.normalize('NFC').trim();if(!s||s.length>max||/^[=+@\-]/.test(s)||/[\x00-\x1f\x7f]/.test(s))throw new Error(label+'不可空白或含有公式、控制字元。');return s;}
function requestId_(v){if(typeof v!=='string'||!/^[A-Za-z0-9-]{16,100}$/.test(v))throw new Error('請求識別碼不正確。');return v;}
function personObject_(row){const updated=row[8]instanceof Date?row[8].toISOString():String(row[8]);return{id:String(row[0]),store:String(row[1]),name:String(row[2]),role:String(row[3]),status:String(row[4]),created:String(row[5]),leftAt:String(row[6]||''),version:Number(row[7]),updated,etag:digest_(JSON.stringify(row))};}
function personnelForStore_(store){return rows_(personnelDb_(),PERSONNEL_HEADERS).filter(r=>String(r[1])===store).map(personObject_);}
function getPersonnel(token){return{people:personnelForStore_(auth_(token))};}

/** 新增或修改人員。改名仍保留相同識別碼與歷史業績。 */
function savePerson(token,input,baseEtag,requestId){
  const store=auth_(token);requestId_(requestId);const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const name=cleanText_(input&&input.name,'姓名',60),role=cleanText_(input&&input.role,'職務',30),sheet=personnelDb_(),rows=rows_(sheet,PERSONNEL_HEADERS);if(!['主管','業務'].includes(role))throw new Error('職務不在選單中。');
    const id=input&&input.id?cleanText_(input.id,'人員識別碼',100):Utilities.getUuid();
    const i=rows.findIndex(r=>String(r[0])===id&&String(r[1])===store),current=i<0?null:personObject_(rows[i]);
    if(current&&rows[i][9]===requestId)return{ok:true,person:current,replayed:true};
    if((current?current.etag:null)!==(baseEtag||null))return{ok:false,conflict:true,current};
    if(rows.some((r,n)=>n!==i&&String(r[1])===store&&String(r[2]).trim()===name&&String(r[4])==='在職'))throw new Error('已有同名的在職人員，請先確認是否為同一人。');
    const now=new Date().toISOString(),row=[id,store,name,role,current?current.status:'在職',current?current.created:now,current?current.leftAt:'',(current?current.version:0)+1,now,requestId];
    sheet.getRange(i<0?sheet.getLastRow()+1:i+2,1,1,PERSONNEL_HEADERS.length).setValues([row]);SpreadsheetApp.flush();return{ok:true,person:personObject_(row)};
  }finally{lock.releaseLock();}
}
function setPersonStatus(token,id,status,baseEtag,requestId){
  const store=auth_(token);requestId_(requestId);if(!['在職','離職'].includes(status))throw new Error('人員狀態不正確。');const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{const sheet=personnelDb_(),rows=rows_(sheet,PERSONNEL_HEADERS),i=rows.findIndex(r=>String(r[0])===id&&String(r[1])===store);if(i<0)throw new Error('找不到這位人員。');const current=personObject_(rows[i]);if(rows[i][9]===requestId)return{ok:true,person:current,replayed:true};if(current.etag!==(baseEtag||null))return{ok:false,conflict:true,current};if(status==='在職'&&rows.some((r,n)=>n!==i&&String(r[1])===store&&String(r[2]).trim()===current.name&&String(r[4])==='在職'))throw new Error('已有同名的在職人員，請先修改其中一人的姓名。');const now=new Date().toISOString(),row=[current.id,store,current.name,current.role,status,current.created,status==='離職'?now:'',current.version+1,now,requestId];sheet.getRange(i+2,1,1,PERSONNEL_HEADERS.length).setValues([row]);SpreadsheetApp.flush();return{ok:true,person:personObject_(row)};}finally{lock.releaseLock();}
}

function monthCell_(value){
  if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,'Asia/Taipei','yyyy-MM');
  const text=String(value||'').trim(),match=text.match(/^(20\d{2})[\/-](\d{1,2})(?:[\/-]\d{1,2})?$/);
  return match?month_(match[1]+'-'+String(Number(match[2])).padStart(2,'0')):month_(text);
}
function rowObject_(row,personMap){let ads;try{ads=JSON.parse(row[9]||'[]');}catch(e){throw new Error('廣告明細格式有誤，請管理者檢查試算表。');}const updated=row[11]instanceof Date?row[11].toISOString():String(row[11]);const person=personMap&&personMap[String(row[3])];return{id:String(row[0]),store:String(row[1]),month:monthCell_(row[2]),personId:String(row[3]),name:person?person.name:String(row[4]),personStatus:person?person.status:'未知',inventory:Number(row[5]),annual:Number(row[6]),target:Number(row[7]),actual:Number(row[8]),ads,version:Number(row[10]),updated,etag:digest_(JSON.stringify(row))};}
function getSnapshot(token,month){const store=auth_(token);month=month_(month);const people=personnelForStore_(store),map=Object.fromEntries(people.map(p=>[p.id,p]));return{records:rows_(db_(),HEADERS).filter(r=>String(r[1])===store&&monthCell_(r[2])===month).map(r=>rowObject_(r,map)),people,platforms:platforms_(),serverTime:new Date().toISOString()};}
function number_(v,label,integer){if(typeof v!=='number'||!Number.isFinite(v)||v<0||v>100000000||(integer&&!Number.isInteger(v)))throw new Error(label+'必須為有效的非負'+(integer?'整數':'數字')+'。');return integer?v:Math.round(v*100)/100;}
function payload_(p,allowed,people){if(!p||typeof p!=='object')throw new Error('資料不完整。');const person=people.find(x=>x.id===p.personId);if(!person)throw new Error('找不到這位人員，請重新整理人事資料。');if(person.status!=='在職'&&!p.allowInactive)throw new Error('這位人員已離職，不能新增當月資料。');const result={personId:person.id,name:person.name,month:month_(p.month),inventory:number_(p.inventory,'庫存數量',true),annual:number_(p.annual,'年度目標'),target:number_(p.target,'本月目標'),actual:number_(p.actual,'目前業績')};if(!Array.isArray(p.ads)||p.ads.length>40)throw new Error('廣告平台資料不正確。');const seen=new Set();result.ads=p.ads.map(a=>{if(!a||!allowed.includes(a.platform))throw new Error('廣告平台不在選單中，請重新整理平台清單。');const other=a.platform==='其他網路平台'?cleanText_(a.other,'其他平台名稱',80):'';const key=a.platform+'|'+other;if(seen.has(key))throw new Error('相同平台請合併為一筆數量。');seen.add(key);return{platform:a.platform,other,quantity:number_(a.quantity,'廣告數量',true)};});return result;}

/** 同一筆月資料採 etag 樂觀鎖 + Apps Script 互斥鎖，不直接覆蓋過期版本。 */
function saveRecord(token,input,baseEtag,requestId){
  const store=auth_(token);requestId_(requestId);const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{const people=personnelForStore_(store),p=payload_(input,platforms_(),people),sheet=db_(),rows=rows_(sheet,HEADERS),id=digest_(JSON.stringify([store,p.month,p.personId])),i=rows.findIndex(r=>String(r[0])===id),map=Object.fromEntries(people.map(x=>[x.id,x])),current=i<0?null:rowObject_(rows[i],map);if(current&&rows[i][12]===requestId)return{ok:true,record:current,replayed:true};if((current?current.etag:null)!==(baseEtag||null))return{ok:false,conflict:true,current};const row=[id,store,p.month,p.personId,p.name,p.inventory,p.annual,p.target,p.actual,JSON.stringify(p.ads),(current?current.version:0)+1,new Date().toISOString(),requestId],rowIndex=i<0?sheet.getLastRow()+1:i+2;sheet.getRange(rowIndex,3).setNumberFormat('@');sheet.getRange(rowIndex,1,1,HEADERS.length).setValues([row]);SpreadsheetApp.flush();return{ok:true,record:rowObject_(row,map)};}finally{lock.releaseLock();}
}
