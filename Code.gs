/**
 * 아라호 관제센터 — Google Apps Script 백엔드 (구글 시트 + Gemini 피드백)
 * ------------------------------------------------------------------
 * 설치 (5분):
 *  1) 새 구글 시트 만들기 → 확장 프로그램 → Apps Script
 *  2) 기존 내용을 지우고 이 파일 전체를 붙여넣기 → 저장
 *  3) 배포 → 새 배포 → 유형: 웹 앱
 *       - 다음 사용자 인증 정보로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *     → 배포 → 권한 허용 → "웹 앱 URL" 복사 (…/exec 로 끝남)
 *  4) index.html 맨 위 API = '여기에 붙여넣기'
 *  5) 시트에 자동으로 생기는 '설정' 탭에서
 *       GEMINI_KEY 칸에 Google AI Studio 키,  PIN 칸에 교사 코드 입력
 * 이후 학생 기록은 '학생' / '기록' 탭에 자동으로 쌓이고,
 * 앱의 선생님 탭에서 실시간으로 볼 수 있습니다.
 */

var SHEET_STUDENTS = '학생';
var SHEET_LOG = '기록';
var SHEET_CONF = '설정';

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name, header){
  var s = ss().getSheetByName(name);
  if(!s){ s = ss().insertSheet(name); if(header) s.appendRow(header); s.setFrozenRows(1); }
  return s;
}

function conf(key){
  var s = sheet(SHEET_CONF, ['항목','값','설명']);
  var v = s.getDataRange().getValues();
  if(v.length < 2){
    s.appendRow(['GEMINI_KEY','', 'Google AI Studio(aistudio.google.com)에서 발급한 키를 B열에 붙여넣기']);
    s.appendRow(['PIN','0726','선생님 탭 입장 코드']);
    s.appendRow(['MODEL','gemini-2.5-flash','피드백에 쓸 모델 (안 되면 gemini-2.0-flash)']);
    v = s.getDataRange().getValues();
  }
  for(var i=1;i<v.length;i++){ if(String(v[i][0]).trim()===key) return String(v[i][1]).trim(); }
  return '';
}

function out(obj, cb){
  var txt = JSON.stringify(obj);
  if(cb){ return ContentService.createTextOutput(cb+'('+txt+')').setMimeType(ContentService.MimeType.JAVASCRIPT); }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 라우팅 ───────── */
function doGet(e){
  var p = (e && e.parameter) || {};
  try{ return out(handle(p), p.callback); }
  catch(err){ return out({ok:false, error:String(err && err.message || err)}, p.callback); }
}
function doPost(e){
  var p = {};
  try{ p = JSON.parse(e.postData.contents || '{}'); }catch(err){ p = (e && e.parameter) || {}; }
  try{ return out(handle(p)); }
  catch(err){ return out({ok:false, error:String(err && err.message || err)}); }
}

function handle(p){
  var a = p.action;
  conf('PIN'); // 설정 탭 보장
  if(a==='ping')      return {ok:true, t:new Date().toISOString(), ai: !!conf('GEMINI_KEY')};
  if(a==='login')     return login(p);
  if(a==='save')      return save(p);
  if(a==='log')       return log(p);
  if(a==='feedback')  return feedback(p);
  if(a==='tlogin')    return tlogin(p);
  if(a==='roster')    return roster(p);
  if(a==='logs')      return logs(p);
  throw new Error('알 수 없는 요청: '+a);
}

/* ───────── 학생 ───────── */
function key(p){ return [String(p.cls||'').trim(), String(p.num||'').trim(), String(p.name||'').trim()].join('|'); }

function findRow(s, k){
  var v = s.getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===k) return i+1; }
  return -1;
}

function login(p){
  if(!p.cls || !p.num || !p.name) throw new Error('반·번호·이름을 모두 입력하세요');
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var k = key(p), r = findRow(s,k), state = null;
  if(r<0){
    s.appendRow([k, p.cls, p.num, p.name, 0, '훈련생', '', '', new Date()]);
  }else{
    var raw = s.getRange(r,8).getValue();
    try{ state = raw ? JSON.parse(raw) : null; }catch(e){ state=null; }
    s.getRange(r,9).setValue(new Date());
  }
  return {ok:true, state:state};
}

function save(p){
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var k = key(p), r = findRow(s,k);
  if(r<0){ s.appendRow([k, p.cls, p.num, p.name, 0, '', '', '', new Date()]); r = s.getLastRow(); }
  var st = String(p.state||'');
  if(st.length > 45000) st = st.slice(0,45000);
  s.getRange(r,5,1,5).setValues([[Number(p.pts||0), String(p.rank||''), String(p.prog||''), st, new Date()]]);
  return {ok:true};
}

function log(p){
  var s = sheet(SHEET_LOG, ['시각','반','번호','이름','종류','항목','내용','자동점검','AI수준','AI피드백']);
  s.appendRow([new Date(), p.cls||'', p.num||'', p.name||'', p.type||'', p.item||'', String(p.text||'').slice(0,5000), p.score||'', p.ai||'', String(p.aifb||'').slice(0,3000)]);
  return {ok:true};
}

/* ───────── Gemini 피드백 (키는 시트 '설정' 탭에만 있음) ───────── */
function feedback(p){
  var k = conf('GEMINI_KEY');
  if(!k) return {ok:false, error:'AI 키가 아직 설정되지 않았어요 (시트 설정 탭)'};
  var models = [conf('MODEL')||'gemini-2.5-flash','gemini-2.0-flash','gemini-1.5-flash'];
  var prompt = "너는 한국 중학교 1학년 과학 논술 코치다. 학생의 글을 아래 기준으로 평가해 JSON으로만 답하라.\n"+
    "채점 관점: "+(p.rubric||'')+"\n학생 글: \"\"\"\n"+(p.text||'')+"\n\"\"\"\n"+
    '출력 형식: {"level":"상|중|하","good":["잘한 점 1~2개"],"improve":["고칠 점 1~2개, 구체적으로"],"next":"학생이 바로 이어 쓸 수 있는 예시 문장 1개"} '+
    '친절하고 짧게, 존댓말로. 정답을 통째로 불러주지 말고 스스로 고치게 힌트를 줘라.';
  var lastErr = '';
  for(var i=0;i<models.length;i++){
    var m = models[i]; if(!m) continue;
    try{
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+m+':generateContent?key='+encodeURIComponent(k),{
        method:'post', contentType:'application/json', muteHttpExceptions:true,
        payload: JSON.stringify({contents:[{parts:[{text:prompt}]}], generationConfig:{responseMimeType:'application/json', temperature:0.4}})
      });
      var code = res.getResponseCode();
      if(code===404) continue;
      if(code!==200){ lastErr = 'HTTP '+code; if(code===429) continue; break; }
      var j = JSON.parse(res.getContentText());
      var txt = (((j.candidates||[])[0]||{}).content||{}).parts; txt = (txt||[]).map(function(x){return x.text||'';}).join('');
      var fb = JSON.parse(txt.replace(/^```json\s*|```\s*$/g,''));
      // 기록
      try{ log({cls:p.cls,num:p.num,name:p.name,type:'AI피드백',item:p.item||'',text:p.text,score:p.score||'',ai:fb.level||'',aifb:JSON.stringify(fb)}); }catch(e){}
      return {ok:true, fb:fb, model:m};
    }catch(err){ lastErr = String(err); }
  }
  return {ok:false, error:'AI 응답 실패 '+lastErr};
}

/* ───────── 선생님 ───────── */
function checkPin(p){ var pin = conf('PIN'); if(!pin || String(p.pin)!==pin) throw new Error('교사 코드가 달라요'); }

function tlogin(p){ checkPin(p); return {ok:true, ai: !!conf('GEMINI_KEY')}; }

function roster(p){
  checkPin(p);
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var v = s.getDataRange().getValues(), rows = [];
  for(var i=1;i<v.length;i++){
    if(p.cls && String(v[i][1])!==String(p.cls)) continue;
    rows.push({cls:v[i][1], num:v[i][2], name:v[i][3], pts:v[i][4], rank:v[i][5], prog:v[i][6], state:v[i][7]||'', seen:v[i][8]?new Date(v[i][8]).toISOString():''});
  }
  return {ok:true, rows:rows};
}

function logs(p){
  checkPin(p);
  var s = sheet(SHEET_LOG, ['시각','반','번호','이름','종류','항목','내용','자동점검','AI수준','AI피드백']);
  var v = s.getDataRange().getValues(), rows = [];
  for(var i=v.length-1;i>=1 && rows.length<800;i--){
    if(p.cls && String(v[i][1])!==String(p.cls)) continue;
    if(p.num && String(v[i][2])!==String(p.num)) continue;
    rows.push({t:v[i][0]?new Date(v[i][0]).toISOString():'', cls:v[i][1], num:v[i][2], name:v[i][3], type:v[i][4], item:v[i][5], text:v[i][6], score:v[i][7], ai:v[i][8], aifb:v[i][9]});
  }
  return {ok:true, rows:rows};
}
