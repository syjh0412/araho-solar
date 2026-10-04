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
 *  5) AI 키와 센터장 코드 — 둘 중 편한 곳에 넣으면 됩니다 (둘 다 있으면 스크립트 속성이 우선)
 *     (가) 스크립트 속성: 왼쪽 ⚙ 프로젝트 설정 → 맨 아래 '스크립트 속성' → 속성 추가
 *          GEMINI_KEY = 발급받은 키 / PIN = 센터장 코드 (예: 0726)
 *     (나) 시트 '설정' 탭 (웹 앱을 한 번 열면 자동 생성): GEMINI_KEY, PIN 의 B칸
 *  6) (선택) 학생 명단: 아무 탭에나 '반 | 번호 | 이름' 머리글로 명단을 붙여넣으면
 *     학생은 반·번호를 고르고 자기 이름을 확인한 뒤 '콜사인'(비밀번호)으로 입장합니다.
 *     콜사인은 첫 입장 때 학생이 정하고 명단 탭의 '콜사인' 열에 저장됩니다.
 * 이후 학생 기록은 '학생'(현재 상태) / '기록'(전체) 탭과
 *   한줄논술①~④ · 시험 문항1~3 · 수배파일 탭(문항별)에 자동으로 쌓이고,
 * 앱의 센터장 탭에서 실시간으로 볼 수 있습니다.
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
  // 1순위: 스크립트 속성(⚙ 프로젝트 설정 → 스크립트 속성). 2순위: 시트 '설정' 탭.
  try{ var sp = PropertiesService.getScriptProperties().getProperty(key); if(sp && String(sp).trim()) return String(sp).trim(); }catch(e){}
  var s = sheet(SHEET_CONF, ['항목','값','설명']);
  var v = s.getDataRange().getValues();
  if(v.length < 2){
    s.appendRow(['GEMINI_KEY','', 'Google AI Studio(aistudio.google.com)에서 발급한 키를 B열에 붙여넣기']);
    s.appendRow(['PIN','0726','센터장 탭 입장 코드']);
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
  if(a==='names')     return rosterPublic(p);
  if(a==='login')     return login(p);
  if(a==='resetcode') return resetCode(p);
  if(a==='save')      return save(p);
  if(a==='log')       return log(p);
  if(a==='feedback')  return feedback(p);
  if(a==='qadd')      return qadd(p);
  if(a==='qlist')     return qlist(p);
  if(a==='qanswer')   return qanswer(p);
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

/* ───────── 명단(학생 명부) ─────────
   시트 아무 탭이든 첫 3줄 안에 '반' '번호' '이름' 머리글이 있으면 명단으로 인식합니다.
   '콜사인' 열이 없으면 자동으로 만들어 학생이 처음 입장할 때 정한 콜사인(비밀번호)을 저장합니다. */
function findRoster(){
  var sheets = ss().getSheets();
  for(var i=0;i<sheets.length;i++){
    var sh = sheets[i]; if(sh.getLastRow()<1) continue;
    var top = sh.getRange(1,1,Math.min(3,sh.getLastRow()),Math.max(1,sh.getLastColumn())).getValues();
    for(var r=0;r<top.length;r++){
      var row = top[r].map(function(x){return String(x).replace(/\s/g,'');});
      var ci=row.indexOf('반'), ni=row.indexOf('번호'), mi=row.indexOf('이름');
      if(ci>=0 && ni>=0 && mi>=0){
        var ki = row.indexOf('콜사인'); if(ki<0) ki = row.indexOf('비밀번호');
        if(ki<0){ ki = row.length; sh.getRange(r+1, ki+1).setValue('콜사인'); }
        return {sh:sh, head:r+1, ci:ci, ni:ni, mi:mi, ki:ki};
      }
    }
  }
  return null;
}
function rosterPublic(p){
  var R = findRoster(); if(!R) return {ok:true, roster:null};
  var v = R.sh.getRange(R.head+1,1,Math.max(0,R.sh.getLastRow()-R.head),Math.max(R.ki+1,R.sh.getLastColumn())).getValues();
  var out = {};
  for(var i=0;i<v.length;i++){ var c=String(v[i][R.ci]).trim(), n=String(v[i][R.ni]).trim(), m=String(v[i][R.mi]).trim();
    if(!c||!n||!m) continue; (out[c]=out[c]||[]).push({num:n,name:m,set:!!String(v[i][R.ki]||'').trim()}); }
  return {ok:true, roster:out};
}
function rosterRow(R, cls, num){
  var v = R.sh.getRange(R.head+1,1,Math.max(0,R.sh.getLastRow()-R.head),Math.max(R.ki+1,R.sh.getLastColumn())).getValues();
  for(var i=0;i<v.length;i++){ if(String(v[i][R.ci]).trim()===String(cls).trim() && String(v[i][R.ni]).trim()===String(num).trim()) return {row:R.head+1+i, name:String(v[i][R.mi]).trim(), code:String(v[i][R.ki]||'').trim()}; }
  return null;
}
function login(p){
  if(!p.cls || !p.num) throw new Error('반과 번호를 입력하세요');
  var R = findRoster(), name = String(p.name||'').trim();
  if(R){
    var row = rosterRow(R, p.cls, p.num);
    if(!row) throw new Error('명단에 없는 반·번호예요. 센터장에게 확인하세요');
    name = row.name;
    var code = String(p.code||'').trim();
    if(!row.code){
      if(code.length<2) throw new Error('처음 입장이에요. 콜사인(2~8자)을 정해 주세요');
      R.sh.getRange(row.row, R.ki+1).setValue(code);
    }else if(code!==row.code){ throw new Error('콜사인이 달라요'); }
  }else if(!name){ throw new Error('이름을 입력하세요'); }
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var k = key({cls:p.cls,num:p.num,name:name}), r = findRow(s,k), state = null;
  if(r<0){ s.appendRow([k, p.cls, p.num, name, 0, '지상 훈련생', '', '', new Date()]); }
  else{ var raw = s.getRange(r,8).getValue(); try{ state = raw ? JSON.parse(raw) : null; }catch(e){ state=null; } s.getRange(r,9).setValue(new Date()); }
  return {ok:true, name:name, state:state};
}
function resetCode(p){
  checkPin(p); var R = findRoster(); if(!R) throw new Error('명단 탭이 없어요');
  var row = rosterRow(R, p.cls, p.num); if(!row) throw new Error('명단에 없어요');
  R.sh.getRange(row.row, R.ki+1).setValue(''); return {ok:true};
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

/* 기록은 전체 '기록' 탭 + 문항별 탭에 나눠 쌓입니다 */
var ITEM_SHEETS = {w_c1:'한줄논술① 태양계 구성원', w_c2:'한줄논술② 행성 분류', w_c3:'한줄논술③ 태양 활동', w_c4:'한줄논술④ 지구 영향'};
function itemSheetName(p){
  var it = String(p.item||''), ty = String(p.type||'');
  for(var k in ITEM_SHEETS){ if(it.indexOf(k)>=0) return ITEM_SHEETS[k]; }
  var m = it.match(/문항\s*([123])/); if(m) return '시험 문항'+m[1];
  if(ty==='수배파일') return '수배파일';
  return '';
}
function log(p){
  var row = [new Date(), p.cls||'', p.num||'', p.name||'', p.type||'', p.item||'', String(p.text||'').slice(0,5000), p.score||'', p.ai||'', String(p.aifb||'').slice(0,3000)];
  var head = ['시각','반','번호','이름','종류','항목','내용','자동점검','AI수준','AI피드백'];
  sheet(SHEET_LOG, head).appendRow(row);
  var name = itemSheetName(p);
  if(name){ var s2 = sheet(name, head); s2.appendRow(row); }
  return {ok:true};
}

/* ───────── 교신(질문) 게시판 ───────── */
var SHEET_Q = '교신', SHEET_A = '응답';
function qid(){ return Utilities.getUuid().slice(0,8); }
function qadd(p){
  if(!p.text || String(p.text).trim().length<5) throw new Error('질문을 조금 더 자세히 적어 주세요');
  var s = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var id = qid();
  s.appendRow([id, new Date(), p.cls||'', p.num||'', p.name||'', p.mission||'', p.phase||'', String(p.text).slice(0,1000), p.anon==='1'?'1':'', 0]);
  return {ok:true, id:id};
}
function qanswer(p){
  if(!p.qid || !p.text || String(p.text).trim().length<2) throw new Error('응답 내용을 적어 주세요');
  var s = sheet(SHEET_A, ['id','qid','시각','반','번호','이름','내용','센터장']);
  s.appendRow([qid(), p.qid, new Date(), p.cls||'', p.num||'', p.name||'', String(p.text).slice(0,1000), p.teacher==='1'?'1':'']);
  // 응답수 갱신
  var q = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var v = q.getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===String(p.qid)){ q.getRange(i+1,10).setValue(Number(v[i][9]||0)+1); break; } }
  return {ok:true};
}
function qlist(p){
  var q = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var a = sheet(SHEET_A, ['id','qid','시각','반','번호','이름','내용','센터장']);
  var qv = q.getDataRange().getValues(), av = a.getDataRange().getValues();
  var ans = {};
  for(var j=1;j<av.length;j++){ var r=av[j]; (ans[String(r[1])]=ans[String(r[1])]||[]).push({t:r[2]?new Date(r[2]).toISOString():'',cls:r[3],num:r[4],name:r[5],text:r[6],teacher:String(r[7])==='1'}); }
  var rows=[];
  for(var i=qv.length-1;i>=1 && rows.length<300;i--){ var x=qv[i];
    if(p.cls && String(x[2])!==String(p.cls)) continue;
    if(p.mission && String(x[5])!==String(p.mission)) continue;
    var anon = String(x[8])==='1';
    rows.push({id:x[0], t:x[1]?new Date(x[1]).toISOString():'', cls:x[2], num:anon?'':x[3], name:anon?'익명 승무원':x[4], mine:(String(x[2])===String(p.cls)&&String(x[3])===String(p.num)), mission:x[5], phase:x[6], text:x[7], answers:ans[String(x[0])]||[]});
  }
  return {ok:true, rows:rows};
}

/* ───────── Gemini 피드백 (키는 시트 '설정' 탭에만 있음) ───────── */
var COACH_RULES =
 "너는 한국 중학교 1학년 과학 수업의 논술 코치 '아라'다. 학생(13세)의 글을 아래 '채점 기준'에 비추어 피드백한다.\n"+
 "규칙:\n"+
 "1) 반드시 중1 교과서 범위 안의 말만 쓴다. 금지: 전리층, 자기권, 코로나 질량 방출, 플라즈마, 태양 복사압, 밀도 공식 변형, 고등학교 수준 용어.\n"+
 "2) 채점 기준에 적힌 '평가 요소'를 하나씩 확인해 있는 것/빠진 것을 가려낸다. 과학적 오개념이 있으면 가장 먼저 짚는다.\n"+
 "3) 정답 문장을 통째로 써 주지 않는다. 빠진 요소는 '자료의 어느 부분을 보면 되는지' 힌트로 준다.\n"+
 "4) 글의 짜임을 본다: 주장(판단·분류) → 자료 속 근거(왜냐하면) → 연결·결론(따라서). 근거가 자료에서 나왔는지 확인한다.\n"+
 "5) 말투는 따뜻하고 짧은 존댓말. 칭찬은 구체적으로 1~2개, 고칠 점은 가장 중요한 것부터 최대 2개.\n"+
 "6) 예상 수준은 채점 기준의 수행 수준(상·중·하)으로 판단하되, 점수 숫자는 말하지 않는다.\n";
function feedback(p){
  var k = conf('GEMINI_KEY');
  if(!k) return {ok:false, error:'AI 키가 아직 설정되지 않았어요 (시트 설정 탭)'};
  var models = [conf('MODEL')||'gemini-2.5-flash','gemini-2.0-flash','gemini-1.5-flash'];
  var prompt = COACH_RULES+
    "\n[채점 기준]\n"+(p.rubric||'')+"\n\n[학생 글]\n\"\"\"\n"+(p.text||'')+"\n\"\"\"\n\n"+
    '다음 JSON 형식으로만 답하라: {"level":"상|중|하","good":["구체적 칭찬 1~2개"],"improve":["가장 중요한 고칠 점부터 최대 2개. 빠진 평가 요소가 있으면 자료의 어디를 보라고 힌트"],"next":"학생이 바로 이어 쓸 수 있는 시작 문장 (빈칸 ___ 포함, 정답은 넣지 않음)"}';
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

/* ───────── 센터장 ───────── */
function checkPin(p){ var pin = conf('PIN'); if(!pin || String(p.pin)!==pin) throw new Error('센터장 코드가 달라요'); }

function tlogin(p){ checkPin(p); return {ok:true, ai: !!conf('GEMINI_KEY')}; }

function roster(p){
  checkPin(p);
  var R=findRoster(), codes={}; if(R){ var rv=R.sh.getRange(R.head+1,1,Math.max(0,R.sh.getLastRow()-R.head),Math.max(R.ki+1,R.sh.getLastColumn())).getValues(); rv.forEach(function(x){ codes[String(x[R.ci]).trim()+'|'+String(x[R.ni]).trim()]=String(x[R.ki]||'').trim(); }); }
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var v = s.getDataRange().getValues(), rows = [];
  for(var i=1;i<v.length;i++){
    if(p.cls && String(v[i][1])!==String(p.cls)) continue;
    rows.push({cls:v[i][1], num:v[i][2], name:v[i][3], pts:v[i][4], rank:v[i][5], prog:v[i][6], state:v[i][7]||'', seen:v[i][8]?new Date(v[i][8]).toISOString():'', code:codes[String(v[i][1]).trim()+'|'+String(v[i][2]).trim()]||''});
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
