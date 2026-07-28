/************************************************************************
 *  디지털 추리 — 실시간 익명 투표 백엔드 (Google Apps Script)
 *  ────────────────────────────────────────────────────────────────
 *  하는 일
 *   1) 교사 휴대폰에 보여줄 "투표 페이지"를 서빙 (QR이 이 주소를 가리킴)
 *   2) 투표를 구글 시트에 기록 (이름 안 받음 = 익명)
 *   3) 강사 화면이 요청하면 집계 결과(JSONP)로 응답
 *   4) 강사 화면이 "지금 이 질문 열기/닫기"를 설정 → 폰이 자동으로 따라감
 *
 *  설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 *        → 배포 → 새 배포 → 웹 앱 → 액세스 "모든 사용자" → /exec URL 복사
 *  (자세한 절차는 '투표연동_설치가이드.md' 참고)
 ************************************************************************/

var SHEET_NAME = 'Votes';

function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function props_(){ return PropertiesService.getScriptProperties(); }
function sheet_(){
  var s = ss_().getSheetByName(SHEET_NAME);
  if(!s){ s = ss_().insertSheet(SHEET_NAME); s.appendRow(['timestamp','room','q','choice','token']); }
  return s;
}

/* ---------- HTTP 진입점 ---------- */
function doGet(e){
  var p = (e && e.parameter) || {}, cb = p.callback, out;
  if (p.action === 'tally')         { out = tally_(p.room, p.q); }
  else if (p.action === 'setActive'){ setActive_(p.room, p.q, p.opts, p.open); out = {ok:true}; }
  else if (p.action === 'getActive'){ out = getActiveState(p.room); }
  else if (p.action === 'vote')     { recordVote(p.room, p.q, p.choice, p.token); out = {ok:true}; }
  else { // 액션 없음 → 투표 페이지 서빙
    return HtmlService.createHtmlOutput(voterHtml_(p.room || 'NOSUNG'))
      .setTitle('디지털 추리 · 투표')
      .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var json = JSON.stringify(out);
  if (cb){ return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT); }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    var d = JSON.parse(e.postData.contents);
    recordVote(d.room, d.q, d.choice, d.token);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ---------- 클라이언트(투표 페이지)에서 google.script.run 으로 호출 ---------- */
function recordVote(room, q, choice, token){
  if(!room || !q || !choice) return {ok:false};
  sheet_().appendRow([new Date(), String(room), String(q), String(choice), String(token||'')]);
  return {ok:true};
}
function getActiveState(room){
  var v = props_().getProperty('active_' + room);
  return v ? JSON.parse(v) : {q:'', opts:'', open:false};
}

/* ---------- 내부 ---------- */
function setActive_(room, q, opts, open){
  props_().setProperty('active_' + room, JSON.stringify({
    q: q || '', opts: opts || '', open: (String(open) === '1'), ts: Date.now()
  }));
}
function tally_(room, q){
  var rows = sheet_().getDataRange().getValues();   // [0]=헤더
  var last = {};                                    // token → 최신 choice (중복 투표는 최신만)
  for (var i = 1; i < rows.length; i++){
    var r = rows[i];
    if (String(r[1]) === String(room) && String(r[2]) === String(q)){
      var tok = r[4] ? String(r[4]) : ('row' + i);
      last[tok] = String(r[3]);
    }
  }
  var counts = {}, total = 0;
  for (var t in last){ counts[last[t]] = (counts[last[t]] || 0) + 1; total++; }
  return {counts: counts, total: total};
}

/* ---------- 운영 도구 (Apps Script 편집기에서 직접 실행) ---------- */
function resetRoom(){                  // 특정 세션의 투표 전부 삭제 (재시연 전 호출)
  var ROOM = 'NOSUNG';                 // ← 필요시 변경
  var sh = sheet_(), rows = sh.getDataRange().getValues();
  for (var i = rows.length; i >= 2; i--){ if (String(rows[i-1][1]) === ROOM){ sh.deleteRow(i); } }
  props_().deleteProperty('active_' + ROOM);
}

/* ================== 투표 페이지 (교사 휴대폰) ================== */
function voterHtml_(room){
  var L = [];
  L.push('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">');
  L.push('<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">');
  L.push('<title>디지털 추리 · 투표</title><style>');
  L.push('*{margin:0;padding:0;box-sizing:border-box;-webkit-tap-highlight-color:transparent}');
  L.push('body{font-family:"Apple SD Gothic Neo","Malgun Gothic",system-ui,sans-serif;background:#0B0F16;color:#EAEEF3;min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:26px 18px 40px}');
  L.push('.case{font-family:ui-monospace,Consolas,monospace;color:#F2B544;letter-spacing:.2em;font-size:13px;margin-bottom:6px}');
  L.push('.brand{font-weight:800;font-size:20px;margin-bottom:22px;color:#EAEEF3}');
  L.push('.card{width:100%;max-width:460px;background:#161D28;border:1px solid #2A323D;border-radius:18px;padding:22px 20px;box-shadow:0 14px 40px rgba(0,0,0,.4)}');
  L.push('.q{font-size:21px;font-weight:800;line-height:1.35;margin-bottom:18px}');
  L.push('.opts{display:flex;flex-direction:column;gap:12px}');
  L.push('.opt{width:100%;text-align:left;font-size:18px;font-weight:700;color:#EAEEF3;background:#0E141C;border:2px solid #2A323D;border-radius:14px;padding:16px 18px;cursor:pointer;transition:.12s}');
  L.push('.opt:active{transform:scale(.985)}');
  L.push('.opt.sel{border-color:#F2B544;background:rgba(242,181,68,.12);color:#F2B544}');
  L.push('.opt .ck{float:right;color:#F2B544;font-weight:900;opacity:0}.opt.sel .ck{opacity:1}');
  L.push('.done{margin-top:16px;text-align:center;color:#43C6A0;font-weight:700;font-size:15px;min-height:20px}');
  L.push('.wait{text-align:center;padding:26px 10px;color:#8B97A4}');
  L.push('.wait .ic{font-size:34px;margin-bottom:12px}.wait .t{font-size:17px;font-weight:700;color:#C7D0DA;line-height:1.4}');
  L.push('.foot{margin-top:20px;color:#5A646F;font-size:12px;font-family:ui-monospace,Consolas,monospace;letter-spacing:.05em;text-align:center}');
  L.push('.dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:#43C6A0;margin-right:5px;vertical-align:middle;animation:bl 1.2s infinite}@keyframes bl{50%{opacity:.25}}');
  L.push('</style></head><body>');
  L.push('<div class="case">DIGITAL DETECTIVE · 투표</div>');
  L.push('<div class="brand">우리 반 그 아이</div>');
  L.push('<div class="card" id="card"><div class="wait"><div class="ic">🔎</div><div class="t">연결 중…</div></div></div>');
  L.push('<div class="foot"><span class="dot"></span>세션 ' + room + ' · 익명 투표 · 이름은 수집하지 않습니다</div>');

  L.push('<script>');
  L.push('var ROOM=' + JSON.stringify(String(room)) + ';');
  L.push('var token=localStorage.getItem("dd_tok"); if(!token){ token="t"+Date.now()+Math.random().toString(36).slice(2,8); localStorage.setItem("dd_tok",token); }');
  L.push('var curQ=null, busy=false;');
  L.push('function esc(s){return String(s).replace(/[&<>\"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}[c];});}');
  L.push('function paint(active){');
  L.push('  var card=document.getElementById("card");');
  L.push('  if(!active||!active.open||!active.q){ curQ=null; card.innerHTML=\'<div class="wait"><div class="ic">🔎</div><div class="t">다음 단서를 기다리는 중…<br>앞 화면을 봐주세요.</div></div>\'; return; }');
  L.push('  var opts=(active.opts||"").split("|").filter(Boolean);');
  L.push('  var mine=localStorage.getItem("dd_v_"+ROOM+"_"+active.q)||"";');
  L.push('  if(curQ!==active.q || card.dataset.q!==active.q){');
  L.push('    var h=\'<div class="q">\'+qText(active.q)+\'</div><div class="opts">\';');
  L.push('    opts.forEach(function(o){ h+=\'<button class="opt\'+(o===mine?" sel":"")+\'" data-o="\'+esc(o)+\'"><span class="ck">✓</span>\'+esc(o)+\'</button>\'; });');
  L.push('    h+=\'</div><div class="done" id="done">\'+(mine?"✔ 제출됨 · 다시 눌러 변경 가능":"")+\'</div>\';');
  L.push('    card.innerHTML=h; card.dataset.q=active.q; curQ=active.q;');
  L.push('    Array.prototype.forEach.call(card.querySelectorAll(".opt"),function(b){ b.addEventListener("click",function(){ submit(b.getAttribute("data-o")); }); });');
  L.push('  } }');
  L.push('function qText(q){ var m={"A1":"여기까지 — 위험 신호가 있습니까?","A2":"이 \\u2018상담 선생님\\u2019의 정체는?","C1":"준호의 상태, 한 단어로?","C2":"이제 무엇으로 보입니까?","C3":"스스로 보냈는데, 피해자일까요?","D1":"지금 누구를 보호하고 싶나요?","D2":"민수를 어떻게 해야 할까요?","D3":"민수는 가해자? 피해자? 둘 다?"}; return m[q]||"투표"; }');
  L.push('function submit(choice){ if(busy)return; busy=true; localStorage.setItem("dd_v_"+ROOM+"_"+curQ,choice);');
  L.push('  Array.prototype.forEach.call(document.querySelectorAll(".opt"),function(b){ b.classList.toggle("sel", b.getAttribute("data-o")===choice); });');
  L.push('  var dn=document.getElementById("done"); if(dn)dn.textContent="제출 중…";');
  L.push('  google.script.run.withSuccessHandler(function(){ busy=false; var d=document.getElementById("done"); if(d)d.textContent="✔ 제출됨 · 다시 눌러 변경 가능"; })');
  L.push('   .withFailureHandler(function(){ busy=false; var d=document.getElementById("done"); if(d)d.textContent="⚠ 전송 실패 · 다시 시도"; })');
  L.push('   .recordVote(ROOM,curQ,choice,token); }');
  L.push('function poll(){ google.script.run.withSuccessHandler(paint).getActiveState(ROOM); }');
  L.push('poll(); setInterval(poll,2000);');
  L.push('</' + 'script></body></html>');
  return L.join('\n');
}
