// 결과 시트는 공유하지 않는다. 이 코드를 Google Apps Script의 Code.gs에 붙여 넣고 웹앱으로 배포한다.
function setupResultsSheet(){
  const book=SpreadsheetApp.getActiveSpreadsheet();
  if(!book)throw Error('결과 시트의 확장 프로그램 메뉴에서 Apps Script를 열어 실행해 주세요.');
  PropertiesService.getScriptProperties().setProperty('RESULTS_SHEET_ID',book.getId());
}
function resultsSheetId(){
  const id=PropertiesService.getScriptProperties().getProperty('RESULTS_SHEET_ID');
  if(!id)throw Error('선생님의 결과 수집 연결이 완료되지 않았습니다.');
  return id;
}
const QUIZ_VERSION='cns-20261021-3';
const QUIZ_ORIGIN='https://goodi0407-commits.github.io';
const QUESTIONS=[{"type":"OX","level":"하","text":"중추신경계통은 뇌와 척수로 구성된다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"뇌와 척수는 감각 정보를 통합하고 반응을 조절하는 중추신경계통을 이룬다.","slide":2},{"type":"OX","level":"하","text":"좌우 대뇌반구를 연결하여 정보를 주고받게 하는 구조는 뇌량이다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"뇌량은 좌우 대뇌반구를 연결한다.","slide":4},{"type":"OX","level":"하","text":"후두엽은 청각 정보를 주로 처리한다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"후두엽은 시각 정보를 처리한다. 청각 정보 처리는 측두엽의 기능이다.","slide":5},{"type":"OX","level":"하","text":"소뇌는 평형과 자세 유지 및 운동의 힘·방향·시기 조절에 관여한다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"소뇌는 균형을 유지하고 운동이 정확하고 부드럽게 이루어지도록 조절한다.","slide":8},{"type":"OX","level":"중","text":"척수의 횡단면에서 회백질은 바깥쪽에 있고 백질은 안쪽에 있다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"척수는 안쪽에 나비 모양의 회백질이 있고, 바깥쪽을 백질이 둘러싼다.","slide":10},{"type":"OX","level":"중","text":"시상하부는 자율신경 및 내분비 기능을 조절하여 항상성 유지에 관여한다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"시상하부는 체온·갈증·식욕·수면과 각성 리듬 등에 관여하고 뇌하수체의 호르몬 분비를 조절한다.","slide":6},{"type":"OX","level":"중","text":"척수의 전근으로 감각섬유가 들어오고 후근으로 운동섬유가 나간다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"후근으로 감각섬유가 들어오고 전근으로 운동섬유가 나간다.","slide":10},{"type":"OX","level":"상","text":"후각 정보는 일차후각피질에 도달하기 전에 반드시 시상을 거친다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"후각은 일차피질에 도달하기 전에 시상을 거치지 않는 예외이다. 시상은 대부분의 감각 정보를 대뇌피질로 중계한다.","slide":6},{"type":"OX","level":"상","text":"척수반사는 의식적 판단을 기다리지 않고 일어나므로 뇌의 조절을 전혀 받지 않는다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"척수반사는 척수 안의 회로에서 신속히 일어나지만 뇌의 조절도 받는다. 의식적 판단이 필요 없다는 것과 뇌의 조절을 받지 않는다는 것은 다르다.","slide":11},{"type":"OX","level":"상","text":"대뇌와 소뇌에서는 표면의 피질뿐 아니라 깊은 곳의 신경핵에도 회백질이 존재한다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"대뇌와 소뇌의 표면은 회백질인 피질이며, 깊은 곳에도 회백질 덩어리인 신경핵이 있다.","slide":3},{"type":"4지선다","level":"하","text":"다음 중 뇌간을 이루는 세 부위가 바르게 묶인 것은?","options":["대뇌·간뇌·소뇌","중뇌·교뇌·연수","시상·시상하부·뇌하수체","전두엽·두정엽·후두엽"],"answer":1,"ex":"뇌간은 중뇌, 교뇌, 연수로 이루어진다.","slide":3},{"type":"4지선다","level":"하","text":"손에 닿은 물체의 촉감과 팔의 위치 감각을 처리하는 중심후회가 있는 엽은?","options":["전두엽","측두엽","두정엽","후두엽"],"answer":2,"ex":"두정엽의 중심후회는 피부감각과 심부감각을 처리한다.","slide":5},{"type":"4지선다","level":"하","text":"뇌척수막을 바깥쪽에서 안쪽으로 바르게 나열한 것은?","options":["경막·지주막·연막","연막·경막·지주막","지주막·연막·경막","경막·연막·지주막"],"answer":0,"ex":"경막이 가장 바깥쪽, 지주막이 중간, 연막이 가장 안쪽에서 뇌와 척수 표면에 밀착한다.","slide":9},{"type":"4지선다","level":"중","text":"뇌간에서 간뇌와 교뇌 사이에 위치하며 동공반사에 관여하는 부위는?","options":["연수","소뇌","시상하부","중뇌"],"answer":3,"ex":"중뇌는 간뇌와 교뇌 사이에 있으며 안구 운동, 동공반사, 시각·청각 자극에 대한 반사에 관여한다.","slide":7},{"type":"4지선다","level":"중","text":"발바닥에 닿은 자갈의 압박 정보를 척수에서 뇌로 전달하는 경로는?","options":["하행전도로","상행전도로","뇌량","전근"],"answer":1,"ex":"상행전도로는 말초의 감각 정보를 척수에서 뇌로 전달한다. 하행전도로는 뇌의 운동 명령을 전달한다.","slide":11},{"type":"4지선다","level":"중","text":"대뇌 깊은 곳의 회백질 덩어리로 운동의 선택과 근긴장 조절에 관여하는 구조는?","options":["뇌량","맥락총","대뇌기저핵","지주막"],"answer":2,"ex":"대뇌기저핵은 대뇌 깊은 곳의 회백질 덩어리이며 운동 선택과 조절 및 근긴장 조절에 관여한다.","slide":4},{"type":"4지선다","level":"중","text":"성인의 척수 끝 높이와 마미에 관한 설명으로 옳은 것은?","options":["척수는 대개 제1~2요추 높이에서 끝나고 그 아래의 신경근 다발이 마미이다.","척수는 대개 천골 끝까지 이어지며 마미는 뇌의 일부이다.","척수는 경추에서 끝나고 마미는 뇌척수막이다.","마미는 척수원추 자체의 다른 이름이다."],"answer":0,"ex":"성인 척수는 대개 제1~2요추 높이에서 끝난다. 마미는 척수원추 아래로 내려가는 척수신경근 다발로 척수 자체와 구분한다.","slide":10},{"type":"4지선다","level":"상","text":"연수 아래쪽의 추체교차에 근거한 설명으로 가장 적절한 것은?","options":["대뇌 운동영역은 오직 같은 쪽 몸만 조절한다.","모든 감각신경이 소뇌 표면에서 교차한다.","교차가 있으므로 척수는 운동 명령을 전달하지 않는다.","추체로 섬유 대부분이 교차하여 대뇌 운동영역은 주로 반대쪽 몸의 운동을 조절한다."],"answer":3,"ex":"연수 아래쪽에서 추체로 섬유의 대부분이 교차한다. 따라서 대뇌 운동영역은 주로 반대쪽 몸의 운동을 조절한다. 모든 운동 경로가 예외 없이 교차한다는 뜻은 아니다.","slide":7},{"type":"4지선다","level":"상","text":"반사궁에서 자극이 전달되는 순서로 옳은 것은?","options":["감각수용기·원심성신경·효과기·반사중추·구심성신경","감각수용기·구심성신경·반사중추·원심성신경·효과기","효과기·구심성신경·반사중추·감각수용기·원심성신경","반사중추·감각수용기·원심성신경·구심성신경·효과기"],"answer":1,"ex":"수용기가 자극을 감지하면 구심성신경이 중추로 전달하고, 반사중추를 거친 뒤 원심성신경이 효과기로 전달한다.","slide":12},{"type":"4지선다","level":"상","text":"슬개건을 두드린 뒤 무릎이 펴지는 과정으로 옳은 것은?","options":["피부 수용기 감지·대퇴사두근 이완·무릎 폄","근방추 감지·대퇴사두근 이완·무릎 굽힘","늘어난 대퇴사두근의 근방추 감지·척수 회로·대퇴사두근 수축·무릎 폄","대뇌의 의식적 판단 완료·대퇴사두근 이완·무릎 폄"],"answer":2,"ex":"슬개건을 두드리면 대퇴사두근이 늘어나고 근방추가 이를 감지한다. 척수 반사 회로를 통해 대퇴사두근이 수축하여 무릎이 펴진다.","slide":12},{"type":"단답형","level":"하","text":"시각 정보를 처리하는 대뇌의 엽 이름을 쓰시오.","answers":["후두엽"],"ex":"후두엽은 눈으로 들어온 글자와 물체의 형태 등 시각 정보를 처리한다.","slide":5},{"type":"단답형","level":"하","text":"척수 안쪽의 나비 모양 영역을 이루며 신경세포체와 시냅스가 많은 물질은?","answers":["회백질"],"ex":"회백질에는 신경세포체와 시냅스가 많다. 척수에서는 안쪽에 분포한다.","slide":[3,10]},{"type":"단답형","level":"하","text":"뇌와 척수 표면에 밀착하는 가장 안쪽 뇌척수막의 이름은?","answers":["연막"],"ex":"뇌척수막은 바깥부터 경막, 지주막, 연막 순이다. 연막은 뇌와 척수 표면에 밀착한다.","slide":9},{"type":"단답형","level":"중","text":"체온·갈증·식욕을 조절하고 뇌하수체와 연결되어 항상성 유지에 관여하는 간뇌의 부위는?","answers":["시상하부"],"ex":"시상하부는 자율신경과 내분비 기능을 조절하며 체온, 갈증, 식욕 등 항상성 유지에 관여한다.","slide":6},{"type":"단답형","level":"중","text":"뇌 안에 있는 공간으로, 뇌척수액이 차 있으며 서로 연결되어 있는 공간을 무엇이라고 하는가?","answers":["뇌실"],"ex":"뇌실은 뇌 내부에 있는 공간으로 뇌척수액이 차 있다. 뇌척수액은 뇌실의 맥락총에서 대부분 생성된다.","slide":9},{"type":"단답형","level":"중","text":"뜨거운 물체에 닿은 손을 통증 자극에서 피하도록 재빨리 굽히는 반사의 이름은?","answers":["굴곡반사","회피반사","철회반사","도피반사"],"ex":"통증 자극을 받으면 사지를 굽혀 자극에서 피하는 굴곡반사가 일어난다. 회피반사, 철회반사, 도피반사도 같은 반응을 가리키는 용어로 인정한다.","slide":12},{"type":"단답형","level":"상","text":"외부 환경이 변하더라도 체온이나 체액의 농도 등 몸의 내부 상태를 비교적 일정하게 유지하려는 성질을 무엇이라고 하는가?","answers":["항상성","생체항상성"],"ex":"항상성은 몸의 내부 환경을 비교적 일정하게 유지하는 성질이다. 시상하부는 체온, 갈증, 식욕 등을 조절하여 항상성 유지에 관여한다.","slide":6},{"type":"단답형","level":"상","text":"뇌의 운동 명령이 척수로 내려온 뒤, 운동섬유가 척수에서 나갈 때 통과하는 신경근의 이름을 쓰시오.","answers":["전근","앞뿌리","운동근"],"ex":"운동섬유는 전근으로 나간다. 감각섬유가 들어오는 후근과 구별해야 한다.","slide":10},{"type":"단답형","level":"상","text":"전두엽에 있는 말 산출 관련 영역으로, 주로 좌측 하전두회에 위치하는 영역의 이름은?","answers":["브로카영역","브로카언어영역","브로카언어중추","브로카중추","브로카"],"ex":"브로카 영역은 말 산출에 관여한다. 전두엽의 중심전회에 있는 일차운동영역과 동일한 부위가 아니며, 언어 기능에는 여러 영역의 연결망이 관여한다.","slide":5},{"type":"단답형","level":"상","text":"뇌간에서 교뇌 아래에 위치하고 척수로 이어지며, 호흡·심장박동·혈관운동과 연하·기침·구토 반사를 조절하는 부위는 무엇인가?","answers":["연수","숨뇌"],"ex":"연수는 교뇌 아래에서 척수로 이어지는 뇌간의 일부이다. 호흡과 심장박동 등 생명 유지에 중요한 기능과 연하·기침·구토 반사를 조절한다.","slide":7}];
function normalizeAnswer(value){return String(value).normalize('NFKC').toLowerCase().replace(/[\s.,。·ㆍ]/g,'');}
function safeCell(value){const text=String(value);return /^[\s]*[=+@-]/.test(text)?"'"+text:text;}
function validateSubmission(p){
  if(!p||p.version!==QUIZ_VERSION)throw Error('문항 버전이 다릅니다. 페이지를 새로고침해 주세요.');
  if(typeof p.attemptId!=='string'||!/^cns-[a-f0-9]{32}$/.test(p.attemptId))throw Error('제출 번호가 올바르지 않습니다.');
  if(typeof p.name!=='string'||!p.name.trim()||p.name.trim().length>60)throw Error('이름을 확인해 주세요.');
  if(!Array.isArray(p.answers)||p.answers.length!==30)throw Error('30문항을 모두 풀어 주세요.');
  let score=0;const wrong=[];
  QUESTIONS.forEach((q,i)=>{
    const a=p.answers[i];
    if(typeof a!=='string'||!a.trim()||a.length>200)throw Error('답안을 확인해 주세요.');
    let right,student,correct;
    if(q.type==='단답형'){right=q.answers.some(x=>normalizeAnswer(x)===normalizeAnswer(a));student=a;correct=q.answers[0];}
    else{if(!/^[0-3]$/.test(a)||Number(a)>=q.options.length)throw Error('선택 답안을 확인해 주세요.');right=Number(a)===q.answer;student=q.options[Number(a)];correct=q.options[q.answer];}
    if(right)score++;else wrong.push({number:i+1,question:q.text,student,correct});
  });
  return {name:p.name.trim(),attemptId:p.attemptId,version:p.version,score,wrong};
}
function saveSubmission(p){
  const result=validateSubmission(p);
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const book=SpreadsheetApp.openById(resultsSheetId());
    const sheet=book.getSheetByName('제출결과');if(!sheet)throw Error('결과 시트가 없습니다.');
    const last=sheet.getLastRow();
    if(last>1){
      const found=sheet.getRange(2,7,last-1,1).createTextFinder(result.attemptId).matchEntireCell(true).findNext();
      if(found)return {ok:true,duplicate:true,attemptId:result.attemptId,score:Number(sheet.getRange(found.getRow(),3).getValue())};
    }
    const details=result.wrong.map(w=>w.number+'번: '+w.question+'\n학생 답: '+w.student+' / 정답: '+w.correct).join('\n\n')||'모두 정답';
    const timestamp=Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
    const row=[timestamp,safeCell(result.name),result.score,result.wrong.length,result.wrong.map(w=>w.number).join(', ')||'없음',safeCell(details),result.attemptId,result.version];
    sheet.getRange(last+1,1,1,row.length).setValues([row]);
    SpreadsheetApp.flush();
    return {ok:true,attemptId:result.attemptId,score:result.score};
  }finally{lock.releaseLock();}
}
function doGet(){return HtmlService.createHtmlOutput('중추신경퀴즈 제출 수신 서비스입니다. 학생 결과는 공개하지 않습니다.');}
function doPost(e){
  let message={type:'cns-quiz-submission',ok:false,attemptId:''};
  try{
    const raw=e&&e.parameter&&e.parameter.payload;
    if(typeof raw!=='string'||raw.length>64000)throw Error('요청 크기를 확인해 주세요.');
    const p=JSON.parse(raw);message.attemptId=typeof p.attemptId==='string'?p.attemptId:'';
    Object.assign(message,saveSubmission(p));
  }catch(error){message.error=String(error.message||'저장하지 못했습니다. 다시 제출해 주세요.').slice(0,160);}
  const data=JSON.stringify(message).replace(/</g,'\\u003c');
  const html='<!doctype html><meta charset="utf-8"><p>제출 처리 결과를 퀴즈에 전달합니다.</p><script>window.top.postMessage('+data+','+JSON.stringify(QUIZ_ORIGIN)+');<\\/script>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
