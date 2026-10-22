// 별도 Apps Script 프로젝트용이다. 프로젝트 설정의 스크립트 속성 RESULTS_SHEET_ID에 교사의 비공개 결과 시트 ID를 설정한다.
// 결과 시트 ID나 학생 답안은 공개 저장소에 넣지 않는다.
function setupResultsSheet(){
  const book=SpreadsheetApp.openById(resultsSheetId());
  if(!book)throw Error('프로젝트 설정의 RESULTS_SHEET_ID 속성을 확인해 주세요.');
  PropertiesService.getScriptProperties().setProperty('RESULTS_SHEET_ID',book.getId());
}
function resultsSheetId(){
  const id=PropertiesService.getScriptProperties().getProperty('RESULTS_SHEET_ID');
  if(!id)throw Error('선생님의 결과 수집 연결이 완료되지 않았습니다.');
  return id;
}
const QUIZ_VERSION='pns-20261022-1';
const QUIZ_ORIGIN='https://goodi0407-commits.github.io';
const QUESTIONS=[{"type":"OX","level":"하","text":"뇌신경은 12쌍이고 척수신경은 31쌍이다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"뇌신경은 12쌍이다. 척수신경은 경신경 8쌍, 흉신경 12쌍, 요신경 5쌍, 천골신경 5쌍, 미골신경 1쌍으로 총 31쌍이다.","slide":[1,10]},{"type":"OX","level":"하","text":"전정와우신경은 청각과 평형감각을 전달하는 운동신경이다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"전정와우신경은 감각신경이다. 와우신경은 청각 정보를, 전정신경은 평형 유지에 필요한 정보를 전달한다.","slide":5},{"type":"OX","level":"중","text":"부신경은 흉쇄유돌근과 승모근을 지배하여 고개를 돌리거나 어깨를 올리는 움직임에 관여한다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"부신경은 흉쇄유돌근과 승모근을 지배하는 운동신경이다.","slide":9},{"type":"OX","level":"중","text":"모든 척수신경은 예외 없이 추간공을 통해 척추 밖으로 나온다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"제1경신경은 후두골과 제1경추 사이로 나온다. 천골 부위에서는 신경가지가 천골공 등을 통해 나온다.","slide":11},{"type":"OX","level":"상","text":"자율신경섬유는 일부 뇌신경과 척수신경을 통해 주행하므로, 자율신경을 이들과 완전히 별개인 세 번째 종류의 신경으로 구분하는 것은 적절하지 않다.","options":["O: 맞다","X: 틀리다"],"answer":0,"ex":"뇌신경과 척수신경은 연결 부위에 따른 구분이다. 자율신경계통은 기능에 따른 구분이므로 같은 기준의 세 종류로 나열하면 오해가 생길 수 있다.","slide":3},{"type":"OX","level":"상","text":"자율신경계통의 조절은 말초에서만 이루어지므로 시상하부·뇌간·척수는 관여하지 않는다.","options":["O: 맞다","X: 틀리다"],"answer":1,"ex":"자율신경계통은 시상하부·뇌간·척수의 중추와 말초의 자율신경 및 신경절이 함께 작용한다.","slide":12},{"type":"4지선다","level":"하","text":"음식의 냄새를 맡을 때 후각 정보를 전달하는 신경은 무엇인가?","options":["후신경","시신경","안면신경","전정와우신경"],"answer":0,"ex":"후신경은 제1뇌신경이며 후각 정보를 전달하는 감각신경이다.","slide":5},{"type":"4지선다","level":"하","text":"외직근을 지배하여 안구를 바깥쪽으로 움직이는 신경은 무엇인가?","options":["동안신경","활차신경","외전신경","설하신경"],"answer":2,"ex":"외전신경은 제6뇌신경이며 외직근을 지배한다.","slide":6},{"type":"4지선다","level":"하","text":"다음 중 부교감신경의 대표적인 작용은 무엇인가?","options":["기관지 확장","심박수 감소","동공 확장","소화관 운동 억제"],"answer":1,"ex":"부교감신경은 심박수를 감소시키고 소화관 운동과 분비를 촉진한다.","slide":15},{"type":"4지선다","level":"중","text":"상사근을 지배하여 안구의 안쪽 회전과 아래쪽 움직임 등에 관여하는 신경은 무엇인가?","options":["동안신경","활차신경","외전신경","부신경"],"answer":1,"ex":"활차신경은 제4뇌신경이며 상사근을 지배한다. 상사근의 구체적인 작용은 안구의 위치에 따라 달라진다.","slide":6},{"type":"4지선다","level":"중","text":"설인신경의 기능을 바르게 설명한 것은 무엇인가?","options":["외직근을 지배하고 안구를 바깥쪽으로 움직인다.","표정근을 지배하고 혀 앞쪽의 미각을 전달한다.","경상인두근을 지배하고 혀 뒤쪽의 미각을 전달한다.","승모근을 지배하고 어깨를 올리는 움직임에 관여한다."],"answer":2,"ex":"설인신경은 경상인두근을 지배하여 삼킴을 돕고, 주로 혀 뒤쪽 3분의 1의 미각을 전달한다.","slide":8},{"type":"4지선다","level":"중","text":"제8경신경이 척추 밖으로 나오는 위치는 어디인가?","options":["후두골과 제1경추 사이","제1경추와 제2경추 사이","제6경추와 제7경추 사이","제7경추와 제1흉추 사이"],"answer":3,"ex":"경추는 7개이지만 경신경은 8쌍이다. 제8경신경은 제7경추와 제1흉추 사이로 나온다.","slide":10},{"type":"4지선다","level":"상","text":"혀 앞쪽 3분의 2에서 촉각과 온도감각은 유지되지만 미각 전달에 문제가 생겼다. 두 감각의 담당 신경을 고려할 때, 미각 전달과 가장 관련이 깊은 신경은 무엇인가?","options":["삼차신경","외전신경","안면신경","설하신경"],"answer":2,"ex":"혀 앞쪽 3분의 2의 일반감각은 삼차신경이, 미각은 주로 안면신경이 전달한다. 두 감각의 전달 경로를 구분해야 한다.","slide":7},{"type":"4지선다","level":"상","text":"척수신경의 신경근과 정보 전달 방향을 바르게 설명한 것은 무엇인가?","options":["감각 정보는 전근으로 들어오고, 운동 명령은 후근으로 나간다.","감각 정보와 운동 명령은 모두 후근을 통해 전달된다.","감각 정보와 운동 명령은 모두 전근을 통해 전달된다.","감각 정보는 후근으로 들어오고, 운동 명령은 전근으로 나간다."],"answer":3,"ex":"감각 정보는 후근을 통해 척수로 들어오고, 운동 명령은 전근을 통해 척수 밖으로 나간다.","slide":11},{"type":"단답형","level":"하","text":"망막에서 받아들인 시각 정보를 뇌로 전달하는 제2뇌신경의 이름을 쓰시오.","answers":["시신경"],"ex":"시신경은 망막에서 받아들인 시각 정보를 뇌로 전달한다.","slide":5},{"type":"단답형","level":"하","text":"혀 근육 대부분의 운동을 지배하여 혀를 내밀거나 움직이는 데 관여하는 제12뇌신경의 이름을 쓰시오.","answers":["설하신경"],"ex":"설하신경은 혀 근육 대부분의 운동을 지배한다.","slide":9},{"type":"단답형","level":"하","text":"척수신경은 감각섬유와 운동섬유를 함께 포함한다. 이처럼 감각과 운동 기능을 함께 가진 신경을 무엇이라고 하는가?","answers":["혼합신경"],"ex":"혼합신경은 감각 정보를 전달하는 섬유와 운동 명령을 전달하는 섬유를 함께 포함한다.","slide":11},{"type":"단답형","level":"하","text":"갑자기 놀라거나 긴장했을 때 심박수를 증가시키고 동공을 확장하는 자율신경은 무엇인가?","answers":["교감신경","교감신경계","교감신경계통"],"ex":"교감신경은 신체가 긴장 상황에 대처하도록 작용한다.","slide":13},{"type":"단답형","level":"상","text":"인두와 후두의 운동·감각에 관여하고, 심장·폐·소화기관 등 여러 내장기관의 감각 전달과 부교감 기능을 함께 수행하는 혼합신경의 이름을 쓰시오.","answers":["미주신경"],"ex":"미주신경은 제10뇌신경으로 감각·운동·부교감신경섬유를 포함한다.","slide":8},{"type":"단답형","level":"하","text":"얼굴의 표정근을 움직여 미소를 짓거나 눈을 감는 데 관여하는 제7뇌신경의 이름을 쓰시오.","answers":["안면신경"],"ex":"안면신경은 표정근의 운동을 담당하며, 혀 앞쪽의 미각 전달에도 관여한다.","slide":7}];
function normalizeAnswer(value){return String(value).normalize('NFKC').toLowerCase().replace(/[\s.,。·ㆍ]/g,'');}
function safeCell(value){const text=String(value);return /^[\s]*[=+@-]/.test(text)?"'"+text:text;}
function validateSubmission(p){
  if(!p||p.version!==QUIZ_VERSION)throw Error('문항 버전이 다릅니다. 페이지를 새로고침해 주세요.');
  if(typeof p.attemptId!=='string'||!/^pns-[a-f0-9]{32}$/.test(p.attemptId))throw Error('제출 번호가 올바르지 않습니다.');
  if(typeof p.name!=='string'||!p.name.trim()||p.name.trim().length>60)throw Error('이름을 확인해 주세요.');
  if(!Array.isArray(p.answers)||p.answers.length!==20)throw Error('20문항을 모두 풀어 주세요.');
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
    const sheet=book.getSheetByName('말초신경_제출결과');if(!sheet)throw Error('결과 시트가 없습니다.');
    if(sheet.getRange(1,1).getValue()!=='제출시각')throw Error('결과 시트의 머리글을 확인해 주세요.');
    const last=sheet.getLastRow();
    if(last>1){
      const found=sheet.getRange(2,7,last-1,1).createTextFinder(result.attemptId).matchEntireCell(true).findNext();
      if(found)return {ok:true,duplicate:true,attemptId:result.attemptId,score:Number(sheet.getRange(found.getRow(),3).getValue())};
    }
    const details=result.wrong.map(w=>w.number+'번: '+w.question+'\n학생 답: '+w.student+' / 정답: '+w.correct).join('\n\n')||'모두 정답';
    const timestamp=Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
    const row=[timestamp,safeCell(result.name),result.score,result.wrong.length,result.wrong.map(w=>w.number).join(', ')||'없음',safeCell(details),result.attemptId,result.version,safeCell(QUESTIONS.map((q,i)=>(i+1)+'번: '+(q.type==='단답형'?p.answers[i]:q.options[Number(p.answers[i])])).join('\n'))];
    const individual=QUESTIONS.map((q,i)=>safeCell(q.type==='단답형'?p.answers[i]:q.options[Number(p.answers[i])]));
    sheet.getRange(last+1,1,1,row.length+individual.length).setValues([row.concat(individual)]);
    SpreadsheetApp.flush();
    return {ok:true,attemptId:result.attemptId,score:result.score};
  }finally{lock.releaseLock();}
}
function doGet(){return ContentService.createTextOutput(JSON.stringify({service:'pns-quiz',version:QUIZ_VERSION,total:QUESTIONS.length,configured:Boolean(PropertiesService.getScriptProperties().getProperty('RESULTS_SHEET_ID'))})).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){
  let message={type:'pns-quiz-submission',ok:false,attemptId:''};
  try{
    const raw=e&&e.parameter&&e.parameter.payload;
    if(typeof raw!=='string'||raw.length>64000)throw Error('요청 크기를 확인해 주세요.');
    const p=JSON.parse(raw);message.attemptId=typeof p.attemptId==='string'?p.attemptId:'';
    Object.assign(message,saveSubmission(p));
  }catch(error){message.error=String(error.message||'저장하지 못했습니다. 다시 제출해 주세요.').slice(0,160);}
  const data=JSON.stringify(message).replace(/</g,'\\u003c');
  const html='<!doctype html><meta charset="utf-8"><p>제출 처리 결과를 퀴즈에 전달합니다.</p><script>window.top.postMessage('+data+','+JSON.stringify(QUIZ_ORIGIN)+');</script>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

