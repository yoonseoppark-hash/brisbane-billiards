/** Brisbane Billiards live Google Sheets sync v8. Keep Script Property SYNC_TOKEN. */
const SHEET_ID='1vZNRkDMLCf5Mm2zpU0onKu5c_t_LBSPslSemI2CBtd4';
const TZ='Australia/Brisbane', FIRST=4, LAST=303, FINE_FIRST=17, FINE_LAST=1002, AVG_FIRST=4, AVG_LAST=999;
// Known preview-only records already present in the staging log; never import them into live sheets.
const BLOCKED_TEST_IDS=['6bb8b0a2-a56f-4214-9121-a549df73a1aa','bd3e1120-d7d6-4026-a93f-dc82e1a5a414','357cbff3-60fa-4bf1-ad75-201d1d504aeb'];
function doGet(){return reply({ok:true,message:'Brisbane Billiards Connected',version:8,handicapRead:true,mode:PropertiesService.getScriptProperties().getProperty('LIVE_ENABLED')==='yes'?'live-write':'locked'});}
function doPost(e){
 const lock=LockService.getScriptLock();let log=null,logRow=0;
 try{
  if(!e||!e.postData||!e.postData.contents)throw Error('Missing request body');
  const d=JSON.parse(e.postData.contents),token=PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if(!token||d.token!==token)return reply({ok:false,message:'Invalid token'});
  if(d.action==='handicaps')return reply({ok:true,action:'handicaps',version:8,handicaps:readConfirmedHandicaps()});
  if(d.action==='test')return reply({ok:true,message:'Authentication successful',version:8});
  if(BLOCKED_TEST_IDS.includes(d.matchId))return reply({ok:true,blocked:true,matchId:d.matchId,version:8,message:'이전 미리보기 테스트 경기는 실제 시트에 반영하지 않았습니다.'});
  if(d.action!=='save'||typeof d.matchId!=='string'||!/^[a-zA-Z0-9_-]{16,100}$/.test(d.matchId))throw Error('Invalid request');
  const payload=JSON.stringify(d.match);if(!payload||payload.length>20000)throw Error('Invalid or oversized match payload');
  const plan=convertMatch(d.match,d.matchId);if(PropertiesService.getScriptProperties().getProperty('LIVE_ENABLED')!=='yes')return reply({ok:false,message:'Live saving is locked'});lock.waitLock(20000);
  const ss=SpreadsheetApp.openById(SHEET_ID);log=ss.getSheetByName('앱전송기록');if(!log)throw Error('앱전송기록 sheet not found');ensureLogHeaders(log);
  logRow=findLogRow(log,d.matchId);
  if(logRow&&log.getRange(logRow,4).getValue()==='저장완료')return reply({ok:true,duplicate:true,matchId:d.matchId,version:8});
  if(!logRow){logRow=log.getLastRow()+1;log.getRange(logRow,1,1,6).setValues([[d.matchId,new Date(),payload,'저장중','','']]);}
  else{log.getRange(logRow,2,1,3).setValues([[new Date(),payload,'저장중']]);log.getRange(logRow,5,1,2).clearContent();}
  SpreadsheetApp.flush();
  ensureTrackingColumns(ss);ensureAutoFineFormula(ss);
  const main=getOrCreateMainMatch(ss,plan,d.matchId);
  const fineCount=appendFineRows(ss,plan,d.matchId);
  const avgRow="기존 자동 연동";
  SpreadsheetApp.flush();
  const loc=(plan.official?'리그경기!':'친선경기!')+main.row;
  const summary=`저장완료 · ${loc} · 에버리지 반영 · 벌금 ${fineCount}건`;
  log.getRange(logRow,4,1,3).setValues([['저장완료',summary,`${loc}; order=${main.order}; avgRow=${avgRow}`]]);SpreadsheetApp.flush();
  return reply({ok:true,matchId:d.matchId,version:8,saved:true,location:loc,fineRows:fineCount});
 }catch(err){try{if(log&&logRow){log.getRange(logRow,4,1,2).setValues([['오류',String(err&&err.message||err).slice(0,500)]]);SpreadsheetApp.flush();}}catch(_){}return reply({ok:false,message:String(err&&err.message||err)});}
 finally{if(lock.hasLock())lock.releaseLock();}
}
function convertMatch(m,id){
 const fail=x=>{throw Error(x);};if(!m||typeof m!=='object')fail('Match data missing');if(m.practice||m.earlyFinish)fail('Practice or incomplete matches cannot be saved');
 if(m.kind!=='3쿠션')fail('Only 3-cushion matches are supported');if(!['공식경기','친선경기'].includes(m.league))fail('Unknown league type');
 const modes={'2':[2,1,'개인전'],'3':[3,1,'개인전'],'4':[4,1,'개인전'],'5':[5,1,'개인전'],'6':[6,1,'개인전'],team:[2,2,'2:2 팀전'],team222:[3,2,'3팀 팀전'],team33:[2,3,'3:3 팀전']};
 const spec=modes[String(m.mode)];if(!spec)fail('Unsupported match mode');const [groups,size,type]=spec,official=m.league==='공식경기';
 if(official&&String(m.mode)!=='2')fail('Official league supports 1:1 only');if(!['대대','중대','중대 1번','중대 2번'].includes(m.table))fail('Unknown table');
 if(!Array.isArray(m.players)||m.players.length!==groups)fail('Player/team count mismatch');
 const groupNames=m.players.map(p=>String(p&&p.name||'').trim());
 if(groupNames.some(n=>!n||n.length>80||/^[=+@-]/.test(n))||new Set(groupNames).size!==groups)fail('Invalid or duplicate player/team names');
 const people=size>1?m.members:groupNames;if(!Array.isArray(people)||people.length!==groups*size)fail('Member count mismatch');
 const clean=people.map(n=>String(n||'').trim());if(clean.some(n=>!n||n.length>80||/^[=+@-]/.test(n))||new Set(clean).size!==clean.length)fail('Invalid or duplicate member names');
 const stats=size>1?m.individuals:m.players;if(!Array.isArray(stats)||stats.length!==clean.length)fail('Individual statistics missing');
 const statRows=clean.map(name=>{const p=stats.find(x=>x&&x.name===name);if(!p)fail('Statistics missing for '+name);const score=Number(p.score),innings=Number(p.innings),average=Number(p.average);if(!Number.isInteger(score)||score<0||!Number.isInteger(innings)||innings<0||!Number.isFinite(average)||average<0)fail('Invalid statistics for '+name);return{name,score,innings,average};});
 const winner=Array.isArray(m.winner)?m.winner.map(String):[];if(winner.length!==1||!groupNames.includes(winner[0]))fail('Winner not confirmed');
 let fines=Array.isArray(m.fines)?m.fines.slice():null;
 if(!fines||fines.length!==groups){if(groups===2){const loser=groupNames.find(n=>n!==winner[0]);fines=[{name:winner[0],rank:1,amount:0},{name:loser,rank:2,amount:official?5:10}];}
 else if(Array.isArray(m.ranking)&&m.ranking.length===groups){const defaults=Array.isArray(m.fineSettings)?m.fineSettings:(groups===3?[0,5,10]:[0,0,0,0,0,0]);fines=m.ranking.map((name,i)=>({name,rank:i+1,amount:Number(defaults[i]||0)}));}
 else fail('Final ranking/fines missing; finish ranking in the app before sending');}
 fines.sort((a,b)=>Number(a.rank)-Number(b.rank));
 if(fines.some((f,i)=>!f||f.rank!==i+1||!groupNames.includes(f.name)||typeof f.amount!=='number'||!Number.isFinite(f.amount)||f.amount<0))fail('Invalid fine settings');
 if(new Set(fines.map(f=>f.name)).size!==groups||fines[0].name!==winner[0]||fines[0].amount!==0)fail('Ranking and winner do not match');if(official&&fines[1].amount!==5)fail('Official loss fine must be $5');
 const dt=new Date(m.date);if(!Number.isFinite(dt.getTime()))fail('Invalid match date');const ymd=Utilities.formatDate(dt,TZ,'yyyy-MM-dd').split('-').map(Number);const day=Math.floor(Date.UTC(ymd[0],ymd[1]-1,ymd[2])/86400000)+25569;
 const table=m.table==='대대'?'대대':'중대',teamLabels=['A팀','B팀','C팀'],rankNames=fines.map(f=>f.name);
 const sheetRanks=size>1?rankNames.map(n=>teamLabels[groupNames.indexOf(n)]):rankNames;
 const charges=[];if(size>1){clean.forEach((name,i)=>{const team=groupNames[Math.floor(i/size)],f=fines.find(x=>x.name===team);if(f&&f.amount>0)charges.push({name,amount:f.amount,rank:f.rank,group:team});});}
 else fines.forEach(f=>{if(f.amount>0)charges.push({name:f.name,amount:f.amount,rank:f.rank,group:f.name});});
 const avgs={};statRows.forEach(p=>avgs[p.name]=p.average);const roster=['Jack','Brian','Steve','Leo','Sion','Hank','Ian','Jo'];
 const note=`APP:${id} | ${m.league} ${type} | ${groupNames.join(' vs ')} | 순위 ${rankNames.join(' > ')} | ${statRows.map(p=>`${p.name} ${p.score}점/${p.innings}이닝 AVG ${p.average.toFixed(3)}`).join(' · ')}`.slice(0,490);
 return{id,official,day,table,groups,size,type,groupNames,people:clean,statRows,winner:winner[0],fines,charges,averages:roster.map(n=>Object.prototype.hasOwnProperty.call(avgs,n)?avgs[n]:''),note,
  friendlyRanks:sheetRanks,friendlyPlayers:clean,friendlyAverages:clean.map(n=>avgs[n]),rankText:rankNames.map((n,i)=>`${i+1}위 ${size>1?teamLabels[groupNames.indexOf(n)]:n}`).join(' > '),officialPlayers:m.players};
}
function getOrCreateMainMatch(ss,p,id){
 const marker='APP:'+id;
 if(p.official){
  const sh=ss.getSheetByName('리그경기');let row=findMarkerRow(sh,19,FIRST,LAST,marker);
  let order=row?Number(sh.getRange(row,18).getValue()):0;if(!row){validateOfficial(ss,p);row=firstEmptyRow(sh,2,FIRST,LAST);if(!row)throw Error('리그경기 has no empty row');}
  if(!order)order=nextOrder(ss);const a=p.officialPlayers[0],b=p.officialPlayers[1];
  sh.getRange(row,18,1,2).setValues([[order,marker]]);
  sh.getRange(row,2,1,5).setValues([[p.day,p.groupNames[0],p.groupNames[1],p.winner,'완료']]);
  sh.getRange(row,10,1,7).setValues([[Number(a.score),Number(a.innings),Number(a.average),Number(b.score),Number(b.innings),Number(b.average),p.table]]);
  return{row,order};
 }
 const sh=ss.getSheetByName('친선경기');let row=findMarkerRow(sh,25,FIRST,LAST,marker);
 let order=row?Number(sh.getRange(row,24).getValue()):0;if(!row){row=firstEmptyRow(sh,1,FIRST,LAST);if(!row)throw Error('친선경기 has no empty row');}
 if(!order)order=nextOrder(ss);
 const layout=sh.getRange(3,4).getValue();
  if(layout!=='테이블' && layout!=='선수1 (A팀)')throw Error('Unknown friendly sheet layout');
  const values=layout==='테이블'?[p.day,p.type,'3쿠션',p.table,...pad(p.friendlyPlayers,6),...pad(p.friendlyRanks,3),'완료',...pad(p.friendlyAverages,6),p.note]:[p.day,p.type,'3쿠션',...pad(p.friendlyPlayers,6),...pad(p.friendlyRanks,3),'완료',...pad(p.friendlyAverages,6),p.table,p.note];
 if(values.length!==21)throw Error('Friendly row width mismatch: '+values.length);
 sh.getRange(row,25).setValue(marker); // idempotency marker first; retry rewrites the full row
 sh.getRange(row,1,1,21).setValues([values]);
 sh.getRange(row,23,1,2).setValues([[p.rankText,order]]);
 sh.getRange(row,22).setValue('정상');
 return{row,order};
}
function appendFineRows(ss,p,id){
 if(p.official)return 0; // existing official fine formula remains authoritative
 const sh=ss.getSheetByName('앱벌금');const order=Number(ss.getSheetByName('친선경기').getRange(findMarkerRow(ss.getSheetByName('친선경기'),25,FIRST,LAST,'APP:'+id),24).getValue());
 let added=0;
 p.charges.forEach((c,i)=>{const key=id+'_'+i;const old=sh.getLastRow()>1?findMarkerRow(sh,6,2,sh.getLastRow(),key):0;const row=old||sh.getLastRow()+1;sh.getRange(row,1,1,6).setValues([[p.day,c.name,c.amount,'앱 친선 / '+p.type+' / '+c.rank+'위',order,key]]);if(!old)added++;});return added;
}
function ensureAutoFineFormula(ss){
 const sh=ss.getSheetByName('벌금_자동집계');let f=sh.getRange('A2').getFormula();
 if(!f||!f.includes('VSTACK(')||f.includes("REGEXMATCH('벌금기록'"))throw Error('Original fine formula required');
 if(!ss.getSheetByName('앱벌금')){const a=ss.insertSheet('앱벌금');a.appendRow(['날짜','이름','벌금원금','비고','입력순서','경기ID']);a.setFrozenRows(1);}
 if(f.includes("'앱벌금'!A2:E"))return;
 const statusCol=ss.getSheetByName('친선경기').getRange(3,4).getValue()==='테이블'?'N':'M';
  const marker=`('친선경기'!${statusCol}4:${statusCol}303="완료")`;
 if(f.split(marker).length!==12)throw Error('Unexpected friendly fine formula');
 const backup=PropertiesService.getScriptProperties();if(!backup.getProperty('ORIGINAL_FINE_FORMULA'))backup.setProperty('ORIGINAL_FINE_FORMULA',f);
 f=f.split(marker).join(marker+`*(LEFT('친선경기'!U4:U303,4)<>"APP:")`);
 const block=`IFERROR(FILTER('앱벌금'!A2:E,'앱벌금'!C2:C>0),{"","","","" ,""}),`;
 sh.getRange('A2').setFormula(f.replace('VSTACK(','VSTACK('+block));
}
function extendFineTotalFormula(ss){
 const sh=ss.getSheetByName('회비·입출금');if(!sh)throw Error('회비·입출금 sheet not found');const c=sh.getRange('B11'),f=c.getFormula();
 if(f&&f.includes('벌금기록')){const updated=f.replace(/\$F\$17:\$F\$305/g,'$F$17:$F$1002').replace(/F17:F305/g,'F17:F1002');if(updated!==f)c.setFormula(updated);}
 else throw Error('회비·입출금!B11 fine total formula not recognised');
}
function ensureTrackingColumns(ss){
 ensureColumn(ss.getSheetByName('리그경기'),19,3,'앱전송ID');
 ensureColumn(ss.getSheetByName('친선경기'),25,3,'앱전송ID');

}
function ensureColumn(sh,col,row,header){if(!sh)throw Error('Missing sheet for '+header);if(sh.getMaxColumns()<col)sh.insertColumnsAfter(sh.getMaxColumns(),col-sh.getMaxColumns());if(!sh.getRange(row,col).getValue())sh.getRange(row,col).setValue(header);}
function ensureLogHeaders(sh){if(!sh.getRange(1,5).getValue())sh.getRange(1,5).setValue('처리 결과');if(!sh.getRange(1,6).getValue())sh.getRange(1,6).setValue('연동 정보');}
function findLogRow(sh,id){const last=sh.getLastRow();if(last<2)return 0;const c=sh.getRange(2,1,last-1,1).createTextFinder(id).matchEntireCell(true).findNext();return c?c.getRow():0;}
function findMarkerRow(sh,col,start,end,marker){const c=sh.getRange(start,col,end-start+1,1).createTextFinder(marker).matchEntireCell(true).findNext();return c?c.getRow():0;}
function firstEmptyRow(sh,col,start,end){const v=sh.getRange(start,col,end-start+1,1).getValues(),i=v.findIndex(r=>r[0]===''||r[0]===null);return i<0?0:start+i;}
function firstEmptyRowBlock(sh,col,width,start,end){const v=sh.getRange(start,col,end-start+1,width).getValues(),i=v.findIndex(r=>r.every(x=>x===''||x===null));return i<0?0:start+i;}
function nextOrder(ss){const a=ss.getSheetByName('리그경기').getRange(`R${FIRST}:R${LAST}`).getValues().flat(),b=ss.getSheetByName('친선경기').getRange(`X${FIRST}:X${LAST}`).getValues().flat();return Math.max(0,...a.concat(b).map(Number).filter(Number.isFinite))+1;}
function pad(a,n){return Array.from({length:n},(_,i)=>a[i]===undefined?'':a[i]);}
function reply(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);}

function validateOfficial(ss,p){
 const settings=ss.getSheetByName('월간리그_설정');const roster=settings.getRange('A9:B17').getValues();
 if(p.people.some(n=>!roster.some(r=>r[0]===n&&r[1]==='참가')))throw Error('공식경기 참가 명단 확인');
 const month=d=>d instanceof Date?Utilities.formatDate(d,TZ,'yyyy-MM'):typeof d==='number'?new Date((d-25569)*86400000).toISOString().slice(0,7):String(d).slice(0,7);
 const rows=ss.getSheetByName('리그경기').getRange('B4:P303').getValues().filter(r=>r[4]==='완료'&&month(r[0])===month(p.day)&&p.people.includes(r[1])&&p.people.includes(r[2]));
 if(rows.length>=Number(settings.getRange('B4').getValue())||rows.some(r=>r[14]===p.table))throw Error('공식경기 상대 또는 테이블 중복');
}

// Read-only: latest confirmed weekly handicap, not the recommendation.
function readConfirmedHandicaps(){
 // The Member Management tab is the app's single source of truth.
 // Its B5:B12 formulas calculate the latest confirmed handicap from the
 // Handicap Change tab, using the same date logic displayed to club admins.
 const sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName('회원관리');
 if(!sh)throw Error('회원관리 sheet not found');
 const names=['Jack','Brian','Steve','Leo','Sion','Hank','Ian','Jo'];
 const rows=sh.getRange(5,1,names.length,2).getValues();
 const result={};
 rows.forEach(([name,value],i)=>{
  if(name!==names[i])throw Error('회원관리 명단 순서/이름 확인 필요: row '+(i+5));
  if(!Number.isInteger(value)||value<=0)throw Error('회원관리 확정 핸디가 비어 있거나 잘못되었습니다: '+name);
  result[name]=value;
 });
 return result;
}
function testHandicapRead(){Logger.log(JSON.stringify(readConfirmedHandicaps()));}