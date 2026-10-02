// Alex Training — Google Sheets sync backend (Google Apps Script)
// Deploy as a Web App bound to your training-log Google Sheet:
//   Extensions > Apps Script > paste this in, replacing any old code > Save
//   Deploy > Manage deployments > pick your existing deployment > Edit (pencil icon)
//     > Version: "New version" > Deploy
//   (Using "New version" on the SAME deployment keeps your existing Web App URL —
//    do NOT create a brand new deployment, the app already has this URL saved.)

const SESSIONS_SHEET = 'Sessions';
const STATE_SHEET = 'DeviceState';
const SESSIONS_HEADERS = ['SessionID','SavedAt','Date','User','ProgramID','ProgramLabel','Week','Deload','DayIndex','DayLabel','Exercise','SwappedFrom','SetNum','Weight','Reps','Notes','Ticked'];
const STATE_HEADERS = ['User','UpdatedAt','StateJSON'];

function getOrCreateSheet(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.type === 'session'){
      handleSessionPost(body);
    } else if(body.type === 'state'){
      handleStatePost(body);
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSessionPost(body){
  const sh = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  const rows = body.rows || [];
  rows.forEach(r=>{
    sh.appendRow([
      body.sessionId, body.sessionId, body.date, body.user, body.program, body.programLabel,
      body.week, body.deload, body.dayIndex, body.dayLabel,
      r.exercise, r.swappedFrom||'', r.set, r.weight, r.reps, r.notes||'', r.ticked
    ]);
  });
}

function handleStatePost(body){
  const sh = getOrCreateSheet(STATE_SHEET, STATE_HEADERS);
  const data = sh.getDataRange().getValues();
  let rowIndex = -1;
  for(let i=1;i<data.length;i++){ if(data[i][0] === body.user){ rowIndex = i+1; break; } }
  if(rowIndex>0){
    sh.getRange(rowIndex,1,1,3).setValues([[body.user, body.updatedAt, body.stateJSON]]);
  } else {
    sh.appendRow([body.user, body.updatedAt, body.stateJSON]);
  }
}

function doGet(e){
  const user = (e.parameter && e.parameter.user) || '';
  const sessions = readSessions(user);
  const state = readState(user);
  return ContentService.createTextOutput(JSON.stringify({sessions, state})).setMimeType(ContentService.MimeType.JSON);
}

function readSessions(user){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SESSIONS_SHEET);
  if(!sh) return [];
  const data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const bySession = {};
  for(let i=1;i<data.length;i++){
    const row = data[i];
    if(user && row[idx.User] !== user) continue;
    const sid = row[idx.SessionID];
    if(!bySession[sid]){
      bySession[sid] = {
        sessionId: sid, savedAt: Number(row[idx.SavedAt])||sid, date: row[idx.Date], user: row[idx.User],
        program: row[idx.ProgramID], programLabel: row[idx.ProgramLabel], week: row[idx.Week]===''?null:row[idx.Week],
        deload: !!row[idx.Deload], dayIndex: row[idx.DayIndex], dayLabel: row[idx.DayLabel], exercises: []
      };
    }
    const rec = bySession[sid];
    let ex = rec.exercises.find(x=> x.name === row[idx.Exercise] && (x.swappedFrom||'') === (row[idx.SwappedFrom]||''));
    if(!ex){
      ex = { name: row[idx.Exercise], swappedFrom: row[idx.SwappedFrom]||null, sets: [], notes: row[idx.Notes]||'' };
      rec.exercises.push(ex);
    }
    ex.sets[Number(row[idx.SetNum])-1] = { weight: row[idx.Weight], reps: row[idx.Reps], done: !!row[idx.Ticked] };
  }
  return Object.keys(bySession).map(k=> bySession[k]);
}

function readState(user){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(STATE_SHEET);
  if(!sh) return null;
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0] === user){ return { updatedAt: Number(data[i][1])||0, stateJSON: data[i][2] }; }
  }
  return null;
}
