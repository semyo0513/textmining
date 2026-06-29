// =============================================
// 텍스트마이닝 웹앱 - Code.gs v3
// 창순기획 | 삼현여자중학교
// =============================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_EMO   = '감정어휘';
const SHEET_STOP  = '불용어';
const TEXT_PREFIX = '분석텍스트_';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('텍스트마이닝 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── 분석 텍스트 시트 목록 ───────────────────────────
function getTextSheetList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheets()
    .map(s => s.getName())
    .filter(n => n.startsWith(TEXT_PREFIX))
    .map(n => ({ name: n, label: n.replace(TEXT_PREFIX, '') }));
}

// ─── 품사 추론 ────────────────────────────────────────
function inferPOS(token) {
  if (token.endsWith('다') && token.length >= 2) return 'verb';
  if (/[기지이고도로히]$/.test(token) && token.length >= 3) return 'adv';
  return 'noun';
}

// ─── 공유 사전 로드 ──────────────────────────────────
function loadSharedDicts(ss) {
  const emoData  = ss.getSheetByName(SHEET_EMO).getDataRange().getValues();
  const stopData = ss.getSheetByName(SHEET_STOP).getDataRange().getValues();
  const emoDict = {};
  for (let i = 1; i < emoData.length; i++)
    if (emoData[i][0]) emoDict[String(emoData[i][0]).trim()] = Number(emoData[i][1]);
  const stopSet = new Set();
  for (let i = 1; i < stopData.length; i++)
    if (stopData[i][0]) stopSet.add(String(stopData[i][0]).trim());
  return { emoDict, stopSet };
}

// ─── 공통 파서 ───────────────────────────────────────
function parseSentences(rows, stopSet, emoDict) {
  const sentences = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const raw = String(r[2]).split(/\s+/).map(t => t.trim()).filter(t => t.length > 1 && t !== 'nan');
    const tokens = raw.filter(t => !stopSet.has(t));
    const tagged = tokens.map(t => ({ t, pos: inferPOS(t) }));
    const emoWords = []; let emoScore = 0;
    tokens.forEach(t => { if (t in emoDict) { emoWords.push([t, emoDict[t]]); emoScore += emoDict[t]; } });
    sentences.push({ no: Number(r[0]), original: String(r[1]), tokens, tagged, stage: String(r[3]), emoScore, emoWords });
  }
  return sentences;
}

// ─── 메인 데이터 (시트 기반) ─────────────────────────
function getAllData(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const { emoDict, stopSet } = loadSharedDicts(ss);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트 없음: ' + sheetName);
  const sentences = parseSentences(sheet.getDataRange().getValues(), stopSet, emoDict);
  return buildResult(sentences, sheetName, sheetName.replace(TEXT_PREFIX, ''));
}

// ─── CSV 업로드 분석 ─────────────────────────────────
// csvText: 클라이언트에서 읽어 온 CSV 문자열
// label:   사용자가 붙인 작품명
function analyzeCSV(csvText, label) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const { emoDict, stopSet } = loadSharedDicts(ss);

  // CSV → rows 파싱 (헤더 포함)
  const lines = csvText.split(/\r?\n/).filter(l => l.trim());
  const rows = [['문장번호','원본문장','형태소 전처리 문장','소설 구성 단계']]; // 헤더 더미
  for (let i = 1; i < lines.length; i++) {
    const cols = parseCSVLine(lines[i]);
    if (cols.length >= 4) rows.push(cols);
  }
  const sentences = parseSentences(rows, stopSet, emoDict);
  return buildResult(sentences, '__csv__:' + label, label);
}

// 간단한 CSV 한 줄 파서 (쌍따옴표 내 쉼표 처리)
function parseCSVLine(line) {
  const result = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { inQ = !inQ; }
    else if (c === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
    else { cur += c; }
  }
  result.push(cur.trim());
  return result;
}

// ─── 결과 빌더 (공통) ────────────────────────────────
function buildResult(sentences, sheetName, label) {
  // 빈도
  const fm = { noun:{}, verb:{}, adv:{}, all:{} };
  sentences.forEach(s => s.tagged.forEach(({ t, pos }) => {
    fm.all[t] = (fm.all[t]||0)+1; fm[pos][t] = (fm[pos][t]||0)+1;
  }));
  const sorted = o => Object.entries(o).sort((a,b)=>b[1]-a[1]);
  const freqAll  = sorted(fm.all);
  const freqNoun = sorted(fm.noun);
  const freqVerb = sorted(fm.verb);
  const freqAdv  = sorted(fm.adv);

  // 공출현
  const coocMap = {};
  sentences.forEach(s => {
    const toks = [...new Set(s.tokens)];
    for (let a = 0; a < toks.length; a++)
      for (let b = a+1; b < toks.length; b++) {
        const key = toks[a] < toks[b] ? toks[a]+'||'+toks[b] : toks[b]+'||'+toks[a];
        coocMap[key] = (coocMap[key]||0)+1;
      }
  });
  const coocList = Object.entries(coocMap).sort((a,b)=>b[1]-a[1]).slice(0,300)
    .map(([k,v]) => { const [s,t] = k.split('||'); return { source:s, target:t, weight:Number(v) }; });

  // 구성단계
  const ORDER = ['발단','전개','위기','절정','결말'];
  const found  = [...new Set(sentences.map(s=>s.stage))];
  const stageOrder = ORDER.filter(s=>found.includes(s));
  if (!stageOrder.length) stageOrder.push(...found);

  const stageEmo = {};
  stageOrder.forEach(st => {
    const ss = sentences.filter(s=>s.stage===st);
    stageEmo[st] = {
      count: ss.length,
      posSum: ss.reduce((a,x)=>a+(x.emoScore>0?x.emoScore:0),0),
      negSum: ss.reduce((a,x)=>a+(x.emoScore<0?x.emoScore:0),0),
      avgScore: ss.length ? ss.reduce((a,x)=>a+x.emoScore,0)/ss.length : 0
    };
  });

  // TF-IDF
  const stageTfidf = {};
  stageOrder.forEach(st => {
    const ss = sentences.filter(s=>s.stage===st);
    const tf = {}; ss.forEach(s=>s.tokens.forEach(t=>{ tf[t]=(tf[t]||0)+1; }));
    stageTfidf[st] = Object.entries(tf).map(([w,f]) => {
      const df = sentences.filter(s=>s.tokens.includes(w)).length;
      const idf = Math.log(sentences.length/(df+1))+1;
      return { word:w, tfidf:f*idf, freq:f, totalFreq:fm.all[w]||0 };
    }).sort((a,b)=>b.tfidf-a.tfidf).slice(0,15);
  });

  return {
    sheetName, label,
    sentences: sentences.map(s=>({
      no:s.no, original:s.original, tokens:s.tokens,
      tagged:s.tagged, stage:s.stage, emoScore:s.emoScore, emoWords:s.emoWords
    })),
    freqAll:  freqAll.slice(0,150),
    freqNoun: freqNoun.slice(0,100),
    freqVerb: freqVerb.slice(0,100),
    freqAdv:  freqAdv.slice(0,100),
    coocList,
    stageEmo, stageTfidf, stageOrder,
    totalStats: {
      totalSentences: sentences.length,
      totalTokenTypes: freqAll.length,
      totalTokenCount: freqAll.reduce((s,[,v])=>s+v,0),
      posCount: sentences.filter(s=>s.emoScore>0).length,
      negCount: sentences.filter(s=>s.emoScore<0).length,
      neuCount: sentences.filter(s=>s.emoScore===0).length,
    }
  };
}

// ─── 인물(키워드) 감정 분석 ── 클라이언트 요청 시 ────
// keywords: ['소년','소녀'] 처럼 사용자가 직접 입력한 목록
// sentences: DATA.sentences (클라이언트에서 전달)
//   → GAS 재쿼리 없이 이미 로드된 DATA 내에서 계산하면 서버왕복 1회 절약
//   → 하지만 GAS ↔ 클라이언트 데이터 전달이 더 빠르므로 sentences는 sheetName+keywords로 서버에서 처리
function getCharEmo(keywords, sheetName, csvData) {
  let sentences;
  if (csvData) {
    // CSV 모드: 이미 buildResult를 거쳤으나 sentences를 재파싱해야 함
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const { emoDict, stopSet } = loadSharedDicts(ss);
    const lines = csvData.split(/\r?\n/).filter(l=>l.trim());
    const rows = [[]];
    for (let i=1;i<lines.length;i++) { const c=parseCSVLine(lines[i]); if(c.length>=4) rows.push(c); }
    sentences = parseSentences(rows, stopSet, emoDict);
  } else {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const { emoDict, stopSet } = loadSharedDicts(ss);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {};
    sentences = parseSentences(sheet.getDataRange().getValues(), stopSet, emoDict);
  }

  const charEmo = {};
  keywords.forEach(ch => {
    const sents = sentences.filter(s => s.tokens.includes(ch));
    const allEmo = []; sents.forEach(s => allEmo.push(...s.emoWords));
    const posW={}, negW={};
    allEmo.forEach(([w,sc])=>{ if(sc>0) posW[w]=(posW[w]||0)+sc; else if(sc<0) negW[w]=(negW[w]||0)+Math.abs(sc); });
    const topSents = sents.filter(s=>s.emoScore!==0)
      .sort((a,b)=>Math.abs(b.emoScore)-Math.abs(a.emoScore)).slice(0,5)
      .map(s=>({ no:s.no, original:s.original, stage:s.stage, emoScore:s.emoScore, emoWords:s.emoWords }));
    charEmo[ch] = {
      sentenceCount: sents.length,
      posSum: allEmo.filter(([,e])=>e>0).reduce((s,[,e])=>s+e,0),
      negSum: allEmo.filter(([,e])=>e<0).reduce((s,[,e])=>s+e,0),
      posWords: Object.entries(posW).sort((a,b)=>b[1]-a[1]).slice(0,15),
      negWords: Object.entries(negW).sort((a,b)=>b[1]-a[1]).slice(0,15),
      topSentences: topSents
    };
  });
  return charEmo;
}

// ─── 연관어 검색 ─────────────────────────────────────
function getAssociation(keyword, sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const { stopSet } = loadSharedDicts(ss);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { keyword, associations:[], sentences:[] };
  const rows = sheet.getDataRange().getValues();
  const coocMap = {}; const kwSents = [];
  for (let i=1;i<rows.length;i++) {
    const r = rows[i]; if(!r[0]) continue;
    const tokens = String(r[2]).split(/\s+/).map(t=>t.trim()).filter(t=>t.length>1&&t!=='nan'&&!stopSet.has(t));
    if (!tokens.includes(keyword)) continue;
    kwSents.push({ no:Number(r[0]), original:String(r[1]), stage:String(r[3]) });
    tokens.forEach(t=>{ if(t!==keyword) coocMap[t]=(coocMap[t]||0)+1; });
  }
  return {
    keyword,
    associations: Object.entries(coocMap).sort((a,b)=>b[1]-a[1]).slice(0,30)
      .map(([word,weight])=>({ word:String(word), weight:Number(weight) })),
    sentences: kwSents
  };
}

// ─── 단어 → 원본문장 조회 ────────────────────────────
function getSentencesByWord(word, sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const { emoDict, stopSet } = loadSharedDicts(ss);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i=1;i<rows.length;i++) {
    const r = rows[i]; if(!r[0]) continue;
    const tokens = String(r[2]).split(/\s+/).map(t=>t.trim()).filter(t=>t.length>1&&t!=='nan'&&!stopSet.has(t));
    if (!tokens.includes(word)) continue;
    const emoWords = tokens.filter(t=>t in emoDict).map(t=>[t,emoDict[t]]);
    result.push({ no:Number(r[0]), original:String(r[1]), stage:String(r[3]),
      tokens, emoWords, emoScore:emoWords.reduce((s,[,e])=>s+e,0) });
  }
  return result;
}
