// =============================================
// 텍스트마이닝 웹앱 - Code.gs v5.0 (Web API Server)
// 창순기획 | 문학 텍스트마이닝
// =============================================

const SHEET_EMO   = '감정어휘';
const SHEET_STOP  = '불용어';
const TEXT_PREFIX = '분석텍스트_';

// 확장 불용어 사전 (노이즈 및 기능어 제거)
const DEFAULT_STOPWORDS = [
  '&quot;', 'quot', 'lt', 'gt', 'amp', 'nbsp',
  '하다', '되다', '있다', '없다', '보다', '오다', '가다', '않다', '이다', '아니다',
  '들다', '내다', '버리다', '놓이다', '이렇다', '어떻다', '그렇다', '저렇다', '게다',
  '허다', '일이', '갈다', '치다', '주다', '받다', '모르다', '알다', '대하다',
  '위하다', '말하다', '생각하다', '나오다', '들어가다', '나가다', '올라가다', '내려가다',
  '가지다', '따르다', '자다', '먹다', '맞다', '이르다', '보이다', '느끼다',
  '것', '수', '등', '때', '곳', '말', '날', '이', '그', '저', '분', '줄', '바',
  '체', '뿐', '채', '만', '중', '후', '전', '점', '씨', '개', '번', '차', '명',
  '자', '쪽', '편', '통', '권', '장', '마리', '원', '년', '월', '일', '시', '분', '초',
  '아버', '어머'
];

// 스프레드시트 안전 획득 함수 (바인딩 시트 or ID/URL)
function getSpreadsheet(customIdOrUrl) {
  if (customIdOrUrl) {
    try {
      if (customIdOrUrl.includes('docs.google.com/spreadsheets')) {
        return SpreadsheetApp.openByUrl(customIdOrUrl);
      } else {
        return SpreadsheetApp.openById(customIdOrUrl);
      }
    } catch (e) {
      console.warn('openById/Url failed:', e);
    }
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    console.warn('getActiveSpreadsheet failed:', e);
  }
  return null;
}

// ─── GET 요청 처리 (Web API: JSON & JSONP) ────────────
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  const callback = p.callback || '';
  const customId = p.spreadsheetId || p.url || '';

  // 1. API 호출 처리 (action 파라미터가 있는 경우 JSON 반환)
  if (action) {
    let result = { success: false };
    try {
      if (action === 'list') {
        result = getTextSheetList(customId);
      } else if (action === 'data') {
        const sheetName = p.sheet || '';
        result = { success: true, data: getAllData(sheetName, customId) };
      } else if (action === 'char') {
        const sheetName = p.sheet || '';
        const chars = (p.chars || '').split(',').filter(Boolean);
        result = { success: true, data: getCharEmo(chars, sheetName, null, customId) };
      } else if (action === 'assoc') {
        const sheetName = p.sheet || '';
        const keyword = p.keyword || '';
        result = { success: true, data: getAssociation(keyword, sheetName, customId) };
      } else if (action === 'sentences') {
        const sheetName = p.sheet || '';
        const word = p.word || '';
        result = { success: true, data: getSentencesByWord(word, sheetName, customId) };
      } else {
        result = { success: false, message: '알 수 없는 action: ' + action };
      }
    } catch (err) {
      result = { success: false, message: err.message, stack: err.stack };
    }

    const jsonStr = JSON.stringify(result);
    if (callback) {
      // JSONP 응답
      return ContentService.createTextOutput(callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      // JSON 응답
      return ContentService.createTextOutput(jsonStr)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. 브라우저 직접 접속 시: index.html이 있으면 렌더링, 없으면 안내 페이지 출력
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('📚 문학 텍스트마이닝 대시보드')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:30px;line-height:1.8;text-align:center;background:#0e1117;color:#fff;min-height:100vh;">' +
      '<h2 style="color:#6366f1;font-size:24px;">📚 문학 텍스트마이닝 API 서버 실행 중</h2>' +
      '<p style="color:#94a3b8;font-size:14px;margin-top:10px;">본 Google Apps Script는 <b>GitHub 저장소에 호스팅된 index.html 웹앱의 백엔드 API</b>로 정상 작동 중입니다.</p>' +
      '<div style="background:#1e2337;padding:16px;border-radius:12px;margin:20px auto;max-width:500px;text-align:left;font-size:13px;border:1px solid #334155;">' +
      '<b>✅ 사용 방법:</b><br>' +
      '1. 깃허브 저장소(GitHub Pages) 또는 로컬에서 <code>index.html</code>을 실행합니다.<br>' +
      '2. 웹앱의 시작 화면에서 본 GAS URL을 통해 구글 시트 데이터를 실시간으로 가져옵니다.' +
      '</div>' +
      '<p style="color:#64748b;font-size:12px;">GAS 엔드포인트: ' + ScriptApp.getService().getUrl() + '</p>' +
      '</div>'
    );
  }
}

// ─── 분석 텍스트 시트 목록 ───────────────────────────
function getTextSheetList(customIdOrUrl) {
  try {
    const ss = getSpreadsheet(customIdOrUrl);
    if (!ss) return { success: false, list: [], sheets: [], message: '스프레드시트를 찾을 수 없습니다. 스프레드시트 URL이나 ID를 확인해 주세요.' };

    const allSheetNames = ss.getSheets().map(s => s.getName());
    
    // 1. '분석텍스트_' 접두사 시트 우선 검색
    let list = allSheetNames
      .filter(n => n.startsWith(TEXT_PREFIX))
      .map(n => ({ name: n, label: n.replace(TEXT_PREFIX, '') }));

    // 2. 접두사 시트가 없는 경우, '감정어휘'/'불용어'를 제외한 일반 시트 모두 포함
    if (!list.length) {
      list = allSheetNames
        .filter(n => n !== SHEET_EMO && n !== SHEET_STOP)
        .map(n => ({ name: n, label: n }));
    }

    return { success: true, list, sheets: list, ssTitle: ss.getName() };
  } catch (err) {
    console.error('getTextSheetList error:', err);
    return { success: false, list: [], sheets: [], message: err.message };
  }
}

// ─── 품사 추론 및 정규화 ─────────────────────────────
function inferPOS(token) {
  if (token.endsWith('다') && token.length >= 2) return 'verb';
  if (/[기지이고도로히]$/.test(token) && token.length >= 3) return 'adv';
  return 'noun';
}

function cleanToken(token) {
  if (!token) return '';
  let t = String(token)
    .replace(/&quot;|quot/g, '')
    .replace(/[.,?!;:~"'`「」『』()<>{}\[\]*^#@$%&+=\/\\|\-]/g, '')
    .trim();
  if (t === '아버') t = '아버지';
  if (t === '어머') t = '어머니';
  return t;
}

// ─── 공유 사전 로드 ──────────────────────────────────
function loadSharedDicts(ss) {
  let emoDict = {}, stopSet = new Set(DEFAULT_STOPWORDS);
  if (ss) {
    try {
      const emoSheet = ss.getSheetByName(SHEET_EMO);
      if (emoSheet) {
        const emoData = emoSheet.getDataRange().getValues();
        for (let i = 1; i < emoData.length; i++) {
          if (emoData[i][0]) emoDict[String(emoData[i][0]).trim()] = Number(emoData[i][1]) || 0;
        }
      }
    } catch (e) {
      console.warn('감정어휘 시트 로드 실패:', e);
    }

    try {
      const stopSheet = ss.getSheetByName(SHEET_STOP);
      if (stopSheet) {
        const stopData = stopSheet.getDataRange().getValues();
        for (let i = 1; i < stopData.length; i++) {
          if (stopData[i][0]) stopSet.add(String(stopData[i][0]).trim());
        }
      }
    } catch (e) {
      console.warn('불용어 시트 로드 실패:', e);
    }
  }
  return { emoDict, stopSet };
}

// ─── 공통 파서 (정밀 형태소 클리닝) ──────────────────
function parseSentences(rows, stopSet, emoDict) {
  const sentences = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && r[0] !== 0) continue;
    const rawTokens = String(r[2] || '').split(/\s+/)
      .map(t => cleanToken(t))
      .filter(t => t.length > 1 && t !== 'nan' && !stopSet.has(t));

    const tagged = rawTokens.map(t => ({ t, pos: inferPOS(t) }));
    const emoWords = [];
    let emoScore = 0;
    rawTokens.forEach(t => {
      if (t in emoDict) {
        emoWords.push([t, emoDict[t]]);
        emoScore += emoDict[t];
      }
    });

    sentences.push({
      no: Number(r[0]),
      original: String(r[1] || ''),
      tokens: rawTokens,
      tagged,
      stage: String(r[3] || '기타').trim(),
      speaker: r[4] ? String(r[4]).trim() : '',
      emoScore,
      emoWords
    });
  }
  return sentences;
}

// ─── 메인 데이터 (시트 기반) ─────────────────────────
function getAllData(sheetName, customIdOrUrl) {
  const ss = getSpreadsheet(customIdOrUrl);
  if (!ss) throw new Error('스프레드시트를 열 수 없습니다. 시트 URL/ID를 확인하거나 내장 모드를 이용해 주세요.');
  const { emoDict, stopSet } = loadSharedDicts(ss);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + sheetName);
  const sentences = parseSentences(sheet.getDataRange().getValues(), stopSet, emoDict);
  return buildResult(sentences, sheetName, sheetName.replace(TEXT_PREFIX, ''));
}

// ─── CSV 업로드 분석 ─────────────────────────────────
function analyzeCSV(csvText, label) {
  const ss = getSpreadsheet();
  let emoDict = {}, stopSet = new Set();
  if (ss) {
    const dicts = loadSharedDicts(ss);
    emoDict = dicts.emoDict;
    stopSet = dicts.stopSet;
  }

  const lines = csvText.split(/\r?\n/).filter(l => l.trim());
  const rows = [['문장번호', '원본문장', '형태소 전처리 문장', '소설 구성 단계']];
  for (let i = 1; i < lines.length; i++) {
    const cols = parseCSVLine(lines[i]);
    if (cols.length >= 4) rows.push(cols);
  }
  const sentences = parseSentences(rows, stopSet, emoDict);
  return buildResult(sentences, '__csv__:' + label, label);
}

// CSV 라인 파서
function parseCSVLine(line) {
  const result = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { inQ = !inQ; }
    else if (c === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
    else { cur += c; }
  }
  result.push(cur.trim());
  return result;
}

// ─── 결과 빌더 (통합 통계 및 TF-IDF 계산) ───────────
function buildResult(sentences, sheetName, label) {
  const fm = { noun: {}, verb: {}, adv: {}, all: {} };
  sentences.forEach(s => s.tagged.forEach(({ t, pos }) => {
    fm.all[t] = (fm.all[t] || 0) + 1;
    if (fm[pos]) fm[pos][t] = (fm[pos][t] || 0) + 1;
  }));

  const sorted = o => Object.entries(o).sort((a, b) => b[1] - a[1]);
  const freqAll  = sorted(fm.all);
  const freqNoun = sorted(fm.noun);
  const freqVerb = sorted(fm.verb);
  const freqAdv  = sorted(fm.adv);

  // 공출현 매핑
  const coocMap = {};
  sentences.forEach(s => {
    const toks = [...new Set(s.tokens)];
    for (let a = 0; a < toks.length; a++) {
      for (let b = a + 1; b < toks.length; b++) {
        const key = toks[a] < toks[b] ? toks[a] + '||' + toks[b] : toks[b] + '||' + toks[a];
        coocMap[key] = (coocMap[key] || 0) + 1;
      }
    }
  });
  const coocList = Object.entries(coocMap).sort((a, b) => b[1] - a[1]).slice(0, 400)
    .map(([k, v]) => {
      const [s, t] = k.split('||');
      return { source: s, target: t, weight: Number(v) };
    });

  // 구성단계
  const ORDER = ['발단', '전개', '위기', '절정', '결말'];
  const found = [...new Set(sentences.map(s => s.stage))];
  const stageOrder = ORDER.filter(s => found.includes(s));
  found.forEach(s => { if (!stageOrder.includes(s) && s) stageOrder.push(s); });
  if (!stageOrder.length) stageOrder.push('전체');

  const stageEmo = {};
  stageOrder.forEach(st => {
    const ss = sentences.filter(s => s.stage === st);
    stageEmo[st] = {
      count: ss.length,
      posSum: ss.reduce((a, x) => a + (x.emoScore > 0 ? x.emoScore : 0), 0),
      negSum: ss.reduce((a, x) => a + (x.emoScore < 0 ? x.emoScore : 0), 0),
      avgScore: ss.length ? ss.reduce((a, x) => a + x.emoScore, 0) / ss.length : 0
    };
  });

  // 구성단계별 TF-IDF
  const stageTfidf = {};
  stageOrder.forEach(st => {
    const ss = sentences.filter(s => s.stage === st);
    const tf = {};
    ss.forEach(s => s.tokens.forEach(t => { tf[t] = (tf[t] || 0) + 1; }));
    stageTfidf[st] = Object.entries(tf).map(([w, f]) => {
      const df = sentences.filter(s => s.tokens.includes(w)).length;
      const idf = Math.log(sentences.length / (df + 1)) + 1;
      return { word: w, tfidf: f * idf, freq: f, totalFreq: fm.all[w] || 0 };
    }).sort((a, b) => b.tfidf - a.tfidf).slice(0, 20);
  });

  const totalWords = freqAll.reduce((s, [, v]) => s + v, 0);

  return {
    sheetName,
    label,
    sentences: sentences.map(s => ({
      no: s.no,
      original: s.original,
      tokens: s.tokens,
      tagged: s.tagged,
      stage: s.stage,
      speaker: s.speaker,
      emoScore: s.emoScore,
      emoWords: s.emoWords
    })),
    freqAll: freqAll.slice(0, 200),
    freqNoun: freqNoun.slice(0, 150),
    freqVerb: freqVerb.slice(0, 150),
    freqAdv: freqAdv.slice(0, 150),
    coocList,
    stageEmo,
    stageTfidf,
    stageOrder,
    totalStats: {
      totalSentences: sentences.length,
      totalTokenTypes: freqAll.length,
      totalTokenCount: totalWords,
      posCount: sentences.filter(s => s.emoScore > 0).length,
      negCount: sentences.filter(s => s.emoScore < 0).length,
      neuCount: sentences.filter(s => s.emoScore === 0).length,
      ttr: sentences.length ? ((freqAll.length / (totalWords || 1)) * 100).toFixed(1) : 0,
      avgSentenceLength: sentences.length ? (totalWords / sentences.length).toFixed(1) : 0
    }
  };
}

// ─── 인물(키워드) 감정 & 관계 분석 ───────────────────
function getCharEmo(keywords, sheetName, csvData) {
  let sentences = [];
  const ss = getSpreadsheet();
  let emoDict = {}, stopSet = new Set();
  if (ss) {
    const d = loadSharedDicts(ss);
    emoDict = d.emoDict;
    stopSet = d.stopSet;
  }

  if (csvData) {
    const lines = csvData.split(/\r?\n/).filter(l => l.trim());
    const rows = [[]];
    for (let i = 1; i < lines.length; i++) {
      const c = parseCSVLine(lines[i]);
      if (c.length >= 4) rows.push(c);
    }
    sentences = parseSentences(rows, stopSet, emoDict);
  } else if (ss && sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) sentences = parseSentences(sheet.getDataRange().getValues(), stopSet, emoDict);
  }

  const charEmo = {};
  keywords.forEach(ch => {
    const sents = sentences.filter(s => s.tokens.includes(ch));
    const allEmo = [];
    sents.forEach(s => allEmo.push(...s.emoWords));
    const posW = {}, negW = {};
    allEmo.forEach(([w, sc]) => {
      if (sc > 0) posW[w] = (posW[w] || 0) + sc;
      else if (sc < 0) negW[w] = (negW[w] || 0) + Math.abs(sc);
    });
    const topSents = sents.filter(s => s.emoScore !== 0)
      .sort((a, b) => Math.abs(b.emoScore) - Math.abs(a.emoScore)).slice(0, 5)
      .map(s => ({ no: s.no, original: s.original, stage: s.stage, emoScore: s.emoScore, emoWords: s.emoWords }));

    charEmo[ch] = {
      sentenceCount: sents.length,
      posSum: allEmo.filter(([, e]) => e > 0).reduce((s, [, e]) => s + e, 0),
      negSum: allEmo.filter(([, e]) => e < 0).reduce((s, [, e]) => s + e, 0),
      posWords: Object.entries(posW).sort((a, b) => b[1] - a[1]).slice(0, 15),
      negWords: Object.entries(negW).sort((a, b) => b[1] - a[1]).slice(0, 15),
      topSentences: topSents
    };
  });

  return charEmo;
}

// ─── 연관어 검색 ─────────────────────────────────────
function getAssociation(keyword, sheetName) {
  const ss = getSpreadsheet();
  let stopSet = new Set();
  if (ss) stopSet = loadSharedDicts(ss).stopSet;
  if (!ss || !sheetName) return { keyword, associations: [], sentences: [] };
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { keyword, associations: [], sentences: [] };

  const rows = sheet.getDataRange().getValues();
  const coocMap = {};
  const kwSents = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && r[0] !== 0) continue;
    const tokens = String(r[2] || '').split(/\s+/).map(t => t.trim()).filter(t => t.length > 1 && t !== 'nan' && !stopSet.has(t));
    if (!tokens.includes(keyword)) continue;
    kwSents.push({ no: Number(r[0]), original: String(r[1] || ''), stage: String(r[3] || '') });
    tokens.forEach(t => { if (t !== keyword) coocMap[t] = (coocMap[t] || 0) + 1; });
  }

  return {
    keyword,
    associations: Object.entries(coocMap).sort((a, b) => b[1] - a[1]).slice(0, 40)
      .map(([word, weight]) => ({ word: String(word), weight: Number(weight) })),
    sentences: kwSents
  };
}

// ─── 단어 → 원본문장 조회 ────────────────────────────
function getSentencesByWord(word, sheetName) {
  const ss = getSpreadsheet();
  let emoDict = {}, stopSet = new Set();
  if (ss) {
    const dicts = loadSharedDicts(ss);
    emoDict = dicts.emoDict;
    stopSet = dicts.stopSet;
  }
  if (!ss || !sheetName) return [];
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && r[0] !== 0) continue;
    const tokens = String(r[2] || '').split(/\s+/).map(t => t.trim()).filter(t => t.length > 1 && t !== 'nan' && !stopSet.has(t));
    if (!tokens.includes(word)) continue;
    const emoWords = tokens.filter(t => t in emoDict).map(t => [t, emoDict[t]]);
    result.push({
      no: Number(r[0]),
      original: String(r[1] || ''),
      stage: String(r[3] || ''),
      tokens,
      emoWords,
      emoScore: emoWords.reduce((s, [, e]) => s + e, 0)
    });
  }
  return result;
}
