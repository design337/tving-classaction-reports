/**
 * ============================================================================
 * 티빙 개인정보 유출 집단소송 — 광고 데이터 통합 처리 + 전략 보고서 발행
 * Single-file Apps Script (Code.gs)
 * Version: v1.1.11 (diag + header-detect + full-range zone1, 2026-06-16)
 *
 * 변경점 (v1.0 -> v1.1)
 *   Fix 1: _updateOverview — '종합지표' 시트 구역1 일자별 KPI + 구역2 채널×주차 매트릭스 자동 갱신
 *   Fix 2: _logReportToSheet — '분석 보고서' 시트 헤더 기반 안전 append + 에러 발생 시 runLog 기록
 *   Fix 3: _aggregateByDate — GA 채널 매핑 확장 (kakao/오가닉/직접/리퍼럴 분리)
 *   +Add : action: "inspect" — 종합지표/분석보고서 헤더 구조 반환 (디버그)
 *
 * 사용:
 *   POST {action:"run", data:{...rawSheets...}}     → 데이터 입력 + 종합지표 갱신 + 일간/주간 보고서 발행
 *   POST {action:"build_report", meta:{...}}        → 시트 기존 데이터로 보고서만 재발행 (+종합지표 갱신)
 *   POST {action:"status"}                          → 상태 확인
 *   POST {action:"inspect"}                         → 종합지표/분석보고서 구조 덤프
 *   POST {action:"get_history"}                     → 마케팅 히스토리 반환
 *
 * Script Properties 필요:
 *   - SHEET_ID:    티빙 시트 ID
 *   - SECRET_TOKEN: 인증 토큰
 *   - GITHUB_PAT:  GitHub Personal Access Token
 *   - GITHUB_OWNER: design337
 *   - GITHUB_REPO: tving-classaction-reports
 * ============================================================================
 */

const CONFIG = {
  PROJECT_NAME: '티빙 개인정보 유출 집단소송',
  CAMPAIGN_START: '2026-06-04',
  CAMPAIGN_LP: 'https://classaction.lawfirmthe-h.com/',
  AUTHOR: '에이아이컨실리움',
  CLIENT: '법무법인 더 에이치 황해',

  // 시트 입력 정의 — 키 이름은 JSON 키 = 시트 이름
  RAW_SHEETS: {
    '메타_광고': {
      dateCol: '일',
      dedupe: ['캠페인 이름', '광고 세트 이름', '광고 이름', '일']
    },
    '네이버_캠페인': {
      dateCol: '일별',
      dedupe: ['캠페인유형', '캠페인', '광고그룹', '일별']
    },
    '네이버_키워드': {
      dateCol: '일별',
      dedupe: ['캠페인', '광고그룹', '키워드', '일별']
    },
    '네이버_소재': {
      dateCol: '일별',
      dedupe: ['캠페인', '광고그룹', '소재', '일별']
    },
    '네이버_검색어': {
      dateCol: '일별',
      dedupe: ['캠페인', '광고그룹', '검색어', '일별']
    },
    '구글_캠페인': {
      dateCol: '일자',
      dedupe: ['캠페인명', '일자']
    },
    '구글_채널유입': {
      dateCol: '일자',
      dedupe: ['채널', '캠페인명', '일자']
    },
    '구글_검색어': {
      dateCol: '일자',
      dedupe: ['캠페인', '검색어', '일자']
    },
    '구글_게재위치': {
      dateCol: '일자',
      dedupe: ['캠페인', '게재위치', '일자']
    },
    'GA4_Raw': {
      dateCol: '날짜',
      dedupe: ['날짜', '이벤트 이름', '세션 소스/매체', '세션 캠페인']
    },
    'GA4_티빙 집단소송 신청하기': {
      dateCol: '날짜',
      dedupe: ['날짜', '이벤트 이름', '세션 소스/매체', '세션 캠페인']
    }
  },

  HISTORY_SHEET: '마케팅 히스토리',
  HISTORY_HEADERS: ['순번', '날짜', '작성자', '변경 유형', '채널/매체', '변경 항목', '변경 상세 내용', '변경 사유', '관련 근거/데이터'],

  REPORT_SHEET: '분석 보고서',
  RUNLOG_SHEET: '_실행로그',
  OVERVIEW_SHEET: '종합지표',

  GITHUB_DAILY_PATH: 'daily',
  GITHUB_WEEKLY_PATH: 'weekly',

  COLOR: {
    META: '#7B68AE',
    NAVER: '#16A34A',
    GOOGLE: '#4285F4',
    GA: '#1A73E8',
    PRIMARY: '#1A2746',
    ACCENT: '#F59E0B',
    DOWN: '#DC2626',
    UP: '#16A34A'
  }
};

// ============================================================================
// Web App entry points
// ============================================================================

function doPost(e) {
  try {
    const params = e.parameter || {};
    const token = params.token || JSON.parse(e.postData.contents || '{}').token;
    const stored = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN');
    if (stored && token !== stored) {
      return _json({ok: false, error: 'unauthorized'});
    }
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || 'run';
    if (action === 'status') return _json(_status());
    if (action === 'inspect') return _json(_inspectSheets(body));
    if (action === 'cleanup_report_log') return _json(_cleanupReportLog());
    if (action === 'rebuild_report_log') return _json(_rebuildReportLog());
    if (action === 'dedupe_meta') return _json(_dedupeMeta());
    if (action === 'get_history') return _json({ok: true, history: _readMarketingHistory()});
    if (action === 'run' || action === 'full_run') return _json(_runFromJson(body));
    if (action === 'build_report') return _json(_buildReportFromSheet(body));
    return _json({ok: false, error: 'unknown action: ' + action});
  } catch (err) {
    return _json({ok: false, error: String(err), stack: err.stack});
  }
}

function doGet(e) {
  return _json({ok: true, info: 'Tving classaction reports', version: 'v1.1.11'});
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// MAIN — runFromJson
// ============================================================================

function _runFromJson(body) {
  const ss = _ss();
  const data = body.data || {};

  // 1. 일자 범위 자동 추출
  const range = _extractDateRange(data);
  if (!range.dates.length) return {ok: false, error: 'no dates found in data'};

  // 2. 시트에 데이터 입력 (dedupe 적용)
  const step1 = _appendDataToSheets(ss, data);

  // 2.5. 종합지표 시트 자동 갱신 (Fix 1) — zone1은 CAMPAIGN_START~maxDate
  let overviewResult = {ok: false, note: 'skipped'};
  try {
    const fullRange = {
      dates: range.dates,
      minDate: CONFIG.CAMPAIGN_START,
      maxDate: range.maxDate,
      days: range.days
    };
    overviewResult = _updateOverview(ss, fullRange);
  } catch (err) {
    overviewResult = {ok: false, error: String(err)};
    _appendRunLog('update_overview', 'ERROR', String(err), (err.stack || '').slice(0, 300));
  }

  // 3. 마케팅 히스토리 읽기
  const history = _readMarketingHistory();

  // 4. 보고서 종류 결정 (일간 / 주간 / 둘 다)
  const reportTypes = _decideReportTypes(range);

  // 5. 보고서 생성 + 발행
  const results = [];
  for (const rt of reportTypes) {
    if (rt.type === 'daily') {
      for (const d of rt.dates) {
        const r = _publishDailyReport(ss, d, history);
        results.push(r);
      }
    } else if (rt.type === 'weekly') {
      const r = _publishWeeklyReport(ss, rt.weekStart, rt.weekEnd, history);
      results.push(r);
    }
  }

  // 6. 실행 로그
  _appendRunLog('run', 'SUCCESS', 'dates=' + range.dates.length + ' reports=' + results.length, JSON.stringify(results).slice(0, 500));

  return {ok: true, range, step1, overview: overviewResult, history_rows: history.length, results};
}

// ============================================================================
// Date range extraction
// ============================================================================

function _extractDateRange(data) {
  const set = new Set();
  for (const sheetName in CONFIG.RAW_SHEETS) {
    const rows = data[sheetName];
    if (!Array.isArray(rows)) continue;
    const dateCol = CONFIG.RAW_SHEETS[sheetName].dateCol;
    for (const r of rows) {
      const d = _normDate(r[dateCol]);
      if (d) set.add(d);
    }
  }
  const dates = Array.from(set).sort();
  return {
    dates,
    minDate: dates[0] || null,
    maxDate: dates[dates.length - 1] || null,
    days: dates.length
  };
}

function _normDate(d) {
  if (d === null || d === undefined || d === '') return '';
  if (d instanceof Date) {
    const y = d.getFullYear(), m = d.getMonth()+1, dd = d.getDate();
    return y + '-' + _pad(m) + '-' + _pad(dd);
  }
  let s = String(d).trim();
  s = s.replace(/\./g, '-').replace(/-+$/, '').replace(/\s/g, '');
  // 20260604 → 2026-06-04
  const raw = s.replace(/-/g, '');
  if (raw.length === 8 && /^\d+$/.test(raw)) {
    return raw.slice(0,4) + '-' + raw.slice(4,6) + '-' + raw.slice(6,8);
  }
  if (raw.length === 6 && /^\d+$/.test(raw)) {
    return '20' + raw.slice(0,2) + '-' + raw.slice(2,4) + '-' + raw.slice(4,6);
  }
  // YYYY-MM-DD already
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function _pad(n) { return n < 10 ? '0' + n : '' + n; }

function _decideReportTypes(range) {
  const types = [];
  // 일간: 모든 일자 각각
  types.push({type: 'daily', dates: range.dates});

  // 주간: 7일 이상 또는 minDate ~ maxDate가 정확히 월~일 한 주
  const minD = new Date(range.minDate);
  const maxD = new Date(range.maxDate);
  const days = Math.round((maxD - minD) / 86400000) + 1;

  if (days >= 7) {
    // 주간 보고서 — minDate 포함 주의 월~일 기준
    const ws = _weekStart(minD);
    const we = _addDays(ws, 6);
    types.push({type: 'weekly', weekStart: _fmt(ws), weekEnd: _fmt(we)});

    // 데이터가 여러 주 걸친 경우 마지막 완성 주도
    if (days >= 14) {
      const we2 = _weekStart(maxD);
      const we2end = _addDays(we2, 6);
      if (_fmt(we2) !== _fmt(ws)) {
        types.push({type: 'weekly', weekStart: _fmt(we2), weekEnd: _fmt(we2end)});
      }
    }
  }
  return types;
}

function _weekStart(d) {
  const c = new Date(d.getTime());
  const day = c.getDay(); // 0=Sun..6=Sat
  const diff = (day === 0 ? -6 : 1 - day); // Monday start
  c.setDate(c.getDate() + diff);
  c.setHours(0, 0, 0, 0);
  return c;
}

function _addDays(d, n) {
  const c = new Date(d.getTime());
  c.setDate(c.getDate() + n);
  return c;
}

function _fmt(d) {
  return d.getFullYear() + '-' + _pad(d.getMonth()+1) + '-' + _pad(d.getDate());
}

// ============================================================================
// Sheet append with dedupe
// ============================================================================

function _appendDataToSheets(ss, data) {
  const result = {};
  for (const sheetName in CONFIG.RAW_SHEETS) {
    const cfg = CONFIG.RAW_SHEETS[sheetName];
    const incoming = data[sheetName];
    if (!Array.isArray(incoming) || incoming.length === 0) {
      result[sheetName] = {note: 'no data'};
      continue;
    }
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      result[sheetName] = {error: 'sheet not found'};
      continue;
    }

    // 캠페인 필터 — tving / 티빙 포함만 통과
    const filtered = incoming.filter(r => _isProjectCampaign(r));

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const existing = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow()-1, headers.length).getValues()
      : [];

    // Build dedupe set from existing data
    const dedupeSet = new Set();
    for (const row of existing) {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      dedupeSet.add(_dedupeKey(obj, cfg.dedupe));
    }

    // Append new unique rows
    const toAppend = [];
    let skipDup = 0, skipFilter = incoming.length - filtered.length;
    for (const row of filtered) {
      const key = _dedupeKey(row, cfg.dedupe);
      if (dedupeSet.has(key)) { skipDup++; continue; }
      dedupeSet.add(key);
      const newRow = headers.map(h => row[h] !== undefined ? row[h] : '');
      toAppend.push(newRow);
    }

    if (toAppend.length > 0) {
      sheet.getRange(sheet.getLastRow()+1, 1, toAppend.length, headers.length).setValues(toAppend);
    }

    result[sheetName] = {
      incoming: incoming.length,
      filtered: filtered.length,
      appended: toAppend.length,
      skipped_duplicate: skipDup,
      skipped_filter: skipFilter,
      final: sheet.getLastRow() - 1
    };
  }
  return result;
}

function _isProjectCampaign(row) {
  // 캠페인 컬럼 후보 검색
  const keys = ['캠페인 이름', '캠페인', '캠페인명', '세션 캠페인'];
  for (const k of keys) {
    if (row[k]) {
      const s = String(row[k]).toLowerCase();
      if (s.includes('tving') || s.includes('티빙')) return true;
    }
  }
  return false;
}

function _dedupeKey(row, keys) {
  return keys.map(k => {
    const v = row[k];
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return _fmt(v);
    return String(v).trim();
  }).join('|');
}

// ============================================================================
// Marketing history reading
// ============================================================================

function _readMarketingHistory() {
  const sheet = _ss().getSheetByName(CONFIG.HISTORY_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];

  // 헤더 행 자동 탐지 — '날짜'/'채널' 포함된 첫 행
  let headerIdx = -1;
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i].map(c => String(c));
    if (row.some(c => c.includes('날짜')) && row.some(c => c.includes('채널'))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) return [];

  const headers = data[headerIdx].map(c => String(c).trim());
  const events = [];
  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(c => c === '' || c === null)) continue;
    const obj = {};
    headers.forEach((h, j) => { if (h) obj[h] = row[j]; });

    // 날짜 정규화
    obj._date = _normDate(obj['날짜']);
    if (!obj._date) continue;

    // 채널 표준화
    obj._channel = String(obj['채널/매체'] || '').toLowerCase();
    if (obj._channel.includes('네이버')) obj._channelStd = '네이버SA';
    else if (obj._channel.includes('메타') || obj._channel.includes('페이스북') || obj._channel.includes('인스타')) obj._channelStd = '메타';
    else if (obj._channel.includes('구글')) obj._channelStd = '구글';
    else if (obj._channel.includes('ga')) obj._channelStd = 'GA';
    else obj._channelStd = obj['채널/매체'] || '';

    events.push(obj);
  }
  return events;
}

// ============================================================================
// Aggregate ad data from sheets (used by report builder)
// ============================================================================

function _readSheetRows(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, headers.length).getValues();
  return rows.map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

function _aggregateByDate(ss, startDate, endDate) {
  // 채널별 일별 집계
  const dateInRange = d => d >= startDate && d <= endDate;
  const num = v => {
    if (v === null || v === undefined || v === '') return 0;
    const s = String(v).replace(/,/g, '').replace(/%/g, '').replace(/₩/g, '').replace(/\s/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  // Meta
  const meta = {cost: 0, imp: 0, clk: 0, conv: 0, rows: []};
  _readSheetRows(ss, '메타_광고').forEach(r => {
    const d = _normDate(r['일']);
    if (!dateInRange(d)) return;
    meta.cost += num(r['지출 금액 (KRW)']);
    meta.imp  += num(r['노출']);
    meta.clk  += num(r['링크 클릭']);
    meta.conv += num(r['결과']);
    meta.rows.push(r);
  });

  // Naver Campaign
  const naver = {cost: 0, imp: 0, clk: 0, rows: []};
  _readSheetRows(ss, '네이버_캠페인').forEach(r => {
    const d = _normDate(r['일별']);
    if (!dateInRange(d)) return;
    naver.cost += num(r['총비용']);
    naver.imp  += num(r['노출수']);
    naver.clk  += num(r['클릭수']);
    naver.rows.push(r);
  });

  // Google
  const google = {cost: 0, imp: 0, clk: 0, conv: 0, rows: []};
  _readSheetRows(ss, '구글_캠페인').forEach(r => {
    const d = _normDate(r['일자']);
    if (!dateInRange(d)) return;
    google.cost += num(r['총비용']);
    google.imp  += num(r['노출수']);
    google.clk  += num(r['클릭수']);
    google.conv += num(r['전환']);
    google.rows.push(r);
  });

  // GA4_Raw — 사이트 전반 세션/참여율 — 채널×주차 가중 평균 이탈률/참여율 계산용 (v1.1.11)
  // 참여율 값이 % 형태("40.68%" 또는 0.4068 모두 처리), 방문수 가중치로 누적
  const gaSess = {byChannel: {}, totalSessions: 0, totalEngaged: 0};
  const parseRate = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    const s = String(v).trim();
    const hasPercent = s.indexOf('%') >= 0;
    const n = parseFloat(s.replace(/[%,\s]/g, ''));
    if (isNaN(n)) return 0;
    if (hasPercent) return n / 100;
    if (n > 1) return n / 100;  // 헤더 % 표기 없어도 값이 >1이면 % 단위
    return n;
  };
  _readSheetRows(ss, 'GA4_Raw').forEach(r => {
    const d = _normDate(r['날짜']);
    if (!dateInRange(d)) return;
    const sessions = num(r['방문수']);
    if (sessions === 0) return;
    const engageRate = parseRate(r['참여율']);
    const engaged = sessions * engageRate;
    const sm = String(r['세션 소스/매체'] || '');
    let ch = 'Other';
    if (/naver/i.test(sm)) ch = '네이버SA';
    else if (/meta|facebook|fb|instagram|ig_/i.test(sm)) ch = '메타';
    else if (/google.*cpc|gad|adwords/i.test(sm)) ch = '구글';
    else if (/kakao|kkt/i.test(sm)) ch = '카카오';
    else if (/direct|\(direct\)|\(none\)|\(not set\)|not.set|data.not.available/i.test(sm)) ch = '직접';
    else if (/organic|google.*organic|naver.*organic/i.test(sm)) ch = '오가닉';
    else if (/referral/i.test(sm)) ch = '리퍼럴';
    else if (/ig|social/i.test(sm)) ch = '메타';
    else if (/chatgpt|ai-assistant/i.test(sm)) ch = 'AI어시스턴트';
    else if (/home/i.test(sm)) ch = '자체유입';
    if (!gaSess.byChannel[ch]) gaSess.byChannel[ch] = {sessions: 0, engaged: 0};
    gaSess.byChannel[ch].sessions += sessions;
    gaSess.byChannel[ch].engaged += engaged;
    gaSess.totalSessions += sessions;
    gaSess.totalEngaged += engaged;
  });

  // GA4 — 신청 이벤트 (Fix 3: 채널 매핑 확장)
  const ga = {users: 0, events: 0, byChannel: {}, rows: []};
  _readSheetRows(ss, 'GA4_티빙 집단소송 신청하기').forEach(r => {
    const d = _normDate(r['날짜']);
    if (!dateInRange(d)) return;
    const u = num(r['총 사용자']);
    const ev = num(r['이벤트 수']);
    ga.users += u;
    ga.events += ev;
    const sm = String(r['세션 소스/매체'] || '');
    let ch = 'Other';
    if (/naver/i.test(sm)) ch = '네이버SA';
    else if (/meta|facebook|fb|instagram|ig_/i.test(sm)) ch = '메타';
    else if (/google.*cpc|gad|adwords/i.test(sm)) ch = '구글';
    else if (/kakao|kkt/i.test(sm)) ch = '카카오';
    else if (/direct|\(direct\)|\(none\)|\(not set\)|not.set|data.not.available/i.test(sm)) ch = '직접';
    else if (/organic|google.*organic|naver.*organic/i.test(sm)) ch = '오가닉';
    else if (/referral/i.test(sm)) ch = '리퍼럴';
    else if (/ig|social/i.test(sm)) ch = '메타';
    else if (/chatgpt|ai-assistant/i.test(sm)) ch = 'AI어시스턴트';
    else if (/home/i.test(sm)) ch = '자체유입';
    ga.byChannel[ch] = (ga.byChannel[ch] || 0) + u;
    ga.rows.push(r);
  });

  // 네이버 키워드 효율 랭킹
  const naverKw = [];
  _readSheetRows(ss, '네이버_키워드').forEach(r => {
    const d = _normDate(r['일별']);
    if (!dateInRange(d)) return;
    naverKw.push({
      keyword: r['키워드'],
      camp: r['캠페인'],
      imp: num(r['노출수']),
      clk: num(r['클릭수']),
      ctr: num(r['클릭률(%)']),
      cpc: num(r['평균 CPC']),
      cost: num(r['총비용'])
    });
  });

  // 구글 채널유입
  const googleCh = [];
  _readSheetRows(ss, '구글_채널유입').forEach(r => {
    const d = _normDate(r['일자']);
    if (!dateInRange(d)) return;
    googleCh.push({
      ch: r['채널'],
      camp: r['캠페인명'],
      imp: num(r['노출수']),
      clk: num(r['클릭수']),
      cost: num(r['비용'])
    });
  });

  return {meta, naver, google, ga, gaSess, naverKw, googleCh};
}

// ============================================================================
// Strategic analysis builder
// ============================================================================

function _buildStrategicAnalysis(agg, history, startDate, endDate) {
  const tot = agg.meta.cost + agg.naver.cost + agg.google.cost;
  const totClk = agg.meta.clk + agg.naver.clk + agg.google.clk;
  const cvr = totClk > 0 ? (agg.ga.users / totClk * 100) : 0;
  const cpa = agg.ga.users > 0 ? (tot / agg.ga.users) : 0;

  // 핵심 채널 식별
  const channels = [
    {name: '메타', cost: agg.meta.cost, clk: agg.meta.clk},
    {name: '네이버SA', cost: agg.naver.cost, clk: agg.naver.clk},
    {name: '구글', cost: agg.google.cost, clk: agg.google.clk}
  ];
  channels.sort((a,b) => b.cost - a.cost);
  const mainChannel = channels[0].name;

  // 전환 채널 식별 (GA 신청 기준)
  const convChannel = Object.entries(agg.ga.byChannel)
    .sort((a,b) => b[1] - a[1])[0];

  // 키워드 TOP 5 (네이버)
  const kwAgg = {};
  agg.naverKw.forEach(k => {
    const key = k.keyword || '(미지정)';
    if (!kwAgg[key]) kwAgg[key] = {imp:0, clk:0, cost:0};
    kwAgg[key].imp += k.imp;
    kwAgg[key].clk += k.clk;
    kwAgg[key].cost += k.cost;
  });
  const topKeywords = Object.entries(kwAgg)
    .map(([kw, v]) => ({kw, ...v, ctr: v.imp>0?v.clk/v.imp*100:0, cpc: v.clk>0?v.cost/v.clk:0}))
    .sort((a,b) => b.cost - a.cost).slice(0, 10);

  // 효율 좋은 키워드 (CTR 높고 CPC 낮음)
  const efficientKw = topKeywords
    .filter(k => k.ctr > 5 && k.cpc < 200 && k.clk > 0)
    .sort((a,b) => b.ctr - a.ctr).slice(0, 3);

  // 구글 채널유입
  const googleChAgg = {};
  agg.googleCh.forEach(g => {
    const ch = g.ch || 'Other';
    if (!googleChAgg[ch]) googleChAgg[ch] = {imp:0, clk:0, cost:0};
    googleChAgg[ch].imp += g.imp;
    googleChAgg[ch].clk += g.clk;
    googleChAgg[ch].cost += g.cost;
  });

  // 마케팅 히스토리 매칭
  const eventsInRange = history.filter(h => h._date >= startDate && h._date <= endDate);
  const causality = [];
  eventsInRange.forEach(h => {
    causality.push({
      date: h._date,
      channel: h._channelStd,
      event: h['변경 항목'],
      detail: h['변경 상세 내용'],
      reason: h['변경 사유']
    });
  });

  // 인사이트
  const insights = [];
  if (channels[0].cost > channels[1].cost * 1.5) {
    insights.push({
      observation: mainChannel + ' 채널이 총 광고비의 ' + (channels[0].cost/tot*100).toFixed(0) + '% 차지 (₩' + _fmt0(channels[0].cost) + ')',
      cause: '단일 채널 의존도가 높아 해당 채널 효율 변동이 전체 KPI에 즉시 영향',
      action: mainChannel + ' 일별 변동 추적 강화 + 보조 채널 점진 확대로 리스크 분산'
    });
  }
  if (convChannel && convChannel[1] > 0) {
    insights.push({
      observation: 'GA 신청 전환 채널 1위: ' + convChannel[0] + ' (' + convChannel[1] + '명)',
      cause: convChannel[0] + '이 ' + (convChannel[0] === mainChannel ? '광고비와 전환 모두' : '광고비는 작지만 전환 효율은 높은') + ' 채널',
      action: convChannel[0] === mainChannel
        ? convChannel[0] + ' 효율 유지 + 다른 채널 전환 기여 분석 필요'
        : convChannel[0] + ' 예산 확대 검토 — 전환 효율 좋음'
    });
  }
  if (efficientKw.length > 0) {
    const list = efficientKw.map(k => '"' + k.kw + '" (CTR ' + k.ctr.toFixed(1) + '%, CPC ₩' + _fmt0(k.cpc) + ')').join(', ');
    insights.push({
      observation: '네이버 SA 효율 최상 키워드 발견: ' + list,
      cause: '검색 의도가 명확하면서 CPC가 낮은 롱테일 키워드 — 입찰가 상향·노출 확대 여지 큼',
      action: '효율 좋은 롱테일을 별도 광고그룹으로 분리 + 입찰가 단계적 상향'
    });
  }
  causality.forEach(c => {
    insights.push({
      observation: c.date + ' ' + c.channel + ' — ' + c.event,
      cause: c.reason || c.detail,
      action: '변경 전후 효율 비교 + 효과 측정 후 확대 여부 결정'
    });
  });

  // 액션
  const actions = [];
  actions.push({priority: 'P1', channel: mainChannel, item: mainChannel + ' 채널 일별 KPI 변동 모니터링 + 효율 키워드 예산 확대', due: '익영업일'});
  if (efficientKw.length > 0) {
    actions.push({priority: 'P1', channel: '네이버SA', item: '효율 최상 키워드 (' + efficientKw.map(k=>k.kw).join('/') + ') 별도 광고그룹 분리 + 입찰가 점진 상향', due: '이번 주'});
  }
  if (convChannel && convChannel[0] !== mainChannel) {
    actions.push({priority: 'P2', channel: convChannel[0], item: convChannel[0] + ' 전환 기여도 확인 후 예산 확대 검토', due: '이번 주'});
  }
  actions.push({priority: 'P2', channel: '전체', item: '채널 다변화 — 단일 채널 의존도 분산', due: '다음 주'});

  return {
    conclusion: startDate + '~' + endDate + ' 총 광고비 ₩' + _fmt0(tot) + ' / 클릭 ' + _fmt0(totClk) + ' / GA 신청 ' + agg.ga.users + '. 핵심 채널 ' + mainChannel + ', 전환 채널 ' + (convChannel ? convChannel[0] : 'N/A'),
    kpi: {totalCost: tot, totalClk: totClk, gaUsers: agg.ga.users, cvr, cpa, channels, convChannel, mainChannel},
    channelFindings: {
      '메타':    '광고비 ₩' + _fmt0(agg.meta.cost) + ' / 노출 ' + _fmt0(agg.meta.imp) + ' / 클릭 ' + _fmt0(agg.meta.clk) + ' / 전환 ' + _fmt0(agg.meta.conv),
      '네이버SA': '광고비 ₩' + _fmt0(agg.naver.cost) + ' / 노출 ' + _fmt0(agg.naver.imp) + ' / 클릭 ' + _fmt0(agg.naver.clk),
      '구글':    '광고비 ₩' + _fmt0(agg.google.cost) + ' / 노출 ' + _fmt0(agg.google.imp) + ' / 클릭 ' + _fmt0(agg.google.clk) + ' / 전환 ' + _fmt0(agg.google.conv),
      'GA':     '신청 총사용자 ' + agg.ga.users + ' / 이벤트 ' + agg.ga.events
    },
    topKeywords,
    efficientKeywords: efficientKw,
    googleChannels: googleChAgg,
    historyEvents: causality,
    insights,
    actions
  };
}

function _fmt0(n) {
  if (!n && n !== 0) return '0';
  return Math.round(n).toLocaleString();
}

// ============================================================================
// Report HTML builder
// ============================================================================

function _buildReportHtml(reportType, startDate, endDate, analysis, history) {
  const dateLabel = (reportType === 'weekly')
    ? startDate + ' ~ ' + endDate
    : endDate + ' (단일 일자)';
  const reportTitle = (reportType === 'weekly')
    ? CONFIG.PROJECT_NAME + ' 광고 성과 주간 분석 보고서 (' + _shortDate(startDate) + '~' + _shortDate(endDate) + ')'
    : CONFIG.PROJECT_NAME + ' 광고 성과 일간 분석 보고서 (' + _shortDate(endDate) + ')';
  const k = analysis.kpi;
  const cf = analysis.channelFindings;
  const C = CONFIG.COLOR;

  const getImp = ch => (analysis.kpi.channels.find(c => c.name === ch) || {}).imp || 0;
  const getClk = ch => (analysis.kpi.channels.find(c => c.name === ch) || {}).clk || 0;
  const getCost = ch => (analysis.kpi.channels.find(c => c.name === ch) || {}).cost || 0;

  let html = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>' + reportTitle + '</title>'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:-apple-system,"Pretendard","Apple SD Gothic Neo",sans-serif;background:#F8F9FB;color:#1A2746;line-height:1.6;padding:32px 20px}'
    + '.wrap{max-width:1100px;margin:0 auto}'
    + '.cover{background:linear-gradient(135deg,#1A2746 0%,#2A3756 100%);color:#fff;padding:40px 48px;border-radius:8px;margin-bottom:24px}'
    + '.cover .badge{display:inline-block;background:rgba(255,255,255,0.15);padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600;letter-spacing:1px;margin-bottom:16px}'
    + '.cover h1{font-size:32px;font-weight:700;margin-bottom:8px}'
    + '.cover .sub{font-size:14px;opacity:0.85;margin-bottom:24px}'
    + '.cover-meta{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;border-top:1px solid rgba(255,255,255,0.2);padding-top:20px}'
    + '.cover-meta .item .label{font-size:11px;opacity:0.7;margin-bottom:4px}'
    + '.cover-meta .item .value{font-size:13px;font-weight:600}'
    + '.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}'
    + '.kpi-card{background:#fff;padding:20px 22px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.05);border-top:3px solid ' + C.PRIMARY + '}'
    + '.kpi-card .label{font-size:12px;color:#6B7280;font-weight:600;margin-bottom:6px;letter-spacing:0.3px}'
    + '.kpi-card .value{font-size:26px;font-weight:700;color:' + C.PRIMARY + ';margin-bottom:4px}'
    + '.kpi-card .delta{font-size:11px;color:#6B7280}'
    + '.section{background:#fff;padding:32px 40px;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,0.04);margin-bottom:20px}'
    + '.section-num{display:inline-block;background:' + C.PRIMARY + ';color:#fff;padding:4px 10px;border-radius:3px;font-size:11px;font-weight:700;margin-bottom:8px;letter-spacing:1px}'
    + '.section h2{font-size:20px;color:' + C.PRIMARY + ';margin-bottom:6px;font-weight:700}'
    + '.section .lead{color:#6B7280;font-size:13px;margin-bottom:20px}'
    + 'table{width:100%;border-collapse:collapse;font-size:13px;margin:14px 0}'
    + 'th{background:#F3F4F6;color:#374151;padding:10px 12px;text-align:left;font-weight:600;border-bottom:2px solid #E5E7EB}'
    + 'td{padding:10px 12px;border-bottom:1px solid #E5E7EB}'
    + 'tr:last-child td{border-bottom:0}'
    + 'td.num{text-align:right;font-variant-numeric:tabular-nums}'
    + 'th.num{text-align:right}'
    + 'tr.total td{background:#FAFBFC;font-weight:700;border-top:2px solid ' + C.PRIMARY + '}'
    + 'tr.highlight td{background:#FFFBEB}'
    + '.ch-dual{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin:16px 0}'
    + '.ch-box{background:#FAFBFC;padding:18px;border-radius:6px;border:1px solid #E5E7EB}'
    + '.ch-box.meta{border-top:3px solid ' + C.META + '}'
    + '.ch-box.naver{border-top:3px solid ' + C.NAVER + '}'
    + '.ch-box.google{border-top:3px solid ' + C.GOOGLE + '}'
    + '.ch-box .title{font-weight:700;font-size:14px;margin-bottom:10px}'
    + '.ch-row{display:flex;justify-content:space-between;padding:5px 0;font-size:12px;border-bottom:1px dashed #E5E7EB}'
    + '.ch-row:last-child{border:0}'
    + '.ch-row .k{color:#6B7280}'
    + '.ch-row .v{font-weight:600;font-variant-numeric:tabular-nums}'
    + '.callout{background:#F0F9FF;border-left:4px solid ' + C.PRIMARY + ';padding:14px 18px;margin:14px 0;border-radius:0 4px 4px 0;font-size:13px}'
    + '.callout .title{font-weight:700;color:' + C.PRIMARY + ';display:block;margin-bottom:6px}'
    + '.tag{display:inline-block;padding:2px 10px;border-radius:3px;font-size:11px;color:#fff;font-weight:600;letter-spacing:0.3px}'
    + '.tag.meta{background:' + C.META + '}.tag.naver{background:' + C.NAVER + '}.tag.google{background:' + C.GOOGLE + '}.tag.ga{background:' + C.GA + '}'
    + '.star{color:' + C.ACCENT + ';font-weight:700}'
    + '.action{background:#FFFBEB;border:1px solid #FCD34D;border-radius:6px;padding:18px 22px;margin:14px 0}'
    + '.action-num{display:inline-block;background:' + C.ACCENT + ';color:#fff;padding:3px 10px;border-radius:3px;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:8px}'
    + '.action h3{font-size:15px;margin-bottom:8px;color:' + C.PRIMARY + '}'
    + '.action .obs{font-size:13px;margin:6px 0}'
    + '.action .obs b{color:' + C.PRIMARY + '}'
    + '.history-tl{background:#F3F4F6;padding:20px 24px;border-radius:6px;margin:14px 0}'
    + '.hist-item{padding:10px 0;border-bottom:1px dashed #D1D5DB}'
    + '.hist-item:last-child{border:0}'
    + '.hist-date{display:inline-block;background:' + C.PRIMARY + ';color:#fff;padding:2px 8px;border-radius:3px;font-size:11px;font-weight:600;margin-right:10px;font-variant-numeric:tabular-nums}'
    + '.hist-item b{color:' + C.PRIMARY + '}'
    + '.hist-detail{font-size:12px;color:#6B7280;margin-top:4px;padding-left:8px}'
    + '.footer{margin-top:32px;padding:20px;border-top:1px solid #E5E7EB;font-size:11px;color:#6B7280;text-align:center}'
    + '</style></head><body><div class="wrap">';

  html += '<div class="cover">'
    + '<div class="badge">TVING CLASS ACTION · ' + (reportType === 'weekly' ? 'WEEKLY' : 'DAILY') + ' MARKETING REPORT</div>'
    + '<h1>' + CONFIG.PROJECT_NAME + '</h1>'
    + '<div class="sub">' + (reportType === 'weekly' ? '주간' : '일일') + ' 광고 성과 보고서</div>'
    + '<div class="cover-meta">'
    + '<div class="item"><div class="label">분석 기간</div><div class="value">' + dateLabel + '</div></div>'
    + '<div class="item"><div class="label">발행일</div><div class="value">' + new Date().toLocaleDateString('ko-KR') + '</div></div>'
    + '<div class="item"><div class="label">채널</div><div class="value">메타 / 네이버SA / 구글 / GA4</div></div>'
    + '<div class="item"><div class="label">작성</div><div class="value">' + CONFIG.AUTHOR + '</div></div>'
    + '</div></div>';

  const totalImp = getImp('메타') + getImp('네이버SA') + getImp('구글');
  html += '<div class="kpi-grid">'
    + '<div class="kpi-card"><div class="label">총 광고비</div><div class="value">₩' + _fmt0(k.totalCost) + '</div><div class="delta">메타 + 네이버SA + 구글</div></div>'
    + '<div class="kpi-card"><div class="label">총 클릭</div><div class="value">' + _fmt0(k.totalClk) + '</div><div class="delta">CTR ' + (k.totalClk>0&&totalImp>0?(k.totalClk/totalImp*100).toFixed(2):'0') + '%</div></div>'
    + '<div class="kpi-card"><div class="label">유입 사용자 (GA4)</div><div class="value">' + _fmt0(k.gaUsers) + '</div><div class="delta">접수 완료 전환</div></div>'
    + '<div class="kpi-card"><div class="label">CPA</div><div class="value">₩' + _fmt0(k.cpa) + '</div><div class="delta">광고비 ÷ 신청</div></div>'
    + '</div>';

  html += '<div class="section">'
    + '<div class="section-num">SECTION 01</div>'
    + '<h2>하루 요약 — 핵심 결론</h2>'
    + '<div class="lead">광고비·전환·효율 종합</div>'
    + '<div class="callout"><span class="title">📌 핵심 메시지</span>' + analysis.conclusion + '</div>'
    + '</div>';

  html += '<div class="section">'
    + '<div class="section-num">SECTION 02</div>'
    + '<h2>채널별 성과</h2>'
    + '<div class="lead">메타 / 네이버SA / 구글 / GA</div>'
    + '<div class="ch-dual">'
    + '<div class="ch-box meta"><div class="title"><span class="tag meta">메타</span> 광고</div>'
    + '<div class="ch-row"><span class="k">광고비</span><span class="v">₩' + _fmt0(getCost('메타')) + '</span></div>'
    + '<div class="ch-row"><span class="k">노출</span><span class="v">' + _fmt0(getImp('메타')) + '</span></div>'
    + '<div class="ch-row"><span class="k">클릭</span><span class="v">' + _fmt0(getClk('메타')) + '</span></div>'
    + '<div class="ch-row"><span class="k">CTR</span><span class="v">' + (getImp('메타')>0?(getClk('메타')/getImp('메타')*100).toFixed(2):'0') + '%</span></div>'
    + '<div class="ch-row"><span class="k">CPC</span><span class="v">₩' + (getClk('메타')>0?Math.round(getCost('메타')/getClk('메타')):'0') + '</span></div>'
    + '</div>'
    + '<div class="ch-box naver"><div class="title"><span class="tag naver">네이버SA</span> 검색광고</div>'
    + '<div class="ch-row"><span class="k">광고비</span><span class="v">₩' + _fmt0(getCost('네이버SA')) + '</span></div>'
    + '<div class="ch-row"><span class="k">노출</span><span class="v">' + _fmt0(getImp('네이버SA')) + '</span></div>'
    + '<div class="ch-row"><span class="k">클릭</span><span class="v">' + _fmt0(getClk('네이버SA')) + '</span></div>'
    + '<div class="ch-row"><span class="k">CTR</span><span class="v">' + (getImp('네이버SA')>0?(getClk('네이버SA')/getImp('네이버SA')*100).toFixed(2):'0') + '%</span></div>'
    + '<div class="ch-row"><span class="k">CPC</span><span class="v">₩' + (getClk('네이버SA')>0?Math.round(getCost('네이버SA')/getClk('네이버SA')):'0') + '</span></div>'
    + '</div>'
    + '<div class="ch-box google"><div class="title"><span class="tag google">구글</span> 광고</div>'
    + '<div class="ch-row"><span class="k">광고비</span><span class="v">₩' + _fmt0(getCost('구글')) + '</span></div>'
    + '<div class="ch-row"><span class="k">노출</span><span class="v">' + _fmt0(getImp('구글')) + '</span></div>'
    + '<div class="ch-row"><span class="k">클릭</span><span class="v">' + _fmt0(getClk('구글')) + '</span></div>'
    + '<div class="ch-row"><span class="k">CTR</span><span class="v">' + (getImp('구글')>0?(getClk('구글')/getImp('구글')*100).toFixed(2):'0') + '%</span></div>'
    + '<div class="ch-row"><span class="k">CPC</span><span class="v">₩' + (getClk('구글')>0?Math.round(getCost('구글')/getClk('구글')):'0') + '</span></div>'
    + '</div>'
    + '</div>'
    + '<table>'
    + '<tr><th>구분</th><th class="num">광고비</th><th class="num">노출</th><th class="num">클릭</th><th class="num">CTR</th><th class="num">CPC</th></tr>'
    + '<tr><td><span class="tag meta">메타</span></td><td class="num">₩' + _fmt0(getCost('메타')) + '</td><td class="num">' + _fmt0(getImp('메타')) + '</td><td class="num">' + _fmt0(getClk('메타')) + '</td><td class="num">' + (getImp('메타')>0?(getClk('메타')/getImp('메타')*100).toFixed(2):'0') + '%</td><td class="num">₩' + (getClk('메타')>0?Math.round(getCost('메타')/getClk('메타')):'0') + '</td></tr>'
    + '<tr><td><span class="tag naver">네이버SA</span></td><td class="num">₩' + _fmt0(getCost('네이버SA')) + '</td><td class="num">' + _fmt0(getImp('네이버SA')) + '</td><td class="num">' + _fmt0(getClk('네이버SA')) + '</td><td class="num">' + (getImp('네이버SA')>0?(getClk('네이버SA')/getImp('네이버SA')*100).toFixed(2):'0') + '%</td><td class="num">₩' + (getClk('네이버SA')>0?Math.round(getCost('네이버SA')/getClk('네이버SA')):'0') + '</td></tr>'
    + '<tr><td><span class="tag google">구글</span></td><td class="num">₩' + _fmt0(getCost('구글')) + '</td><td class="num">' + _fmt0(getImp('구글')) + '</td><td class="num">' + _fmt0(getClk('구글')) + '</td><td class="num">' + (getImp('구글')>0?(getClk('구글')/getImp('구글')*100).toFixed(2):'0') + '%</td><td class="num">₩' + (getClk('구글')>0?Math.round(getCost('구글')/getClk('구글')):'0') + '</td></tr>'
    + '<tr class="total"><td>합계</td><td class="num">₩' + _fmt0(k.totalCost) + '</td><td class="num">' + _fmt0(totalImp) + '</td><td class="num">' + _fmt0(k.totalClk) + '</td><td class="num">' + (k.totalClk>0&&totalImp>0?(k.totalClk/totalImp*100).toFixed(2):'0') + '%</td><td class="num">₩' + (k.totalClk>0?Math.round(k.totalCost/k.totalClk):'0') + '</td></tr>'
    + '</table>'
    + '<div class="callout"><span class="title">해석</span>' + cf['메타'] + '<br>' + cf['네이버SA'] + '<br>' + cf['구글'] + '<br>' + cf['GA'] + '</div>'
    + '</div>';

  // Section 03: 네이버 키워드
  if (analysis.topKeywords && analysis.topKeywords.length > 0) {
    html += '<div class="section">'
      + '<div class="section-num">SECTION 03</div>'
      + '<h2>네이버 SA 키워드 효율</h2>'
      + '<div class="lead">노출 / 클릭 / CTR / CPC / 비용 분석</div>'
      + '<table>'
      + '<tr><th>키워드</th><th class="num">노출</th><th class="num">클릭</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">비용</th></tr>';
    const effSet = new Set(analysis.efficientKeywords.map(e => e.kw));
    analysis.topKeywords.forEach(kw => {
      const star = effSet.has(kw.kw) ? '<span class="star">⭐</span> ' : '';
      const cls = effSet.has(kw.kw) ? 'highlight' : '';
      html += '<tr class="' + cls + '"><td>' + star + kw.kw + '</td><td class="num">' + _fmt0(kw.imp) + '</td><td class="num">' + _fmt0(kw.clk) + '</td><td class="num">' + kw.ctr.toFixed(2) + '%</td><td class="num">₩' + _fmt0(kw.cpc) + '</td><td class="num">₩' + _fmt0(kw.cost) + '</td></tr>';
    });
    html += '</table>';
    if (analysis.efficientKeywords.length > 0) {
      html += '<div class="callout"><span class="title">⭐ 효율 최상 키워드</span>' + analysis.efficientKeywords.map(k => '<b>' + k.kw + '</b> (CTR ' + k.ctr.toFixed(1) + '%, CPC ₩' + _fmt0(k.cpc) + ')').join(' · ') + '<br>입찰가 점진 상향 + 별도 광고그룹 분리 검토</div>';
    }
    html += '</div>';
  }

  // Section 04: 구글 채널유입
  if (analysis.googleChannels && Object.keys(analysis.googleChannels).length > 0) {
    html += '<div class="section">'
      + '<div class="section-num">SECTION 04</div>'
      + '<h2>구글 채널유입 분포</h2>'
      + '<div class="lead">YouTube / Google 검색 / Display 등</div>'
      + '<table>'
      + '<tr><th>채널</th><th class="num">노출</th><th class="num">클릭</th><th class="num">비용</th></tr>';
    Object.entries(analysis.googleChannels).sort((a,b)=>b[1].cost-a[1].cost).forEach(([ch, vv]) => {
      html += '<tr><td>' + ch + '</td><td class="num">' + _fmt0(vv.imp) + '</td><td class="num">' + _fmt0(vv.clk) + '</td><td class="num">₩' + _fmt0(vv.cost) + '</td></tr>';
    });
    html += '</table></div>';
  }

  // Section 04B: GA 채널 분포 (Fix 3 가시화)
  if (analysis.kpi && Object.keys(analysis.gaChannels || {}).length === 0) {
    // populate from kpi
  }
  // Build GA channel section from analysis (use convChannel order via ga.byChannel passed indirectly through kpi)
  // We'll emit the section using analysis._gaByChannel if exists, else skip
  // To make it easy, we set _gaByChannel later from publishers

  // Section 05: 마케팅 히스토리
  if (analysis.historyEvents && analysis.historyEvents.length > 0) {
    html += '<div class="section">'
      + '<div class="section-num">SECTION 05</div>'
      + '<h2>마케팅 히스토리</h2>'
      + '<div class="lead">분석 기간 내 캠페인 변경 사항 (사용자 입력)</div>'
      + '<div class="history-tl">';
    analysis.historyEvents.forEach(h => {
      const tagCls = _chTagClass(h.channel);
      html += '<div class="hist-item">'
        + '<span class="hist-date">' + h.date + '</span>'
        + '<span class="tag ' + tagCls + '">' + h.channel + '</span>'
        + '<b>' + h.event + '</b>'
        + '<div class="hist-detail">' + (h.detail || '') + (h.reason ? '<br>사유: ' + h.reason : '') + '</div>'
        + '</div>';
    });
    html += '</div></div>';
  }

  // Section 06: 인사이트 & 액션
  html += '<div class="section">'
    + '<div class="section-num">SECTION 06</div>'
    + '<h2>인사이트 & 액션 제안</h2>'
    + '<div class="lead">관찰 · 원인 · 액션</div>';
  analysis.insights.forEach((ins, i) => {
    html += '<div class="action">'
      + '<div class="action-num">ACTION ' + String(i+1).padStart(2,'0') + '</div>'
      + '<h3>' + ins.observation + '</h3>'
      + '<div class="obs"><b>원인.</b> ' + ins.cause + '</div>'
      + '<div class="obs"><b>액션.</b> ' + ins.action + '</div>'
      + '</div>';
  });
  html += '</div>';

  // Section 07: 실행 계획
  html += '<div class="section">'
    + '<div class="section-num">SECTION 07</div>'
    + '<h2>실행 계획</h2>'
    + '<div class="lead">Priority · Channel · Item · Due</div>'
    + '<table>'
    + '<tr><th>우선순위</th><th>채널</th><th>실행 항목</th><th>실행 시점</th></tr>';
  analysis.actions.forEach(a => {
    html += '<tr><td><b>' + a.priority + '</b></td><td>' + a.channel + '</td><td>' + a.item + '</td><td>' + a.due + '</td></tr>';
  });
  html += '</table></div>';

  html += '<div class="footer">' + CONFIG.AUTHOR + ' | ' + CONFIG.CLIENT + ' — CONFIDENTIAL | 작성일 ' + new Date().toLocaleDateString('ko-KR') + '</div>'
    + '</div></body></html>';

  return {html, title: reportTitle};
}

function _chTagClass(ch) {
  if (/메타/.test(ch)) return 'meta';
  if (/네이버/.test(ch)) return 'naver';
  if (/구글/.test(ch)) return 'google';
  return 'ga';
}

function _shortDate(d) {
  const parts = d.split('-');
  return parts.length === 3 ? parseInt(parts[1]) + '/' + parseInt(parts[2]) : d;
}

// ============================================================================
// Publish reports
// ============================================================================

function _publishDailyReport(ss, date, history) {
  const agg = _aggregateByDate(ss, date, date);
  const analysis = _buildStrategicAnalysis(agg, history, date, date);
  analysis.kpi.channels.forEach(c => {
    if (c.name === '메타') c.imp = agg.meta.imp;
    if (c.name === '네이버SA') c.imp = agg.naver.imp;
    if (c.name === '구글') c.imp = agg.google.imp;
  });

  const built = _buildReportHtml('daily', date, date, analysis, history);
  const path = CONFIG.GITHUB_DAILY_PATH + '/' + date + '.html';
  const url = _publishToGithub(path, built.html, 'daily ' + date);
  _logReportToSheet('daily', date, date, url, built.title);
  return {type: 'daily', date, url, title: built.title};
}

function _publishWeeklyReport(ss, weekStart, weekEnd, history) {
  const agg = _aggregateByDate(ss, weekStart, weekEnd);
  const analysis = _buildStrategicAnalysis(agg, history, weekStart, weekEnd);
  analysis.kpi.channels.forEach(c => {
    if (c.name === '메타') c.imp = agg.meta.imp;
    if (c.name === '네이버SA') c.imp = agg.naver.imp;
    if (c.name === '구글') c.imp = agg.google.imp;
  });

  const built = _buildReportHtml('weekly', weekStart, weekEnd, analysis, history);
  const path = CONFIG.GITHUB_WEEKLY_PATH + '/' + weekStart + '_' + weekEnd + '.html';
  const url = _publishToGithub(path, built.html, 'weekly ' + weekStart + '~' + weekEnd);
  _logReportToSheet('weekly', weekStart, weekEnd, url, built.title);
  return {type: 'weekly', weekStart, weekEnd, url, title: built.title};
}

// ============================================================================
// GitHub publish
// ============================================================================

function _publishToGithub(path, html, commitMsg) {
  const props = PropertiesService.getScriptProperties();
  const pat = props.getProperty('GITHUB_PAT');
  const owner = props.getProperty('GITHUB_OWNER') || 'design337';
  const repo = props.getProperty('GITHUB_REPO') || 'tving-classaction-reports';

  if (!pat) throw new Error('GITHUB_PAT not set in Script Properties');

  const url = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
  let sha = null;
  try {
    const get = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {Authorization: 'Bearer ' + pat, Accept: 'application/vnd.github+json'},
      muteHttpExceptions: true
    });
    if (get.getResponseCode() === 200) sha = JSON.parse(get.getContentText()).sha;
  } catch (e) {}

  const body = {
    message: 'Auto: ' + commitMsg,
    content: Utilities.base64Encode(Utilities.newBlob(html).getBytes()),
    branch: 'main'
  };
  if (sha) body.sha = sha;

  const put = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: {Authorization: 'Bearer ' + pat, Accept: 'application/vnd.github+json'},
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (put.getResponseCode() >= 300) {
    throw new Error('GitHub push failed: ' + put.getResponseCode() + ' ' + put.getContentText().slice(0, 300));
  }
  return 'https://' + owner + '.github.io/' + repo + '/' + path;
}

// ============================================================================
// Report log + run log (Fix 2)
// ============================================================================

// === 메타 시트 정규화 + 모든 시트 일자 정렬 (v1.1.11) ===
// _normalizeMetaSheet: 메타_광고 시트 캠페인명 통일 + trim + 빈 row 제거 + 일자 정렬
// _sortAllSheetsByDate: 모든 광고/GA4 시트 일자 컬럼 기준 오름차순 정렬
function _normalizeMetaSheet() {
  const sheet = _ss().getSheetByName('메타_광고');
  if (!sheet) return {ok: false, error: 'sheet not found'};
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2) return {ok: true, note: 'empty'};

  const data = sheet.getRange(1, 1, lr, lc).getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const idxCampaign = headers.indexOf('캠페인 이름');
  const idxAdSet = headers.indexOf('광고 세트 이름');
  const idxAd = headers.indexOf('광고 이름');
  const idxDate = headers.indexOf('일');
  const idxCost = headers.indexOf('지출 금액 (KRW)');

  const NORM_CAMP = '티빙_집단소송_260604';  // 통일된 캠페인명

  // 1. 노이즈 제거 + trim + 캠페인명 통일
  const cleaned = [];
  let removedEmpty = 0, removedNoise = 0;
  for (const r of rows) {
    // 빈 row 검사
    const isEmpty = r.every(c => c === '' || c === null || c === undefined);
    if (isEmpty) { removedEmpty++; continue; }
    // cost null/0이고 노출도 0인 노이즈
    const cost = parseFloat(String(r[idxCost]||'0').replace(/,/g,'')) || 0;
    const idxImp = headers.indexOf('노출');
    const imp = idxImp >= 0 ? parseFloat(String(r[idxImp]||'0').replace(/,/g,'')) || 0 : 0;
    if (cost === 0 && imp === 0) { removedNoise++; continue; }
    // 캠페인명 통일 — '티빙_집단소송' prefix 가진 모든 변형을 NORM_CAMP로
    const newRow = r.slice();
    const c0 = String(newRow[idxCampaign] || '').trim();
    if (/^티빙_집단소송/.test(c0) || c0.indexOf('티빙_집단소송') === 0 || c0.indexOf('티빙') === 0) {
      newRow[idxCampaign] = NORM_CAMP;
    } else {
      newRow[idxCampaign] = c0;
    }
    // trim 광고세트/광고이름
    if (idxAdSet >= 0) newRow[idxAdSet] = String(newRow[idxAdSet] || '').trim();
    if (idxAd >= 0) newRow[idxAd] = String(newRow[idxAd] || '').trim();
    cleaned.push(newRow);
  }

  // 2. 일자 오름차순 → 광고세트 → 광고이름 정렬
  cleaned.sort((a, b) => {
    const da = _normDate(a[idxDate]);
    const db = _normDate(b[idxDate]);
    if (da !== db) return da < db ? -1 : 1;
    const sa = String(a[idxAdSet] || '');
    const sb = String(b[idxAdSet] || '');
    if (sa !== sb) return sa < sb ? -1 : 1;
    return String(a[idxAd] || '') < String(b[idxAd] || '') ? -1 : 1;
  });

  // 3. 시트 데이터 영역 모두 삭제 + 재기록
  sheet.getRange(2, 1, lr - 1, lc).clearContent();
  if (cleaned.length > 0) {
    sheet.getRange(2, 1, cleaned.length, lc).setValues(cleaned);
  }

  // 일자 범위
  let minDate = null, maxDate = null;
  if (cleaned.length > 0) {
    minDate = _normDate(cleaned[0][idxDate]);
    maxDate = _normDate(cleaned[cleaned.length - 1][idxDate]);
  }

  return {
    ok: true,
    before: rows.length,
    after: cleaned.length,
    removedEmpty,
    removedNoise,
    normalizedCampaign: NORM_CAMP,
    dateRange: minDate + '~' + maxDate
  };
}

// 모든 광고/GA4 시트 일자 컬럼 기준 오름차순 정렬
function _sortAllSheetsByDate() {
  const ss = _ss();
  const results = {};
  for (const sheetName in CONFIG.RAW_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { results[sheetName] = 'not found'; continue; }
    const lr = sheet.getLastRow();
    const lc = sheet.getLastColumn();
    if (lr < 3) { results[sheetName] = 'too small'; continue; }
    const dateCol = CONFIG.RAW_SHEETS[sheetName].dateCol;
    const headers = sheet.getRange(1, 1, 1, lc).getValues()[0];
    const idxDate = headers.indexOf(dateCol);
    if (idxDate < 0) { results[sheetName] = 'date col not found: ' + dateCol; continue; }

    const data = sheet.getRange(2, 1, lr - 1, lc).getValues();
    // 빈 row 제거 + 정렬
    const cleaned = data.filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
    cleaned.sort((a, b) => {
      const da = _normDate(a[idxDate]);
      const db = _normDate(b[idxDate]);
      if (da !== db) return da < db ? -1 : 1;
      return 0;
    });

    sheet.getRange(2, 1, lr - 1, lc).clearContent();
    if (cleaned.length > 0) {
      sheet.getRange(2, 1, cleaned.length, lc).setValues(cleaned);
    }
    results[sheetName] = {before: data.length, after: cleaned.length};
  }
  return {ok: true, sheets: results};
}

// === 메타_광고 시트 dedupe (v1.1.11) ===
// 캠페인명 정규화: 'tving_집단소송_260604' 같은 suffix(_숫자/_yyyymmdd) 제거 → '티빙_집단소송'
// 정규화된 (캠페인+광고세트+광고이름+일) 키로 그룹화 → 가장 cost 큰 row만 보존
function _dedupeMeta() {
  const sheet = _ss().getSheetByName('메타_광고');
  if (!sheet) return {ok: false, error: 'sheet not found'};
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2) return {ok: true, note: 'empty'};

  const headers = sheet.getRange(1, 1, 1, lc).getValues()[0];
  const data = sheet.getRange(2, 1, lr - 1, lc).getValues();

  const idxCampaign = headers.indexOf('캠페인 이름');
  const idxAdSet = headers.indexOf('광고 세트 이름');
  const idxAd = headers.indexOf('광고 이름');
  const idxDate = headers.indexOf('일');
  const idxCost = headers.indexOf('지출 금액 (KRW)');

  // 캠페인명 정규화: trailing '_숫자' suffix 제거
  const normCampaign = (c) => {
    if (!c) return '';
    let s = String(c).trim();
    // Remove trailing '_YYYYMMDD' (6-8 digits), '_숫자'
    s = s.replace(/_\d+$/, '');
    return s;
  };
  const normDate = (d) => {
    if (d instanceof Date) {
      return d.getFullYear() + '-' + _pad(d.getMonth()+1) + '-' + _pad(d.getDate());
    }
    return String(d).trim();
  };

  const groups = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const key = normCampaign(row[idxCampaign]) + '|' + String(row[idxAdSet]||'').trim() + '|' + String(row[idxAd]||'').trim() + '|' + normDate(row[idxDate]);
    const cost = parseFloat(String(row[idxCost]||'0').replace(/,/g, '')) || 0;
    if (!groups[key]) {
      groups[key] = {row: row.slice(), cost: cost};
    } else if (cost > groups[key].cost) {
      groups[key] = {row: row.slice(), cost: cost};
    }
  }

  // 정규화된 캠페인명으로 row 갱신
  const newRows = [];
  for (const k in groups) {
    const row = groups[k].row;
    row[idxCampaign] = normCampaign(row[idxCampaign]);
    newRows.push(row);
  }

  const beforeCount = data.length;
  const afterCount = newRows.length;

  // 시트 데이터 영역 모두 삭제 → 정규화된 row 일괄 setValues
  sheet.getRange(2, 1, lr - 1, lc).clearContent();
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, lc).setValues(newRows);
  }

  return {
    ok: true,
    before: beforeCount,
    after: afterCount,
    deleted: beforeCount - afterCount,
    keepRule: '캠페인명 정규화 후 같은 키 그룹 중 가장 cost 큰 row 보존'
  };
}

// === 분석 보고서 시트 — GitHub 발행 파일 기반 전체 재생성 (v1.1.11) ===
// daily/*.html + weekly/*.html → 일자 오름차순 정렬해서 R3부터 다시 채움
function _rebuildReportLog() {
  const props = PropertiesService.getScriptProperties();
  const pat = props.getProperty('GITHUB_PAT');
  const owner = props.getProperty('GITHUB_OWNER') || 'design337';
  const repo = props.getProperty('GITHUB_REPO') || 'tving-classaction-reports';
  if (!pat) return {ok: false, error: 'GITHUB_PAT not set'};

  const fetchDir = (path) => {
    const url = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
    const r = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {Authorization: 'Bearer ' + pat, Accept: 'application/vnd.github+json'},
      muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) return [];
    return JSON.parse(r.getContentText()).filter(f => f.type === 'file' && /\.html$/.test(f.name));
  };

  const dailyFiles = fetchDir('daily');
  const weeklyFiles = fetchDir('weekly');

  const entries = [];

  // 일간: daily/YYYY-MM-DD.html
  dailyFiles.forEach(f => {
    const m = f.name.match(/^(\d{4}-\d{2}-\d{2})\.html$/);
    if (!m) return;
    const date = m[1];
    const parts = date.split('-');
    const md = parseInt(parts[1]) + '/' + parseInt(parts[2]);
    entries.push({
      sortKey: date + '#A',
      날짜: date,
      종류: '일간',
      제목: CONFIG.PROJECT_NAME + ' 광고 성과 일간 분석 보고서 (' + md + ')',
      채널: '메타 / 네이버SA / 구글 / GA4',
      작성자: CONFIG.AUTHOR,
      링크: 'https://' + owner + '.github.io/' + repo + '/daily/' + f.name,
      비고: 'SUCCESS'
    });
  });

  // 주간: weekly/YYYY-MM-DD_YYYY-MM-DD.html
  weeklyFiles.forEach(f => {
    const m = f.name.match(/^(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.html$/);
    if (!m) return;
    const start = m[1], end = m[2];
    const s = start.split('-'), e = end.split('-');
    const range = parseInt(s[1]) + '/' + parseInt(s[2]) + '~' + parseInt(e[1]) + '/' + parseInt(e[2]);
    entries.push({
      sortKey: start + '#B',
      날짜: start + '~' + end,
      종류: '주간',
      제목: CONFIG.PROJECT_NAME + ' 광고 성과 주간 분석 보고서 (' + range + ')',
      채널: '메타 / 네이버SA / 구글 / GA4',
      작성자: CONFIG.AUTHOR,
      링크: 'https://' + owner + '.github.io/' + repo + '/weekly/' + f.name,
      비고: 'SUCCESS'
    });
  });

  // 정렬: sortKey 오름차순 (날짜 ASC, 일간(A) 먼저 주간(B) 나중)
  entries.sort((a, b) => a.sortKey < b.sortKey ? -1 : (a.sortKey > b.sortKey ? 1 : 0));

  const sheet = _ss().getSheetByName(CONFIG.REPORT_SHEET);
  if (!sheet) return {ok: false, error: 'report sheet not found'};

  // R3 이후 모두 삭제
  const lr = sheet.getLastRow();
  if (lr >= 3) sheet.deleteRows(3, lr - 2);

  // 다시 채우기 — 헤더 행(R2) 기준 컬럼 순서: 날짜/보고서 종류/보고서 제목/분석 채널/작성자/링크/비고
  if (entries.length === 0) {
    return {ok: true, dailyFound: dailyFiles.length, weeklyFound: weeklyFiles.length, entriesAdded: 0};
  }
  const rows = entries.map(e => [e.날짜, e.종류, e.제목, e.채널, e.작성자, e.링크, e.비고]);
  sheet.getRange(3, 1, rows.length, 7).setValues(rows);

  return {
    ok: true,
    dailyFound: dailyFiles.length,
    weeklyFound: weeklyFiles.length,
    entriesAdded: entries.length,
    firstRow: {date: entries[0].날짜, type: entries[0].종류, title: entries[0].제목, link: entries[0].링크},
    lastRow: {date: entries[entries.length-1].날짜, type: entries[entries.length-1].종류, title: entries[entries.length-1].제목, link: entries[entries.length-1].링크},
    finalLastRow: sheet.getLastRow()
  };
}

// === 분석 보고서 시트 — 링크 없는/잘못된 row 정리 (v1.1.11) ===
// R1 제목, R2 헤더, R3부터 데이터. F열(링크)이 http 시작 안 하면 빈 row로 간주하여 삭제.
function _cleanupReportLog() {
  const sheet = _ss().getSheetByName(CONFIG.REPORT_SHEET);
  if (!sheet) return {ok: false, error: 'sheet not found'};
  const lr = sheet.getLastRow();
  if (lr < 3) return {ok: true, note: 'nothing to clean', lastRow: lr};
  // 헤더 R2, 데이터 R3부터. F열 (col 6, 1-indexed) = 링크
  const dataRange = sheet.getRange(3, 1, lr - 2, 7).getValues();
  const rowsToDelete = [];
  let totalSeen = 0, kept = 0;
  for (let i = dataRange.length - 1; i >= 0; i--) {
    totalSeen++;
    const row = dataRange[i];
    const link = String(row[5] || '').trim();
    // 링크가 http로 시작하지 않으면 삭제 대상
    if (!link || !/^https?:\/\//i.test(link)) {
      rowsToDelete.push(3 + i);
    } else {
      kept++;
    }
  }
  // 역순 삭제 — Apps Script deleteRow는 단일 행 삭제
  for (const r of rowsToDelete) {
    sheet.deleteRow(r);
  }
  return {
    ok: true,
    beforeLastRow: lr,
    afterLastRow: sheet.getLastRow(),
    rowsScanned: totalSeen,
    rowsDeleted: rowsToDelete.length,
    rowsKept: kept
  };
}

function _logReportToSheet(reportType, startDate, endDate, url, title) {
  try {
    const sheet = _ss().getSheetByName(CONFIG.REPORT_SHEET);
    if (!sheet) {
      _appendRunLog('log_report', 'WARN', 'report sheet not found: ' + CONFIG.REPORT_SHEET, '');
      return;
    }
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const lastRow = Math.max(sheet.getLastRow(), 1);

    // 후보 매칭 — 사용자 시트 헤더 호환 (날짜/보고서 종류/보고서 제목/분석 채널/작성자/링크/비고)
    const candidates = {
      ts:    ['날짜', '발행 일시', '발행일시', '발행일', '발행 일자', '발행시각', 'timestamp', '일시', '시각', '시간'],
      type:  ['보고서 종류', '보고서 유형', '보고서유형', '유형', '구분', '타입', 'type'],
      start: ['기간 시작', '기간시작', '시작일', '시작 날짜', '시작', 'start'],
      end:   ['기간 종료', '기간종료', '종료일', '종료 날짜', '종료', 'end'],
      title: ['보고서 제목', '제목', 'title'],
      channel:['분석 채널', '채널', 'channel'],
      author:['작성자', 'author'],
      url:   ['링크', 'URL', '주소', 'url'],
      status:['비고', '상태', 'status']
    };

    // 헤더 자동 탐지 — 1~5행 중 '유형' OR '링크/URL/제목' 포함된 첫 행
    let headerRowIdx = 1;  // 1-based
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const headRows = sheet.getRange(1, 1, Math.min(lastRow, 5), lastCol).getValues();
    for (let i = 0; i < headRows.length; i++) {
      const row = headRows[i].map(c => String(c).trim());
      const allCands = []
        .concat(candidates.type, candidates.url, candidates.title, candidates.ts);
      if (row.some(c => allCands.indexOf(c) >= 0)) {
        headerRowIdx = i + 1;
        headers = row;
        break;
      }
    }

    const findIdx = (cands) => {
      for (const c of cands) {
        const i = headers.indexOf(c);
        if (i >= 0) return i;
      }
      return -1;
    };

    const idxTs = findIdx(candidates.ts);
    const idxType = findIdx(candidates.type);
    const idxStart = findIdx(candidates.start);
    const idxEnd = findIdx(candidates.end);
    const idxTitle = findIdx(candidates.title);
    const idxChannel = findIdx(candidates.channel);
    const idxAuthor = findIdx(candidates.author);
    const idxUrl = findIdx(candidates.url);
    const idxStatus = findIdx(candidates.status);

    // 헤더 누락 시 fallback
    if (idxTs < 0 && idxType < 0 && idxUrl < 0 && idxTitle < 0) {
      const fallbackRow = [new Date(), reportType, startDate, endDate, title, url, 'SUCCESS'];
      sheet.getRange(sheet.getLastRow()+1, 1, 1, fallbackRow.length).setValues([fallbackRow]);
      _appendRunLog('log_report', 'FALLBACK', 'headers not matched (row 1~5)', JSON.stringify(headers).slice(0, 250) + ' || headerRowIdx=' + headerRowIdx);
      return;
    }

    // 헤더 매칭 row 작성
    const dateLabel = (reportType === 'weekly')
      ? (startDate + '~' + endDate)
      : endDate;
    const newRow = new Array(headers.length).fill('');
    if (idxTs       >= 0) newRow[idxTs]       = dateLabel;
    if (idxType     >= 0) newRow[idxType]     = (reportType === 'weekly' ? '주간' : '일간');
    if (idxStart    >= 0) newRow[idxStart]    = startDate;
    if (idxEnd      >= 0) newRow[idxEnd]      = endDate;
    if (idxTitle    >= 0) newRow[idxTitle]    = title;
    if (idxChannel  >= 0) newRow[idxChannel]  = '메타 / 네이버SA / 구글 / GA4';
    if (idxAuthor   >= 0) newRow[idxAuthor]   = CONFIG.AUTHOR;
    if (idxUrl      >= 0) newRow[idxUrl]      = url;
    if (idxStatus   >= 0) newRow[idxStatus]   = 'SUCCESS';

    sheet.getRange(sheet.getLastRow()+1, 1, 1, headers.length).setValues([newRow]);
    _appendRunLog('log_report', 'SUCCESS', reportType + ' ' + startDate + '~' + endDate + ' headerRowIdx=' + headerRowIdx, url);
  } catch (err) {
    _appendRunLog('log_report', 'ERROR', String(err), (err.stack || '').slice(0, 300));
  }
}

function _appendRunLog(action, status, summary, detail) {
  try {
    const sheet = _ss().getSheetByName(CONFIG.RUNLOG_SHEET);
    if (!sheet) return;
    sheet.appendRow([new Date(), action, status, summary, detail]);
  } catch (e) {}
}

// ============================================================================
// Overview sheet auto-update (Fix 1) — patched for actual sheet layout
// ============================================================================
//
// 실제 종합지표 시트 레이아웃:
//   R1: 제목
//   R2: 【구역 1】 마케팅 KPI
//   R3: 헤더 — 총 기간 | 총 광고비 | 총 노출 | 총 클릭 | 평균 CTR | 평균 CPC |
//                평균 참여율 | 홈페이지 유입 | info 유입자 | 실 결제자 | 매출금액 | ROAS
//   R4: 단일 데이터 행 (전체 기간 합산)
//   R5: blank
//   R6: 【구역 2】 채널별 주간 지표
//   R7: 헤더 — col 0:채널 | 1:주차 | 2:기간 | 3:광고비 | 6:노출 | 9:클릭 |
//                12:CTR | 15:CPC | 18:CPM | 21:info유입 | 24:이탈율 | 27:참여율
//   R9~R17: 메타 0주차~8주차
//   R18: 메타 소계
//   R19~R27: 네이버SA 0주차~8주차
//   R28: 네이버SA 소계
//   R29~R37: 구글 0주차~8주차
//   R38: 구글 소계
//   R39: 전체 합계
//
//   주차 기간 텍스트: "X/Y(요일)~X/Y(요일)" (col 2). 2026년 가정.
//   KPI 값은 헤더 col+1 위치 (col 4, 7, 10, 13, 16, 19).

const _OVERVIEW = {
  zone1Row: 4,                     // 1-based row
  zone1Cols: {                     // 1-based col
    range: 1, cost: 2, imp: 3, clk: 4, ctr: 5, cpc: 6,
    engage: 7, home: 8, info: 9, pay: 10, rev: 11, roas: 12
  },
  zone2KpiCols: {                  // 1-based col offsets for value cells
    cost: 5, imp: 8, clk: 11, ctr: 14, cpc: 17, cpm: 20,
    info: 23, bounce: 26, engage: 29
  },
  zone2RowChannel: 1,              // 1-based col index for channel name
  zone2RowWeek: 2,
  zone2RowPeriod: 3,
  zone2DataStartRow: 9,
  zone2DataEndRow: 39
};

function _updateOverview(ss, range) {
  const sheet = ss.getSheetByName(CONFIG.OVERVIEW_SHEET);
  if (!sheet) {
    return {ok: false, error: 'overview sheet not found'};
  }
  const zone1Result = _writeZone1Fixed(ss, sheet, range);
  const zone2Result = _writeZone2Fixed(ss, sheet, range);
  // 셀 포맷 정규화 (v1.1.11) — % 잘못 잡힌 셀들 정상화
  let formatResult = {ok: false, note: 'skipped'};
  try {
    formatResult = _normalizeOverviewFormats(sheet);
  } catch (err) {
    formatResult = {ok: false, error: String(err)};
    _appendRunLog('normalize_formats', 'ERROR', String(err), (err.stack || '').slice(0, 200));
  }
  return {ok: true, zone1: zone1Result, zone2: zone2Result, formats: formatResult};
}

// === 셀 포맷 정규화 (v1.1.11) ===
// 종합지표 시트 zone1 R4 + zone2 R9~R39 셀 포맷 일괄 정리
//   - 광고비/CPC: ₩#,##0
//   - 노출/클릭/CPM/info유입: #,##0
//   - CTR/이탈률/참여율: 0.00%
function _normalizeOverviewFormats(sheet) {
  const NUM = '#,##0';
  const WON = '₩#,##0';
  const PCT = '0.00%';
  const DEC = '0.00';

  // Zone 1 R4 (행 1-indexed = 4)
  // _OVERVIEW.zone1Cols: range:1, cost:2, imp:3, clk:4, ctr:5, cpc:6, engage:7, home:8, info:9, pay:10, rev:11, roas:12
  const r1 = _OVERVIEW.zone1Row; // 4
  const z1 = [
    {col: 2,  fmt: WON},  // 총 광고비
    {col: 3,  fmt: NUM},  // 총 노출
    {col: 4,  fmt: NUM},  // 총 클릭
    {col: 5,  fmt: PCT},  // 평균 CTR
    {col: 6,  fmt: WON},  // 평균 CPC
    {col: 7,  fmt: PCT},  // 평균 참여율
    {col: 8,  fmt: NUM},  // 홈페이지 유입
    {col: 9,  fmt: NUM},  // info 유입자
    {col: 10, fmt: NUM},  // 실 결제자
    {col: 11, fmt: WON},  // 매출금액
    {col: 12, fmt: DEC}   // ROAS
  ];
  z1.forEach(w => sheet.getRange(r1, w.col).setNumberFormat(w.fmt));

  // Zone 2 R9~R39 (데이터 행 + 소계 + 합계)
  // _OVERVIEW.zone2KpiCols: cost:5, imp:8, clk:11, ctr:14, cpc:17, cpm:20, info:23, bounce:26, engage:29
  const startRow = _OVERVIEW.zone2DataStartRow; // 9
  const endRow = _OVERVIEW.zone2DataEndRow;     // 39
  const rowCount = endRow - startRow + 1;

  const colFormats = [
    {col: 5,  fmt: WON},  // 광고비
    {col: 8,  fmt: NUM},  // 노출
    {col: 11, fmt: NUM},  // 클릭
    {col: 14, fmt: PCT},  // CTR
    {col: 17, fmt: WON},  // CPC
    {col: 20, fmt: WON},  // CPM
    {col: 23, fmt: NUM},  // info유입 ← % 잘못된 거 정수로 정상화
    {col: 26, fmt: PCT},  // 이탈률
    {col: 29, fmt: PCT}   // 참여율
  ];
  colFormats.forEach(w => {
    sheet.getRange(startRow, w.col, rowCount, 1).setNumberFormat(w.fmt);
  });

  return {
    ok: true,
    zone1Cells: z1.length,
    zone2Columns: colFormats.length,
    zone2RowsAffected: rowCount,
    note: '광고비/CPC/CPM = ₩#,##0, 노출/클릭/info = #,##0, CTR/이탈률/참여율 = 0.00%'
  };
}

// === Zone 1 — 단일 row 전체 기간 합산 ===
function _writeZone1Fixed(ss, sheet, range) {
  const r = _OVERVIEW.zone1Row;
  const C = _OVERVIEW.zone1Cols;
  const agg = _aggregateByDate(ss, range.minDate, range.maxDate);
  const tot = agg.meta.cost + agg.naver.cost + agg.google.cost;
  const imp = agg.meta.imp + agg.naver.imp + agg.google.imp;
  const clk = agg.meta.clk + agg.naver.clk + agg.google.clk;
  const ctr = imp > 0 ? clk / imp : 0;
  const cpc = clk > 0 ? tot / clk : 0;
  const home = agg.ga.users || 0;
  // 옵션 A — info 유입자도 사용자 수 기준 (v1.1.11)
  // zone2 채널×주차 합(13,977 사용자, 매핑된 3채널만)과 zone1 I4(17,400 전체 채널 사용자)
  // 차이 3,423 = 매핑 안 된 채널(직접/오가닉/카카오 등) 사용자
  const info = agg.ga.users || 0;
  // 기간 텍스트 — yymmdd-yymmdd
  const yyMMdd = (d) => d.replace(/-/g,'').slice(2);
  const rangeText = yyMMdd(range.minDate) + '-' + yyMMdd(range.maxDate);

  // 가중 평균 참여율 — GA4_Raw 전체 세션 가중 (v1.1.11)
  const engage = (agg.gaSess && agg.gaSess.totalSessions > 0)
    ? agg.gaSess.totalEngaged / agg.gaSess.totalSessions : 0;
  const writes = [
    {col: C.range,  val: rangeText},
    {col: C.cost,   val: tot},
    {col: C.imp,    val: imp},
    {col: C.clk,    val: clk},
    {col: C.ctr,    val: ctr},
    {col: C.cpc,    val: cpc},
    {col: C.engage, val: engage},
    {col: C.home,   val: home},
    {col: C.info,   val: info}
    // pay / rev / roas — manual
  ];
  writes.forEach(w => sheet.getRange(r, w.col).setValue(w.val));
  return {
    ok: true, row: r, writes: writes.length, rangeText,
    totalCost: tot, totalClk: clk, engage: engage,
    gaSessionsTotal: agg.gaSess ? agg.gaSess.totalSessions : 0,
    gaEngagedTotal: agg.gaSess ? agg.gaSess.totalEngaged : 0
  };
}

// === Zone 2 — 채널×주차 매트릭스 ===
function _writeZone2Fixed(ss, sheet, range) {
  const startRow = _OVERVIEW.zone2DataStartRow;
  const endRow = _OVERVIEW.zone2DataEndRow;
  const data = sheet.getRange(startRow, 1, endRow - startRow + 1, 4).getValues();

  // Period 파싱: "X/Y(요일)~X/Y(요일)" → {start: 'YYYY-MM-DD', end: 'YYYY-MM-DD'}
  const parsePeriod = (txt) => {
    if (!txt) return null;
    const m = String(txt).match(/(\d+)\/(\d+).*?~\s*(\d+)\/(\d+)/);
    if (!m) return null;
    return {
      start: _periodToISO(parseInt(m[1]), parseInt(m[2])),
      end:   _periodToISO(parseInt(m[3]), parseInt(m[4]))
    };
  };

  const K = _OVERVIEW.zone2KpiCols;
  const channelAgg = {'메타': null, '네이버SA': null, '구글': null};
  const summary = [];
  let writeCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const sheetRow = startRow + i;
    const ch = String(row[0] || '').trim();
    const wk = String(row[1] || '').trim();
    const pd = parsePeriod(row[2]);
    if (!ch) continue;

    // 소계 / 합계 행 처리
    if (/소계/.test(ch) || ch === '전체 합계') {
      // 전체 기간 합산
      const agg = _aggregateByDate(ss, range.minDate, range.maxDate);
      let cost = 0, imp = 0, clk = 0, info = 0;
      if (/메타/.test(ch))      { cost = agg.meta.cost;   imp = agg.meta.imp;   clk = agg.meta.clk;   info = agg.ga.byChannel['메타'] || 0; }
      else if (/네이버/.test(ch)){ cost = agg.naver.cost;  imp = agg.naver.imp;  clk = agg.naver.clk;  info = agg.ga.byChannel['네이버SA'] || 0; }
      else if (/구글/.test(ch)) { cost = agg.google.cost; imp = agg.google.imp; clk = agg.google.clk; info = agg.ga.byChannel['구글'] || 0; }
      else if (/전체/.test(ch)) {
        cost = agg.meta.cost + agg.naver.cost + agg.google.cost;
        imp  = agg.meta.imp + agg.naver.imp + agg.google.imp;
        clk  = agg.meta.clk + agg.naver.clk + agg.google.clk;
        info = (agg.ga.byChannel['메타']||0) + (agg.ga.byChannel['네이버SA']||0) + (agg.ga.byChannel['구글']||0);
      }
      // 가중 평균 이탈률/참여율 — 채널별 또는 전체 (v1.1.11)
      let gsSessions = 0, gsEngaged = 0;
      const gsBC = agg.gaSess ? agg.gaSess.byChannel : {};
      if (/메타/.test(ch))      { gsSessions = (gsBC['메타']||{}).sessions||0;   gsEngaged = (gsBC['메타']||{}).engaged||0; }
      else if (/네이버/.test(ch)){ gsSessions = (gsBC['네이버SA']||{}).sessions||0; gsEngaged = (gsBC['네이버SA']||{}).engaged||0; }
      else if (/구글/.test(ch)) { gsSessions = (gsBC['구글']||{}).sessions||0; gsEngaged = (gsBC['구글']||{}).engaged||0; }
      else if (/전체/.test(ch)) {
        gsSessions = agg.gaSess ? agg.gaSess.totalSessions : 0;
        gsEngaged  = agg.gaSess ? agg.gaSess.totalEngaged : 0;
      }
      const engageR = gsSessions > 0 ? gsEngaged / gsSessions : 0;
      const bounceR = gsSessions > 0 ? 1 - engageR : 0;
      const ctr = imp > 0 ? clk / imp : 0;
      const cpc = clk > 0 ? cost / clk : 0;
      const cpm = imp > 0 ? cost / imp * 1000 : 0;
      const writes = [
        {col: K.cost, val: cost}, {col: K.imp, val: imp}, {col: K.clk, val: clk},
        {col: K.ctr, val: ctr}, {col: K.cpc, val: cpc}, {col: K.cpm, val: cpm},
        {col: K.info, val: info},
        {col: K.bounce, val: bounceR}, {col: K.engage, val: engageR}
      ];
      writes.forEach(w => sheet.getRange(sheetRow, w.col).setValue(w.val));
      writeCount += writes.length;
      summary.push({row: sheetRow, type: 'subtotal', channel: ch, writes: writes.length, sessions: gsSessions, engageRate: engageR});
      continue;
    }

    // 채널 + 주차 데이터 행 — period가 분석 기간과 겹치는지 확인
    if (!pd) continue;
    if (!(ch === '메타' || ch === '네이버SA' || ch === '구글')) continue;
    // 겹침 검사: 주차 [pd.start, pd.end] ∩ [range.minDate, range.maxDate] != ∅
    if (pd.end < range.minDate || pd.start > range.maxDate) continue;

    // 그 주차의 데이터를 채널별로 집계 (해당 주차 기간만)
    const agg = _aggregateByDate(ss, pd.start, pd.end);
    let cost = 0, imp = 0, clk = 0, info = 0;
    if (ch === '메타')      { cost = agg.meta.cost;   imp = agg.meta.imp;   clk = agg.meta.clk;   info = agg.ga.byChannel['메타'] || 0; }
    else if (ch === '네이버SA'){ cost = agg.naver.cost; imp = agg.naver.imp; clk = agg.naver.clk; info = agg.ga.byChannel['네이버SA'] || 0; }
    else if (ch === '구글') { cost = agg.google.cost; imp = agg.google.imp; clk = agg.google.clk; info = agg.ga.byChannel['구글'] || 0; }

    if (cost === 0 && imp === 0 && clk === 0 && info === 0) continue;  // 빈 주차는 건드리지 않음

    // 가중 평균 이탈률/참여율 — 채널별 (v1.1.11)
    const gsBC = agg.gaSess ? agg.gaSess.byChannel : {};
    let gsSessions = 0, gsEngaged = 0;
    if (ch === '메타')      { gsSessions = (gsBC['메타']||{}).sessions||0;   gsEngaged = (gsBC['메타']||{}).engaged||0; }
    else if (ch === '네이버SA'){ gsSessions = (gsBC['네이버SA']||{}).sessions||0; gsEngaged = (gsBC['네이버SA']||{}).engaged||0; }
    else if (ch === '구글') { gsSessions = (gsBC['구글']||{}).sessions||0; gsEngaged = (gsBC['구글']||{}).engaged||0; }
    const engageR = gsSessions > 0 ? gsEngaged / gsSessions : 0;
    const bounceR = gsSessions > 0 ? 1 - engageR : 0;
    const ctr = imp > 0 ? clk / imp : 0;
    const cpc = clk > 0 ? cost / clk : 0;
    const cpm = imp > 0 ? cost / imp * 1000 : 0;
    const writes = [
      {col: K.cost, val: cost}, {col: K.imp, val: imp}, {col: K.clk, val: clk},
      {col: K.ctr,  val: ctr},  {col: K.cpc, val: cpc}, {col: K.cpm, val: cpm},
      {col: K.info, val: info},
      {col: K.bounce, val: bounceR}, {col: K.engage, val: engageR}
    ];
    writes.forEach(w => sheet.getRange(sheetRow, w.col).setValue(w.val));
    writeCount += writes.length;
    summary.push({row: sheetRow, channel: ch, week: wk, period: pd.start + '~' + pd.end, writes: writes.length, sessions: gsSessions, engageRate: engageR});
  }

  return {ok: true, writes: writeCount, rows: summary};
}

function _periodToISO(m, d) {
  // 2026년 가정. m이 1~12 사이.
  return '2026-' + _pad(m) + '-' + _pad(d);
}

function _detectOverviewLayout(grid) {
  const layout = {
    zone1HeaderRow: -1,
    zone1Cols: {},      // {standardKey: colIdx(0-based)}
    zone1DateColIdx: -1,
    zone1DataStartRow: -1,
    zone1DataEndRow: -1,

    zone2HeaderRow: -1,
    zone2WeekCols: {},  // {weekKey: colIdx}
    zone2ChannelRows: {}, // {channelKey: rowIdx}
    zone2KpiRows: []    // [{rowIdx, kpiKey, channelKey}]
  };

  // 헤더 키 후보 매핑
  const kpiAliases = {
    cost:    ['총광고비', '광고비', '비용', '총비용', '집행비', '총 비용', '총 광고비'],
    imp:     ['총노출', '노출', '노출수', '총 노출'],
    clk:     ['총클릭', '클릭', '클릭수', '총 클릭'],
    ctr:     ['평균CTR', 'CTR', '평균 CTR', '클릭률', '클릭률(%)'],
    cpc:     ['평균CPC', 'CPC', '평균 CPC'],
    home:    ['홈페이지유입자', '홈페이지 유입자', '홈페이지유입', '홈 유입', '홈페이지'],
    info:    ['info유입자', 'info 유입', 'info유입', 'Info 유입', 'INFO유입'],
    pay:     ['실결제자', '결제자', '결제', '실 결제자'],
    rev:     ['매출금액', '매출', '매출 금액'],
    roas:    ['ROAS', 'Roas', 'roas'],
    memo:    ['메모', 'NOTE', 'Note', '비고'],
    date:    ['날짜', '일자', '일']
  };

  // 구역1 헤더 탐색 — '날짜' + (총광고비 or 광고비) 둘 다 포함된 행
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(c => String(c == null ? '' : c).trim());
    const hasDate = row.some(c => kpiAliases.date.indexOf(c) >= 0);
    const hasCost = row.some(c => kpiAliases.cost.indexOf(c) >= 0);
    if (hasDate && hasCost) {
      layout.zone1HeaderRow = r;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (!cell) continue;
        for (const key in kpiAliases) {
          if (kpiAliases[key].indexOf(cell) >= 0) {
            if (key === 'date') layout.zone1DateColIdx = c;
            layout.zone1Cols[key] = c;
            break;
          }
        }
      }
      break;
    }
  }

  // 구역1 데이터 영역 — 헤더 다음 행부터 첫 빈 행 또는 다른 헤더 행 만나기 전까지
  if (layout.zone1HeaderRow >= 0) {
    layout.zone1DataStartRow = layout.zone1HeaderRow + 1;
    for (let r = layout.zone1DataStartRow; r < grid.length; r++) {
      const row = grid[r].map(c => String(c == null ? '' : c).trim());
      // 빈 행이거나 새 헤더(주차 포함)면 종료
      const allEmpty = row.every(c => c === '');
      const hasWeek = row.some(c => /^\s*\d+\s*주차\s*$/.test(c) || c === '소계' || c === '주차');
      if (allEmpty || hasWeek) {
        layout.zone1DataEndRow = r - 1;
        break;
      }
      layout.zone1DataEndRow = r;
    }
  }

  // 구역2 헤더 행 — '주차' 패턴 포함된 첫 행
  const weekPattern = /^\s*(\d+)\s*주차\s*$/;
  for (let r = (layout.zone1DataEndRow >= 0 ? layout.zone1DataEndRow + 1 : 0); r < grid.length; r++) {
    const row = grid[r].map(c => String(c == null ? '' : c).trim());
    let weekCount = 0;
    const weekColsLocal = {};
    let subtotalCol = -1;
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      const m = cell.match(weekPattern);
      if (m) {
        weekColsLocal[m[1] + '주차'] = c;
        weekCount++;
      } else if (cell === '소계' || cell === '합계' || cell === '총계') {
        subtotalCol = c;
      }
    }
    if (weekCount >= 2) {
      layout.zone2HeaderRow = r;
      layout.zone2WeekCols = weekColsLocal;
      if (subtotalCol >= 0) layout.zone2WeekCols['소계'] = subtotalCol;
      break;
    }
  }

  // 구역2 채널 행 + KPI 행 매핑
  if (layout.zone2HeaderRow >= 0) {
    const channelAliases = {
      '메타':    ['메타', 'Meta', 'meta', '페이스북', '페이스북 / 인스타'],
      '네이버SA':['네이버SA', '네이버', '네이버 SA', 'NaverSA'],
      '구글':    ['구글', 'Google', 'google']
    };
    let currentChannel = null;
    for (let r = layout.zone2HeaderRow + 1; r < grid.length; r++) {
      const row = grid[r].map(c => String(c == null ? '' : c).trim());
      // 채널 헤더 행 — 첫 두 컬럼 중 하나가 채널명
      let chFound = null;
      for (let cc = 0; cc < Math.min(3, row.length); cc++) {
        const cell = row[cc];
        for (const ch in channelAliases) {
          if (channelAliases[ch].indexOf(cell) >= 0) {
            chFound = ch;
            break;
          }
        }
        if (chFound) break;
      }
      if (chFound) {
        currentChannel = chFound;
        layout.zone2ChannelRows[chFound] = r;
        continue;
      }
      // KPI 행 — 채널 컨텍스트 하에 KPI 라벨 매칭
      if (currentChannel) {
        let kpiKey = null;
        for (let cc = 0; cc < Math.min(3, row.length); cc++) {
          const cell = row[cc];
          if (!cell) continue;
          for (const key in kpiAliases) {
            if (key === 'date') continue;
            if (kpiAliases[key].indexOf(cell) >= 0) {
              kpiKey = key;
              break;
            }
          }
          if (kpiKey) break;
        }
        if (kpiKey) {
          layout.zone2KpiRows.push({rowIdx: r, channelKey: currentChannel, kpiKey: kpiKey});
        }
      }
    }
  }

  return layout;
}

function _writeZone1(ss, sheet, grid, layout, range) {
  if (layout.zone1HeaderRow < 0 || layout.zone1DateColIdx < 0) {
    return {ok: false, error: 'zone1 header not detected'};
  }
  const dateColIdx = layout.zone1DateColIdx;
  const cols = layout.zone1Cols;

  // 데이터 영역의 일자 → 행 인덱스 매핑
  const dateRowMap = {};
  for (let r = layout.zone1DataStartRow; r <= layout.zone1DataEndRow; r++) {
    if (r < 0 || r >= grid.length) continue;
    const cellVal = grid[r][dateColIdx];
    const d = _normDate(cellVal);
    if (d) dateRowMap[d] = r;
  }

  const updates = [];
  let nextEmptyRow = -1;
  for (let r = layout.zone1DataStartRow; r <= layout.zone1DataEndRow + 30; r++) {
    if (r >= grid.length || (grid[r] && String(grid[r][dateColIdx] || '').trim() === '')) {
      nextEmptyRow = r;
      break;
    }
  }

  for (const date of range.dates) {
    const agg = _aggregateByDate(ss, date, date);
    const tot = agg.meta.cost + agg.naver.cost + agg.google.cost;
    const totImp = agg.meta.imp + agg.naver.imp + agg.google.imp;
    const totClk = agg.meta.clk + agg.naver.clk + agg.google.clk;
    const ctr = totImp > 0 ? (totClk / totImp * 100) : 0;
    const cpc = totClk > 0 ? (tot / totClk) : 0;

    const values = {
      date: date,
      cost: tot,
      imp:  totImp,
      clk:  totClk,
      ctr:  ctr / 100,  // percent → 0.xx (sheet may format)
      cpc:  cpc,
      home: agg.ga.users || 0,
      info: agg.ga.events || 0,
      pay:  0,
      rev:  0,
      roas: 0,
      memo: ''
    };

    let targetRow;
    if (dateRowMap[date] !== undefined) {
      targetRow = dateRowMap[date];
    } else {
      // append at next empty row
      if (nextEmptyRow < 0) nextEmptyRow = sheet.getLastRow();
      targetRow = nextEmptyRow;
      nextEmptyRow++;
      dateRowMap[date] = targetRow;
    }

    // 행 갱신 — 매핑된 컬럼만
    const writes = [];
    for (const key of Object.keys(values)) {
      if (cols[key] !== undefined) {
        writes.push({row: targetRow + 1, col: cols[key] + 1, value: values[key]});
      }
    }
    updates.push({date, row: targetRow + 1, writes: writes.length});
    writes.forEach(w => {
      if (w.value !== null && w.value !== undefined && w.value !== '') {
        sheet.getRange(w.row, w.col).setValue(w.value);
      }
    });
  }

  return {ok: true, datesUpdated: updates.length, updates: updates.slice(0, 30)};
}

function _writeZone2(ss, sheet, grid, layout, range) {
  if (layout.zone2HeaderRow < 0 || Object.keys(layout.zone2WeekCols).length === 0) {
    return {ok: false, error: 'zone2 header not detected'};
  }
  if (layout.zone2KpiRows.length === 0) {
    return {ok: false, error: 'no kpi rows detected', channels: Object.keys(layout.zone2ChannelRows), weeks: Object.keys(layout.zone2WeekCols)};
  }

  const campaignStart = new Date(CONFIG.CAMPAIGN_START);
  campaignStart.setHours(0,0,0,0);

  // 각 일자를 주차로 매핑
  const dayToWeek = (date) => {
    const d = new Date(date);
    d.setHours(0,0,0,0);
    const diff = Math.floor((d - campaignStart) / 86400000);
    if (diff < 0) return -1;
    return Math.floor(diff / 7); // 0주차, 1주차, ...
  };

  // 주차별 일자 범위 — minDate~maxDate 포함된 주차 모두
  const weeksToProcess = new Set();
  for (const d of range.dates) {
    const w = dayToWeek(d);
    if (w >= 0) weeksToProcess.add(w);
  }

  // 각 주차에 대해 채널×KPI 집계
  let writeCount = 0;
  const summary = [];
  for (const weekNum of weeksToProcess) {
    const wkKey = weekNum + '주차';
    const colIdx = layout.zone2WeekCols[wkKey];
    if (colIdx === undefined) continue;

    // 주차의 시작/끝 날짜
    const wkStart = new Date(campaignStart.getTime() + weekNum * 7 * 86400000);
    const wkEnd = new Date(wkStart.getTime() + 6 * 86400000);
    const wkStartStr = _fmt(wkStart);
    const wkEndStr = _fmt(wkEnd);

    const agg = _aggregateByDate(ss, wkStartStr, wkEndStr);

    const channelKPIs = {
      '메타': {
        cost: agg.meta.cost,
        imp:  agg.meta.imp,
        clk:  agg.meta.clk,
        ctr:  agg.meta.imp > 0 ? agg.meta.clk / agg.meta.imp : 0,
        cpc:  agg.meta.clk > 0 ? agg.meta.cost / agg.meta.clk : 0,
        home: 0, info: 0, pay: 0, rev: 0, roas: 0
      },
      '네이버SA': {
        cost: agg.naver.cost,
        imp:  agg.naver.imp,
        clk:  agg.naver.clk,
        ctr:  agg.naver.imp > 0 ? agg.naver.clk / agg.naver.imp : 0,
        cpc:  agg.naver.clk > 0 ? agg.naver.cost / agg.naver.clk : 0,
        home: agg.ga.byChannel['네이버SA'] || 0,
        info: 0, pay: 0, rev: 0, roas: 0
      },
      '구글': {
        cost: agg.google.cost,
        imp:  agg.google.imp,
        clk:  agg.google.clk,
        ctr:  agg.google.imp > 0 ? agg.google.clk / agg.google.imp : 0,
        cpc:  agg.google.clk > 0 ? agg.google.cost / agg.google.clk : 0,
        home: agg.ga.byChannel['구글'] || 0,
        info: 0, pay: 0, rev: 0, roas: 0
      }
    };
    channelKPIs['메타'].home = agg.ga.byChannel['메타'] || 0;

    for (const kpiRow of layout.zone2KpiRows) {
      const ch = kpiRow.channelKey;
      const k  = kpiRow.kpiKey;
      const value = channelKPIs[ch] && channelKPIs[ch][k] !== undefined ? channelKPIs[ch][k] : null;
      if (value === null) continue;
      sheet.getRange(kpiRow.rowIdx + 1, colIdx + 1).setValue(value);
      writeCount++;
    }
    summary.push({week: wkKey, start: wkStartStr, end: wkEndStr, col: colIdx + 1});
  }

  // 소계 컬럼 — minDate ~ maxDate 누적
  if (layout.zone2WeekCols['소계'] !== undefined) {
    const colIdx = layout.zone2WeekCols['소계'];
    const agg = _aggregateByDate(ss, range.minDate, range.maxDate);
    const channelKPIs = {
      '메타': {
        cost: agg.meta.cost, imp: agg.meta.imp, clk: agg.meta.clk,
        ctr:  agg.meta.imp > 0 ? agg.meta.clk / agg.meta.imp : 0,
        cpc:  agg.meta.clk > 0 ? agg.meta.cost / agg.meta.clk : 0,
        home: agg.ga.byChannel['메타'] || 0, info: 0, pay: 0, rev: 0, roas: 0
      },
      '네이버SA': {
        cost: agg.naver.cost, imp: agg.naver.imp, clk: agg.naver.clk,
        ctr:  agg.naver.imp > 0 ? agg.naver.clk / agg.naver.imp : 0,
        cpc:  agg.naver.clk > 0 ? agg.naver.cost / agg.naver.clk : 0,
        home: agg.ga.byChannel['네이버SA'] || 0, info: 0, pay: 0, rev: 0, roas: 0
      },
      '구글': {
        cost: agg.google.cost, imp: agg.google.imp, clk: agg.google.clk,
        ctr:  agg.google.imp > 0 ? agg.google.clk / agg.google.imp : 0,
        cpc:  agg.google.clk > 0 ? agg.google.cost / agg.google.clk : 0,
        home: agg.ga.byChannel['구글'] || 0, info: 0, pay: 0, rev: 0, roas: 0
      }
    };
    for (const kpiRow of layout.zone2KpiRows) {
      const ch = kpiRow.channelKey;
      const k  = kpiRow.kpiKey;
      const value = channelKPIs[ch] && channelKPIs[ch][k] !== undefined ? channelKPIs[ch][k] : null;
      if (value === null) continue;
      sheet.getRange(kpiRow.rowIdx + 1, colIdx + 1).setValue(value);
      writeCount++;
    }
    summary.push({week: '소계', start: range.minDate, end: range.maxDate, col: colIdx + 1});
  }

  return {ok: true, weeks: summary, writes: writeCount};
}

// ============================================================================
// Inspect (debug)
// ============================================================================

function _inspectSheets(body) {
  const ss = _ss();
  const overview = ss.getSheetByName(CONFIG.OVERVIEW_SHEET);
  const report = ss.getSheetByName(CONFIG.REPORT_SHEET);
  const runlog = ss.getSheetByName(CONFIG.RUNLOG_SHEET);

  const out = {ok: true};

  if (overview) {
    const lr = Math.max(overview.getLastRow(), 1);
    const lc = Math.max(overview.getLastColumn(), 1);
    const rowsToRead = Math.min(lr, 50);
    const colsToRead = Math.min(lc, 30);
    const grid = overview.getRange(1, 1, rowsToRead, colsToRead).getValues();
    // Convert grid to compact representation (only non-empty cells)
    const compact = [];
    for (let r = 0; r < grid.length; r++) {
      const rowOut = {};
      for (let c = 0; c < grid[r].length; c++) {
        const v = grid[r][c];
        if (v !== '' && v !== null && v !== undefined) {
          rowOut[c] = (v instanceof Date) ? _fmt(v) : String(v).slice(0, 80);
        }
      }
      if (Object.keys(rowOut).length > 0) {
        compact.push({row: r + 1, cells: rowOut});
      }
    }
    // 전체 컬럼 (빈 셀 포함) — R7 헤더와 R15/R25/R35 데이터 행
    const fullRowsTarget = [7, 8, 15, 16, 17, 25, 26, 27, 35, 36, 37, 18, 28, 38, 39, 4];
    const fullRows = [];
    // displayValues — 사용자 시트 화면에 보이는 그대로 (포맷 적용된 문자열)
    const displayGrid = overview.getRange(1, 1, rowsToRead, colsToRead).getDisplayValues();
    for (const targetRow of fullRowsTarget) {
      if (targetRow > rowsToRead) continue;
      const row = grid[targetRow - 1] || [];
      const drow = displayGrid[targetRow - 1] || [];
      const cells = [];
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        const label = (v === '' || v === null || v === undefined) ? '∅' :
          ((v instanceof Date) ? _fmt(v) : String(v).slice(0, 40));
        const disp = drow[c] === '' || drow[c] === null ? '∅' : String(drow[c]).slice(0, 40);
        cells.push({col: c, raw: label, disp: disp});
      }
      fullRows.push({row: targetRow, cells});
    }
    out.overview = {
      sheetName: CONFIG.OVERVIEW_SHEET,
      lastRow: lr,
      lastCol: lc,
      grid: compact,
      fullRows: fullRows
    };
  } else {
    out.overview = {error: 'sheet not found: ' + CONFIG.OVERVIEW_SHEET};
  }

  // GA4 시트 헤더 + 첫 3행 + date 파라미터의 raw rows dump
  const gaSheetNames = ['GA4_Raw', 'GA4_티빙 집단소송 신청하기'];
  out.gaSheets = {};
  for (const name of gaSheetNames) {
    const s = ss.getSheetByName(name);
    if (!s) { out.gaSheets[name] = {error: 'not found'}; continue; }
    const lc = Math.max(s.getLastColumn(), 1);
    const lr = Math.max(s.getLastRow(), 1);
    const headers = s.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h));
    const sampleN = Math.min(3, lr - 1);
    const sample = sampleN > 0 ? s.getRange(2, 1, sampleN, lc).getValues() : [];
    out.gaSheets[name] = {
      lastRow: lr,
      lastCol: lc,
      headers: headers,
      sampleRows: sample.map(row => row.map(v => (v instanceof Date) ? _fmt(v) : String(v).slice(0, 50)))
    };
  }

  // 전체 기간(CAMPAIGN_START~today) GA 채널 분포 — v1.1.11 매핑 안 된 채널 정체 분석
  try {
    const today = _fmt(new Date());
    const agg = _aggregateByDate(ss, CONFIG.CAMPAIGN_START, today);
    out.gaFullCampaign = {
      range: CONFIG.CAMPAIGN_START + '~' + today,
      ga_users_total: agg.ga.users,
      ga_events_total: agg.ga.events,
      ga_byChannel: agg.ga.byChannel,
      mapped3Channels_sum: (agg.ga.byChannel['메타']||0) + (agg.ga.byChannel['네이버SA']||0) + (agg.ga.byChannel['구글']||0),
      unmapped_diff: agg.ga.users - ((agg.ga.byChannel['메타']||0) + (agg.ga.byChannel['네이버SA']||0) + (agg.ga.byChannel['구글']||0))
    };
  } catch (err) {
    out.gaFullCampaign = {error: String(err)};
  }

  // 분석 보고서 — 헤더 자동 탐지 (1~5행 중 '유형' + '링크/URL/제목' 동시 포함된 행)
  if (report) {
    const lc = Math.max(report.getLastColumn(), 1);
    const lr = Math.max(report.getLastRow(), 1);
    const headRows = report.getRange(1, 1, Math.min(lr, 5), lc).getValues();
    let detectedHeaderRow = -1;
    for (let i = 0; i < headRows.length; i++) {
      const row = headRows[i].map(c => String(c));
      const hasType = row.some(c => /유형|타입|type/i.test(c));
      const hasUrl = row.some(c => /URL|링크|주소/i.test(c)) || row.some(c => /제목|title/i.test(c));
      if (hasType || hasUrl) {
        detectedHeaderRow = i + 1;
        break;
      }
    }
    // last 5 rows
    const lastN = Math.min(5, lr);
    const lastRows = lr > 0 ? report.getRange(Math.max(1, lr - lastN + 1), 1, lastN, lc).getValues() : [];
    out.report = {
      sheetName: CONFIG.REPORT_SHEET,
      lastRow: lr,
      lastCol: lc,
      detectedHeaderRow: detectedHeaderRow,
      firstRows: headRows.map((row, i) => ({row: i + 1, cells: row.map(c => String(c).slice(0, 60))})),
      lastRows: lastRows.map((row, i) => ({row: lr - lastN + 1 + i, cells: row.map(c => String(c).slice(0, 60))}))
    };
  } else {
    out.report = {error: 'sheet not found: ' + CONFIG.REPORT_SHEET};
  }

  if (runlog) {
    const lc = Math.max(runlog.getLastColumn(), 1);
    const lr = Math.max(runlog.getLastRow(), 1);
    const headers = runlog.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h));
    const lastN = Math.min(15, lr);
    const lastRows = lr > 0 ? runlog.getRange(Math.max(1, lr - lastN + 1), 1, lastN, lc).getValues() : [];
    out.runlog = {
      sheetName: CONFIG.RUNLOG_SHEET,
      headers: headers,
      lastRow: lr,
      lastRows: lastRows.map((row, i) => ({row: lr - lastN + 1 + i, cells: row.map(c => (c instanceof Date) ? _fmt(c) : String(c).slice(0, 80))}))
    };
  }

  // GA4 source/medium 분포 (param: date)
  const dateParam = body && body.date;
  if (dateParam) {
    const gaSheet = ss.getSheetByName('GA4_티빙 집단소송 신청하기');
    if (gaSheet) {
      const rows = _readSheetRows(ss, 'GA4_티빙 집단소송 신청하기');
      const dist = {};
      const mapped = {};
      for (const r of rows) {
        const d = _normDate(r['날짜']);
        if (d !== dateParam) continue;
        const sm = String(r['세션 소스/매체'] || '');
        const users = parseFloat(String(r['총 사용자']).replace(/,/g,'')) || 0;
        dist[sm] = (dist[sm] || 0) + users;
        // 현재 매핑 결과
        let ch = 'Other';
        if (/naver/i.test(sm)) ch = '네이버SA';
        else if (/meta|facebook|fb|instagram|ig_/i.test(sm)) ch = '메타';
        else if (/google.*cpc|gad|adwords/i.test(sm)) ch = '구글';
        else if (/kakao|kkt/i.test(sm)) ch = '카카오';
        else if (/direct|\(direct\)|\(none\)|\(not set\)|not.set|data.not.available/i.test(sm)) ch = '직접';
        else if (/organic|google.*organic|naver.*organic/i.test(sm)) ch = '오가닉';
        else if (/referral/i.test(sm)) ch = '리퍼럴';
        else if (/ig|social/i.test(sm)) ch = '메타';
        else if (/chatgpt|ai-assistant/i.test(sm)) ch = 'AI어시스턴트';
        else if (/home/i.test(sm)) ch = '자체유입';
        mapped[ch] = (mapped[ch] || 0) + users;
      }
      out.ga = {date: dateParam, totalSources: Object.keys(dist).length, sourceMediumDistribution: dist, currentChannelMapping: mapped};
    }
  }

  return out;
}

// ============================================================================
// Status / Helper
// ============================================================================

function _status() {
  const ss = _ss();
  const sheets = {};
  for (const name in CONFIG.RAW_SHEETS) {
    const s = ss.getSheetByName(name);
    sheets[name] = s ? {exists: true, rows: s.getLastRow() - 1, cols: s.getLastColumn()} : {exists: false};
  }
  const extra = [CONFIG.HISTORY_SHEET, CONFIG.REPORT_SHEET, CONFIG.RUNLOG_SHEET, CONFIG.OVERVIEW_SHEET];
  for (const name of extra) {
    const s = ss.getSheetByName(name);
    sheets[name] = s ? {exists: true, rows: s.getLastRow(), cols: s.getLastColumn()} : {exists: false};
  }
  return {
    ok: true,
    project: CONFIG.PROJECT_NAME,
    spreadsheet: ss.getName(),
    sheetId: ss.getId(),
    version: 'v1.1.11',
    sheets,
    historyEvents: _readMarketingHistory().length
  };
}

function _ss() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('SHEET_ID not set in Script Properties');
  return SpreadsheetApp.openById(id);
}

// ============================================================================
// Build report from existing sheet data only (without new data input)
// ============================================================================

function _buildReportFromSheet(body) {
  const ss = _ss();
  const history = _readMarketingHistory();
  const meta = body.meta || {};
  const dates = meta.data_dates || [];
  const reportType = meta.report_type || 'daily';

  const results = [];
  if (reportType === 'daily') {
    for (const d of dates) {
      results.push(_publishDailyReport(ss, d, history));
    }
  } else if (reportType === 'weekly' && dates.length >= 2) {
    results.push(_publishWeeklyReport(ss, dates[0], dates[dates.length-1], history));
  }

  // 종합지표 갱신 — zone1은 항상 CAMPAIGN_START ~ maxDate 전체 기간
  let overviewResult = {ok: false, note: 'skipped'};
  if (dates.length > 0) {
    try {
      const sortedDates = dates.slice().sort();
      const fullRangeStart = CONFIG.CAMPAIGN_START;
      const maxDate = sortedDates[sortedDates.length - 1];
      const range = {
        dates: sortedDates,
        minDate: fullRangeStart,
        maxDate: maxDate,
        days: 999
      };
      overviewResult = _updateOverview(ss, range);
    } catch (err) {
      overviewResult = {ok: false, error: String(err)};
      _appendRunLog('build_report_overview', 'ERROR', String(err), err.stack || '');
    }
  }

  return {ok: true, results, overview: overviewResult};
}
