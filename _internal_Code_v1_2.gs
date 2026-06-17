/**
 * ============================================================================
 * 티빙 개인정보 유출 집단소송 — 광고 데이터 통합 처리 + 전략 보고서 발행
 * Single-file Apps Script (Code.gs)
 * Version: v1.1.12 (diag + header-detect + full-range zone1, 2026-06-16)
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
  return _json({ok: true, info: 'Tving classaction reports', version: 'v1.1.12'});
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
    // GA4_티빙 집단소송 신청하기 = 신청 의도자 전수 (모든 채널 — 카카오/오가닉/리퍼럴/직접/AI/광고)
    // GA4_Raw = 광고 캠페인 트래픽 추적 (티빙 캠페인 매칭만, 필터 유지)
    const skipFilterForGA = sheetName === 'GA4_티빙 집단소송 신청하기';
    const filtered = skipFilterForGA ? incoming.slice() : incoming.filter(r => _isProjectCampaign(r));

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

  // GA4_Raw — 사이트 전반 세션/참여율 — 채널×주차 가중 평균 이탈률/참여율 계산용 (v1.1.12)
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
    {name: '메타', cost: agg.meta.cost, clk: agg.meta.clk, imp: agg.meta.imp},
    {name: '네이버SA', cost: agg.naver.cost, clk: agg.naver.clk, imp: agg.naver.imp},
    {name: '구글', cost: agg.google.cost, clk: agg.google.clk, imp: agg.google.imp}
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




// ============================================================================
// Publish reports
// ============================================================================



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

// === 메타 시트 active 필터링 (v1.1.12) ===
// cost > 0 OR 노출 > 0 인 row만 보존 (실제 운영된 광고)
function _filterActiveMeta() {
  const sheet = _ss().getSheetByName('메타_광고');
  if (!sheet) return {ok: false, error: 'sheet not found'};
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2) return {ok: true, note: 'empty'};

  const data = sheet.getRange(1, 1, lr, lc).getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const idxCost = headers.indexOf('지출 금액 (KRW)');
  const idxImp = headers.indexOf('노출');
  const idxDate = headers.indexOf('일');
  if (idxCost < 0 || idxImp < 0) return {ok: false, error: 'cost/imp column not found'};

  const num = v => {
    if (v === null || v === undefined || v === '') return 0;
    const s = String(v).replace(/,/g,'');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  const active = rows.filter(r => num(r[idxCost]) > 0 || num(r[idxImp]) > 0);

  const dateDist = {};
  for (const r of active) {
    const d = _normDate(r[idxDate]);
    dateDist[d] = (dateDist[d] || 0) + 1;
  }

  sheet.getRange(2, 1, lr - 1, lc).clearContent();
  if (active.length > 0) {
    sheet.getRange(2, 1, active.length, lc).setValues(active);
  }

  return {
    ok: true,
    before: rows.length,
    after: active.length,
    removed: rows.length - active.length,
    rule: 'cost > 0 OR 노출 > 0',
    dateDistribution: dateDist
  };
}

// === 메타 시트 정규화 + 모든 시트 일자 정렬 (v1.1.12) ===
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

  // 메타 정규화 후 다른 시트들도 일자 정렬
  let sortResult;
  try { sortResult = _sortAllSheetsByDate(); } catch (e) { sortResult = {error: String(e)}; }

  return {
    ok: true,
    before: rows.length,
    after: cleaned.length,
    removedEmpty,
    removedNoise,
    normalizedCampaign: NORM_CAMP,
    dateRange: minDate + '~' + maxDate,
    sortAllSheets: sortResult
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

// === 메타_광고 시트 dedupe (v1.1.12) ===
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

// === 분석 보고서 시트 — GitHub 발행 파일 기반 전체 재생성 (v1.1.12) ===
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

// === 분석 보고서 시트 — 링크 없는/잘못된 row 정리 (v1.1.12) ===
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
  // 셀 포맷 정규화 (v1.1.12) — % 잘못 잡힌 셀들 정상화
  let formatResult = {ok: false, note: 'skipped'};
  try {
    formatResult = _normalizeOverviewFormats(sheet);
  } catch (err) {
    formatResult = {ok: false, error: String(err)};
    _appendRunLog('normalize_formats', 'ERROR', String(err), (err.stack || '').slice(0, 200));
  }
  return {ok: true, zone1: zone1Result, zone2: zone2Result, formats: formatResult};
}

// === 셀 포맷 정규화 (v1.1.12) ===
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
  // v1.2.1.3 — zone1은 종합. info유입 = GA 사용자 중 직접 제외 = 전체 - 직접
  // home = ga.users 그대로 (홈유입 — 모든 사용자)
  // info = 직접 제외 사용자 = 신청자(직접 제외)
  const home = agg.ga.users || 0;
  const directU = (agg.ga.byChannel && agg.ga.byChannel['직접']) || 0;
  const info = Math.max(0, (agg.ga.users || 0) - directU);
  // 기간 텍스트 — yymmdd-yymmdd
  const yyMMdd = (d) => d.replace(/-/g,'').slice(2);
  const rangeText = yyMMdd(range.minDate) + '-' + yyMMdd(range.maxDate);

  // 가중 평균 참여율 — GA4_Raw 전체 세션 가중 (v1.1.12)
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
      // 가중 평균 이탈률/참여율 — 채널별 또는 전체 (v1.1.12)
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

    // 가중 평균 이탈률/참여율 — 채널별 (v1.1.12)
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

  // 전체 기간(CAMPAIGN_START~today) GA 채널 분포 — v1.1.12 매핑 안 된 채널 정체 분석
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
    version: 'v1.1.12',
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


// ============================================================================
// v1.2.0 — DoD/WoW + Kakao + Device + Meta 광고세트 분석 헬퍼
// ============================================================================

// DoD: 전일 데이터 재집계
function _aggregatePreviousDate(ss, dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() - 1);
  const prev = _fmt(d);
  return _aggregateByDate(ss, prev, prev);
}

// WoW: 전주 데이터 재집계 (weekStart, weekEnd 동일 기간 만큼 이전)
function _aggregatePreviousWeek(ss, startStr, endStr) {
  const s = new Date(startStr + 'T00:00:00');
  const e = new Date(endStr + 'T00:00:00');
  s.setDate(s.getDate() - 7);
  e.setDate(e.getDate() - 7);
  return _aggregateByDate(ss, _fmt(s), _fmt(e));
}

// 키워드 변동 (curr vs prev)
function _keywordDiff(currKw, prevKw) {
  const cMap = {};
  currKw.forEach(k => { cMap[k.kw] = k; });
  const pMap = {};
  (prevKw || []).forEach(k => { pMap[k.kw] = k; });
  const out = currKw.map(c => {
    const p = pMap[c.kw];
    let dod = '유지';
    let dodCls = 'p2';
    if (!p) { dod = '▲ 신규 등장'; dodCls = 'p1'; }
    else {
      const dClk = p.clk > 0 ? Math.round((c.clk - p.clk) / p.clk * 100) : (c.clk > 0 ? 999 : 0);
      if (Math.abs(dClk) >= 10) {
        dod = (dClk > 0 ? '▲ +' : '▼ ') + dClk + '%';
        dodCls = dClk > 0 ? 'p1' : 'p3';
      }
    }
    return Object.assign({}, c, {dod, dodCls});
  });
  return out;
}

// 카카오 source별 신청자 분석 (재방문 추정)
function _kakaoReturnAnalysis(ss, startDate, endDate) {
  const num = v => { const n = parseFloat(String(v||'').replace(/[,%\s₩]/g,'')); return isNaN(n) ? 0 : n; };
  const dateInRange = d => d >= startDate && d <= endDate;
  const sources = {};
  _readSheetRows(ss, 'GA4_티빙 집단소송 신청하기').forEach(r => {
    const d = _normDate(r['날짜']);
    if (!dateInRange(d)) return;
    const sm = String(r['세션 소스/매체'] || '');
    if (!/kakao|kkt/i.test(sm)) return;
    if (!sources[sm]) sources[sm] = {users:0, events:0, sessTime:0, sessCount:0};
    const u = num(r['총 사용자']);
    sources[sm].users += u;
    sources[sm].events += num(r['이벤트 수']);
    const t = num(r['평균 세션 시간']);
    if (t > 0) { sources[sm].sessTime += t * u; sources[sm].sessCount += u; }
  });
  const out = Object.entries(sources).map(([sm, v]) => ({
    sm, users: v.users, events: v.events,
    avgSessTime: v.sessCount > 0 ? Math.round(v.sessTime / v.sessCount) : 0,
    isReVisit: /crm_alimtalk|2nd_alimtalk|kakao_menu/i.test(sm),
    isStrong: /2nd_alimtalk|crm_alimtalk/i.test(sm)
  })).sort((a,b) => b.users - a.users);
  const total = out.reduce((s,r) => s + r.users, 0);
  const reVisitUsers = out.filter(r => r.isReVisit).reduce((s,r) => s + r.users, 0);
  const reVisitRate = total > 0 ? Math.round(reVisitUsers / total * 100) : 0;
  return {sources: out, total, reVisitRate};
}

// 네이버 sa_mo vs sa_pc 디바이스 분석
function _deviceMobilePcAnalysis(ss, startDate, endDate) {
  const num = v => { const n = parseFloat(String(v||'').replace(/[,%\s₩]/g,'')); return isNaN(n) ? 0 : n; };
  const dateInRange = d => d >= startDate && d <= endDate;
  const dev = {mo: {users:0, events:0}, pc: {users:0, events:0}, referral: {users:0, events:0}};
  _readSheetRows(ss, 'GA4_티빙 집단소송 신청하기').forEach(r => {
    const d = _normDate(r['날짜']);
    if (!dateInRange(d)) return;
    const sm = String(r['세션 소스/매체'] || '');
    const u = num(r['총 사용자']);
    const ev = num(r['이벤트 수']);
    if (/Naver\s*\/\s*sa_mo/i.test(sm)) { dev.mo.users += u; dev.mo.events += ev; }
    else if (/Naver\s*\/\s*sa_pc/i.test(sm)) { dev.pc.users += u; dev.pc.events += ev; }
    else if (/naver/i.test(sm) && /referral/i.test(sm)) { dev.referral.users += u; dev.referral.events += ev; }
  });
  return dev;
}

// 메타 캠페인 분석 — 광고세트별 분리 + 마케팅 히스토리 자동 매칭
function _metaCampaignAnalysis(ss, startDate, endDate, history) {
  const num = v => { const n = parseFloat(String(v||'').replace(/[,%\s₩]/g,'')); return isNaN(n) ? 0 : n; };
  const dateInRange = d => d >= startDate && d <= endDate;
  const ads = {};
  _readSheetRows(ss, '메타_광고').forEach(r => {
    const d = _normDate(r['일']);
    if (!dateInRange(d)) return;
    const adName = String(r['광고 이름'] || '').trim();
    const setName = String(r['광고 세트 이름'] || '').trim();
    if (!adName) return;
    if (!ads[adName]) ads[adName] = {adName, setName, cost:0, imp:0, clk:0, rows:0};
    ads[adName].cost += num(r['지출 금액 (KRW)']);
    ads[adName].imp += num(r['노출']);
    ads[adName].clk += num(r['링크 클릭']);
    ads[adName].rows++;
  });

  // 마케팅 히스토리에서 의도/시작일 매칭
  const intentMap = {};
  const startMap = {};
  (history || []).forEach(h => {
    const detail = String(h['변경 상세 내용'] || '');
    const item = String(h['변경 항목'] || '');
    const date = h._date || '';
    // tving_xxx_yymmdd 패턴 추출
    const matches = detail.match(/tving_\w+_\d{6}/g) || [];
    matches.forEach(m => {
      // 첫 등장한 row가 시작일 + 의도
      if (!startMap[m]) startMap[m] = date;
      if (!intentMap[m]) {
        // detail에서 의도 추출: "메인 소재로 ...", "보조 소재 ..." 등
        let intent = item || detail.slice(0, 120);
        // "tving_XXX(설명)" 패턴이 있으면 설명 추출
        const desc = detail.match(new RegExp(m + '\\s*[\\(（]([^\\)）]+)[\\)）]'));
        if (desc) intent = desc[1].slice(0, 80);
        intentMap[m] = intent;
      }
    });
    // base 이름(_260604 제거)도 매칭
    const baseMatches = detail.match(/tving_\w+(?=_\d{6}|\s|"|”|\)|“)/g) || [];
    baseMatches.forEach(base => {
      if (!intentMap[base]) {
        const desc = detail.match(new RegExp(base + '\\s*[\\(（]([^\\)）]+)[\\)）]'));
        if (desc) intentMap[base] = desc[1].slice(0, 80);
      }
    });
  });

  // ads에 intent + startDate 매핑
  const list = Object.values(ads).map(a => {
    const adKey = a.adName;
    const baseKey = a.adName.replace(/_\d{6}$/, '');
    return Object.assign({}, a, {
      ctr: a.imp > 0 ? a.clk / a.imp * 100 : 0,
      cpc: a.clk > 0 ? a.cost / a.clk : 0,
      cpm: a.imp > 0 ? a.cost / a.imp * 1000 : 0,
      intent: intentMap[adKey] || intentMap[baseKey] || '',
      startDate: startMap[adKey] || ''
    });
  });

  // 광고세트별 그룹
  const setGroups = {};
  list.forEach(a => {
    if (!setGroups[a.setName]) setGroups[a.setName] = [];
    setGroups[a.setName].push(a);
  });
  Object.keys(setGroups).forEach(s => setGroups[s].sort((a,b) => b.cost - a.cost));

  // 미집행 광고 (cost=0이지만 마케팅 히스토리에 등록된 것)
  const inactive = [];
  (history || []).forEach(h => {
    const detail = String(h['변경 상세 내용'] || '');
    const matches = detail.match(/tving_\w+_\d{6}/g) || [];
    matches.forEach(m => {
      if (!ads[m] && !inactive.find(x => x.adName === m)) {
        inactive.push({adName: m, intent: intentMap[m] || '', startDate: startMap[m] || ''});
      }
    });
  });

  const total = list.reduce((acc, a) => ({
    cost: acc.cost + a.cost, imp: acc.imp + a.imp, clk: acc.clk + a.clk
  }), {cost: 0, imp: 0, clk: 0});

  return {ads: list, setGroups, inactive, total};
}

// 채널 태그 클래스 (기존 _chTagClass 보완)
function _chTagClassV2(ch) {
  if (/메타|meta|fb|ig/i.test(ch)) return 'meta';
  if (/네이버|naver/i.test(ch)) return 'naver';
  if (/구글|google/i.test(ch)) return 'google';
  if (/카카오|kakao/i.test(ch)) return 'meta';
  return 'ga';
}

// DoD percent helper
function _dod(curr, prev) {
  if (!prev || prev === 0) return curr > 0 ? '▲ 신규' : '—';
  const d = (curr - prev) / prev * 100;
  const sign = d >= 0 ? '▲ +' : '▼ ';
  return sign + d.toFixed(1) + '%';
}
function _dodCls(curr, prev) {
  if (!prev || prev === 0) return curr > 0 ? 'up' : '';
  return curr >= prev ? 'up' : 'down';
}
function _dodAbs(curr, prev) {
  const d = curr - prev;
  const sign = d >= 0 ? '+' : '';
  return sign + _fmt0(d);
}


// ============================================================================
// v1.2.0 — _buildReportHtml (replaces v1.1.x)
// 출력: v1.2.2 샘플과 동일 양식 (cover + TOC + 10섹션 + 조건부 구글)
// ============================================================================

function _buildReportHtml(reportType, startDate, endDate, analysis, history) {
  const ss = _ss();
  const k = analysis.kpi;
  const C = CONFIG.COLOR;
  const isWeekly = reportType === 'weekly';
  const dateLabel = isWeekly ? (startDate + ' ~ ' + endDate) : endDate;
  const reportTitle = isWeekly
    ? '티빙집단소송 광고 성과 분석 보고서 (' + _shortDate(startDate) + '~' + _shortDate(endDate) + ')'
    : '티빙집단소송 광고 성과 분석 보고서 (' + _shortDate(endDate) + ')';
  const reportNum = analysis._totalCost || k.totalCost;
  const reportClicks = k.totalClk;

  // 채널별 비용
  const meta = (analysis.kpi.channels.find(c => c.name === '메타') || {});
  const naver = (analysis.kpi.channels.find(c => c.name === '네이버SA') || {});
  const google = (analysis.kpi.channels.find(c => c.name === '구글') || {});
  const showGoogle = (google.cost || 0) > 0;

  // 전일/전주 비교
  const prev = analysis._prev || null;
  const prevMeta = prev ? (prev.kpi.channels.find(c => c.name === '메타') || {}) : {};
  const prevNaver = prev ? (prev.kpi.channels.find(c => c.name === '네이버SA') || {}) : {};
  const prevGoogle = prev ? (prev.kpi.channels.find(c => c.name === '구글') || {}) : {};

  // GA 채널 (직접 제외 분류)
  const gaBy = analysis._gaByChannel || {};
  const prevGaBy = (prev && prev._gaByChannel) || {};

  // 분류 함수 — agg.ga.byChannel은 이미 한국어 채널명으로 집계됨
  function classifyChannels(by) {
    const out = {naver:0, meta:0, kakao:0, organic:0, referral:0, ai:0, etc:0};
    Object.entries(by).forEach(([key, u]) => {
      if (key === '네이버SA' || /naver\s*\/\s*sa_/i.test(key)) out.naver += u;
      else if (key === '메타' || /Meta\s*\/\s*fb_ig/i.test(key)) out.meta += u;
      else if (key === '카카오' || /kakao/i.test(key)) out.kakao += u;
      else if (key === '오가닉' || /organic/i.test(key)) out.organic += u;
      else if (key === '리퍼럴' || /referral|blog_txt/i.test(key)) out.referral += u;
      else if (key === 'AI어시스턴트' || /chatgpt|perplexity|ai-assistant/i.test(key)) out.ai += u;
      else if (key === '자체유입' || /social|home|ig/i.test(key)) out.etc += u;
    });
    return out;
  }
  const cls = classifyChannels(gaBy);
  const prevCls = classifyChannels(prevGaBy);
  const totalAd = cls.naver + cls.meta;
  const totalNonAd = cls.kakao + cls.organic + cls.referral + cls.ai + cls.etc;
  const prevTotalAd = prevCls.naver + prevCls.meta;
  const prevTotalNonAd = prevCls.kakao + prevCls.organic + prevCls.referral + prevCls.ai + prevCls.etc;

  // direct 트래픽 (제외 대상)
  let directUsers = 0;
  Object.entries(gaBy).forEach(([key, u]) => {
    if (key === '직접' || /\(direct\)|\(not set\)|data not available/i.test(key)) directUsers += u;
  });

  // 광고 CPA
  const adCost = (meta.cost||0) + (naver.cost||0) + (google.cost||0);
  const adCPA = totalAd > 0 ? adCost / totalAd : 0;
  const prevAdCost = (prevMeta.cost||0) + (prevNaver.cost||0) + (prevGoogle.cost||0);
  const prevAdCPA = prevTotalAd > 0 ? prevAdCost / prevTotalAd : 0;

  // 메타 캠페인 분석
  const metaAds = analysis._metaCampaigns || {ads:[], setGroups:{}, inactive:[], total:{cost:0,imp:0,clk:0}};
  const kakaoData = analysis._kakao || {sources:[], total:0, reVisitRate:0};
  const deviceData = analysis._device || {mo:{users:0}, pc:{users:0}, referral:{users:0}};
  const prevDevice = (prev && prev._device) || {mo:{users:0}, pc:{users:0}, referral:{users:0}};

  const cssBlock = _v12CssBlock();
  const periodLabel = isWeekly ? '주간' : '일간';
  const subtitle = isWeekly
    ? ('광고 성과 주간 분석 보고서 (' + startDate + ' ~ ' + endDate + ')')
    : ('광고 성과 일간 분석 보고서 (' + endDate + ')');

  let html = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta name="robots" content="noindex, nofollow"><title>' + reportTitle + '</title>' + cssBlock + '</head><body><div class="page-wrap">';

  // COVER
  html += '<section class="cover">'
    + '<span class="wf-tag">TVING CLASS ACTION · ' + (isWeekly ? 'WEEKLY' : 'DAILY') + ' REPORT</span>'
    + '<span class="mode-badge">' + dateLabel + '</span>'
    + '<h1>티빙 개인정보 유출 집단소송</h1>'
    + '<div class="subtitle">' + subtitle + '</div>'
    + '<div class="kpi-grid">'
    + '<div class="kpi-cover"><div class="label">총 광고비</div><div class="value">₩' + _fmt0(adCost) + ' <span class="unit">KRW</span></div></div>'
    + '<div class="kpi-cover"><div class="label">총 노출</div><div class="value">' + _fmt0((meta.imp||0) + (naver.imp||0) + (google.imp||0)) + '</div></div>'
    + '<div class="kpi-cover"><div class="label">총 클릭</div><div class="value">' + _fmt0(reportClicks) + '</div></div>'
    + '<div class="kpi-cover"><div class="label">신청자<span style="font-size:6pt;opacity:0.7;"> (직접 제외)</span></div><div class="value">' + _fmt0(totalAd + totalNonAd) + '</div></div>'
    + '<div class="kpi-cover"><div class="label">CPA</div><div class="value">₩' + _fmt0((totalAd + totalNonAd) > 0 ? adCost / (totalAd + totalNonAd) : 0) + ' <span class="unit">/신청</span></div></div>'
    + '</div>'
    + '<div class="meta">발행 · ' + new Date().toLocaleDateString('ko-KR') + ' · ' + CONFIG.AUTHOR + ' · ' + CONFIG.CLIENT + '</div>'
    + '</section>';

  // TOC
  const tocItems = [
    '<li><strong>핵심 결론</strong> — ' + (isWeekly ? '전주' : '어제') + ' 대비 종합 평가</li>',
    '<li><strong>' + periodLabel + ' KPI 스냅샷</strong> — ' + (isWeekly ? 'WoW' : 'DoD') + ' ▲▼ 비교</li>',
    '<li><strong>채널별 광고 성과</strong> — ' + (showGoogle ? '메타 / 네이버SA / 구글' : '메타 / 네이버SA') + ' + 전일 대비</li>',
    '<li><strong>네이버 SA 키워드 효율</strong> — TOP + ' + (isWeekly ? 'WoW' : 'DoD') + '</li>',
    '<li><strong>GA4 신청 채널 분석</strong> — 광고 vs 비광고 (직접 제외) + ' + (isWeekly ? 'WoW' : 'DoD') + '</li>',
    '<li><strong>카카오톡 재신청 패턴</strong> — Source별 + ' + (isWeekly ? 'WoW' : 'DoD') + '</li>',
    '<li><strong>메타 광고 캠페인 성과</strong> — 광고세트별 광고비/CTR/CPC + ' + (isWeekly ? 'WoW' : 'DoD') + '</li>',
    '<li><strong>모바일 vs PC</strong> — 디바이스별 + ' + (isWeekly ? 'WoW' : 'DoD') + '</li>',
    '<li><strong>마케팅 히스토리</strong> — 기간 내 캠페인 변경</li>',
    '<li><strong>인사이트 & 액션 플랜</strong> — P1/P2/P3</li>'
  ];
  html += '<section class="toc"><h2>목차 / Contents</h2><ol>' + tocItems.join('') + '</ol></section>';

  // SECTION 01 — 핵심 결론
  const naverShare = totalAd > 0 ? Math.round(cls.naver / totalAd * 1000) / 10 : 0;
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 01</div><div class="sec-title">핵심 결론</div>'
    + '<div class="sec-subtitle">' + dateLabel + ' · ' + (isWeekly ? '전주' : '어제') + ' 대비 종합 평가</div>'
    + '</div><div class="section-body">'
    + '<div class="callout"><span class="title">📌 핵심 메시지</span>'
    + dateLabel + ' 광고비 <b>₩' + _fmt0(adCost) + '</b>(' + _dod(adCost, prevAdCost) + ') '
    + '클릭 <b>' + _fmt0(reportClicks) + '</b>(' + _dod(reportClicks, (prev?prev.kpi.totalClk:0)) + ') '
    + '<b>신청자 ' + _fmt0(totalAd + totalNonAd) + '명*</b>(' + _dod(totalAd + totalNonAd, prevTotalAd + prevTotalNonAd) + '). '
    + '핵심 전환 채널 <span class="tag naver">네이버SA</span> ' + _fmt0(cls.naver) + '명. '
    + '<span class="tag meta">메타</span> ' + _fmt0(cls.meta) + '명 — <b>일부 신규 소재는 머신러닝 학습 진행 중</b>. '
    + '카카오 알림톡 ' + _fmt0(cls.kakao) + '명.'
    + '</div>'
    + '<div class="section-narrative">CPA <b>₩' + _fmt0((totalAd + totalNonAd) > 0 ? adCost / (totalAd + totalNonAd) : 0) + '</b>(광고비 / 신청자, 직접 제외). '
    + '내역: 광고 채널 ' + _fmt0(totalAd) + '명(네이버SA ' + _fmt0(cls.naver) + ' / 메타 ' + _fmt0(cls.meta) + '), '
    + '비광고 채널 ' + _fmt0(totalNonAd) + '명(카카오/오가닉/리퍼럴/AI/기타 합산).<br/>'
    + '<span style="font-size:7pt;color:#94A3B8;">* 신청자는 GA 직접(direct) 채널 ' + _fmt0(directUsers) + '명 제외 — 일부는 카카오 알림톡 UTM 누락분 추정</span>'
    + '</div></div></section>';

  // SECTION 02 — KPI 스냅샷
  const totalImp = (meta.imp||0) + (naver.imp||0) + (google.imp||0);
  const prevTotalImp = (prevMeta.imp||0) + (prevNaver.imp||0) + (prevGoogle.imp||0);
  const ctr = totalImp > 0 ? reportClicks / totalImp * 100 : 0;
  const prevCtr = prevTotalImp > 0 ? (prev?prev.kpi.totalClk:0) / prevTotalImp * 100 : 0;
  const cpc = reportClicks > 0 ? adCost / reportClicks : 0;
  const prevCpc = (prev?prev.kpi.totalClk:0) > 0 ? prevAdCost / prev.kpi.totalClk : 0;

  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 02</div><div class="sec-title">' + periodLabel + ' KPI 스냅샷</div>'
    + '<div class="sec-subtitle">' + dateLabel + ' vs ' + (isWeekly ? '전주' : '전일') + ' ' + (isWeekly ? 'WoW' : 'DoD') + '</div>'
    + '</div><div class="section-body"><div class="kpi-grid">'
    + _v12KpiCard('총 광고비', '₩' + _fmt0(adCost), adCost, prevAdCost, '₩' + _fmt0(prevAdCost))
    + _v12KpiCard('총 노출', _fmt0(totalImp), totalImp, prevTotalImp, _fmt0(prevTotalImp))
    + _v12KpiCard('총 클릭', _fmt0(reportClicks), reportClicks, (prev?prev.kpi.totalClk:0), _fmt0(prev?prev.kpi.totalClk:0))
    + _v12KpiCardPt('평균 CTR', ctr.toFixed(2) + '%', ctr, prevCtr)
    + _v12KpiCard('평균 CPC', '₩' + _fmt0(cpc), cpc, prevCpc, '₩' + _fmt0(prevCpc), true)
    + _v12KpiCard('신청자*', _fmt0(totalAd + totalNonAd), totalAd + totalNonAd, prevTotalAd + prevTotalNonAd, _fmt0(prevTotalAd + prevTotalNonAd) + ' (직접 제외)')
    + _v12KpiCard('CPA', '₩' + _fmt0((totalAd + totalNonAd) > 0 ? adCost / (totalAd + totalNonAd) : 0), (totalAd + totalNonAd) > 0 ? adCost / (totalAd + totalNonAd) : 0, (prevTotalAd + prevTotalNonAd) > 0 ? prevAdCost / (prevTotalAd + prevTotalNonAd) : 0, '광고비 / 신청자', true)
    + _v12KpiCard('광고 채널 신청', _fmt0(totalAd), totalAd, prevTotalAd, _fmt0(prevTotalAd) + ' (메타+네이버SA)')
    + '</div>'
    + '<p style="font-size:7pt;color:#94A3B8;margin-top:10px;">* 신청자는 GA 직접(direct) 채널 ' + _fmt0(directUsers) + '명 제외 — 일부는 카카오 알림톡 UTM 누락분 추정. CPA는 광고비 ÷ 전체 신청자(직접 제외).</p>'
    + '</div></section>';

  // SECTION 03 — 채널별 광고 성과
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 03</div><div class="sec-title">채널별 광고 성과</div>'
    + '<div class="sec-subtitle">' + (showGoogle ? '메타 / 네이버SA / 구글' : '메타 / 네이버SA') + ' — ' + (isWeekly ? 'WoW' : 'DoD') + '</div>'
    + '</div><div class="section-body"><table>'
    + '<tr><th rowspan="2">채널</th><th colspan="2" class="num">광고비</th><th colspan="2" class="num">클릭</th><th colspan="2" class="num">신청자</th><th colspan="2" class="num">CPA</th></tr>'
    + '<tr><th class="num">' + _shortDate(endDate) + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th><th class="num">' + _shortDate(endDate) + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th><th class="num">' + _shortDate(endDate) + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th><th class="num">' + _shortDate(endDate) + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th></tr>';

  function chRow(name, tagCls, curr, prev, currConv, prevConv, mlNote) {
    const cpa = currConv > 0 ? curr.cost / currConv : 0;
    const pcpa = prevConv > 0 ? prev.cost / prevConv : 0;
    return '<tr>'
      + '<td><span class="tag ' + tagCls + '">' + name + '</span>' + (mlNote ? ' <span style="font-size:7pt;color:#94A3B8;">*ML 학습중</span>' : '') + '</td>'
      + '<td class="num">₩' + _fmt0(curr.cost||0) + '</td><td class="num ' + _dodCls(curr.cost||0, prev.cost||0) + '">' + _dod(curr.cost||0, prev.cost||0) + '</td>'
      + '<td class="num">' + _fmt0(curr.clk||0) + '</td><td class="num ' + _dodCls(curr.clk||0, prev.clk||0) + '">' + _dod(curr.clk||0, prev.clk||0) + '</td>'
      + '<td class="num">' + _fmt0(currConv) + '</td><td class="num ' + _dodCls(currConv, prevConv) + '">' + _dod(currConv, prevConv) + '</td>'
      + '<td class="num">₩' + _fmt0(cpa) + '</td><td class="num ' + _dodCls(pcpa, cpa) + '">' + _dod(cpa, pcpa) + '</td>'
      + '</tr>';
  }
  html += chRow('메타', 'meta', meta, prevMeta, cls.meta, prevCls.meta, true);
  html += chRow('네이버SA', 'naver', naver, prevNaver, cls.naver, prevCls.naver, false);
  if (showGoogle) html += chRow('구글', 'google', google, prevGoogle, 0, 0, false);
  html += '<tr class="total"><td>광고 합계</td>'
    + '<td class="num">₩' + _fmt0(adCost) + '</td><td class="num ' + _dodCls(adCost, prevAdCost) + '">' + _dod(adCost, prevAdCost) + '</td>'
    + '<td class="num">' + _fmt0(reportClicks) + '</td><td class="num ' + _dodCls(reportClicks,(prev?prev.kpi.totalClk:0)) + '">' + _dod(reportClicks,(prev?prev.kpi.totalClk:0)) + '</td>'
    + '<td class="num">' + _fmt0(totalAd) + '</td><td class="num ' + _dodCls(totalAd,prevTotalAd) + '">' + _dod(totalAd,prevTotalAd) + '</td>'
    + '<td class="num">₩' + _fmt0(adCPA) + '</td><td class="num ' + _dodCls(prevAdCPA, adCPA) + '">' + _dod(adCPA, prevAdCPA) + '</td>'
    + '</tr></table>'
    + '<div class="callout"><span class="title">🤖 메타 머신러닝 학습 컨텍스트</span>메타는 Advantage+ 머신러닝 알고리즘이 오디언스/시간대를 최적화 중. 일부 신규 소재(CI 계열)는 <b>수요 적은 시간대 노출을 자체 절감</b> 중이라 노출/클릭이 낮게 보일 수 있음. <b>학습 완료(약 7~14일 데이터 누적) 후 효율 자동 회복 예상</b> — 지금 끄지 말고 예산 유지가 정답.</div>';
  if (!showGoogle) html += '<div class="section-narrative">구글 광고는 현재 미집행 상태로 본 보고서에서 제외.</div>';
  html += '</div></section>';

  // SECTION 04 — 네이버 키워드
  if (analysis.topKeywords && analysis.topKeywords.length > 0) {
    const kwDiff = _keywordDiff(analysis.topKeywords, (prev && prev.topKeywords) || []);
    html += '<section class="section"><div class="section-header">'
      + '<div class="sec-num">SECTION 04</div><div class="sec-title">네이버 SA 키워드 효율</div>'
      + '<div class="sec-subtitle">TOP ' + kwDiff.length + ' — 실측 + ' + (isWeekly?'WoW':'DoD') + '</div>'
      + '</div><div class="section-body"><table>'
      + '<tr><th>키워드</th><th class="num">노출</th><th class="num">클릭</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">비용</th><th>' + (isWeekly?'WoW':'DoD') + '</th></tr>';
    const effSet = new Set(analysis.efficientKeywords.map(e => e.kw));
    kwDiff.forEach(kw => {
      const star = effSet.has(kw.kw) ? ' ⭐' : '';
      html += '<tr><td>' + kw.kw + star + '</td>'
        + '<td class="num">' + _fmt0(kw.imp) + '</td>'
        + '<td class="num">' + _fmt0(kw.clk) + '</td>'
        + '<td class="num">' + kw.ctr.toFixed(2) + '%</td>'
        + '<td class="num">₩' + _fmt0(kw.cpc) + '</td>'
        + '<td class="num">₩' + _fmt0(kw.cost) + '</td>'
        + '<td><span class="tag ' + kw.dodCls + '">' + kw.dod + '</span></td>'
        + '</tr>';
    });
    html += '</table>';
    if (analysis.efficientKeywords.length > 0) {
      html += '<div class="callout"><span class="title">⭐ 효율 최상 키워드 (CTR 기준)</span>'
        + analysis.efficientKeywords.map(k => '<b>' + k.kw + '</b> CTR ' + k.ctr.toFixed(1) + '% / CPC ₩' + _fmt0(k.cpc)).join(' · ')
        + '. 입찰가 단계적 상향 + 별도 그룹 분리 권장.</div>';
    }
    html += '</div></section>';
  }

  // SECTION 05 — GA 신청 채널 (직접 제외)
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 05</div><div class="sec-title">GA4 신청 채널 분석</div>'
    + '<div class="sec-subtitle">광고 vs 비광고 — 직접(direct) 제외 + ' + (isWeekly?'WoW':'DoD') + '</div>'
    + '</div><div class="section-body">'
    + '<p style="font-size:8pt;color:#64748B;">※ GA `(not set)`/`(direct)/(none)`/`(data not available)` 등 미식별 트래픽 ' + _fmt0(directUsers) + '명 분석 대상 제외</p>'
    + '<table>'
    + '<tr><th>구분</th><th>채널</th><th class="num">' + _shortDate(endDate) + '</th><th class="num">비중</th><th class="num">' + (isWeekly?'전주':'전일') + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th><th class="num">광고비</th><th class="num">CPA</th></tr>';
  const totalAll = totalAd + totalNonAd;
  function gaRow(label, channel, curr, prev, cost, hasGroup) {
    const cpa = curr > 0 && cost > 0 ? cost / curr : 0;
    const pct = totalAll > 0 ? (curr / totalAll * 100).toFixed(1) : '0';
    return '<tr>' + (hasGroup ? '<td>' + hasGroup + '</td>' : '') + '<td>' + channel + '</td>'
      + '<td class="num">' + _fmt0(curr) + '</td><td class="num">' + pct + '%</td>'
      + '<td class="num">' + _fmt0(prev) + '</td><td class="num ' + _dodCls(curr, prev) + '">' + _dod(curr, prev) + '</td>'
      + '<td class="num">₩' + _fmt0(cost) + '</td><td class="num">' + (cpa > 0 ? '₩' + _fmt0(cpa) : '₩0') + '</td></tr>';
  }
  html += '<tr><td rowspan="2"><b>광고</b></td><td><span class="tag naver">네이버SA</span></td><td class="num">' + _fmt0(cls.naver) + '</td><td class="num">' + (totalAll>0?(cls.naver/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevCls.naver) + '</td><td class="num ' + _dodCls(cls.naver, prevCls.naver) + '">' + _dod(cls.naver, prevCls.naver) + '</td><td class="num">₩' + _fmt0(naver.cost||0) + '</td><td class="num">₩' + _fmt0(cls.naver>0?(naver.cost||0)/cls.naver:0) + '</td></tr>';
  html += '<tr><td><span class="tag meta">메타</span></td><td class="num">' + _fmt0(cls.meta) + '</td><td class="num">' + (totalAll>0?(cls.meta/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevCls.meta) + '</td><td class="num ' + _dodCls(cls.meta, prevCls.meta) + '">' + _dod(cls.meta, prevCls.meta) + '</td><td class="num">₩' + _fmt0(meta.cost||0) + '</td><td class="num">₩' + _fmt0(cls.meta>0?(meta.cost||0)/cls.meta:0) + '</td></tr>';
  if (showGoogle) html += '<tr><td><span class="tag google">구글</span></td><td class="num">0</td><td class="num">0%</td><td class="num">0</td><td class="num">—</td><td class="num">₩' + _fmt0(google.cost||0) + '</td><td class="num">—</td></tr>';
  html += '<tr class="total"><td colspan="2">광고 소계</td><td class="num">' + _fmt0(totalAd) + '</td><td class="num">' + (totalAll>0?(totalAd/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevTotalAd) + '</td><td class="num ' + _dodCls(totalAd,prevTotalAd) + '">' + _dod(totalAd,prevTotalAd) + '</td><td class="num">₩' + _fmt0(adCost) + '</td><td class="num">₩' + _fmt0(adCPA) + '</td></tr>';
  html += '<tr><td rowspan="5"><b>비광고<br/>(직접 제외)</b></td><td>네이버 referral</td><td class="num">' + _fmt0(cls.referral) + '</td><td class="num">' + (totalAll>0?(cls.referral/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevCls.referral) + '</td><td class="num ' + _dodCls(cls.referral,prevCls.referral) + '">' + _dod(cls.referral,prevCls.referral) + '</td><td class="num">₩0</td><td class="num">₩0</td></tr>';
  html += '<tr><td>오가닉 검색</td><td class="num">' + _fmt0(cls.organic) + '</td><td class="num">' + (totalAll>0?(cls.organic/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevCls.organic) + '</td><td class="num ' + _dodCls(cls.organic,prevCls.organic) + '">' + _dod(cls.organic,prevCls.organic) + '</td><td class="num">₩0</td><td class="num">₩0</td></tr>';
  html += '<tr><td><b>카카오</b></td><td class="num">' + _fmt0(cls.kakao) + '</td><td class="num">' + (totalAll>0?(cls.kakao/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevCls.kakao) + '</td><td class="num ' + _dodCls(cls.kakao,prevCls.kakao) + '">' + _dod(cls.kakao,prevCls.kakao) + '</td><td class="num">₩0</td><td class="num">₩0</td></tr>';
  html += '<tr><td>기타 (home/social)</td><td class="num">' + _fmt0(cls.etc) + '</td><td class="num">' + (totalAll>0?(cls.etc/totalAll*100).toFixed(1):'0') + '%</td><td class="num">—</td><td class="num">—</td><td class="num">₩0</td><td class="num">₩0</td></tr>';
  html += '<tr><td>AI 어시스턴트</td><td class="num">' + _fmt0(cls.ai) + '</td><td class="num">' + (totalAll>0?(cls.ai/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevCls.ai) + '</td><td class="num ' + _dodCls(cls.ai,prevCls.ai) + '">' + _dod(cls.ai,prevCls.ai) + '</td><td class="num">₩0</td><td class="num">₩0</td></tr>';
  html += '<tr class="total"><td colspan="2">비광고 소계</td><td class="num">' + _fmt0(totalNonAd) + '</td><td class="num">' + (totalAll>0?(totalNonAd/totalAll*100).toFixed(1):'0') + '%</td><td class="num">' + _fmt0(prevTotalNonAd) + '</td><td class="num ' + _dodCls(totalNonAd,prevTotalNonAd) + '">' + _dod(totalNonAd,prevTotalNonAd) + '</td><td class="num">₩0</td><td class="num">₩0</td></tr>';
  html += '<tr class="total"><td colspan="2"><b>전체 (직접 제외)</b></td><td class="num"><b>' + _fmt0(totalAll) + '</b></td><td class="num">100%</td><td class="num">' + _fmt0(prevTotalAd+prevTotalNonAd) + '</td><td class="num ' + _dodCls(totalAll, prevTotalAd+prevTotalNonAd) + '">' + _dod(totalAll, prevTotalAd+prevTotalNonAd) + '</td><td class="num">₩' + _fmt0(adCost) + '</td><td class="num">₩' + _fmt0(totalAll>0?adCost/totalAll:0) + '</td></tr>';
  html += '</table>'
    + '<div class="callout"><span class="title">📌 측정 메모 — 직접(direct) 트래픽 일부는 카카오 알림톡일 가능성</span>GA `(direct)` 트래픽 <b>' + _fmt0(directUsers) + '명</b> 중 일부는 카카오 알림톡 발송 시 <b>UTM 누락된 분으로 추정</b>. 측정 정확도를 위해 향후 모든 알림톡 발송 시 UTM 필수 부착 필요.</div>'
    + '</div></section>';

  // SECTION 06 — 카카오 재신청
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 06</div><div class="sec-title">카카오톡 재신청 패턴 분석</div>'
    + '<div class="sec-subtitle">Source 세분화 + ' + (isWeekly?'WoW':'DoD') + ' + 재방문 추정</div>'
    + '</div><div class="section-body"><table>'
    + '<tr><th>Kakao source</th><th>의미</th><th class="num">' + _shortDate(endDate) + ' 사용자</th><th class="num">' + (isWeekly?'전주':'전일') + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th><th class="num">이벤트</th><th>재방문 신호</th></tr>';
  const prevKakao = (prev && prev._kakao && prev._kakao.sources) || [];
  const prevKakaoMap = {};
  prevKakao.forEach(s => { prevKakaoMap[s.sm] = s.users; });
  kakaoData.sources.forEach(s => {
    const prevU = prevKakaoMap[s.sm] || 0;
    let label = '카카오 일반';
    if (/crm_alimtalk/i.test(s.sm)) label = 'CRM 알림톡 (기존 신청자 대상)';
    else if (/2nd_alimtalk/i.test(s.sm)) label = '<b>2차 알림톡 (재신청 명시)</b>';
    else if (/kakao_menu/i.test(s.sm)) label = '카카오 메뉴 직접 진입';
    else if (/defect/i.test(s.sm)) label = '결함 메인';
    const signal = s.isStrong ? '<span class="tag p1">매우 강력</span>' : (s.isReVisit ? '<span class="tag p2">강력</span>' : '—');
    html += '<tr><td>' + s.sm + '</td><td>' + label + '</td><td class="num">' + s.users + '</td><td class="num">' + prevU + '</td><td class="num ' + _dodCls(s.users, prevU) + '">' + _dod(s.users, prevU) + '</td><td class="num">' + s.events + '</td><td>' + signal + '</td></tr>';
  });
  const prevKakaoTotal = prevKakao.reduce((s,k)=>s+k.users, 0);
  html += '<tr class="total"><td colspan="2">카카오 전체</td><td class="num">' + kakaoData.total + '</td><td class="num">' + prevKakaoTotal + '</td><td class="num ' + _dodCls(kakaoData.total, prevKakaoTotal) + '">' + _dod(kakaoData.total, prevKakaoTotal) + '</td><td class="num">—</td><td>—</td></tr>';
  html += '</table>'
    + '<div class="callout"><span class="title">🔁 재방문 추정</span>전체 카카오 ' + kakaoData.total + '명 중 재방문 추정률 <b>' + kakaoData.reVisitRate + '%</b>. crm_alimtalk + 2nd_alimtalk + kakao_menu가 기존 신청자 대상 발송 → 재방문 가능성 높음.</div>'
    + '<div class="callout callout-red"><span class="title">⚠️ 측정 한계 — 표시 ' + kakaoData.total + '명은 UTM 부착 발송분만</span>실제 카카오 알림톡 신청자는 표시 수치보다 클 가능성. <b>일부 알림톡 발송 시 UTM 누락</b>으로 GA에서 `(direct)`로 분류된 트래픽 존재. 측정 정확도 개선을 위해 <b>모든 알림톡 발송 시 UTM 자동 부착 필수</b>.</div>'
    + '</div></section>';

  // SECTION 07 — 메타 광고 캠페인
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 07</div><div class="sec-title">메타 광고 캠페인 성과</div>'
    + '<div class="sec-subtitle">광고세트별 분리 + 마케팅 히스토리 매칭</div>'
    + '</div><div class="section-body">';
  if (Object.keys(metaAds.setGroups).length === 0) {
    html += '<p>메타 광고 데이터 없음.</p>';
  } else {
    html += '<p style="font-size:8pt;color:#64748B;">※ 캠페인 `티빙_집단소송_260604` 내 광고세트별 운영. 의도 라벨은 마케팅 히스토리 시트 기준.</p>';
    Object.entries(metaAds.setGroups).forEach(([setName, ads]) => {
      const setLabel = setName.replace(/\(.*$/, '').trim();
      html += '<h3 style="font-size:10pt;color:#5B4894;margin-top:8px;">광고세트: ' + setName + '</h3>'
        + '<table><tr><th>광고 이름</th><th>소재 의도 (마케팅 히스토리)</th><th class="num">광고비</th><th class="num">노출</th><th class="num">클릭</th><th class="num">CTR</th><th class="num">CPC</th></tr>';
      ads.forEach(a => {
        const star = a.cost === metaAds.total.cost && metaAds.ads.length > 1 ? '' : '';
        const status = a.cost === 0 ? '<span class="tag p3">미집행</span>' :
                       (a.cpm > 2500 && a.ctr < 8) ? '<span class="tag p3">학습 절감</span>' :
                       (a.ctr > 10 && a.cpc < 30) ? '<span class="tag p1">압도 1위</span>' : '';
        html += '<tr><td>' + a.adName + '</td>'
          + '<td>' + (a.intent || '<span style="color:#94A3B8;font-size:8pt;">미매칭</span>') + '</td>'
          + '<td class="num">₩' + _fmt0(a.cost) + '</td>'
          + '<td class="num">' + _fmt0(a.imp) + '</td>'
          + '<td class="num">' + _fmt0(a.clk) + '</td>'
          + '<td class="num">' + a.ctr.toFixed(2) + '%</td>'
          + '<td class="num">₩' + _fmt0(a.cpc) + '</td>'
          + '</tr>';
      });
      html += '</table>';
    });
    html += '<div class="callout"><span class="title">🤖 머신러닝 학습 단계 안내</span>광고세트 모수 <b>2560 → 2060 재학습</b> (6/15 전환). 일부 소재의 DoD 감소 또는 저예산 노출은 Advantage+ 알고리즘이 학습 단계에서 효율 낮은 시간대 노출을 자체 절감한 결과. 신규 소재는 7~14일 학습 데이터 누적이 필요.</div>';
    if (metaAds.inactive.length > 0) {
      html += '<div class="callout callout-red"><span class="title">⚠️ 미집행 광고 ' + metaAds.inactive.length + '종</span>';
      metaAds.inactive.forEach(i => {
        html += '<br><b>' + i.adName + '</b> — ' + (i.intent || '의도 미매칭') + (i.startDate ? ' (등록 ' + i.startDate + ')' : '');
      });
      html += '<br>광고 세트 활성화 누락 또는 예산 미배정 가능성. 메타 광고 매니저 점검 필요.</div>';
    }
  }
  html += '</div></section>';

  // SECTION 08 — 모바일 vs PC
  const moTotal = deviceData.mo.users;
  const pcTotal = deviceData.pc.users;
  const naverTotal = moTotal + pcTotal;
  const moCost = naverTotal > 0 ? (naver.cost||0) * moTotal / naverTotal : 0;
  const pcCost = naverTotal > 0 ? (naver.cost||0) * pcTotal / naverTotal : 0;
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 08</div><div class="sec-title">모바일 vs PC</div>'
    + '<div class="sec-subtitle">디바이스별 신청 효율 + ' + (isWeekly?'WoW':'DoD') + '</div>'
    + '</div><div class="section-body"><table>'
    + '<tr><th>채널</th><th>디바이스</th><th class="num">' + _shortDate(endDate) + ' 신청</th><th class="num">' + (isWeekly?'전주':'전일') + '</th><th class="num">' + (isWeekly?'WoW':'DoD') + '</th><th class="num">광고비</th><th class="num">CPA</th><th class="num">비중</th></tr>'
    + '<tr><td rowspan="2">네이버SA</td><td><b>모바일 (sa_mo)</b></td><td class="num">' + _fmt0(moTotal) + '</td><td class="num">' + _fmt0(prevDevice.mo.users) + '</td><td class="num ' + _dodCls(moTotal, prevDevice.mo.users) + '">' + _dod(moTotal, prevDevice.mo.users) + '</td><td class="num">₩' + _fmt0(moCost) + '</td><td class="num">₩' + _fmt0(moTotal>0?moCost/moTotal:0) + '</td><td class="num">' + (naverTotal>0?(moTotal/naverTotal*100).toFixed(1):'0') + '%</td></tr>'
    + '<tr><td>PC (sa_pc)</td><td class="num">' + _fmt0(pcTotal) + '</td><td class="num">' + _fmt0(prevDevice.pc.users) + '</td><td class="num ' + _dodCls(pcTotal, prevDevice.pc.users) + '">' + _dod(pcTotal, prevDevice.pc.users) + '</td><td class="num">₩' + _fmt0(pcCost) + '</td><td class="num">₩' + _fmt0(pcTotal>0?pcCost/pcTotal:0) + '</td><td class="num">' + (naverTotal>0?(pcTotal/naverTotal*100).toFixed(1):'0') + '%</td></tr>'
    + '<tr class="total"><td colspan="2">네이버SA 합계</td><td class="num">' + _fmt0(naverTotal) + '</td><td class="num">' + _fmt0(prevDevice.mo.users + prevDevice.pc.users) + '</td><td class="num ' + _dodCls(naverTotal, prevDevice.mo.users+prevDevice.pc.users) + '">' + _dod(naverTotal, prevDevice.mo.users+prevDevice.pc.users) + '</td><td class="num">₩' + _fmt0(naver.cost||0) + '</td><td class="num">₩' + _fmt0(naverTotal>0?(naver.cost||0)/naverTotal:0) + '</td><td class="num">100%</td></tr>'
    + '<tr><td colspan="2">네이버 referral (모바일 검색)</td><td class="num">' + _fmt0(deviceData.referral.users) + '</td><td class="num">' + _fmt0(prevDevice.referral.users) + '</td><td class="num ' + _dodCls(deviceData.referral.users, prevDevice.referral.users) + '">' + _dod(deviceData.referral.users, prevDevice.referral.users) + '</td><td class="num">₩0</td><td class="num">—</td><td class="num">—</td></tr>'
    + '</table></div></section>';

  // SECTION 09 — 마케팅 히스토리
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 09</div><div class="sec-title">마케팅 히스토리</div>'
    + '<div class="sec-subtitle">기간 내 캠페인 변경 사항 (마케팅 히스토리 시트)</div>'
    + '</div><div class="section-body">';
  if (history && history.length > 0) {
    // 최근 9개
    const recent = history.slice().sort((a,b) => (b._date||'').localeCompare(a._date||'')).slice(0, 9);
    html += '<table><tr><th>일자</th><th>채널</th><th>변경 유형</th><th>변경 내용</th><th>의도/사유</th></tr>';
    recent.forEach(h => {
      const tagCls = _chTagClassV2(h._channelStd);
      html += '<tr><td>' + (h._date||'').slice(5) + '</td>'
        + '<td><span class="tag ' + tagCls + '">' + h._channelStd + '</span></td>'
        + '<td>' + (h['변경 내용 변경 유형'] || h['변경 유형'] || '') + '</td>'
        + '<td>' + String(h['변경 항목']||'').slice(0,80) + '</td>'
        + '<td>' + String(h['변경 사유']||'').slice(0,100) + '</td></tr>';
    });
    html += '</table>';
  } else {
    html += '<p class="note-empty">기간 내 등록된 캠페인 변경 이벤트 없음.</p>';
  }
  html += '</div></section>';

  // SECTION 10 — 인사이트 & 액션
  html += '<section class="section"><div class="section-header">'
    + '<div class="sec-num">SECTION 10</div><div class="sec-title">인사이트 & 액션 플랜</div>'
    + '<div class="sec-subtitle">관찰 · 원인 · 액션 · 우선순위</div>'
    + '</div><div class="section-body">';
  html += '<h3>🔍 인사이트</h3>';
  (analysis.insights || []).forEach((ins, i) => {
    html += '<div class="callout"><span class="title">[' + (i+1) + '] ' + ins.observation + '</span>'
      + '<b>원인:</b> ' + (ins.cause||'') + '<br/><b>액션:</b> ' + (ins.action||'') + '</div>';
  });
  // 추가 고정 인사이트 — 카카오 UTM
  html += '<div class="callout callout-red"><span class="title">[' + ((analysis.insights||[]).length+1) + '] 카카오 알림톡 UTM 누락 — 측정 정확도 문제</span>'
    + '표시된 카카오 신청 ' + kakaoData.total + '명은 UTM 부착 발송분만. 일부 알림톡에 UTM이 빠져 GA에서 <b>(direct) 트래픽 ' + _fmt0(directUsers) + '명에 혼입</b>된 것으로 추정.<br/>'
    + '<b>원인:</b> 알림톡 발송 도구의 UTM 자동 부착 미설정<br/>'
    + '<b>액션:</b> 모든 알림톡 템플릿 URL에 UTM 자동 부착 — utm_source=kakao / utm_medium=crm_alimtalk / utm_campaign=캠페인명'
    + '</div>';

  html += '<h3 style="margin-top:20px;">⚡ 액션 플랜</h3>'
    + '<table><tr><th>우선순위</th><th>채널</th><th>실행 항목</th><th>실행 시점</th></tr>';
  (analysis.actions || []).forEach(a => {
    html += '<tr><td><span class="tag ' + (a.priority==='P1'?'p1':a.priority==='P2'?'p2':'p3') + '">' + a.priority + '</span></td>'
      + '<td>' + a.channel + '</td><td>' + a.item + '</td><td>' + a.due + '</td></tr>';
  });
  // 고정 액션 추가
  html += '<tr><td><span class="tag p1">P1</span></td><td>카카오</td><td><b>알림톡 발송 시 UTM 자동 부착</b> — 측정 정확도 개선</td><td>익영업일</td></tr>';
  html += '<tr><td><span class="tag p1">P1</span></td><td>메타</td><td><b>머신러닝 학습 단계 인내</b> — 노출 감소 소재 끄지 말고 예산 유지 (7~14일 안정화)</td><td>지속</td></tr>';
  html += '</table></div></section>';

  // FOOTER
  html += '<div class="footer">' + CONFIG.AUTHOR + ' · ' + CONFIG.CLIENT + ' · CONFIDENTIAL · ' + dateLabel + ' ' + periodLabel + ' 분석 보고서'
    + '<span class="pgnum"> — 인쇄용 A4</span></div>'
    + '</div></body></html>';

  return {html, title: reportTitle};
}

// KPI 카드 (DoD up=긍정, lowerBetter=true면 down이 긍정)
function _v12KpiCard(label, value, curr, prev, prevText, lowerBetter) {
  const d = curr - prev;
  const pct = prev > 0 ? (d / prev * 100).toFixed(1) : '0';
  const sign = d >= 0 ? '▲' : '▼';
  const isPositive = lowerBetter ? d < 0 : d > 0;
  const cls = (d === 0) ? '' : (isPositive ? 'up' : 'down');
  return '<div class="kpi-card"><div class="label">' + label + '</div><div class="value">' + value + '</div>'
    + '<div class="delta ' + cls + '">' + sign + ' ' + Math.abs(parseFloat(pct)) + '% (' + (d >= 0 ? '+' : '') + _fmt0(d) + ')</div>'
    + '<div class="delta-note">' + ((typeof prevText === 'string') ? prevText : '') + '</div></div>';
}
function _v12KpiCardPt(label, value, curr, prev) {
  const d = curr - prev;
  const sign = d >= 0 ? '▲' : '▼';
  const cls = (d === 0) ? '' : (d > 0 ? 'up' : 'down');
  return '<div class="kpi-card"><div class="label">' + label + '</div><div class="value">' + value + '</div>'
    + '<div class="delta ' + cls + '">' + sign + ' ' + Math.abs(d).toFixed(2) + 'pt</div>'
    + '<div class="delta-note">전일 ' + prev.toFixed(2) + '%</div></div>';
}

// CSS 블록 — v1.2.2 양식 그대로
function _v12CssBlock() {
  return _V12_CSS;
}

// CSS 상수 (별도 정의 — _V12_CSS)


const _V12_CSS = `<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&display=swap');
*, *::before, *::after { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: 'Noto Sans KR', 'Noto Sans CJK KR', -apple-system, sans-serif;
  font-size: 9pt;
  line-height: 1.55;
  color: #1E293B;
  background: #FFFFFF;
}
.page-wrap { max-width: 900px; margin: 0 auto; padding: 24px 18px; }
@page { size: A4; margin: 14mm 12mm 16mm 12mm; }
@page :first { margin: 0; }

/* ---- Cover ---- */
.cover {
  page-break-after: always;
  break-after: page;
  background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
  color: #FFFFFF;
  padding: 48px 36px;
  border-radius: 8px;
  margin-bottom: 28px;
}
.cover .wf-tag,
.cover .mode-badge {
  display: inline-block;
  padding: 3px 11px;
  border-radius: 12px;
  font-size: 7.5pt;
  font-weight: 700;
  letter-spacing: 0.4px;
}
.cover .wf-tag { background: rgba(123,104,174,0.20); color: #D6CEF0; margin-bottom: 6px; }
.cover .mode-badge { background: rgba(59,130,246,0.20); color: #BFD7FF; margin-left: 6px; margin-bottom: 6px; }
.cover h1 {
  margin: 6px 0 4px;
  font-size: 28pt;
  font-weight: 900;
  letter-spacing: -0.6px;
  color: #FFFFFF;
}
.cover .subtitle { font-size: 12pt; font-weight: 300; color: #CBD5E1; margin-bottom: 14px; }
.cover .divider {
  width: 48px;
  height: 2px;
  background: #7B68AE;
  margin: 14px 0 22px;
}
.cover .kpi-row {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 12px;
  margin: 20px 0 18px;
}
.cover .kpi-grid {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 8px;
}
.cover .kpi-cover {
  background: rgba(255,255,255,0.06);
  border-left: 3px solid #7B68AE;
  padding: 10px 12px;
  border-radius: 0 4px 4px 0;
}
.cover .kpi-cover .label { font-size: 6.5pt; color: #94A3B8; font-weight: 500; letter-spacing: 0.4px; }
.cover .kpi-cover .value { font-size: 15pt; font-weight: 700; color: #FFFFFF; margin-top: 2px; }
.cover .kpi-cover .unit  { font-size: 7.5pt; font-weight: 400; color: #CBD5E1; }
.cover .meta { font-size: 7.5pt; color: #94A3B8; margin-top: 14px; }

/* ---- TOC ---- */
.toc {
  page-break-after: always;
  break-after: page;
  padding: 12px 4px 24px;
}
.toc h2 {
  font-size: 13pt;
  font-weight: 700;
  color: #1E293B;
  border-bottom: 2px solid #7B68AE;
  padding-bottom: 6px;
  margin-bottom: 14px;
}
.toc ol { margin: 0; padding-left: 20px; }
.toc li { margin: 4px 0; font-size: 9.5pt; color: #5A6478; }
.toc li strong { color: #1E293B; }

/* ---- Section header (dark box) ---- */
.section {
  page-break-inside: avoid;
  break-inside: avoid;
  margin: 22px 0 12px;
}
.section-header {
  background: #1E293B;
  padding: 10px 14px;
  border-radius: 4px;
  display: flex;
  align-items: baseline;
  gap: 10px;
  flex-wrap: wrap;
}
.section-header .sec-num {
  color: #7B68AE;
  font-size: 7.5pt;
  font-weight: 700;
  letter-spacing: 1px;
}
.section-header .sec-title {
  color: #FFFFFF;
  font-size: 11pt;
  font-weight: 700;
}
.section-header .sec-subtitle {
  color: #CBD5E1;
  font-size: 8pt;
  font-weight: 400;
}
.section-body { padding: 12px 4px 6px; }

/* ---- KPI cards (in body) ---- */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 10px;
  margin: 10px 0 14px;
}
.kpi-card {
  background: #F8FAFC;
  border: 1px solid #E2E8F0;
  border-left: 3px solid #7B68AE;
  padding: 10px 12px;
  border-radius: 4px;
}
.kpi-card .label { font-size: 7pt; color: #5A6478; font-weight: 500; letter-spacing: 0.4px; }
.kpi-card .value { font-size: 14pt; font-weight: 700; color: #1E293B; margin-top: 2px; }
.kpi-card .unit  { font-size: 7.5pt; font-weight: 400; color: #94A3B8; margin-top: 1px; }
.kpi-card .delta { display: inline-block; margin-left: 4px; font-size: 7.5pt; font-weight: 700; }

/* ---- Tables ---- */
table {
  width: 100%;
  border-collapse: collapse;
  font-size: 8pt;
  margin: 6px 0 12px;
  page-break-inside: avoid;
}
thead th {
  background: #334155;
  color: #FFFFFF;
  font-size: 7.5pt;
  font-weight: 600;
  padding: 6px 8px;
  text-align: left;
  letter-spacing: 0.2px;
}
tbody td {
  padding: 5px 8px;
  border-bottom: 1px solid #E2E8F0;
  color: #1E293B;
  vertical-align: middle;
}
tbody tr:nth-child(odd)  td { background: #FFFFFF; }
tbody tr:nth-child(even) td { background: #F8FAFC; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
tr.total td {
  background: #1E293B !important;
  color: #FFFFFF;
  font-weight: 700;
}

/* ---- Delta ---- */
.up   { color: #C0392B; font-weight: 700; }
.down { color: #1E3A8A; font-weight: 700; }
.up-bad   { color: #C0392B; font-weight: 700; }
.flat { color: #94A3B8; }
.delta-note { font-size: 7pt; color: #94A3B8; font-style: italic; margin-left: 4px; }

/* ---- Boxes ---- */
.callout {
  background: #E8E4F2;
  border-left: 3px solid #7B68AE;
  padding: 10px 14px;
  border-radius: 0 4px 4px 0;
  margin: 8px 0 12px;
  font-size: 8.5pt;
  line-height: 1.7;
}
.callout strong { color: #5B4894; }
.callout .title { font-size: 9pt; font-weight: 700; color: #5B4894; display: block; margin-bottom: 4px; }
.callout-red {
  background: #FDF0EE;
  border-left: 3px solid #C0392B;
  padding: 9px 12px;
  margin: 8px 0;
  border-radius: 0 4px 4px 0;
  font-size: 8.5pt;
}
.callout-red .title { color: #C0392B; font-weight: 700; }
.callout-green {
  background: #EAF4EE;
  border-left: 3px solid #1A6644;
  padding: 9px 12px;
  margin: 8px 0;
  border-radius: 0 4px 4px 0;
  font-size: 8.5pt;
}
.callout-green .title { color: #1A6644; font-weight: 700; }

/* ---- Tags / priorities ---- */
.tag {
  display: inline-block;
  padding: 1px 8px;
  border-radius: 10px;
  font-size: 7pt;
  font-weight: 700;
  margin-right: 4px;
  letter-spacing: 0.3px;
}
.tag.p1 { background: #FDF0EE; color: #C0392B; }
.tag.p2 { background: #FEF3C7; color: #B45309; }
.tag.p3 { background: #E8E4F2; color: #5B4894; }
.tag.meta  { background: #E8E4F2; color: #5B4894; }
.tag.naver { background: #EAF4EE; color: #1A6644; }
.tag.ga    { background: #DBEAFE; color: #1E40AF; }

/* ---- Misc ---- */
p { margin: 4px 0; }
.note-empty { color: #94A3B8; font-size: 8pt; font-style: italic; }
.section-narrative { font-size: 8.5pt; color: #5A6478; margin: 6px 0 10px; line-height: 1.7; }
.divider-soft { border: 0; border-top: 1px solid #E2E8F0; margin: 16px 0; }

/* ---- Footer ---- */
.footer {
  border-top: 1px solid #E2E8F0;
  text-align: center;
  margin-top: 28px;
  padding: 14px 0 6px;
  color: #94A3B8;
  font-size: 6.5pt;
  letter-spacing: 0.3px;
}
.footer .pgnum { margin-left: 6px; }

@media print {
  body { background: #FFFFFF; }
  .page-wrap { padding: 0; max-width: 100%; }
  .cover, .section-header, thead th, tr.total td, .callout, .callout-red, .callout-green, .kpi-card {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}
</style>
`;


// ============================================================================
// v1.2.0 — _publishDailyReport / _publishWeeklyReport overrides
// 전일/전주 데이터 + 카카오/디바이스/메타 캠페인 분석 추가
// ============================================================================

function _publishDailyReport(ss, date, history) {
  const agg = _aggregateByDate(ss, date, date);
  const analysis = _buildStrategicAnalysis(agg, history, date, date);

  // v1.2.0: 전일 데이터
  const prevAgg = _aggregatePreviousDate(ss, date);
  const prevAnalysis = _buildStrategicAnalysis(prevAgg, history, _yesterday(date), _yesterday(date));
  prevAnalysis._gaByChannel = prevAgg.ga.byChannel || {};
  prevAnalysis._kakao = _kakaoReturnAnalysis(ss, _yesterday(date), _yesterday(date));
  prevAnalysis._device = _deviceMobilePcAnalysis(ss, _yesterday(date), _yesterday(date));

  analysis._prev = prevAnalysis;
  analysis._gaByChannel = agg.ga.byChannel || {};
  analysis._kakao = _kakaoReturnAnalysis(ss, date, date);
  analysis._device = _deviceMobilePcAnalysis(ss, date, date);
  analysis._metaCampaigns = _metaCampaignAnalysis(ss, date, date, history);

  const built = _buildReportHtml('daily', date, date, analysis, history);
  const path = 'daily/' + date + '.html';
  const url = _publishToGithub(path, built.html, '[bot] daily report ' + date + ' v1.2');
  _logReportToSheet('daily', date, date, url, built.title);
  return {date, url, title: built.title};
}

function _publishWeeklyReport(ss, weekStart, weekEnd, history) {
  const agg = _aggregateByDate(ss, weekStart, weekEnd);
  const analysis = _buildStrategicAnalysis(agg, history, weekStart, weekEnd);

  // v1.2.0: 전주 데이터
  const prevAgg = _aggregatePreviousWeek(ss, weekStart, weekEnd);
  const ws = new Date(weekStart + 'T00:00:00'); ws.setDate(ws.getDate() - 7);
  const we = new Date(weekEnd + 'T00:00:00'); we.setDate(we.getDate() - 7);
  const prevStart = _fmt(ws), prevEnd = _fmt(we);
  const prevAnalysis = _buildStrategicAnalysis(prevAgg, history, prevStart, prevEnd);
  prevAnalysis._gaByChannel = prevAgg.ga.byChannel || {};
  prevAnalysis._kakao = _kakaoReturnAnalysis(ss, prevStart, prevEnd);
  prevAnalysis._device = _deviceMobilePcAnalysis(ss, prevStart, prevEnd);

  analysis._prev = prevAnalysis;
  analysis._gaByChannel = agg.ga.byChannel || {};
  analysis._kakao = _kakaoReturnAnalysis(ss, weekStart, weekEnd);
  analysis._device = _deviceMobilePcAnalysis(ss, weekStart, weekEnd);
  analysis._metaCampaigns = _metaCampaignAnalysis(ss, weekStart, weekEnd, history);

  const built = _buildReportHtml('weekly', weekStart, weekEnd, analysis, history);
  const path = 'weekly/' + weekStart + '_' + weekEnd + '.html';
  const url = _publishToGithub(path, built.html, '[bot] weekly report ' + weekStart + '~' + weekEnd + ' v1.2');
  _logReportToSheet('weekly', weekStart, weekEnd, url, built.title);
  return {weekStart, weekEnd, url, title: built.title};
}

function _yesterday(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() - 1);
  return _fmt(d);
}

function _shortDate(d) {
  const parts = String(d).split('-');
  return parts.length === 3 ? parseInt(parts[1]) + '/' + parseInt(parts[2]) : d;
}

function _chTagClass(ch) {
  if (/메타/.test(ch)) return 'meta';
  if (/네이버/.test(ch)) return 'naver';
  if (/구글/.test(ch)) return 'google';
  if (/카카오/.test(ch)) return 'meta';
  return 'ga';
}
