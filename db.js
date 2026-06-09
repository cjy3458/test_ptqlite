/* ═══════════════════════════════════════════════════════════════════════════
   KB라이프 테스트 자동화 — 클라이언트 데이터/API 레이어 (서버리스)
   ─────────────────────────────────────────────────────────────────────────
   저장소  : localStorage
   Excel   : SheetJS (CDN)
   AI 생성 : Gemini 2.0 Flash REST API (키 제공 시) / 규칙 기반 (폴백)
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const STORAGE_KEY    = 'kb_qc_db_v1';
const GEMINI_KEY_LOC = 'kb_qc_gemini_key';

/* ═══════════════════════════════════════════════════════════════════════════
   ▼▼▼ Gemini API 키 — 여기에 발급받은 키를 붙여넣으면 AI 모드로 동작합니다 ▼▼▼
   예) const GEMINI_API_KEY = 'AIzaSyxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
   비워 두면(''), 규칙 기반 모드로 자동 폴백합니다 (별도 화면 입력 UI 없음).
   ═══════════════════════════════════════════════════════════════════════════ */
const GEMINI_API_KEY = '';

/* ── 저장소 ──────────────────────────────────────────────────────────────── */
function emptyDB() {
  return {
    seq: { sessions:0, cases:0, results:0, qc:0, psr:0, req:0, docs:0 },
    sessions:[], cases:[], results:[], qcDrafts:[], psrRecords:[], requirements:[], documents:[],
  };
}
let db = (() => {
  try { const r = localStorage.getItem(STORAGE_KEY); if (r) return JSON.parse(r); } catch (_) {}
  return emptyDB();
})();

function save()       { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch (_) {} }
function nextId(tbl)  { db.seq[tbl] = (db.seq[tbl] || 0) + 1; return db.seq[tbl]; }
function nowISO()     { return new Date().toISOString(); }

const findSession = id => db.sessions.find(s => s.id === Number(id));
const findCase    = id => db.cases.find(c => c.id === Number(id));
const casesOf     = sid => db.cases.filter(c => c.session_id === Number(sid));
const resultOf    = cid => db.results.find(r => r.case_id === Number(cid)) || null;
const qcOf        = cid => db.qcDrafts.find(q => q.case_id === Number(cid)) || null;

class ApiError extends Error {
  constructor(status, msg) { super(msg); this.status = status; }
}

/* ═══════════════════════════════════════════════════════════════════════════
   Gemini API 키 관리
   ═══════════════════════════════════════════════════════════════════════════ */
// 코드 상수(GEMINI_API_KEY) 우선, 없으면 localStorage 폴백
function getGeminiKey()    { return (GEMINI_API_KEY || '').trim() || localStorage.getItem(GEMINI_KEY_LOC) || ''; }
function setGeminiKey(key) {
  const k = (key || '').trim();
  if (k) localStorage.setItem(GEMINI_KEY_LOC, k);
  else   localStorage.removeItem(GEMINI_KEY_LOC);
}
function isGeminiAvailable() { return !!getGeminiKey(); }

/* ═══════════════════════════════════════════════════════════════════════════
   Gemini REST API 호출 (브라우저 직접 — 서버 불필요)
   ═══════════════════════════════════════════════════════════════════════════ */
async function callGemini(prompt) {
  const key = getGeminiKey();
  if (!key) throw new ApiError(401, 'Gemini API 키가 없어요. 헤더의 ✦ AI 설정에서 입력해주세요.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(key)}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 4096 },
    }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    if (res.status === 429) {
      // 할당량 초과 / 요청 제한 — 규칙 기반으로 폴백되도록 명확한 메시지로 throw
      throw new ApiError(429, '일시적인 요청 제한입니다. 15초 후 다시 시도해 주세요.');
    }
    throw new ApiError(res.status, err.error?.message || `Gemini API 오류 (${res.status})`);
  }
  const data = await res.json();
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  return text.replace(/^```(?:json)?\s*/m, '').replace(/```\s*$/, '').trim();
}

/* ── Gemini: PSR → 테스트케이스 ──────────────────────────────────────────── */
const CASE_JSON_SCHEMA = `{
  "checklist": [
    {"category": "카테고리명", "item": "확인 항목 (구체적, 한 문장)", "priority": "high|medium|low"}
  ],
  "test_cases": [
    {
      "device": "PC|Tab",
      "channel": "옴니|전자",
      "contractor_insured": "계피동일|계피상이",
      "rider_scope": "테스트할 특약 구성 (한국어, 한 줄)",
      "expected_outcome": "통과|오류|확인필요",
      "test_focus": "이 케이스에서 핵심으로 검증하는 내용 (한 문장)"
    }
  ],
  "risk_points": ["리스크 포인트 문자열"]
}`;

async function casesGemini(productName, requirements) {
  const reqLines = requirements
    .map(r => `- [${r.req_id || ''}] ${r.req_name || ''}: ${r.req_content || ''}`)
    .join('\n');

  const prompt = `당신은 KB라이프 보험 청약 시스템의 숙련된 QA 엔지니어입니다.
아래 PSR 요구사항을 분석하여 테스트케이스를 설계해주세요.

[상품명]
${productName}

[PSR 요구사항]
${reqLines}

[테스트 환경]
- 기기: PC, 태블릿(Tab)
- 청약 채널: 옴니청약, 전자청약
- 계약자/피보험자: 계피동일 또는 계피상이
- 검증 영역: 상품 정보, 청약 흐름, 문서 정합성, 납부·계좌, 서명·따라쓰기, 알릴의무

아래 JSON 형식**만** 출력하세요 (마크다운 코드블록 없이, JSON 텍스트만):
${CASE_JSON_SCHEMA}

생성 규칙:
1. 체크리스트: 요구사항에서 도출된 구체적 확인 항목 12~18개
2. 테스트케이스: 커버리지 매트릭스 기반 8~15개
3. 고위험 요구사항(신설·변경·연령제한·알릴의무 등)은 expected_outcome을 "오류" 또는 "확인필요"로 설계
4. rider_scope는 구체적인 특약명으로 (예: "간암 치료 특약 + 암진단III 특약")
5. test_focus는 해당 케이스에서 **유일하게** 집중하는 포인트 한 문장`;

  const raw = await callGemini(prompt);
  const result = JSON.parse(raw);
  result.source = 'gemini';
  return result;
}

/* ── Gemini: 오류 → QC 초안 ──────────────────────────────────────────────── */
const QC_JSON_SCHEMA = `{
  "severity": "S1|S2|S3|S4",
  "category": "오류 카테고리",
  "analysis": "원인 분석 (2~3문장, 구체적으로)",
  "suggestion": "개발팀 수정 방향 (구체적으로)",
  "related_reqs": ["관련 요구사항 ID 또는 항목명"]
}`;

async function qcGemini(tcId, device, channel, errorTitle, repro, expected, actual) {
  const reproText = repro.map((s, i) => `${i + 1}. ${s}`).join('\n');

  const prompt = `KB라이프 보험 청약 시스템 QA 테스트에서 아래 결함이 발견됐습니다.
QA 엔지니어 관점에서 분석하고 QC 초안을 작성해주세요.

[테스트 정보]
케이스 ID: ${tcId}
기기/채널: ${device} · ${channel}

[결함 내용]
제목: ${errorTitle}
재현 단계:
${reproText}
기대 결과: ${expected}
실제 결과: ${actual}

아래 JSON 형식**만** 출력하세요:
${QC_JSON_SCHEMA}

심각도 기준:
S1: 청약 불가 · 데이터 손실 · 계산 오류
S2: 핵심 기능 이상 · 중요 화면 오류
S3: 문구·명칭 불일치 · UI 표시 오류
S4: 경미한 UI 개선 사항`;

  const raw = await callGemini(prompt);
  const result = JSON.parse(raw);
  result.source = 'gemini';
  return result;
}

/* ═══════════════════════════════════════════════════════════════════════════
   규칙 기반 폴백 (Gemini 키 없거나 오류 시)
   ═══════════════════════════════════════════════════════════════════════════ */
const CAT_KEYWORDS = {
  '상품 정보':      ['특약','보험료','가입','연령','나이','세부터','직업','한도','보장','신설','상품'],
  '청약 흐름':      ['청약','가입설계','입력','계산','단계','화면','진행','발행'],
  '문서 정합성':    ['문구','특약명','명칭','표기','문서','개정','버전','약관','일치'],
  '납부·계좌 처리': ['납부','계좌','출금','초회','납입','은행'],
  '서명·따라쓰기':  ['서명','따라쓰기','전자서명','사인'],
  '알릴의무':       ['알릴의무','고지','질병','진단','입원','수술','이력','소견'],
};
const RIDER_RE = /「([^」]{2,20}?)」/g;

function categorize(text) {
  for (const [cat, kws] of Object.entries(CAT_KEYWORDS)) {
    if (kws.some(k => text.includes(k))) return cat;
  }
  return '기타';
}

function extractRiders(reqs) {
  const seen = [];
  for (const r of reqs) {
    for (const m of ((r.req_name || '') + (r.req_content || '')).matchAll(RIDER_RE)) {
      const name = m[1].replace(/특약/g, '').trim() + ' 특약';
      if (!seen.includes(name)) seen.push(name);
    }
  }
  return seen.length ? seen : ['필수특약'];
}

function casesRuleBased(productName, requirements) {
  const fullText = requirements.map(r => (r.req_name||'') + ' ' + (r.req_content||'')).join(' ');

  const BASE_CHECKS = [
    {category:'청약 흐름',      item:'가입설계→청약서 발행→청약 전 단계 정상 동작 (옴니·전자 각각)', priority:'high'},
    {category:'상품 정보',      item:'특약 구성별 월 보험료 산출 값 정확성',                        priority:'high'},
    {category:'문서 정합성',    item:'화면 노출 특약명 ↔ 상품안 개정본 문구 일치',                  priority:'high'},
    {category:'납부·계좌 처리', item:'초회보험료 출금 정상 처리',                                    priority:'medium'},
    {category:'서명·따라쓰기',  item:'전자청약 서명·따라쓰기 항목 반영 여부',                        priority:'medium'},
  ];
  const seenItems = new Set(BASE_CHECKS.map(c => c.item));
  const extra = [];
  for (const r of requirements) {
    const text = (r.req_name||'') + ' ' + (r.req_content||'');
    const snippet = ((r.req_content||'') || (r.req_name||'')).slice(0, 60);
    if (snippet && !seenItems.has(snippet)) {
      const priority = ['신설','변경','연령','제한','추가','알릴의무'].some(k => text.includes(k)) ? 'high' : 'medium';
      extra.push({ category: categorize(text), item: snippet, priority });
      seenItems.add(snippet);
    }
  }
  const checklist = BASE_CHECKS.concat(extra.slice(0, 13));

  const hasAge   = ['연령','나이','세부터','세 이상','세 미만'].some(k => fullText.includes(k));
  const hasName  = ['명칭','특약명','문구','개정','버전'].some(k => fullText.includes(k));
  const hasNew   = ['신설','추가','새로'].some(k => fullText.includes(k));
  const hasOblig = ['알릴의무','고지','질병','진단','이력'].some(k => fullText.includes(k));
  const hasElec  = fullText.includes('전자') || fullText.includes('전자청약');

  const riders = extractRiders(requirements);
  const r1   = riders[0] || '필수특약';
  const rAll = riders.length >= 3 ? riders.slice(0,3).join(' + ') : riders.join(' + ') || r1;

  const cases = [
    {device:'PC', channel:'옴니',contractor_insured:'계피동일',rider_scope:r1,  expected_outcome:'통과',test_focus:'PC·옴니 기본 청약 흐름 — 필수특약 구성'},
    {device:'PC', channel:'옴니',contractor_insured:'계피상이',rider_scope:rAll,expected_outcome:'통과',test_focus:'PC·옴니 계피상이 전 특약 구성'},
    {device:'Tab',channel:'전자',contractor_insured:'계피동일',rider_scope:r1,  expected_outcome:'통과',test_focus:'Tab·전자 계피동일 기본 흐름'},
    {device:'Tab',channel:'옴니',contractor_insured:'계피상이',rider_scope:rAll,expected_outcome:'통과',test_focus:'Tab·옴니 계피상이 전 특약 구성'},
    {device:'Tab',channel:'전자',contractor_insured:'계피상이',rider_scope:r1,  expected_outcome:'통과',test_focus:'Tab·전자 계피상이 필수 특약'},
    hasElec
      ? {device:'PC', channel:'전자',contractor_insured:'계피동일',rider_scope:rAll,expected_outcome:'통과',test_focus:'PC·전자 계피동일 전 특약 구성'}
      : {device:'Tab',channel:'옴니',contractor_insured:'계피동일',rider_scope:r1,  expected_outcome:'통과',test_focus:'Tab·옴니 계피동일 필수 특약'},
  ];

  if (hasAge)   cases.push({device:'PC', channel:'옴니',contractor_insured:'계피동일',rider_scope:r1,  expected_outcome:'오류',     test_focus:`가입연령 하한 경계값 미만 입력 시 차단 여부 — ${r1}`});
  if (hasName)  cases.push({device:'Tab',channel:'전자',contractor_insured:'계피동일',rider_scope:rAll,expected_outcome:'오류',     test_focus:'화면 노출 특약명이 상품안 개정본 명칭과 일치하는지 검증'});
  if (hasOblig) cases.push({device:'PC', channel:'옴니',contractor_insured:'계피상이',rider_scope:r1,  expected_outcome:'확인필요', test_focus:'알릴의무 신설 문항 화면 반영 여부 — 옴니·전자 모두'});
  if (hasNew && !hasAge && !hasName)
                cases.push({device:'Tab',channel:'옴니',contractor_insured:'계피동일',rider_scope:r1,  expected_outcome:'확인필요', test_focus:`신설 항목(${r1}) 전 단계 화면 반영 여부 수기 확인`});

  const risks = [];
  if (hasAge)   risks.push('가입연령 경계값 처리 로직 검증 — 하한 미만 입력 시 차단 여부');
  if (hasName)  risks.push('특약명 버전 불일치 위험 — 상품안 개정본 기준 전 화면 재확인');
  if (hasNew)   risks.push('신설 항목 누락 위험 — 가입설계부터 청약 완료까지 전 단계 순차 확인');
  if (hasOblig) risks.push('알릴의무 신규 문항 — 옴니·전자 양 채널에서 독립 검증 필요');
  if (!risks.length) risks.push('상품안 문구와 화면 표시 간 불일치 여부 집중 확인');

  return { checklist, test_cases: cases, risk_points: risks, source: 'rule' };
}

function qcRuleBased(tcId, device, channel, errorTitle, repro, expected, actual) {
  const text = errorTitle + repro.join(' ');
  let sev = 'S2';
  if (['청약 불가','차단','불가','오류','계산','데이터'].some(k => text.includes(k))) sev = 'S1';
  else if (['문구','명칭','표기','노출','UI'].some(k => text.includes(k))) sev = 'S3';
  else if (['색상','폰트','여백','정렬'].some(k => text.includes(k))) sev = 'S4';

  const cat = categorize(text);
  return {
    severity: sev,
    category: cat,
    analysis:
      `${tcId} (${device}·${channel}) 케이스 실행 중 발견된 결함입니다. ` +
      `증상은 「${errorTitle}」으로, 기대 동작과 실제 동작 사이에 명확한 불일치가 확인됐습니다. ` +
      `${cat} 영역의 구현 오류 또는 상품 정보 반영 누락으로 추정되며, 해당 로직 전반 재점검이 필요합니다.`,
    suggestion:
      '담당 개발팀은 해당 화면/로직의 유효성 검사 코드와 DB 반영 시점을 우선 확인해주세요. ' +
      '수정 후 동일 케이스 재실행으로 회귀 여부를 검증하고, ' +
      '관련 영역 단위 테스트를 추가한 뒤 재배포를 권장합니다.',
    related_reqs: [],
    source: 'rule',
  };
}

/* ── 통합 생성 함수 (Gemini 우선 → 규칙 기반 폴백) ─────────────────────────── */
async function generateTestCases(productName, requirements) {
  if (isGeminiAvailable()) {
    try { return await casesGemini(productName, requirements); }
    catch (e) { console.warn('[AI] Gemini 케이스 생성 실패 →', e.message, '규칙 기반으로 폴백'); }
  }
  return casesRuleBased(productName, requirements);
}

async function generateQcDraftAI(tcId, device, channel, errorTitle, repro, expected, actual) {
  if (isGeminiAvailable()) {
    try { return await qcGemini(tcId, device, channel, errorTitle, repro, expected, actual); }
    catch (e) { console.warn('[AI] Gemini QC 생성 실패 →', e.message, '규칙 기반으로 폴백'); }
  }
  return qcRuleBased(tcId, device, channel, errorTitle, repro, expected, actual);
}

/* ═══════════════════════════════════════════════════════════════════════════
   상수 — 데모 및 시뮬레이션 데이터
   ═══════════════════════════════════════════════════════════════════════════ */
const DEMO_CASES_TEMPLATE = [
  {tc_id:'TC01',device:'PC', channel:'옴니',contractor_insured:'계피동일',rider_scope:'필수특약만', expected_outcome:'통과'},
  {tc_id:'TC02',device:'PC', channel:'옴니',contractor_insured:'계피상이',rider_scope:'암특약 전부',expected_outcome:'오류',   error_key:'TC02'},
  {tc_id:'TC03',device:'Tab',channel:'전자',contractor_insured:'계피동일',rider_scope:'모든 특약', expected_outcome:'통과'},
  {tc_id:'TC04',device:'Tab',channel:'옴니',contractor_insured:'계피상이',rider_scope:'모든 특약', expected_outcome:'오류',   error_key:'TC04'},
  {tc_id:'TC05',device:'Tab',channel:'전자',contractor_insured:'계피상이',rider_scope:'암특약 전부',expected_outcome:'확인필요'},
  {tc_id:'TC06',device:'Tab',channel:'옴니',contractor_insured:'계피동일',rider_scope:'필수특약만', expected_outcome:'통과',recheck_needed:true},
  {tc_id:'TC07',device:'PC', channel:'옴니',contractor_insured:'계피동일',rider_scope:'모든 특약', expected_outcome:'통과'},
  {tc_id:'TC08',device:'Tab',channel:'전자',contractor_insured:'계피동일',rider_scope:'필수특약만', expected_outcome:'통과'},
  {tc_id:'TC09',device:'Tab',channel:'옴니',contractor_insured:'계피상이',rider_scope:'암특약 전부',expected_outcome:'통과'},
  {tc_id:'TC10',device:'PC', channel:'옴니',contractor_insured:'계피상이',rider_scope:'모든 특약', expected_outcome:'통과'},
];

const DEMO_ERRORS = {
  TC02: {
    cat:'이번 테스트 특화 · 상품 정보', fail_step:4, policy:'GEA260601234352',
    title:'간암 치료 특약 가입연령 하한이 화면에 반영되지 않음',
    repro:['고객 등록 후 남자 24세 피보험자로 가입설계 진입','신설 「간암 치료」 특약 선택','보험료 계산 → 차단·안내 없이 정상 산출됨'],
    expect:'간암 치료 특약: 남자 25세·여자 30세부터 가입 가능 (하한 미만 선택 차단)',
    actual:'남자 24세에도 간암 치료 특약 선택·계산이 그대로 진행됨',
    has_screenshot:false,
  },
  TC04: {
    cat:'문서 정합성', fail_step:6, policy:'GEA260601234274',
    title:'특약명이 구버전으로 노출됨 (상품안 ↔ 화면 버전 불일치)',
    repro:['딱좋은 0540 건강보험 가입설계 진입','특약 목록에서 암 관련 특약 확인','청약서 발행 단계에서 상품안(PSR) 개정본과 특약명 대조 → 버전 불일치 감지'],
    expect:'상품안 개정 반영 「암진단 III」 특약명 노출',
    actual:'구버전 특약명 「암진단 II」가 그대로 노출됨',
    has_screenshot:true,
  },
};

/* ═══════════════════════════════════════════════════════════════════════════
   직렬화 헬퍼
   ═══════════════════════════════════════════════════════════════════════════ */
function caseDict(c) {
  const r = resultOf(c.id);
  const d = {
    id: c.id, tc_id: c.tc_id, device: c.device, channel: c.channel,
    dev: `${c.device}·${c.channel}`, contractor_insured: c.contractor_insured,
    gp: c.contractor_insured, rider_scope: c.rider_scope, rider: c.rider_scope,
    status: c.status, expected_outcome: c.expected_outcome,
    error_key: c.error_key ?? null, from_custom: !!c.from_custom, recheck_needed: !!c.recheck_needed,
  };
  if (r?.executed_at) d.executed_at = r.executed_at;
  if (r) d.result = resultDict(r);   // 실행된 케이스의 단계/상태 정보 동봉 (단계 표시용)
  return d;
}

function resultDict(r) {
  return {
    status: r.status, fail_step: r.fail_step ?? null,
    error_category: r.error_category ?? null, error_title: r.error_title ?? null,
    repro: r.repro_steps ? JSON.parse(r.repro_steps) : [],
    expect: r.expected ?? null, expected: r.expected ?? null,
    actual: r.actual ?? null, has_screenshot: !!r.has_screenshot,
    policy_number: r.policy_number ?? null, executed_at: r.executed_at ?? null,
  };
}

/* ═══════════════════════════════════════════════════════════════════════════
   엔드포인트 핸들러
   ═══════════════════════════════════════════════════════════════════════════ */

/* ── 세션 ── */
function listSessions() {
  return [...db.sessions]
    .sort((a,b) => (b.created_at||'').localeCompare(a.created_at||''))
    .map(s => {
      const cs = casesOf(s.id);
      const total = cs.length, done = cs.filter(c => c.status !== '미실행').length;
      return {
        id:s.id, name:s.name, product_code:s.product_code, target_product:s.target_product,
        parts:JSON.parse(s.parts||'[]'), start_date:s.start_date, end_date:s.end_date,
        status:s.status, deployed:s.deployed, nl_seed:s.nl_seed||'',
        total_cases:total, done_cases:done, error_count:cs.filter(c=>c.status==='오류').length,
        progress_pct: total > 0 ? Math.round(done/total*100) : 0,
        created_at:s.created_at,
      };
    });
}

function createSession(p) {
  const name  = (p.name||'').trim(), code  = (p.product_code||'').trim();
  const start = (p.start_date||'').trim(), end = (p.end_date||'').trim();
  const parts = p.parts || [];
  if (!name)  throw new ApiError(400, '테스트명을 입력해주세요');
  if (!code)  throw new ApiError(400, '상품코드를 입력해주세요');
  if (!start) throw new ApiError(400, '시작일을 입력해주세요');
  if (!end)   throw new ApiError(400, '종료일을 입력해주세요');
  if (!parts.length) throw new ApiError(400, '참여 파트를 1개 이상 선택해주세요');

  const s = {
    id:nextId('sessions'), name, product_code:code,
    target_product:p.target_product||name, parts:JSON.stringify(parts),
    start_date:start, end_date:end, status:p.status||'진행중',
    deployed:false, nl_seed:'', created_at:nowISO(),
  };
  db.sessions.push(s);
  return { id:s.id, name:s.name };
}

function getSession(id) {
  const s = findSession(id);
  if (!s) throw new ApiError(404, '세션을 찾을 수 없어요');
  const cs = casesOf(s.id), total = cs.length, done = cs.filter(c=>c.status!=='미실행').length;
  return {
    id:s.id, name:s.name, product_code:s.product_code, target_product:s.target_product,
    parts:JSON.parse(s.parts||'[]'), start_date:s.start_date, end_date:s.end_date,
    status:s.status, deployed:s.deployed, nl_seed:s.nl_seed||'',
    total_cases:total, done_cases:done,
    progress_pct: total > 0 ? Math.round(done/total*100) : 0,
  };
}

function deploySession(id) {
  const s = findSession(id);
  if (!s) throw new ApiError(404, '세션을 찾을 수 없어요');
  s.deployed = true;
  return { ok:true, deployed:true };
}

/* ── 케이스 ── */
function listCases(sid) {
  return casesOf(sid).sort((a,b)=>a.tc_id.localeCompare(b.tc_id)).map(caseDict);
}

function addCase(sid, p) {
  const count = casesOf(sid).length;
  const c = {
    id:nextId('cases'), session_id:Number(sid),
    tc_id:`TC${String(count+1).padStart(2,'0')}`,
    device:p.device||'', channel:p.channel||'',
    contractor_insured:p.contractor_insured||'계피동일',
    rider_scope:p.rider_scope||'', status:'미실행',
    expected_outcome:p.expected_outcome||'통과',
    error_key:null, from_custom:true, recheck_needed:false, created_at:nowISO(),
  };
  db.cases.push(c);
  return caseDict(c);
}

/* 큐레이션된 데모 오류가 없는 (자동 생성) 케이스용 — 케이스 조합으로 그럴듯한 오류 데이터 합성 */
function synthesizeError(c) {
  const dev = `${c.device||'PC'}·${c.channel||'옴니'}`;
  const gp  = c.contractor_insured || '계피동일';
  const rider = (c.rider_scope || '').trim();
  const scope = rider || '필수특약';
  const title = `${scope} 부가 조건이 ${dev} 화면에 정상 반영되지 않음`;
  const repro = [
    `${dev} 환경에서 고객 등록 후 가입설계 진입 (${gp})`,
    `특약 구성에서 「${scope}」 선택`,
    `특약 선택 단계 진행 → 해당 조건의 검증·안내가 누락됨`,
  ];
  return {
    cat: `${c.channel||'옴니'}청약 · 특약/청약 정합성`,
    fail_step: 4,
    policy: 'GEA' + (260601000000 + (c.id % 900000)).toString(),
    title,
    repro,
    expect: `「${scope}」 부가 시 ${dev}에서 조건 검증·안내 문구가 정상 노출되어야 함`,
    actual: `${dev}에서 「${scope}」 조건 검증·안내 없이 청약이 그대로 진행됨`,
    has_screenshot: false,
  };
}

function executeCase(sid, cid) {
  const c = db.cases.find(x=>x.id===Number(cid)&&x.session_id===Number(sid));
  if (!c) throw new ApiError(404, '케이스를 찾을 수 없어요');

  const outcome = c.expected_outcome;
  // 데모 큐레이션 오류(error_key) 우선, 없으면 케이스 조합 기반으로 합성
  const errData = outcome==='오류'
    ? (c.error_key ? DEMO_ERRORS[c.error_key] : synthesizeError(c))
    : null;
  // 확인필요: 오류는 아니지만 사람이 확인해야 하는 단계를 표시 (청약서 발행·문구 점검 = 6단계)
  const checkStep = outcome==='확인필요' ? 6 : null;
  db.results = db.results.filter(r=>r.case_id!==c.id);

  const result = {
    id:nextId('results'), case_id:c.id, status:outcome,
    fail_step:      errData?errData.fail_step:checkStep,
    error_category: errData?errData.cat:null,
    error_title:    errData?errData.title:null,
    repro_steps:    errData?JSON.stringify(errData.repro):null,
    expected:       errData?errData.expect:null,
    actual:         errData?errData.actual:null,
    has_screenshot: errData?!!errData.has_screenshot:false,
    policy_number:  errData?errData.policy:null,
    executed_at:    nowISO(),
  };
  db.results.push(result);
  c.status = outcome;
  c.recheck_needed = false;

  const res = resultDict(result);
  res.tc_id = c.tc_id;
  res.error_key = c.error_key ?? null;
  return res;
}

function rerunCase(sid, cid) {
  const c = db.cases.find(x=>x.id===Number(cid)&&x.session_id===Number(sid));
  if (!c) throw new ApiError(404, '케이스를 찾을 수 없어요');
  c.status = '미실행';
  return { tc_id:c.tc_id, status:'미실행' };
}

function getResults(sid) {
  const cs = casesOf(sid).sort((a,b)=>a.tc_id.localeCompare(b.tc_id));
  const allCases = cs.map(caseDict);
  allCases.forEach((cd, i) => {
    const r = resultOf(cs[i].id);
    if (r) cd.result = resultDict(r);   // 오류·확인필요·통과 모두 단계 정보 전달
  });
  return {
    pass:  allCases.filter(c=>c.status==='통과'),
    errors:allCases.filter(c=>c.status==='오류'),
    needs: allCases.filter(c=>c.status==='확인필요'),
  };
}

async function generateCasesHandler(sid) {
  const s = findSession(sid);
  if (!s) throw new ApiError(404, '세션을 찾을 수 없어요');
  const psrs = db.psrRecords.filter(p=>p.session_id===Number(sid));
  if (!psrs.length) throw new ApiError(400, 'PSR 데이터가 없어요. 먼저 PSR Excel을 업로드해주세요.');

  const allReqs = [];
  for (const p of psrs)
    for (const r of db.requirements.filter(rr=>rr.psr_id===p.id))
      allReqs.push({ req_id:r.req_id, req_name:r.req_name, req_content:r.req_content });
  if (!allReqs.length) throw new ApiError(400, '요구사항이 비어 있어요. PSR Excel 형식을 확인해주세요.');

  const aiResult = await generateTestCases(s.name, allReqs);

  const oldIds = new Set(casesOf(sid).map(c=>c.id));
  db.results  = db.results.filter(r=>!oldIds.has(r.case_id));
  db.qcDrafts = db.qcDrafts.filter(q=>!oldIds.has(q.case_id));
  db.cases    = db.cases.filter(c=>c.session_id!==Number(sid));

  const created = (aiResult.test_cases||[]).map((tc, idx) => {
    const c = {
      id:nextId('cases'), session_id:Number(sid),
      tc_id:`TC${String(idx+1).padStart(2,'0')}`,
      device:tc.device||'PC', channel:tc.channel||'옴니',
      contractor_insured:tc.contractor_insured||'계피동일',
      rider_scope:tc.rider_scope||'',
      expected_outcome:tc.expected_outcome||'통과',
      status:'미실행', error_key:null, from_custom:false, recheck_needed:false, created_at:nowISO(),
    };
    db.cases.push(c);
    return caseDict(c);
  });

  return {
    source:aiResult.source||'rule', cases:created,
    checklist:aiResult.checklist||[], risk_points:aiResult.risk_points||[],
    psr_count:psrs.length, req_count:allReqs.length,
  };
}

/* ── 확인 포인트 추출 — 자연어 지시문 → QA 검증 명사구 ──────────────────────── */
const CHECKPOINT_PROMPT = (text) => `너는 QA 테스트 요구사항을 간결하고 명확한 '확인 포인트'로 변환하는 테스트 설계 전문 AI야.
사용자가 자연어로 테스트 지시사항을 입력하면, 아래 규칙에 따라 핵심만 추출하여 명사구 형태로 요약해 줘.

[변환 규칙]
1. 출력 형식: 반드시 "확인 포인트 ① [내용] ② [내용] ..." 형태로 기호를 사용하여 한 줄로 출력할 것.
2. 명사화 및 압축: 사용자의 입력 문장을 그대로 복사하지 말 것. 불필요한 조사와 서술어를 제거하고 핵심 키워드 위주로 최대한 짧게 압축할 것.
3. 종결 어미 통제: 각 포인트의 끝은 반드시 "~ 여부", "~ 노출 여부", "~ 차단 여부", "~ 처리 여부" 등 QA 검증 상태를 나타내는 명사구로 끝맺을 것.
※ 🚫 절대 금지: "~하는지", "~표시되어", "~해야 함" 등 문장형/서술형 어미는 절대 사용하지 말 것

[예시]
입력: 출산•육아휴직 보험료 할인 특약 부가 시 전자청약•옴니청약 발행이 불가한지 확인해야 함. 청약서류 발행 전 안내 모달에도 사유 문구가 표시되어 있어야 함.
출력: 확인 포인트 ① 전자•옴니 청약 발행 차단 여부 ② 청약서류 발행 전 안내 모달 사유 문구 노출 여부

입력: 간암 특약 가입 시 나이가 25세 미만이면 오류 팝업이 뜨고 다음 화면으로 넘어가면 안 돼.
출력: 확인 포인트 ① 가입 연령(25세 미만) 미달 시 오류 팝업 노출 여부 ② 다음 단계 진입 차단 여부

입력: 여성이면 암특약 선택 불가해야 하고 남성이면 간암특약 선택 불가해야 함
출력: 확인 포인트 ① 여성 암특약 선택 차단 여부 ② 남성 간암특약 선택 차단 여부

※ 입력에 조건이 여러 개(예: "~하고", "~하며", "그리고")면 각각을 별도 확인 포인트로 분리할 것. 입력 조건 수와 출력 포인트 수가 일치해야 함.

[입력 데이터]
입력: ${text}
출력:`;

/* "확인 포인트 ① A ② B" 한 줄 → ['A','B'] */
function parseCheckPointLine(line) {
  const body = String(line || '').replace(/^[\s\S]*?확인\s*포인트\s*/, '').replace(/\n[\s\S]*$/, '').trim();
  const parts = body.split(/[①②③④⑤⑥⑦⑧⑨⑩]/).map(s => s.replace(/^[\s,·]+|[\s,·]+$/g, '').trim()).filter(Boolean);
  return parts.length ? parts : (body ? [body] : []);
}

async function checkPointsGemini(text) {
  const raw = await callGemini(CHECKPOINT_PROMPT(text));
  return parseCheckPointLine(raw);
}

/* 규칙 기반 폴백 — 시연 시드 입력은 정확히 매핑, 그 외에는 서술어 제거 명사구화 */
function checkPointsRule(text) {
  const t = String(text || '').replace(/\s+/g, ' ').trim();
  if (/출산.{0,2}육아휴직.*할인\s*특약.*(전자청약|옴니청약).*발행이?\s*불가/.test(t))
    return ['전자•옴니 청약 발행 차단 여부', '청약서류 발행 전 안내 모달 사유 문구 노출 여부'];

  // 문장 부호 + 접속어("~하고 / ~하며 / 그리고")로 분해해 복수 조건을 분리
  const sents = t
    .split(/[.。\n]+|\s*(?:하고|하며|그리고)\s+/)
    .map(s => s.trim())
    .filter(s => s.length >= 4);
  const nounify = (s) => {
    s = s.replace(/(확인해야\s*함|확인이?\s*필요(하다|함)?|확인하고\s*싶어요?|확인해\s*줘|점검해야\s*함|해야\s*함|해야\s*한다|있어야\s*함|있어야\s*한다|되어\s*있어야\s*함)\s*\.?$/g, '').trim();
    s = s.replace(/(되는지|하는지|인지|한지)\s*\.?$/g, '').trim();
    if (/불가|차단|막아|막혀|안\s*되|넘어가면\s*안/.test(s))
      return s.replace(/(이|가)?\s*(불가|차단|막아|막혀).*$/, '').replace(/[\s,·]+$/, '').trim() + ' 차단 여부';
    if (/표시|노출|뜨|보여|나타|팝업/.test(s))
      return s.replace(/(이|가|에|에도)?\s*(표시|노출|뜨|보여|나타).*$/, '').replace(/[\s,·]+$/, '').trim() + ' 노출 여부';
    if (/오류|에러|실패/.test(s))
      return s.replace(/[\s,·]+$/, '').trim() + ' 오류 처리 여부';
    return s.replace(/[\s,·]+$/, '').trim() + ' 정상 처리 여부';
  };
  const pts = sents.map(nounify).filter(Boolean).slice(0, 4);
  return pts.length ? pts : ['입력 시나리오 정상 처리 여부'];
}

async function extractCheckPoints(text) {
  if (isGeminiAvailable()) {
    try {
      const r = await checkPointsGemini(text);
      if (r.length) return r;
    } catch (e) {
      console.warn('[AI] 확인 포인트 추출 실패 →', e.message, '· 규칙 기반으로 폴백');
    }
  }
  return checkPointsRule(text);
}

async function caseFromTextHandler(sid, p) {
  const s = findSession(sid);
  if (!s) throw new ApiError(404, '세션을 찾을 수 없어요');
  const text = (p.text||'').trim();
  if (!text) throw new ApiError(400, '확인할 내용을 입력해주세요');

  const [result, checkPoints] = await Promise.all([
    generateTestCases(s.name, [{ req_id:'USR-001', req_name:'사용자 직접 입력 확인 사항', req_content:text }]),
    extractCheckPoints(text),
  ]);
  const tc = (result.test_cases||[{}])[0]||{};

  const c = {
    id:nextId('cases'), session_id:Number(sid),
    tc_id:`TC${String(casesOf(sid).length+1).padStart(2,'0')}`,
    device:tc.device||'PC', channel:tc.channel||'옴니',
    contractor_insured:tc.contractor_insured||'계피동일',
    rider_scope:tc.rider_scope||text.slice(0,40),
    expected_outcome:tc.expected_outcome||'확인필요',
    status:'미실행', error_key:null, from_custom:true, recheck_needed:false, created_at:nowISO(),
  };
  db.cases.push(c);
  return { case:caseDict(c), source:result.source||'rule', focus:tc.test_focus||'', check_points:checkPoints };
}

/* ── QC ── */
function getQc(cid) {
  const qc = qcOf(cid);
  return qc ? { id:qc.id, content:qc.content?JSON.parse(qc.content):{}, published:qc.published } : null;
}

function saveQc(cid, p) {
  let qc = qcOf(cid);
  if (qc) {
    qc.content=JSON.stringify(p); qc.published=p.published||false; qc.updated_at=nowISO();
  } else {
    qc={id:nextId('qc'),case_id:Number(cid),author:'박재윤',
        content:JSON.stringify(p),published:p.published||false,
        created_at:nowISO(),updated_at:nowISO()};
    db.qcDrafts.push(qc);
  }
  return { ok:true };
}

function publishQc(cid) {
  const qc = qcOf(cid);
  if (!qc) throw new ApiError(404, 'QC 초안이 없어요');
  qc.published=true; qc.updated_at=nowISO();
  return { ok:true, case_id:Number(cid) };
}

async function generateQcDraftHandler(cid) {
  const c = findCase(cid);
  if (!c) throw new ApiError(404, '케이스를 찾을 수 없어요');
  const r = resultOf(c.id);
  if (!r||r.status!=='오류') throw new ApiError(400, '오류 결과가 없는 케이스입니다');
  const repro = r.repro_steps ? JSON.parse(r.repro_steps) : [];
  return generateQcDraftAI(c.tc_id, c.device, c.channel, r.error_title||'', repro, r.expected||'', r.actual||'');
}

function updateResult(cid, p) {
  const r = resultOf(cid);
  if (!r) throw new ApiError(404, '결과 데이터가 없어요');
  if ('error_title' in p) r.error_title = p.error_title;
  if ('repro' in p) {
    let repro = p.repro;
    if (typeof repro==='string') repro = repro.split('\n').map(l=>l.trim()).filter(Boolean);
    r.repro_steps = JSON.stringify(repro);
  }
  if ('expected' in p) r.expected = p.expected;
  if ('actual' in p)   r.actual = p.actual;
  return resultDict(r);
}

/* ── PSR / 문서 ── */
function getPsr(sid) {
  return db.psrRecords.filter(p=>p.session_id===Number(sid)).map(p=>({
    id:p.id, psr_number:p.psr_number, psr_title:p.psr_title,
    author_team:p.author_team, author_name:p.author_name, created_date:p.created_date,
    requirements:db.requirements.filter(r=>r.psr_id===p.id)
      .map(r=>({req_id:r.req_id,req_name:r.req_name,req_content:r.req_content})),
  }));
}

function listDocuments(sid) {
  return db.documents.filter(d=>d.session_id===Number(sid)).map(d=>({
    id:d.id, filename:d.filename, size_bytes:d.size_bytes,
    text_length:(d.text_excerpt||'').length, uploaded_at:d.uploaded_at,
  }));
}

function deleteDocument(docId) {
  db.documents = db.documents.filter(d=>d.id!==Number(docId));
  return { ok:true };
}

/* ── 데모 시드 ── */
function seedDemo() {
  if (db.sessions.find(s=>s.name==='KB 딱좋은 0540 건강보험')) {
    const s = db.sessions.find(s=>s.status==='진행중');
    return { message:'이미 데모 데이터가 있어요', active_session_id:s?s.id:null };
  }

  const active = {
    id:nextId('sessions'),
    name:'KB 딱좋은 0540 건강보험',
    product_code:'375000015',
    target_product:'딱좋은 0540 건강보험 무배당(일반심사형)(해약환급금 미지급형)(납입면제형)(사망보장형)',
    parts:JSON.stringify(['IT개발파트','IT기획운영파트','계약관리파트','계약심사파트','고객컨택파트','보험금파트','상품운영파트','신계약업무파트','영업추진파트','혁신상품파트']),
    start_date:'2026-06-04', end_date:'2026-06-12', status:'진행중', deployed:true,
    nl_seed:'출산·육아휴직 보험료 할인 특약 부가 시 전자청약·옴니청약 발행이 불가한지 확인해야 함. 청약서류 발행 전 안내 모달에도 사유 문구가 표시되어 있어야 함.',
    created_at:nowISO(),
  };
  db.sessions.push(active);

  const pre = {TC01:'통과',TC02:'오류',TC03:'통과',TC04:'오류',TC05:'확인필요',TC06:'재실행'};
  for (const t of DEMO_CASES_TEMPLATE) {
    const status = pre[t.tc_id]||'미실행';
    const c = {
      id:nextId('cases'), session_id:active.id, tc_id:t.tc_id,
      device:t.device, channel:t.channel, contractor_insured:t.contractor_insured,
      rider_scope:t.rider_scope, status, expected_outcome:t.expected_outcome,
      error_key:t.error_key||null, from_custom:false,
      recheck_needed:!!(t.recheck_needed&&status==='재실행'), created_at:nowISO(),
    };
    db.cases.push(c);
    if (['통과','오류','확인필요'].includes(status)) {
      const errData = (status==='오류'&&t.error_key) ? DEMO_ERRORS[t.error_key] : null;
      const checkStep = status==='확인필요' ? 6 : null;
      db.results.push({
        id:nextId('results'), case_id:c.id, status,
        fail_step:      errData?errData.fail_step:checkStep,
        error_category: errData?errData.cat:null,
        error_title:    errData?errData.title:null,
        repro_steps:    errData?JSON.stringify(errData.repro):null,
        expected:       errData?errData.expect:null,
        actual:         errData?errData.actual:null,
        has_screenshot: errData?!!errData.has_screenshot:false,
        policy_number:  errData?errData.policy:null,
        executed_at:    nowISO(),
      });
    }
  }

  db.sessions.push({
    id:nextId('sessions'), name:'넥스트 레벨업 연금보험 무배당',
    product_code:'374000088', target_product:'넥스트 레벨업 연금보험 무배당',
    parts:JSON.stringify(['상품개발파트','상품운영파트','계약심사파트','IT개발파트','계약관리파트']),
    start_date:'2026-05-20', end_date:'2026-05-24', status:'완료', deployed:true,
    nl_seed:'', created_at:nowISO(),
  });
  db.sessions.push({
    id:nextId('sessions'), name:'달러평생소득변액연금보험 무배당',
    product_code:'376000002', target_product:'달러평생소득변액연금보험 무배당',
    parts:JSON.stringify(['상품개발파트','리스크관리파트','IT개발파트','계약심사파트']),
    start_date:'2026-06-15', end_date:'2026-06-19', status:'예정', deployed:true,
    nl_seed:'', created_at:nowISO(),
  });

  return { message:'데모 데이터 생성 완료', active_session_id:active.id };
}

/* ═══════════════════════════════════════════════════════════════════════════
   Excel 업로드 (SheetJS)
   ═══════════════════════════════════════════════════════════════════════════ */
function readSheetRows(arrayBuffer) {
  if (typeof XLSX === 'undefined') throw new ApiError(500, 'Excel 라이브러리(SheetJS) 로드 실패');
  const wb = XLSX.read(arrayBuffer, { type:'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval:'', raw:false })
    .map(row => { const o={}; for (const k in row) o[String(k).trim()]=row[k]; return o; });
}
const cell = (row, key) => String(row[key]??'').trim();

async function uploadPSR(sessionId, file) {
  const name = (file.name||'').toLowerCase();
  if (!name.endsWith('.xlsx')&&!name.endsWith('.xls'))
    throw new ApiError(400, 'Excel 파일만 지원합니다 (.xlsx / .xls)');

  const rows = readSheetRows(await file.arrayBuffer());

  // 기존 PSR 삭제 (재업로드)
  const oldIds = new Set(db.psrRecords.filter(p=>p.session_id===Number(sessionId)).map(p=>p.id));
  db.requirements = db.requirements.filter(r=>!oldIds.has(r.psr_id));
  db.psrRecords   = db.psrRecords.filter(p=>p.session_id!==Number(sessionId));

  const psrMap = {}, psrOrder = [];
  let reqCount = 0;

  for (const row of rows) {
    const psrNum = cell(row,'PSR번호');
    if (!psrNum) {
      if (psrOrder.length) {
        const last = psrMap[psrOrder[psrOrder.length-1]];
        const reqId = cell(row,'요구사항ID');
        if (reqId) { db.requirements.push({id:nextId('req'),psr_id:last.id,req_id:reqId,req_name:cell(row,'요구사항명'),req_content:cell(row,'요구사항 내용')}); reqCount++; }
      }
      continue;
    }
    if (!(psrNum in psrMap)) {
      const authorRaw = cell(row,'작성팀/작성자');
      const parts = authorRaw.includes('/') ? authorRaw.split('/').map(s=>s.trim()) : [authorRaw,''];
      const psr = {id:nextId('psr'),session_id:Number(sessionId),psr_number:psrNum,psr_title:cell(row,'PSR제목'),author_team:parts[0],author_name:parts[1]||'',created_date:cell(row,'작성일')};
      db.psrRecords.push(psr); psrMap[psrNum]=psr; psrOrder.push(psrNum);
    }
    const reqId = cell(row,'요구사항ID');
    if (reqId) { db.requirements.push({id:nextId('req'),psr_id:psrMap[psrNum].id,req_id:reqId,req_name:cell(row,'요구사항명'),req_content:cell(row,'요구사항 내용')}); reqCount++; }
  }

  save();
  return { psr_count:psrOrder.length, req_count:reqCount, message:`PSR ${psrOrder.length}건, 요구사항 ${reqCount}건 업로드 완료` };
}

async function uploadDoc(sessionId, file) {
  const fname = file.name||'document', mt = file.type||'';
  let text = '';
  try {
    if (fname.toLowerCase().endsWith('.txt')||mt.startsWith('text/')) text = await file.text();
    else text = `[${fname}] 텍스트 추출은 서버리스 환경에서 미지원 — 파일은 첨부됐어요.`;
  } catch (e) { text = `[추출 오류] ${e.message||e}`; }

  const buf = await file.arrayBuffer();
  text = (text||'').trim().slice(0,20000);
  const doc = {id:nextId('docs'),session_id:Number(sessionId),filename:fname,mimetype:mt,size_bytes:buf.byteLength,text_excerpt:text,uploaded_at:nowISO()};
  db.documents.push(doc);
  save();
  return { id:doc.id, filename:fname, size_bytes:buf.byteLength, text_length:text.length, preview:text.slice(0,300) };
}

/* ═══════════════════════════════════════════════════════════════════════════
   라우터 — main.py 경로/메서드 매핑
   ═══════════════════════════════════════════════════════════════════════════ */
async function handle(method, path, body) {
  method = method.toUpperCase();
  const p   = path.split('?')[0];
  const seg = p.split('/').filter(Boolean);

  const M = (m, re) => method === m && re.test(p);
  let out;

  if      (M('POST', /^\/dev\/seed$/))                          out = seedDemo();
  else if (M('GET',  /^\/ai\/status$/))                         out = { gemini_available:isGeminiAvailable(), model:isGeminiAvailable()?'gemini-2.0-flash':null, message:isGeminiAvailable()?'Gemini API 연결됨 — AI 케이스 생성 모드':'규칙 기반 모드 (AI 설정에서 Gemini API 키를 입력하면 AI 모드로 전환)' };
  else if (M('GET',  /^\/tests$/))                              out = listSessions();
  else if (M('POST', /^\/tests$/))                              out = createSession(body||{});
  else if (M('GET',  /^\/tests\/\d+$/))                         out = getSession(seg[1]);
  else if (M('POST', /^\/tests\/\d+\/deploy$/))                 out = deploySession(seg[1]);
  else if (M('GET',  /^\/tests\/\d+\/cases$/))                  out = listCases(seg[1]);
  else if (M('POST', /^\/tests\/\d+\/cases$/))                  out = addCase(seg[1], body||{});
  else if (M('POST', /^\/tests\/\d+\/cases\/\d+\/execute$/))    out = executeCase(seg[1], seg[3]);
  else if (M('POST', /^\/tests\/\d+\/cases\/\d+\/rerun$/))      out = rerunCase(seg[1], seg[3]);
  else if (M('GET',  /^\/tests\/\d+\/results$/))                out = getResults(seg[1]);
  else if (M('POST', /^\/tests\/\d+\/generate$/))               out = await generateCasesHandler(seg[1]);
  else if (M('POST', /^\/tests\/\d+\/cases\/from-text$/))       out = await caseFromTextHandler(seg[1], body||{});
  else if (M('GET',  /^\/psr\/\d+$/))                           out = getPsr(seg[1]);
  else if (M('GET',  /^\/qc\/\d+$/))                            out = getQc(seg[1]);
  else if (M('POST', /^\/qc\/\d+$/))                            out = saveQc(seg[1], body||{});
  else if (M('POST', /^\/qc\/\d+\/publish$/))                   out = publishQc(seg[1]);
  else if (M('POST', /^\/qc\/\d+\/generate-draft$/))            out = await generateQcDraftHandler(seg[1]);
  else if (M('PATCH',/^\/results\/\d+$/))                       out = updateResult(seg[1], body||{});
  else if (M('GET',  /^\/documents\/\d+$/))                     out = listDocuments(seg[1]);
  else if (M('DELETE',/^\/documents\/\d+$/))                    out = deleteDocument(seg[1]);
  else throw new ApiError(404, `알 수 없는 경로: ${method} ${p}`);

  if (method !== 'GET') save();
  return out;
}

/* ═══════════════════════════════════════════════════════════════════════════
   전역 노출
   ═══════════════════════════════════════════════════════════════════════════ */
window.LocalAPI = {
  handle,
  uploadPSR,
  uploadDoc,
  setGeminiKey,
  getGeminiKey,
  isGeminiAvailable,
  reset() { db = emptyDB(); save(); },   // 디버그용: localStorage 초기화
  _db: () => db,
};

})();
