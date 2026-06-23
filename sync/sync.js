/**
 * 구글 시트 → coway_edu_portal.html 동기화 스크립트
 * 실행: node sync.js  (또는 동기화.bat 더블클릭)
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const HTML_PATH = path.join(__dirname, '..', 'coway_edu_portal.html');
const CSV_PATH = path.join(__dirname, 'data.csv.csv');

const GAS_URL = 'https://script.google.com/macros/s/AKfycbx6xPSeHA76yQwHEYPH_Yfi5VoZxfN18weHTF9oO5QGPKmTJiDnVgPeGWTf22qiRRqN/exec';

// HTML에서 API 키와 스프레드시트 ID를 읽어옴
function readApiConfig() {
  const html = fs.readFileSync(HTML_PATH, 'utf8');
  const keyMatch = html.match(/const SHEETS_API_KEY\s*=\s*'([^']+)'/);
  const idMatch = html.match(/const SPREADSHEET_ID\s*=\s*'([^']+)'/);
  const gidMatch = html.match(/const SHEET_GID\s*=\s*(\d+)/);
  return {
    apiKey: keyMatch ? keyMatch[1] : null,
    spreadsheetId: idMatch ? idMatch[1] : null,
    sheetGid: gidMatch ? parseInt(gidMatch[1]) : null,
  };
}

// Apps Script 웹앱에서 JSON 데이터 가져오기
async function fetchFromGAS() {
  const json = await httpsGet(GAS_URL);
  const data = JSON.parse(json);
  if (data.error) throw new Error('GAS error: ' + data.error);
  return data.data || [];
}

function httpsGet(url) {
  return new Promise((resolve, reject) => {
    const follow = (u) => {
      https.get(u, (res) => {
        // 리다이렉트 처리 (301/302/307)
        if ([301, 302, 307, 308].includes(res.statusCode) && res.headers.location) {
          return follow(res.headers.location);
        }
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => {
          if (res.statusCode !== 200) {
            reject(new Error(`HTTP ${res.statusCode}: ${data.slice(0, 200)}`));
          } else {
            resolve(data);
          }
        });
      }).on('error', reject);
    };
    follow(url);
  });
}

// API 키 없이 공개 구글 시트에서 CSV 직접 다운로드
async function fetchCsvFromGoogleSheets(spreadsheetId, sheetGid) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${sheetGid}`;
  const csv = await httpsGet(url);
  if (!csv || csv.trim().length === 0) throw new Error('빈 응답');
  // HTML 오류 페이지 감지 (로그인 페이지로 리다이렉트된 경우)
  if (csv.trimStart().startsWith('<!')) throw new Error('시트가 비공개이거나 접근 불가');
  return csv;
}

const CAT_CLASS_MAP = {
  '온보딩': 'cat-onboarding',
  '지식공유': 'cat-knowledge',
  '지식공유회': 'cat-knowledge',
  'IP전략': 'cat-ip',
  '특허 온라인 교육': 'cat-ip',
  '실패사례': 'cat-failure',
  'PM기본': 'cat-pm',
  'PM교육': 'cat-pm',
  '직무실무': 'cat-failure',
  '제품실무': 'cat-failure',
  '제품 실무': 'cat-failure',
  'AI교육': 'cat-knowledge',
};

function getCatClass(cat) {
  return CAT_CLASS_MAP[cat] || 'cat-knowledge';
}

function extractMonth(dateStr) {
  if (!dateStr) return null;
  const m4 = dateStr.match(/20\d{2}[.\-\/](\d{1,2})[.\-\/]/);
  if (m4) return parseInt(m4[1]);
  const m2 = dateStr.match(/^(\d{1,2})[\/\-]/);
  if (m2) return parseInt(m2[1]);
  const mk = dateStr.match(/(\d{1,2})월/);
  if (mk) return parseInt(mk[1]);
  return null;
}

function esc(s) {
  return String(s || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\r?\n/g, ' ');
}

function escHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// "이름(소속팀), 이름(소속팀)" 형식의 명단 파싱 → JS 객체 리터럴 문자열 반환
// 괄호 없이 공백이 포함된 항목(설명 텍스트 등)은 이름으로 간주하지 않음
function parseAttendees(str, status = 'done') {
  if (!str || !str.trim()) return '';
  const items = str.split(',').map(s => {
    s = s.trim();
    if (!s) return null;
    const m = s.match(/^(.+?)\((.+?)\)\s*$/);
    if (m) return `{name:'${esc(m[1].trim())}', dept:'${esc(m[2].trim())}', status:'${status}'}`;
    if (s.includes(' ')) return null;
    return `{name:'${esc(s)}', dept:'', status:'${status}'}`;
  }).filter(Boolean);
  return items.join(', ');
}


// 날짜 문자열이 오늘보다 과거인지 판단
function isCourseInPast(dateStr) {
  const m = dateStr.match(/20(\d{2})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (!m) return false;
  const courseDate = new Date(2000 + parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  courseDate.setHours(23, 59, 59);
  return courseDate < new Date();
}

// courses JS 코드 생성 — CSV 데이터 기반으로 전체 재생성
function generateCoursesBlock(rows) {
  let id = 1;
  const lines = rows
    .filter(row => row['교육명'])
    .map(row => {
      const cat = row['구분'] || '';
      const date = row['일시'] || '';
      const time = row['시간'] || '';
      const dateStr = time ? `${date} ${time}` : date;

      const past = isCourseInPast(date);

      // 이수자(수료자) 명단 + 현장 추가 이수자 명단
      const attendeesStr = parseAttendees(
        row['이수자(수료자) 명단'] || row['이수자(수료자)명단'] || row['이수자 명단'] || row['이수자명단'] || ''
      );
      const extraStr = parseAttendees(
        row['현장 추가 이수자(수료자) 명단'] || row['현장추가이수자명단'] || '', 'extra'
      );
      const allAttendeeParts = [attendeesStr, extraStr].filter(Boolean);
      const attendeesArr = allAttendeeParts.length ? `[${allAttendeeParts.join(', ')}]` : '[]';
      // 이수자 수: 명시 컬럼 우선, 없으면 명단+현장 추가 합산
      const extraCount = parseInt(row['현장 추가 이수자 (명)'] || row['현장추가이수자(명)'] || '') ||
                         (extraStr ? extraStr.split('status:').length - 1 : 0);
      const doneCount = (parseInt(row['이수자 (명)'] || row['이수자(명)'] || '') ||
                        (attendeesStr ? attendeesStr.split('status:').length - 1 : 0)) + extraCount;

      // 사전 신청자 명단 + 인원수
      const preStr = parseAttendees(
        row['사전 신청자 명단'] || row['사전신청자 명단'] || row['사전신청자명단'] || ''
      );
      const preArr = preStr ? `[${preStr}]` : '[]';
      // 사전 신청자 수: 명시 컬럼 우선, 없으면 명단에서 카운트 (설명 텍스트는 이미 제외됨)
      const preCount = parseInt(row['사전 신청자 (명)'] || row['사전신청자(명)'] || '') ||
                       (preStr ? preStr.split('status:').length - 1 : 0);

      // 교육비 집행 유무 (CSV 직접 반영)
      const budgetRaw = row['교육비 집행 유무'] || '';
      const budget = budgetRaw.includes('완료') ? '집행완료' : '미집행';

      // 폐강/취소 여부
      const progRaw = row['교육 진행 유무'] || '';
      const canceled = progRaw.includes('폐강') || progRaw.includes('취소');

      // 교육 진행 상태
      const status = (past || budgetRaw.includes('완료')) ? 'complete' : 'progress';

      // 신청 URL (CSV 컬럼, 없으면 기본 주소)
      const applyUrl = (row['지금 바로 신청하기 (사이트 주소)'] || '').trim() || 'https://edu.coway.com/emp';

      return `      { id:${id++}, cat:'${esc(cat)}', catClass:'${getCatClass(cat)}', ` +
        `name:'${esc(row['교육명'])}', pre:${preCount}, done:${doneCount}, ` +
        `date:'${esc(dateStr.trim())}', location:'${esc(row['장소'])}', ` +
        `instructor:'${esc(row['강사'])}', target:'전체', status:'${status}', canceled:${canceled}, budget:'${budget}', ` +
        `applyUrl:'${esc(applyUrl)}', attendees:${attendeesArr}, preAttendees:${preArr} }`;
    });

  return '    // §COURSES_START§\n    const courses = [\n' +
         lines.join(',\n') +
         '\n    ];\n    // §COURSES_END§';
}

// 종합 대시보드 HTML 생성
function generateDashboardBlock(rows) {
  const validRows = rows.filter(r => r['교육명']);
  const total = validRows.length;

  // 완료 / 예정 분류
  const completedRows = validRows.filter(r =>
    isCourseInPast(r['일시'] || '') || (r['교육비 집행 유무'] || '').includes('완료')
  );
  const completedCount = completedRows.length;
  const progressCount = total - completedCount;

  // 총 이수 인원
  const totalDone = validRows.reduce((s, r) => s + (parseInt(r['이수자 (명)'] || '') || 0), 0);

  // 평균 이수율 (사전신청 대비)
  const rateRows = completedRows.filter(r =>
    parseInt(r['이수자 (명)'] || '') > 0 && parseInt(r['사전 신청자 (명)'] || '') > 0
  );
  let avgRate = 0;
  if (rateRows.length) {
    const sum = rateRows.reduce((s, r) =>
      s + parseInt(r['이수자 (명)']) / parseInt(r['사전 신청자 (명)']) * 100, 0);
    avgRate = Math.round(sum / rateRows.length);
  }

  // 카테고리별 과정 수
  const catCounts = {};
  validRows.forEach(r => {
    const cat = r['구분'] || '기타';
    catCounts[cat] = (catCounts[cat] || 0) + 1;
  });
  const catEntries = Object.entries(catCounts).sort((a, b) => b[1] - a[1]);

  // 도넛 차트 SVG
  const CIRC = 314.16;
  const CAT_COLORS = ['#1565C0', '#E53935', '#7B1FA2', '#039BE5', '#F57C00', '#00ACC1', '#43A047', '#FF8F00', '#D84315'];
  let donutOffset = 0;
  const donutCircles = catEntries.map(([, count], i) => {
    const len = (count / total) * CIRC;
    const c = `                    <circle cx="70" cy="70" r="50" fill="none" stroke="${CAT_COLORS[i % CAT_COLORS.length]}" stroke-width="22" stroke-dasharray="${len.toFixed(1)} ${CIRC}" stroke-dashoffset="${donutOffset === 0 ? 0 : (-donutOffset).toFixed(1)}" />`;
    donutOffset += len;
    return c;
  }).join('\n');
  const legendItems = catEntries.map(([cat, count], i) =>
    `                  <div class="legend-item"><div class="legend-dot" style="background:${CAT_COLORS[i % CAT_COLORS.length]}"></div>${escHtml(cat)} (${count}개)</div>`
  ).join('\n');

  // Top 5 과정 — 이수자/사전신청자 구분 집계, 합산 기준 정렬
  const courseRanking = validRows
    .map(r => ({
      name: r['교육명'] || '',
      done: parseInt(r['이수자 (명)'] || '') || 0,
      pre:  parseInt(r['사전 신청자 (명)'] || '') || 0,
    }))
    .filter(r => r.done + r.pre > 0)
    .sort((a, b) => b.done - a.done || (b.pre - a.pre))
    .slice(0, 5);
  const maxTotal = courseRanking[0] ? Math.max(courseRanking[0].done, 1) : 1;
  const barItems = courseRanking.length > 0
    ? courseRanking.map(c => {
        const total = c.done + c.pre;
        const doneW = Math.round(c.done / maxTotal * 100);
        const preW  = Math.round(c.pre  / maxTotal * 100);
        return `              <div class="bar-item">
                <div class="bar-label" style="display:flex;justify-content:space-between;align-items:baseline">
                  <span>${escHtml(c.name)}</span>
                  <span style="font-size:11px;white-space:nowrap;margin-left:8px">
                    <span style="color:#1565C0;font-weight:700">이수 ${c.done}명</span>
                    <span style="color:#aaa;margin:0 3px">/</span>
                    <span style="color:#43A047;font-weight:700">사전 ${c.pre}명</span>
                  </span>
                </div>
                <div class="bar-track" style="display:flex;gap:2px">
                  <div class="bar-fill" style="width:${doneW}%;background:#1565C0;border-radius:3px 0 0 3px"></div>
                  <div class="bar-fill" style="width:${preW}%;background:#43A047;border-radius:0 3px 3px 0;opacity:0.7"></div>
                </div>
              </div>`;
      }).join('\n')
    : '              <div style="color:var(--text-muted);font-size:13px;padding:16px 0">데이터가 없습니다.</div>';

  // 소속팀 랭킹 — 전 과정 통합 합산
  // CSV 오타 팀명 → 표준 팀명 매핑 (여기에 추가 가능)
  const TEAM_ALIASES = {
    'TechTalenTF':  'TechTalentTF',
    'TechTelentTF': 'TechTalentTF',
  };

  function canonicalDept(d) {
    d = d.trim().replace(/\s+/g, ' ');
    return TEAM_ALIASES[d] || d;
  }

  const deptAccum = {}; // canonicalName -> { done, pre }

  function addDeptCount(raw, field) {
    (raw || '').split(',').forEach(item => {
      item = item.trim();
      if (!item) return;
      const m = item.match(/^(.+?)\((.+?)\)\s*$/);
      if (!m) return;
      const dept = canonicalDept(m[2]);
      if (!dept) return;
      if (!deptAccum[dept]) deptAccum[dept] = { done: 0, pre: 0 };
      deptAccum[dept][field]++;
    });
  }

  validRows.forEach(r => {
    addDeptCount(r['이수자(수료자) 명단'] || r['이수자(수료자)명단'] || r['이수자 명단'] || '', 'done');
    addDeptCount(r['사전 신청자 명단'] || r['사전신청자 명단'] || '', 'pre');
  });

  const deptRanking = Object.entries(deptAccum)
    .map(([dept, v]) => ({ display: dept, done: v.done, pre: v.pre }))
    .sort((a, b) => b.done - a.done || b.pre - a.pre)
    .slice(0, 5);

  const rankClasses = ['rank-1', 'rank-2', 'rank-3', 'rank-other', 'rank-other'];
  const rankItems = deptRanking.length > 0
    ? deptRanking.map(({ display, done, pre }, i) => `
                <div class="rank-item">
                  <div class="rank-num ${rankClasses[i]}">${i + 1}</div>
                  <div class="rank-name">${escHtml(display)}</div>
                  <div class="rank-count" style="font-size:11px">
                    <span style="color:#43A047;font-weight:700">사전 신청 ${pre}명</span>
                    <span style="color:#aaa;margin:0 3px">/</span>
                    <span style="color:#1565C0;font-weight:700">이수 ${done}명</span>
                  </div>
                </div>`).join('')
    : '<div style="color:var(--text-muted);font-size:13px;padding:16px 0">참여 데이터가 없습니다.</div>';

  // 예산 집행 현황
  const budgetDone = validRows.filter(r => (r['교육비 집행 유무'] || '').includes('완료')).length;
  // 이미 완료된 과정 중 교육비 미집행 건수
  const budgetPendingCompleted = validRows.filter(r =>
    isCourseInPast(r['일시'] || '') &&
    !(r['교육비 집행 유무'] || '').includes('완료')
  ).length;

  // 이슈 감지 ① 폐강 과정: 과거 날짜 + 신청자 1~7명 + 이수자 0명, 또는 진행유무에 폐강/취소 명시
  const cancelledCourses = validRows.filter(r => {
    const pre  = parseInt(r['사전 신청자 (명)'] || '') || 0;
    const done = parseInt(r['이수자 (명)'] || '') || 0;
    const prog = r['교육 진행 유무'] || '';
    const past = isCourseInPast(r['일시'] || '');
    return prog.includes('폐강') || prog.includes('취소') ||
           (past && pre > 0 && pre < 8 && done === 0);
  });

  // 이슈 감지 ② 신청 미달 위험 과정: 미래 일정 + 신청자 1~7명
  const atRiskCourses = validRows.filter(r => {
    const pre  = parseInt(r['사전 신청자 (명)'] || '') || 0;
    const done = parseInt(r['이수자 (명)'] || '') || 0;
    const prog = r['교육 진행 유무'] || '';
    const past = isCourseInPast(r['일시'] || '');
    const 상시 = (r['일시'] || '').includes('상시');
    return !past && !상시 && pre > 0 && pre < 8 && done === 0 &&
           !prog.includes('폐강') && !prog.includes('취소');
  });

  // 오피니언 아이템 동적 생성
  const opinionItems = [];

  opinionItems.push(`
                <div class="opinion-item">
                  <div class="opinion-icon">✅</div>
                  <div class="opinion-text"><strong>수료 완료 과정:</strong> 현재까지 총 <strong>${completedCount}개</strong> 과정이 수료 완료되었으며, 누적 이수자 <strong>${totalDone}명</strong>이 확정되었습니다.</div>
                </div>`);

  opinionItems.push(`
                <div class="opinion-item">
                  <div class="opinion-icon">✅</div>
                  <div class="opinion-text"><strong>이수율 달성 현황:</strong> 수료 과정의 사전 신청자 대비 평균 이수율이 <strong>${avgRate}%</strong>로 집계되었습니다.</div>
                </div>`);

  opinionItems.push(`
                <div class="opinion-item">
                  <div class="opinion-icon">${budgetPendingCompleted > 0 ? '⚠️' : '✅'}</div>
                  <div class="opinion-text"><strong>예산 집행 현황:</strong> 전체 ${total}개 과정 중 집행 완료 <strong>${budgetDone}개</strong>${budgetPendingCompleted > 0 ? `, 교육 완료 후 미집행 <strong style="color:#E53935">${budgetPendingCompleted}건</strong> — 예산 집행 처리가 필요합니다.` : ' — 완료 과정의 예산이 모두 집행되었습니다.'}</div>
                </div>`);

  if (cancelledCourses.length > 0) {
    const names = cancelledCourses.map(r => `<strong>${escHtml(r['교육명'])}</strong>`).join(', ');
    opinionItems.push(`
                <div class="opinion-item">
                  <div class="opinion-icon">🚫</div>
                  <div class="opinion-text"><strong>폐강 과정 (${cancelledCourses.length}건):</strong> 최소 신청인원(8명) 미달로 ${names} 과정이 폐강 처리되었습니다. 재개설 여부 검토가 필요합니다.</div>
                </div>`);
  }

  if (atRiskCourses.length > 0) {
    const names = atRiskCourses.map(r => {
      const pre = parseInt(r['사전 신청자 (명)'] || '') || 0;
      return `<strong>${escHtml(r['교육명'])}</strong> (현재 ${pre}명)`;
    }).join(', ');
    opinionItems.push(`
                <div class="opinion-item">
                  <div class="opinion-icon">⚠️</div>
                  <div class="opinion-text"><strong>신청 미달 위험 과정 (${atRiskCourses.length}건):</strong> 최소 신청인원(8명)에 미치지 못한 예정 과정이 있습니다. ${names} — 추가 홍보 또는 일정 조정이 필요합니다.</div>
                </div>`);
  }

  const opinionHTML = opinionItems.join('');

  const syncDate = new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' });

  return `        <!-- §DASHBOARD_START§ -->
        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-icon blue">📚</div>
            <div>
              <div class="stat-label">총 개설 과정</div>
              <div class="stat-value">${total}<span style="font-size:14px;font-weight:500;color:var(--text-secondary)">개</span></div>
              <div class="stat-sub">상/하반기 통합</div>
            </div>
          </div>
          <div class="stat-card">
            <div class="stat-icon red">✅</div>
            <div>
              <div class="stat-label">진행완료 / 진행예정</div>
              <div class="stat-value">${completedCount}<span style="font-size:14px;color:var(--text-secondary)">/</span>${progressCount}<span style="font-size:14px;font-weight:500;color:var(--text-secondary)">개</span></div>
              <div class="stat-sub">일정 기준 자동 분류</div>
            </div>
          </div>
          <div class="stat-card">
            <div class="stat-icon green">🎓</div>
            <div>
              <div class="stat-label">실제 이수 완료 인원</div>
              <div class="stat-value">${totalDone}<span style="font-size:14px;font-weight:500;color:var(--text-secondary)">명</span></div>
              <div class="stat-sub">CSV 이수자 명단 기준</div>
            </div>
          </div>
          <div class="stat-card">
            <div class="stat-icon orange">%</div>
            <div>
              <div class="stat-label">평균 수료 이수율</div>
              <div class="stat-value">${avgRate}<span style="font-size:14px;font-weight:500;color:var(--text-secondary)">%</span></div>
              <div class="stat-sub">사전 신청 대비 이수율</div>
            </div>
          </div>
        </div>

        <div class="two-col">
          <div class="card">
            <div class="card-header">
              <div class="card-title"><span class="card-title-icon">🥧</span> 카테고리별 과정 분포</div>
            </div>
            <div class="card-body">
              <div class="donut-wrap">
                <div class="donut-container">
                  <svg width="140" height="140" viewBox="0 0 140 140">
                    <circle cx="70" cy="70" r="50" fill="none" stroke="#E3F0FF" stroke-width="22" />
${donutCircles}
                  </svg>
                  <div class="donut-center">
                    <div class="donut-total">${total}</div>
                    <div class="donut-label">전체 과정</div>
                  </div>
                </div>
                <div class="legend">
${legendItems}
                </div>
              </div>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <div class="card-title"><span class="card-title-icon">📊</span> 인기 교육 과정 Top 5</div>
              <span class="card-sub"><span style="color:#1565C0">■</span> 이수자 &nbsp;<span style="color:#43A047">■</span> 사전신청자</span>
            </div>
            <div class="card-body">
${barItems}
            </div>
          </div>
        </div>

        <div class="two-col">
          <div class="card">
            <div class="card-header">
              <div class="card-title"><span class="card-title-icon">🏆</span> 참여도 높은 소속팀 Top 5</div>
              <span class="card-sub">전 과정 사전신청+수료 합산 순</span>
            </div>
            <div class="card-body">
              <div class="rank-list">
${rankItems}
              </div>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <div class="card-title"><span class="card-title-icon">💡</span> R&D 수료 및 분석 오피니언</div>
              <span class="card-sub">${syncDate} 기준</span>
            </div>
            <div class="card-body">
              <div class="opinion-list">
${opinionHTML}
              </div>
            </div>
          </div>
        </div>
        <!-- §DASHBOARD_END§ -->`;
}

// 연구원 모드 월 탭 버튼 생성
function generateTabsBlock(rows) {
  const today = new Date();
  const currentMonth = today.getMonth() + 1;

  // CSV에서 실제 데이터가 있는 월 수집
  const dataMonths = new Set();
  rows.forEach(row => {
    const m = extractMonth(row['일시'] || '');
    if (m) dataMonths.add(m);
  });
  // 현재 월 ~ 12월 항상 포함 (상시 학습 포함)
  for (let m = currentMonth; m <= 12; m++) dataMonths.add(m);

  const sortedMonths = Array.from(dataMonths).sort((a, b) => a - b);

  const tabs = sortedMonths.map(m => {
    const isActive = m === currentMonth;
    const isPast = m < currentMonth;
    const tag = isActive
      ? ` <span class="month-tag">운영중</span>`
      : isPast
        ? ` <span class="month-tag past-tag">완료</span>`
        : '';
    const cls = isActive ? ' active' : '';
    return `      <button class="month-tab-btn${cls}" onclick="switchMonth(${m}, this)">${m}월${tag}</button>`;
  });

  return '      <!-- §TABS_START§ -->\n' + tabs.join('\n') + '\n      <!-- §TABS_END§ -->';
}

// monthlyEdu JS 코드 생성
function generateMonthlyBlock(rows) {
  const groups = {};
  const alwaysCourses = []; // 상시 학습 과정 별도 수집

  rows.forEach(row => {
    const dateStr = row['일시'] || '';
    const time = row['시간'] || '';
    const entry = {
      cat: row['구분'] || '',
      name: row['교육명'] || '',
      desc: row['세부 내용'] || '',
      instructor: row['강사'] || '',
      date: (time ? `${dateStr} ${time}` : dateStr).trim(),
      location: row['장소'] || '',
      seats: parseInt(row['사전 신청자 (명)'] || row['정원']) || null,
      eduHours: row['교육 시간(H)'] || '',
      applyUrl: (row['지금 바로 신청하기 (사이트 주소)'] || '').trim() || 'https://edu.coway.com/emp',
      canceled: (row['교육 진행 유무'] || '').includes('폐강') || (row['교육 진행 유무'] || '').includes('취소'),
      cancelReason: ((row['교육 진행 유무'] || '').match(/\(([^)]+)\)/) || [])[1] || '',
    };

    // '상시' 포함 일시 → 12월까지 매월 반복
    if (dateStr.includes('상시')) {
      alwaysCourses.push(entry);
      return;
    }

    const m = extractMonth(dateStr) || row['__month'];
    if (!m) { console.log(`  ⚠️  월 파싱 실패: "${dateStr}"`); return; }
    if (!groups[m]) groups[m] = [];
    groups[m].push(entry);
  });

  // 상시 학습 과정: 1월 ~ 12월 모든 월에 추가
  for (let m = 1; m <= 12; m++) {
    if (!groups[m]) groups[m] = [];
    alwaysCourses.forEach(c => groups[m].push(c));
  }

  const monthEntries = [];
  for (let m = 1; m <= 12; m++) {
    const list = groups[m] || [];
    if (!list.length) {
      monthEntries.push(
        `      ${m}: { title:'${m}월 수강 신청 과정', total:0, period:'신청 기간: 미정', ` +
        `desc:'${m}월 교육 계획을 수립 중입니다. 추후 공지 예정입니다.', courses:[] }`
      );
      continue;
    }
    const courseItems = list.map(c =>
      `        { cat:'${esc(c.cat)}', catClass:'${getCatClass(c.cat)}', name:'${esc(c.name)}', ` +
      `desc:'${esc(c.desc)}', instructor:'${esc(c.instructor)}', date:'${esc(c.date)}', ` +
      `location:'${esc(c.location)}', seats:${c.seats || 'null'}, eduHours:'${esc(c.eduHours || '')}', ` +
      `canceled:${c.canceled || false}, cancelReason:'${esc(c.cancelReason || '')}', target:'전체', applyType:'open', applyUrl:'${esc(c.applyUrl)}' }`
    ).join(',\n');
    monthEntries.push(
      `      ${m}: {\n        title:'${m}월 수강 신청 과정', total:${list.length},\n` +
      `        period:'신청 기간: 미정',\n        desc:'${m}월 교육 계획입니다.',\n` +
      `        courses:[\n${courseItems},\n        ]\n      }`
    );
  }

  return '    // §MONTHLY_START§\n    const monthlyEdu = {\n' +
         monthEntries.join(',\n') +
         '\n    };\n    // §MONTHLY_END§';
}

// CSV 파싱 (첫 번째 빈 행 자동 스킵, 헤더 자동 탐색)
function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim().replace(/,/g, '') !== '');
  if (lines.length < 2) return [];
  // 헤더 행 탐색 (첫 번째 비어있지 않은 행)
  const headerLine = lines[0];
  const headers = splitCSVLine(headerLine).map(h => h.trim());
  console.log(`  → 컬럼: ${headers.join(', ')}`);

  // 순차 월 추적 (상시 학습 등 날짜 없는 항목에 이전 월 적용)
  let lastMonth = null;
  return lines.slice(1).map(line => {
    const vals = splitCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || '').trim(); });
    // 교육명 없는 행 스킵
    if (!obj['교육명'] && !obj['구분']) return null;
    // 월 추적
    const m = extractMonth(obj['일시'] || '');
    if (m) lastMonth = m;
    else if (lastMonth) obj['__month'] = lastMonth;
    return obj;
  }).filter(r => r && Object.values(r).some(v => v !== ''));
}

function splitCSVLine(line) {
  const result = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else { inQ = !inQ; }
    } else if (ch === ',' && !inQ) {
      result.push(cur); cur = '';
    } else { cur += ch; }
  }
  result.push(cur);
  return result;
}

async function main() {
  console.log('─────────────────────────────────────');
  console.log('  Coway 교육포털 구글시트 동기화');
  console.log(`  실행 시각: ${new Date().toLocaleString('ko-KR')}`);
  console.log('─────────────────────────────────────');

  let rows = [];

  // ① Apps Script 웹앱에서 JSON 데이터 가져오기
  console.log('① Apps Script에서 데이터 다운로드 중...');
  try {
    rows = await fetchFromGAS();
    console.log(`  → Apps Script에서 ${rows.length}개 행 다운로드 완료`);
  } catch (e1) {
    console.warn(`  ⚠️  Apps Script 연결 실패: ${e1.message}`);
    // ② CSV 직접 다운로드 시도
    console.log('② 구글 시트에서 CSV 다운로드 시도...');
    try {
      const config = readApiConfig();
      if (!config.spreadsheetId) throw new Error('HTML에서 스프레드시트 ID를 찾을 수 없음');
      const csvText = await fetchCsvFromGoogleSheets(config.spreadsheetId, config.sheetGid);
      fs.writeFileSync(CSV_PATH, csvText, 'utf8');
      rows = parseCSV(csvText);
      console.log(`  → 구글 시트에서 ${rows.length}개 행 다운로드 완료`);
    } catch (e2) {
      console.warn(`  ⚠️  CSV 다운로드도 실패: ${e2.message}`);
      console.log('  → 로컬 캐시 CSV 파일로 대체 시도...');
      if (!fs.existsSync(CSV_PATH)) {
        console.error('\n❌ 로컬 CSV 캐시도 없음:', CSV_PATH);
        process.exit(1);
      }
      const csvText = fs.readFileSync(CSV_PATH, 'utf8');
      rows = parseCSV(csvText);
      console.log(`  → 캐시에서 ${rows.length}개 행 확인`);
    }
  }

  if (!rows.length) {
    console.error('\n❌ 데이터 없음 — 스프레드시트가 비어 있습니다.\n');
    process.exit(1);
  }

  // 3단계: HTML 업데이트
  console.log('③ HTML 파일 업데이트 중...');

  let html = fs.readFileSync(HTML_PATH, 'utf8');

  html = html.replace(
    /\/\/ §COURSES_START§[\s\S]*?\/\/ §COURSES_END§/,
    generateCoursesBlock(rows)
  );
  html = html.replace(
    /\/\/ §MONTHLY_START§[\s\S]*?\/\/ §MONTHLY_END§/,
    generateMonthlyBlock(rows)
  );
  html = html.replace(
    /<!-- §TABS_START§ -->[\s\S]*?<!-- §TABS_END§ -->/,
    generateTabsBlock(rows)
  );
  html = html.replace(
    /<!-- §DASHBOARD_START§ -->[\s\S]*?<!-- §DASHBOARD_END§ -->/,
    generateDashboardBlock(rows)
  );

  fs.writeFileSync(HTML_PATH, html, 'utf8');

  // 완료 출력
  const groups = {};
  rows.forEach(r => {
    const m = extractMonth(r['일시'] || '');
    if (m) groups[m] = (groups[m] || 0) + 1;
  });
  console.log('\n✅ 동기화 완료!');
  console.log('   월별 업데이트:');
  Object.keys(groups).sort((a, b) => a - b).forEach(m => {
    console.log(`   → ${m}월: ${groups[m]}개 과목`);
  });
  console.log('\n   브라우저에서 HTML 파일을 새로고침하세요.\n');
}

main().catch(e => { console.error('오류:', e.message); process.exit(1); });
