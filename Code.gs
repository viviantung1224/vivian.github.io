// ═══════════════════════════════════════════════════════════════════════
//  2Fish 營運看板  ─  Google Apps Script API
// ───────────────────────────────────────────────────────────────────────
//  【部署步驟】
//  1. 開啟任一張試算表 → 擴充功能 → Apps Script
//  2. 把這份程式碼全部貼上
//  3. 左上角命名為「2Fish Dashboard API」
//  4. 點「部署」→「新增部署」
//     - 類型：網頁應用程式
//     - 執行身分：我（你的帳號）
//     - 存取權限：所有人
//  5. 授權 → 複製部署網址 → 貼到 dashboard.html 的 API_URL 變數
//  6. 之後每次修改程式，記得「管理部署」→「編輯」→「建立新版本」
// ═══════════════════════════════════════════════════════════════════════

const SHEET_IDS = {
  finance: '10LsL0_aWd9ieF5lVt_9yeZxZpFZM4uso5ZixL50nXbI',
  project: '16dRT1f62Pn6oWPP8CzdSfUKH5hG65JwWyH_pdK4jw64',
  editing: '1wHRwYxVXJxXxd8I-D37bu4-MsC8NJvyUmn8pzlC6UOA',
  team:    '1IH8gSevjUxA4NNekYa0sWIpO_6-v8NM54tqkjuPLHjg',
};

const CACHE_KEY = 'dashboard_v1';
const CACHE_TTL = 300; // 5 分鐘

// ─── 進入點 ─────────────────────────────────────────────────────────────
function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  const refresh = e && e.parameter && e.parameter.refresh === '1';

  try {
    if (!refresh) {
      const cached = CacheService.getScriptCache().get(CACHE_KEY);
      if (cached) { out.setContent(cached); return out; }
    }
    const data = buildAll();
    const json = JSON.stringify({ ok: true, updated: new Date().toISOString(), data });
    CacheService.getScriptCache().put(CACHE_KEY, json, CACHE_TTL);
    out.setContent(json);
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, error: err.message, stack: err.stack }));
  }
  return out;
}

// ─── 工具 ────────────────────────────────────────────────────────────────
function c(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim();
  return ['-','—','null','None','undefined','N/A',''].includes(s) ? '' : s;
}
function n(v) { const x = parseFloat(v); return isNaN(x) ? 0 : x; }
function dt(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  const s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0,10) : '';
}
function rows(ws, from) {
  const lr = ws.getLastRow(), lc = ws.getLastColumn();
  if (lr < from || lc === 0) return [];
  return ws.getRange(from, 1, lr - from + 1, lc).getValues();
}
function ss(id)   { return SpreadsheetApp.openById(id); }
function sh(id, name) {
  const w = ss(id).getSheetByName(name);
  if (!w) throw new Error('找不到頁籤：' + name);
  return w;
}

// ─── 主建構 ──────────────────────────────────────────────────────────────
function buildAll() {
  const staffMap  = buildStaff();
  const statusMap = buildStatus();
  const finMap    = buildFinance();
  const videoMap  = buildVideo();
  const trend     = buildTrend(videoMap);

  const STATUS_ORDER = { active:0, waiting:1, rwaiting:2, pending:3, no:4, term:5, closed:6 };
  const projects = [];

  for (const [pname, staff] of Object.entries(staffMap)) {
    if (!statusMap[pname]) continue;
    const brand = staff._brand || '';
    const fin   = finMap[brand]  || { total:0, received:0, rule:'' };
    const vid   = videoMap[pname] || { demand:0, supply:0, lastPublished:null, latestR:'' };
    projects.push({
      name: pname,
      pm: staff.pm, planner: staff.planner,
      publisher: staff.publisher, pm_asst: staff.pm_asst,
      sr_editor: staff.sr_editor, director: staff.director, producer: staff.producer,
      status:   statusMap[pname],
      progress: vid.latestR,
      start:    staff.start,
      renewals: fin.renewals || 0,
      total:    Math.round(fin.total),
      received: Math.round(fin.received),
      rule:     fin.rule || '—',
      video_demand:   vid.demand,
      video_supply:   vid.supply,
      last_published: vid.lastPublished,
      payments: [],
    });
  }

  projects.sort((a,b) =>
    (STATUS_ORDER[a.status]??9) - (STATUS_ORDER[b.status]??9) ||
    a.name.localeCompare(b.name, 'zh')
  );

  return { projects, trend };
}

// ─── 人員 ─────────────────────────────────────────────────────────────────
function buildStaff() {
  const map = {};
  for (const r of rows(sh(SHEET_IDS.project, '0.專案客戶資訊總表'), 3)) {
    if (!r[0]) continue;
    const pname = c(r[17]); if (!pname) continue;
    map[pname] = {
      _brand: c(r[3]), pm: c(r[18]), planner: c(r[19]),
      start: dt(r[21]), publisher: c(r[22]),
      pm_asst: c(r[25]), sr_editor: c(r[26]), director: c(r[27]), producer: c(r[28]),
    };
  }
  return map;
}

// ─── 狀態 ─────────────────────────────────────────────────────────────────
function buildStatus() {
  const STA = {
    '續約':'active','等結案':'waiting','續約等結案':'rwaiting',
    'Pending中':'pending','提前解約':'term','解約':'term','不續約':'no',
  };
  const map = {};
  for (const r of rows(sh(SHEET_IDS.project, '專案總覽（戰時管理）'), 3)) {
    if (!r[0]) continue;
    const pname = c(r[2]); if (!pname) continue;
    map[pname] = STA[c(r[12])] || 'active';
  }
  return map;
}

// ─── 財務 ─────────────────────────────────────────────────────────────────
function buildFinance() {
  const finSs = ss(SHEET_IDS.finance);
  const map = {};

  function add(wsName, startRow, brandCol, totalCol, ruleCol) {
    const w = finSs.getSheetByName(wsName);
    if (!w) return;
    for (const r of rows(w, startRow)) {
      const brand = c(r[brandCol]); if (!brand) continue;
      if (!map[brand]) map[brand] = { total:0, received:0, rule:'', renewals:0 };
      map[brand].total += n(r[totalCol]);
      if (ruleCol !== null) {
        const rule = c(r[ruleCol]);
        if (rule && !map[brand].rule) map[brand].rule = rule.replace(/\n/g, '；');
      }
    }
  }

  add('簽約客戶資訊＆收款紀錄總表', 3, 3, 21, 17);  // 新約
  add('第一次續約',                  3, 3, 23, 18);  // 第一次續約
  add('第二次(含)以上續約',          3, 3, 19, null); // 第二次以上

  // 已收款
  for (const r of rows(finSs.getSheetByName('收款流水帳'), 1)) {
    if (!r[0] || c(r[0]) === '案號') continue;
    const brand = c(r[1]); if (!brand || brand === '品牌名稱') continue;
    if (r[12] && n(r[14]) > 0) {
      if (!map[brand]) map[brand] = { total:0, received:0, rule:'', renewals:0 };
      map[brand].received += n(r[14]);
    }
  }

  return map;
}

// ─── 影片統計 ─────────────────────────────────────────────────────────────
function buildVideo() {
  const map = {};
  for (const r of rows(sh(SHEET_IDS.editing, '剪輯排程總表'), 3)) {
    const pname = c(r[0]); if (!pname) continue;
    if (!map[pname]) map[pname] = { demand:0, supply:0, lastPublished:null, _latestR:'', _dates:[] };
    map[pname].demand++;
    const published = r[16];
    if (published === true || published === 1 || published === 'TRUE') {
      map[pname].supply++;
      const ds = dt(r[11]); if (ds) map[pname]._dates.push(ds);
    }
    const batch = c(r[2]);
    if (batch && batch > map[pname]._latestR) map[pname]._latestR = batch;
  }
  for (const v of Object.values(map)) {
    v.lastPublished = v._dates.length ? v._dates.sort().pop() : null;
    v.latestR = v._latestR.replace(/^R0*(\d+)/, 'R$1');
    delete v._dates; delete v._latestR;
  }
  return map;
}

// ─── 月趨勢 ───────────────────────────────────────────────────────────────
function buildTrend(videoMap) {
  // 最近 6 個月剪輯量，從剪輯排程計算
  const now = new Date();
  const months = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const ym = Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM');
    months.push({ label: ym.replace('-','/'), ym, editing:0, active:0 });
  }

  for (const r of rows(sh(SHEET_IDS.editing, '剪輯排程總表'), 3)) {
    if (r[16] !== true && r[16] !== 1) continue;
    const ds = dt(r[11]); if (!ds) continue;
    const ym = ds.slice(0,7);
    const m = months.find(x => x.ym === ym);
    if (m) m.editing++;
  }

  return {
    labels:   months.map(m => m.label),
    editing:  months.map(m => m.editing),
    shooting: months.map(() => 0), // 需接拍攝排程時更新
    active:   months.map(() => 0), // 由前端計算
  };
}
