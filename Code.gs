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

// ─── 進入點（支援 JSONP，解決 CORS 問題） ──────────────────────────────
function doGet(e) {
  const refresh  = e && e.parameter && e.parameter.refresh === '1';
  const callback = e && e.parameter && e.parameter.callback; // JSONP callback name

  let body;
  try {
    if (!refresh) {
      const cached = CacheService.getScriptCache().get(CACHE_KEY);
      if (cached) {
        body = cached;
      }
    }
    if (!body) {
      const data = buildAll();
      body = JSON.stringify({ ok: true, updated: new Date().toISOString(), data });
      CacheService.getScriptCache().put(CACHE_KEY, body, CACHE_TTL);
    }
  } catch (err) {
    body = JSON.stringify({ ok: false, error: err.message });
  }

  // JSONP 模式：包成 callback(...)
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // 一般 JSON 模式
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
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
// cols = 要讀的欄數（從第1欄起算），不傳則讀到最後一欄
function rows(ws, from, cols) {
  const lr = ws.getLastRow();
  const lc = cols || ws.getLastColumn();
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
  const { videoMap, trend } = buildVideoAndTrend();

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
      video_ac_done:  vid.ac_done,
      video_abandoned:vid.abandoned,
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

// ─── 人員（只讀前 29 欄） ──────────────────────────────────────────────────
function buildStaff() {
  const map = {};
  for (const r of rows(sh(SHEET_IDS.project, '0.專案客戶資訊總表'), 3, 29)) {
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

// ─── 狀態（只讀前 13 欄） ──────────────────────────────────────────────────
function buildStatus() {
  const STA = {
    '續約':'active','等結案':'waiting','續約等結案':'rwaiting',
    'Pending中':'pending','提前解約':'term','解約':'term','不續約':'no',
  };
  const map = {};
  for (const r of rows(sh(SHEET_IDS.project, '專案總覽（戰時管理）'), 3, 13)) {
    if (!r[0]) continue;
    const pname = c(r[2]); if (!pname) continue;
    map[pname] = STA[c(r[12])] || 'active';
  }
  return map;
}

// ─── 財務（限制欄寬） ──────────────────────────────────────────────────────
function buildFinance() {
  const finSs = ss(SHEET_IDS.finance);
  const map = {};

  function add(wsName, startRow, brandCol, totalCol, ruleCol, maxCols) {
    const w = finSs.getSheetByName(wsName);
    if (!w) return;
    for (const r of rows(w, startRow, maxCols)) {
      const brand = c(r[brandCol]); if (!brand) continue;
      if (!map[brand]) map[brand] = { total:0, received:0, rule:'', renewals:0 };
      map[brand].total += n(r[totalCol]);
      if (ruleCol !== null) {
        const rule = c(r[ruleCol]);
        if (rule && !map[brand].rule) map[brand].rule = rule.replace(/\n/g, '；');
      }
    }
  }

  add('簽約客戶資訊＆收款紀錄總表', 3, 3, 21, 17, 22);  // 新約：只讀前 22 欄
  add('第一次續約',                  3, 3, 23, 18, 24);  // 第一次續約：前 24 欄
  add('第二次(含)以上續約',          3, 3, 19, null, 20); // 第二次以上：前 20 欄

  // 已收款（只讀前 15 欄）
  for (const r of rows(finSs.getSheetByName('收款流水帳'), 1, 15)) {
    if (!r[0] || c(r[0]) === '案號') continue;
    const brand = c(r[1]); if (!brand || brand === '品牌名稱') continue;
    if (r[12] && n(r[14]) > 0) {
      if (!map[brand]) map[brand] = { total:0, received:0, rule:'', renewals:0 };
      map[brand].received += n(r[14]);
    }
  }

  return map;
}

// ─── 影片統計 + 月趨勢（合併成一次讀取，只讀前 17 欄） ────────────────────
function buildVideoAndTrend() {
  const now = new Date();
  const months = [];
  for (let i = 5; i >= 0; i--) {
    const d  = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const ym = Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM');
    months.push({ label: ym.replace('-','/'), ym, editing: 0 });
  }

  const videoMap = {};

  for (const r of rows(sh(SHEET_IDS.editing, '剪輯排程總表'), 3, 17)) {
    const pname = c(r[0]); if (!pname) continue;
    if (!videoMap[pname]) videoMap[pname] = { demand:0, supply:0, ac_done:0, abandoned:0, lastPublished:null, _latestR:'', _dates:[] };

    videoMap[pname].demand++;
    const batch     = c(r[2]);
    const planDate  = r[11];
    const acDone    = r[12]; // M欄 AC已完成
    const abandoned = r[7];  // H欄 確定放棄
    const published = r[16];

    if (batch && batch > videoMap[pname]._latestR) videoMap[pname]._latestR = batch;
    if (acDone    === true || acDone    === 1) videoMap[pname].ac_done++;
    if (abandoned === true || abandoned === 1) videoMap[pname].abandoned++;

    if (published === true || published === 1) {
      videoMap[pname].supply++;
      const ds = dt(planDate);
      if (ds) {
        videoMap[pname]._dates.push(ds);
        const ym = ds.slice(0,7);
        const m = months.find(x => x.ym === ym);
        if (m) m.editing++;
      }
    }
  }

  for (const v of Object.values(videoMap)) {
    v.lastPublished = v._dates.length ? v._dates.sort().pop() : null;
    v.latestR = v._latestR.replace(/^R0*(\d+)/, 'R$1');
    delete v._dates; delete v._latestR;
  }

  const trend = {
    labels:   months.map(m => m.label),
    editing:  months.map(m => m.editing),
    shooting: months.map(() => 0),
    active:   months.map(() => 0),
  };

  return { videoMap, trend };
}
