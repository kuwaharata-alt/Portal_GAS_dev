/** ======================================================================
 * 案件一覧表示用 
 * ====================================================================== */
const WEBVIEW = {
  SS_KEY: 'Project',

  SHEET_PROJECT: '案件情報',
  SHEET_HQ: '案件管理表_本社-現地',
  SHEET_WH: '案件管理表_倉庫',

  HEADER_ROW_PROJECT: 1,
  HEADER_ROW_HQ: 2,
  HEADER_ROW_WH: 2,

  KEY_HEADER: '見積番号',
  CUSTOMER_HEADER: '顧客名',
  MAX_ROWS: 5000,
};

const VIEW_CONFIG = {
  SV: {
    type: 'SV',
    label: '案件管理表(本社-現地)',
    detailSheetName: WEBVIEW.SHEET_HQ,
    detailHeaderRow: WEBVIEW.HEADER_ROW_HQ,
  },
  CL: {
    type: 'CL',
    label: '案件管理表(倉庫)',
    detailSheetName: WEBVIEW.SHEET_WH,
    detailHeaderRow: WEBVIEW.HEADER_ROW_WH,
  },
};

/** ======================================================================
 * 一覧取得
 * params:
 *  - type: SV / CL
 *  - q: 検索文字列
 *  - members: [担当者名, ...]
 *  - category: ｶﾃｺﾞﾘｰ
 * ====================================================================== */
function api_list(params) {
  try {
    const p = params || {};
    const type = String(p.type || 'SV').toUpperCase();
    const cfg = VIEW_CONFIG[type] || VIEW_CONFIG.SV;

    const q = normalizeSearchText_(p.q || '');
    const memberFilters = Array.isArray(p.members)
      ? p.members.map(nz_).filter(Boolean)
      : [];
    const categoryFilter = nz_(p.category);
    const statusFilters = Array.isArray(p.statuses)
      ? p.statuses.map(nz_).filter(Boolean)
      : [];

    const ss = getSS_(WEBVIEW.SS_KEY);
    const shProject = ss.getSheetByName(WEBVIEW.SHEET_PROJECT);
    const shDetail = ss.getSheetByName(cfg.detailSheetName);

    if (!shProject) throw new Error(`シートが見つかりません: ${WEBVIEW.SHEET_PROJECT}`);
    if (!shDetail) throw new Error(`シートが見つかりません: ${cfg.detailSheetName}`);

    const projectObj = readSheetObjects_(shProject, WEBVIEW.HEADER_ROW_PROJECT, WEBVIEW.KEY_HEADER);
    const detailObj = readSheetObjects_(shDetail, cfg.detailHeaderRow, WEBVIEW.KEY_HEADER);

    const headers = ['ステータス', '案件概要', '案件情報', '担当SE', '作業期間', '備考'];
    const rows = [];
    const projectMap = projectObj.rowMap;

    (detailObj.rows || []).forEach(dt => {
      const key = nz_(dt[WEBVIEW.KEY_HEADER]);
      if (!key) return;

      const pj = projectMap[key] || {};
      const category = getRowCategory_(type, pj, dt);
      const memberNames = getRowMemberNames_(pj, dt);
      const statusFilterValue = getStatusFilterValue_(dt);

      if (memberFilters.length > 0) {
        const hit = memberFilters.some(name => memberNames.includes(name));
        if (!hit) return;
      }

      if (categoryFilter && category !== categoryFilter) return;

      if (statusFilters.length > 0) {
        if (!statusFilters.includes(statusFilterValue || '')) return;
      }

      if (q) {
        const searchText = normalizeSearchText_(
          [
            key,
            category,
            statusFilterValue,
            memberNames.join(' '),
            ...Object.values(pj),
            ...Object.values(dt),
          ].join('\n')
        );
        if (searchText.indexOf(q) === -1) return;
      }

      rows.push(buildListRow_(type, key, pj, dt));
    });

    return {
      ok: true,
      type,
      title: cfg.label,
      keyHeader: WEBVIEW.KEY_HEADER,
      customerHeader: WEBVIEW.CUSTOMER_HEADER,
      headers,
      rows,
      total: rows.length,
      page: 1,
      pageSize: rows.length,
    };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err),
    };
  }
}

/** ======================================================================
 * フィルタ用マスタ
 * ====================================================================== */
function api_getListFilterMaster(params) {
  try {
    const p = params || {};
    const type = String(p.type || 'SV').toUpperCase();
    const cfg = VIEW_CONFIG[type] || VIEW_CONFIG.SV;

    const ss = getSS_(WEBVIEW.SS_KEY);
    const shProject = ss.getSheetByName(WEBVIEW.SHEET_PROJECT);
    const shDetail = ss.getSheetByName(cfg.detailSheetName);

    if (!shProject) throw new Error(`シートが見つかりません: ${WEBVIEW.SHEET_PROJECT}`);
    if (!shDetail) throw new Error(`シートが見つかりません: ${cfg.detailSheetName}`);

    const projectObj = readSheetObjects_(shProject, WEBVIEW.HEADER_ROW_PROJECT, WEBVIEW.KEY_HEADER);
    const detailObj = readSheetObjects_(shDetail, cfg.detailHeaderRow, WEBVIEW.KEY_HEADER);

    const projectMap = projectObj.rowMap;
    const memberSet = new Set();
    const categorySet = new Set();
    const statusSet = new Set();

    (detailObj.rows || []).forEach(dt => {
      const key = nz_(dt[WEBVIEW.KEY_HEADER]);
      if (!key) return;

      const pj = projectMap[key] || {};
      const category = getRowCategory_(type, pj, dt);
      const statusFilterValue = getStatusFilterValue_(dt);
      const memberNames = getRowMemberNames_(pj, dt);

      if (category) categorySet.add(category);
      if (statusFilterValue) statusSet.add(statusFilterValue);

      memberNames.forEach(name => {
        if (name) memberSet.add(name);
      });
    });

    return {
      ok: true,
      members: Array.from(memberSet).sort((a, b) => a.localeCompare(b, 'ja')),
      categories: Array.from(categorySet).sort((a, b) => a.localeCompare(b, 'ja')),
      statuses: Array.from(statusSet).sort((a, b) => a.localeCompare(b, 'ja')),
    };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err),
      members: [],
      categories: [],
      statuses: [],
    };
  }
}

function buildListRow_(type, key, pj, dt) {
  const customer = nz_(dt['顧客名']) || nz_(pj['顧客名']);
  const category = getRowCategory_(type, pj, dt);
  const memberNames = getRowMemberNames_(pj, dt);
  const statusFilterValue = getStatusFilterValue_(dt);

  const statusHtml = buildStatusHtml_(dt, type);
  const summaryHtml = buildSummaryHtml_(key, customer, pj, dt, type);
  const caseInfoHtml = buildCaseInfoHtml_(pj, dt, type);
  const memberHtml = buildMemberHtml_(pj, dt, type);
  const periodHtml = buildPeriodHtml_(dt, type);
  const remarksHtml = buildRemarksHtml_(dt, type);

  return {
    key,
    customer,
    category,
    memberNames,
    statusFilterValue,
    cells: [
      statusHtml,
      summaryHtml,
      caseInfoHtml,
      memberHtml,
      periodHtml,
      remarksHtml,
    ],
  };
}

/** -----------------------------
 * 各列HTML
 * ----------------------------- */
function buildStatusHtml_(dt, type) {
  const raw = nz_(dt['ステータス']);
  const parsed = parseStatusParts_(raw);
  const main = resolveMainStatusDisplay_(parsed, dt);
  const icons = buildStatusMiniIcons_(dt);
  const chipClass = getStatusChipClass_(main.kind, main.status);

  return `
    <div class="status-wrap">
      <div class="status-main">
        <span class="status-chip ${chipClass}">${escapeHtml_(main.displayText || '未設定')}</span>
      </div>
      <div class="status-mini-area">
        ${icons}
      </div>
    </div>
  `;
}

function parseStatusParts_(raw) {
  const s = String(raw || '').replace(/\r/g, '').trim();
  const m = s.match(/^【([^】]+)】\s*\n?\s*(.*)$/);

  return {
    raw: s,
    kind: m ? m[1].trim() : '',
    status: m ? m[2].trim() : s
  };
}

function resolveMainStatusDisplay_(parsed, dt) {
  const kind = parsed.kind;
  const rawStatus = parsed.status;

  if (kind === '案件管理') {
    return {
      kind: '案件管理',
      status: rawStatus || '未設定',
      displayText: rawStatus || '未設定'
    };
  }

  if (kind === 'ドキュメント') {
    const st = rawStatus || '未設定';
    return {
      kind: 'ドキュメント',
      status: st,
      displayText: `ドキュメント\n${st}`
    };
  }

  if (kind === '社内作業') {
    const st = nz_(dt['社内作業ステータス']) || rawStatus || '未設定';
    return {
      kind: '社内作業中',
      status: st,
      displayText: `社内作業中\n${st}`
    };
  }

  if (kind === '現地作業') {
    const st = nz_(dt['現地作業ステータス']) || rawStatus || '未設定';
    return {
      kind: '現地作業中',
      status: st,
      displayText: `現地作業中\n${st}`
    };
  }

  if (kind === '倉庫作業') {
    const st = nz_(dt['倉庫作業ステータス']) || rawStatus || '未設定';
    return {
      kind: '倉庫作業中',
      status: st,
      displayText: `倉庫作業中\n${st}`
    };
  }

  return {
    kind: kind || '',
    status: rawStatus || '未設定',
    displayText: rawStatus || '未設定'
  };
}

function buildStatusMiniIcons_(dt) {
  const items = [];

  const internal = shortStatusLabel_(dt['社内作業ステータス']);
  const onsite = shortStatusLabel_(dt['現地作業ステータス']);
  const wh = shortStatusLabel_(dt['倉庫作業ステータス']);

  if (internal) {
    items.push(miniStatusHtml_('内：', internal, 'mini-internal'));
  }
  if (onsite) {
    items.push(miniStatusHtml_('外：', onsite, 'mini-onsite'));
  }
  if (wh) {
    items.push(miniStatusHtml_('倉：', wh, 'mini-warehouse'));
  }

  return items.join('');
}

function shortStatusLabel_(v) {
  const s = nz_(v);

  if (!s) return '';

  if (s === '対応無') return '対応無';
  if (s === '機器入荷待ち') return '未入荷';
  if (s === '作業待ち') return '未着手';
  if (s === '対応中') return '対応中';
  if (s === '機器出荷待ち') return '出荷待';
  if (s === '完了') return '完了';

  return s;
}

function getStatusChipClass_(kind, status) {
  const k = nz_(kind);
  const s = nz_(status);

  if (k === '案件管理') {
    if (s.includes('クローズ')) return 'st-close';
    if (s.includes('キャンセル')) return 'st-cancel';
    if (s.includes('保守')) return 'st-hold';
    if (s.includes('完了')) return 'st-case-done';
    if (s.includes('済')) return 'st-case-done';
    if (s.includes('調整中')) return 'st-case-progress';
    return 'st-case';
  }

  if (k === 'ドキュメント') {
    if (s.includes('対応中')) return 'st-progress';
    return 'st-doc';
  }

  if (k === '社内作業中') {
    if (s.includes('完了')) return 'st-internal-done';
    if (s.includes('機器出荷待ち')) return 'st-wait';
    if (s.includes('対応中') || s.includes('セットアップ中')) return 'st-progress';
    if (s.includes('対応無') || s.includes('セットアップ中')) return 'st-nothing';
    return 'st-internal';
  }

  if (k === '現地作業中') {
    if (s.includes('完了')) return 'st-onsite-done';
    if (s.includes('対応中') || s.includes('作業中')) return 'st-onsite-progress';
    return 'st-onsite';
  }

  if (k === '倉庫作業中') {
    if (s.includes('完了')) return 'st-warehouse-done';
    if (s.includes('対応中')) return 'st-warehouse-progress';
    return 'st-warehouse';
  }

  return 'st-default';
}

function miniStatusHtml_(label, value, cls) {
  return `
    <span class="status-mini ${cls}">
      <span class="status-mini-label">${escapeHtml_(label)}</span>
      <span class="status-mini-value">${escapeHtml_(value)}</span>
    </span>
  `;
}

function buildSummaryHtml_(key, customer, pj, dt, type) {
  const summary = nz_(pj['案件概要']);
  const estimateUrl = getEstimateUrl_(pj, dt);

  const headText = `【${escapeHtml_(key)}】${escapeHtml_(customer)}`;
  const headHtml = estimateUrl
    ? `<a class="case-summary-link" href="${escapeHtml_(estimateUrl)}" target="_blank" rel="noopener noreferrer">${headText}</a>`
    : headText;

  return `
    <div class="case-summary-wrap">
      <div class="case-summary-head">${headHtml}</div>
      ${summary ? `<div class="case-summary-body">${escapeHtml_(summary)}</div>` : ''}
    </div>
  `;
}

function getEstimateUrl_(pj, dt) {
  // 1) 案件情報シートの「見積」列を最優先
  const pjUrl =
    firstNonEmpty_(
      pj['__link_見積'],
      pj['見積'],
      dt['__link_見積'],
      dt['見積']
    );

  if (isUrlLike_(pjUrl)) return pjUrl;

  return '';
}

function isUrlLike_(v) {
  const s = nz_(v);
  return /^https?:\/\/\S+$/i.test(s);
}

function buildCaseInfoHtml_(pj, dt, type) {
  const rows = [];
  const workCategory = getRowCategory_(type, pj, dt);

  const workplace = type === 'SV'
    ? nz_(dt['作業形態'])
    : '倉庫作業';

  const place = type === 'SV'
    ? joinNonEmpty_([dt['現地']])
    : '';

  const due = type === 'SV'
    ? (nz_(pj['検収予定']) || nz_(dt['検収予定']))
    : nz_(dt['検収予定']);

  rows.push(infoLineHtml_('作業カテゴリ', workCategory));
  rows.push(infoLineHtml_('作業形態', workplace));
  if (place) rows.push(infoLineHtml_('作業場所', place));
  rows.push(infoLineHtml_('検収予定', due));

  return `<div class="case-info-wrap">${rows.join('')}</div>`;
}

function buildMemberHtml_(pj, dt, type) {
  const supervisor = firstNonEmpty_(pj['監督者']);
  const manager = firstNonEmpty_(dt['管理者'], pj['管理者']);
  const supports = uniqueNonEmpty_([
    firstNonEmpty_(dt['サポート1'], pj['サポート']),
    (dt['サポート1'], pj['サポート2']),
    dt['サポート3'],
  ]);

  const bp = firstNonEmpty_(dt['BP利用'], pj['BP利用']);

  const rows = [];
  rows.push(seLineHtml_('監督者', supervisor));
  rows.push(seLineHtml_('管理者', manager));
  rows.push(seLineHtml_('サポート', supports.join(' 、 ')));
  rows.push(seLineHtml_('BP利用', bp));

  return `<div class="case-se-wrap">${rows.join('')}</div>`;
}

function buildPeriodHtml_(dt, type) {
  if (type === 'SV') {
    const caseText = '未設定';
    const inText = rangeText_(dt['社内作業開始日'], dt['社内作業終了日']);
    const outText = rangeText_(dt['現地作業開始日'], dt['現地作業終了日']);
    const warn = (!nz_(dt['完了予定']) && !nz_(dt['検収予定'])) ? '完了予定未設定' : '';

    return `
      <div class="wp-wrap-simple">
        ${wpRowHtml_('案件', caseText)}
        ${wpRowHtml_('社内', inText || '-')}
        ${wpRowHtml_('社外', outText || '-')}
        ${warn ? `<div class="wp-info-simple">${escapeHtml_(warn)}</div>` : ''}
      </div>
    `;
  }

  const inText = rangeText_(dt['社内作業開始日'], dt['社内作業終了日']);
  const due = nz_(dt['案件完了予定']) || nz_(dt['検収予定']);

  return `
    <div class="wp-wrap-simple">
      ${wpRowHtml_('案件', due || '未設定')}
      ${wpRowHtml_('社内', inText || '-')}
    </div>
  `;
}

function buildRemarksHtml_(dt, type) {
  const rows = [];

  if (type === 'SV') {
    rows.push(remarksRowHtml_('案件', nz_(dt['案件コメント※案件全体の状況を更新']) || ''));
    rows.push(remarksRowHtml_('社内', nz_(dt['コメント']) || ''));
    rows.push(remarksRowHtml_('社外', nz_(dt['コメント_2']) || ''));
  } else {
    rows.push(remarksRowHtml_('案件', nz_(dt['案件コメント※都度案件状況を更新']) || ''));
    rows.push(remarksRowHtml_('社内', ''));
  }

  return `<div class="remarks-list">${rows.join('')}</div>`;
}

/** -----------------------------
 * 行情報
 * ----------------------------- */
function getRowCategory_(type, pj, dt) {
  if (type === 'SV') return nz_(pj['ｶﾃｺﾞﾘｰ']);
  return nz_(dt['ｶﾃｺﾞﾘｰ']);
}

function getRowMemberNames_(pj, dt) {
  const supervisor = splitMemberNames_(pj['監督者']);
  const manager = splitMemberNames_(firstNonEmpty_(dt['管理者'], pj['管理者']));
  const support1 = splitMemberNames_(firstNonEmpty_(dt['サポート1'], pj['サポート']));
  const support2 = splitMemberNames_(dt['サポート2']);
  const support3 = splitMemberNames_(dt['サポート3']);

  return uniqueNonEmpty_([
    ...supervisor,
    ...manager,
    ...support1,
    ...support2,
    ...support3,
  ]);
}

/** -----------------------------
 * HTML helper
 * ----------------------------- */
function infoLineHtml_(label, value) {
  return `
    <div class="info-line">
      <span class="info-tag">${escapeHtml_(label)}</span>
      <span class="info-value">${escapeHtml_(value)}</span>
    </div>
  `;
}

function seLineHtml_(label, value) {
  return `
    <div class="info-line">
      <span class="info-tag se-tag">${escapeHtml_(label)}</span>
      <span class="info-value">${escapeHtml_(value)}</span>
    </div>
  `;
}

function wpRowHtml_(label, value) {
  return `
    <div class="wp-row-simple">
      <div class="wp-label-simple">${escapeHtml_(label)}</div>
      <div class="wp-lane-simple">
        <div class="wp-bar-simple"><span class="wp-bar-text-simple">${escapeHtml_(value)}</span></div>
      </div>
    </div>
  `;
}

function remarksRowHtml_(label, text) {
  return `
    <div class="remarks-row">
      <div class="remarks-label">${escapeHtml_(label)}</div>
      <div class="remarks-text">${escapeHtml_(text || '')}</div>
    </div>
  `;
}

function statusClass_(raw) {
  const s = String(raw || '');
  if (!s) return 'status-pill status-na';
  if (/close|完了/i.test(s)) return 'status-pill status-done';
  if (/対応中|作業中|進行/i.test(s)) return 'status-pill status-doing';
  if (/未設定|未着手|未/i.test(s)) return 'status-pill status-todo';
  return 'status-pill status-assign-progress';
}

function rangeText_(start, end) {
  const s = nz_(start);
  const e = nz_(end);
  if (s && e) return `${s} ～ ${e}`;
  if (s) return `${s} ～`;
  if (e) return `～ ${e}`;
  return '';
}

function parseStatus_(text) {
  const s = String(text || '').trim();
  const m = s.match(/^【([^】]+)】\s*(.*)$/);

  return {
    kind: m ? m[1] : '',
    label: m ? m[2] : s
  };
}

/** -----------------------------
 * シート読込
 * ----------------------------- */
function readSheetObjects_(sheet, headerRow, keyHeader) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < headerRow || lastCol < 1) {
    return { headers: [], rows: [], rowMap: {} };
  }

  const readRows = Math.min(lastRow, WEBVIEW.MAX_ROWS);

  const range = sheet.getRange(1, 1, readRows, lastCol);
  const values = range.getDisplayValues();

  const headers = (values[headerRow - 1] || []).map(normalizeHeader_);
  const map = buildHeaderMap0_(headers);

  const estimateLinkIdx = map['見積'];
  const estimateRichValues = estimateLinkIdx != null
    ? sheet.getRange(1, estimateLinkIdx + 1, readRows, 1).getRichTextValues()
    : null;

  const keyIdx = map[keyHeader];
  if (keyIdx == null) {
    throw new Error(`ヘッダー「${keyHeader}」が見つかりません: ${sheet.getName()}`);
  }

  const rows = [];
  const rowMap = {};

  for (let r = headerRow; r < values.length; r++) {
    const row = values[r] || [];
    const key = stringifyCell_(row[keyIdx]);
    if (!key) continue;

    const obj = {};

    headers.forEach((h, i) => {
      if (!h) return;

      const cellText = stringifyCell_(row[i]);
      obj[h] = cellText;

      if (h === '見積' && estimateRichValues) {
        const rt = estimateRichValues[r] ? estimateRichValues[r][0] : null;
        const linkUrl = getCellLinkUrl_(rt);
        if (linkUrl) {
          obj[`__link_${h}`] = linkUrl;
        }
      }
    });

    if (sheet.getName() === WEBVIEW.SHEET_HQ) {
      const c23 = stringifyCell_(row[22]);
      const c29 = stringifyCell_(row[28]);
      obj['コメント'] = c23;
      obj['コメント_2'] = c29;
      obj['案件コメント※案件全体の状況を更新'] = stringifyCell_(row[6]);
    }

    if (sheet.getName() === WEBVIEW.SHEET_WH) {
      obj['案件コメント※都度案件状況を更新'] = stringifyCell_(row[6]);
    }

    rows.push(obj);
    rowMap[key] = obj;
  }

  return { headers, rows, rowMap };
}

function getCellLinkUrl_(rt) {
  if (!rt) return '';

  // セル全体にリンク
  const direct = rt.getLinkUrl();
  if (direct) return direct;

  // 部分リンク対応
  const runs = rt.getRuns ? rt.getRuns() : [];
  for (let i = 0; i < runs.length; i++) {
    const url = runs[i].getLinkUrl();
    if (url) return url;
  }

  return '';
}
/** -----------------------------
 * common
 * ----------------------------- */
function normalizeHeader_(v) {
  return String(v == null ? '' : v).replace(/\r?\n/g, '').trim();
}

function buildHeaderMap0_(headers) {
  const map = {};
  (headers || []).forEach((h, i) => {
    if (h) map[h] = i;
  });
  return map;
}

function stringifyCell_(v) {
  return String(v == null ? '' : v).trim();
}

function normalizeSearchText_(v) {
  return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' ');
}

function nz_(v) {
  return String(v == null ? '' : v).trim();
}

function firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i++) {
    const s = nz_(arguments[i]);
    if (s) return s;
  }
  return '';
}

function joinNonEmpty_(arr) {
  return (arr || []).map(nz_).filter(Boolean).join(' / ');
}

function uniqueNonEmpty_(arr) {
  return Array.from(new Set((arr || []).map(nz_).filter(Boolean)));
}

function splitMemberNames_(v) {
  const s = nz_(v);
  if (!s) return [];
  return s
    .split(/[\/／,、，\n\r\t]/)
    .map(nz_)
    .filter(Boolean);
}

function stripHtml_(s) {
  return String(s || '').replace(/<[^>]*>/g, ' ');
}

function escapeHtml_(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getStatusFilterValue_(dt) {
  const raw = nz_(dt['ステータス']);
  const parsed = parseStatusParts_(raw);
  const main = resolveMainStatusDisplay_(parsed, dt);
  return nz_(main.displayText || main.status || raw);
}