// ==UserScript==
// @name         Trưởng Bưu Cục
// @namespace    local.800best.truongbuucuc
// @version      1.8.0
// @description  Bản chia sẻ cho Trưởng bưu cục: màn Thống kê tổng quan 1 bưu cục (thẻ chỉ số + bảng shipper, gồm Chỉ số giao hàng / Chỉ số lấy hàng / Chỉ số TikTok Shop). Tự dò bưu cục đang quản lý khi cài lần đầu. v1.1: bỏ GTC chuyến (2 dải) thay bằng "Cần đạt KPI POD2400" (KPI% tự cài ở Cài đặt), thêm thẻ Tỉ lệ hoàn (cuối dải Giao thường), đổi được màu nền (🎨). v1.2: bỏ áp quy tắc lọc "ship đi làm" của tool chính vào bảng này, phông chữ to hơn ~50%, bảng shipper giãn hết chiều cao còn trống (bớt cuộn). v1.3: sửa đúng nguyên nhân thiếu "Đơn mới" trong bảng — gộp thêm các tài khoản có quét đơn nhưng không khớp tên với danh sách "đi làm hôm nay" (trước đây bị loại âm thầm khỏi bảng), cỡ chữ giảm lại 20% so với v1.2. v1.4: thêm nút "📅 Ngày khác" xem lại 9 ngày gần nhất (lấy từ Kanban vì API hôm nay không tra được ngày cũ) — bảng ngày cũ rút gọn (không có Chưa quét giao/Đang giao/Tỉ lệ hoàn/Backlog do Kanban không trả đủ trường), vẫn có Lấy hàng. v1.4.1: sửa bảng ngày cũ bị thiếu đơn — gộp các đơn chưa gán tên shipper (ví dụ chưa có người lấy/phát) vào 1 dòng "(Chưa gán shipper)" thay vì bỏ qua hoàn toàn. v1.5: bảng "hôm nay" cũng thêm dòng "(Chưa quét giao)" cho đồng bộ với bảng ngày cũ — TỔNG của bảng giờ khớp thẳng với thẻ "Tổng đơn" phía trên, không cần cộng thêm số ở thẻ "Chưa quét giao" nữa. v1.5.1: sửa TỔNG bảng bị CỘNG THỪA (913 > 895 thẻ) — danh sách "đi làm hôm nay" đôi khi liệt kê trùng 1 tài khoản nhiều dòng, khiến đơn của người đó bị tính 2 lần; giờ chỉ giữ 1 dòng cho mỗi người. v1.6: watermark đổi sang dạng huy hiệu (to hơn ~50%, nền trắng mờ + viền, màu theo theme đang chọn), thêm dòng email dưới tên. v1.6.1: sửa tiếp TỔNG bảng vẫn lệch thẻ (918 vs 895) — cột "Tổng đơn giao"/"Đã ký nhận" từng ưu tiên lấy số báo sẵn trong bảng nhân sự thay vì số tự đếm lại từ đơn hàng (2 nguồn không khớp nhau); giờ luôn dùng số tự đếm khi đã khớp được người, chỉ ước lượng tạm khi không khớp được ai. v1.7: rà soát lại toàn bộ công thức bằng toán (không chỉ vá theo báo lỗi) — bỏ hẳn số ước lượng từ bảng nhân sự (nguồn gây lệch), thêm dòng "(Không khớp người quét)" cho đơn đã quét nhưng đọc không ra tên. Từ bản này, TỔNG bảng = thẻ "Tổng đơn" là một đẳng thức LUÔN đúng theo cấu trúc code, không còn phụ thuộc dữ liệu thực tế của từng ngày. v1.8: tự kiểm tra bản mới trên GitHub (không cần vào Tampermonkey) — khi có bản mới sẽ hiện banner "🆕 Đã có bản mới…" ngay trong tool kèm nút "⬆ Cập nhật" (bấm 1 lần là Tampermonkey tự hiện hộp cài bản mới), cộng thêm chấm sáng trên nút mở tool để biết có bản mới cả khi chưa mở.
// @match        https://v9.800best.com/sea-web/*
// @run-at       document-start
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        none
// @updateURL    https://raw.githubusercontent.com/Erick1939/best-tbc/main/800best-truongbuucuc.user.js
// @downloadURL  https://raw.githubusercontent.com/Erick1939/best-tbc/main/800best-truongbuucuc.user.js
// ==/UserScript==

(() => {
  'use strict';

  /* ------------------------------------------------------------------ *
   * Kiểm tra phiên bản mới (từ GitHub) — nhớ bump SCRIPT_VERSION cùng lúc với
   * @version ở đầu file mỗi lần đổi code, nếu không banner cập nhật sẽ sai.
   * ------------------------------------------------------------------ */
  const SCRIPT_VERSION = '1.8.0';
  const RAW_URL = 'https://raw.githubusercontent.com/Erick1939/best-tbc/main/800best-truongbuucuc.user.js';
  const UPDATE_CHECK_MS = 6 * 60 * 60 * 1000; // 6 tiếng/lần, đỡ gọi GitHub liên tục
  // So 2 chuỗi "x.y.z": trả về true nếu a mới hơn b
  function verNewer(a, b) {
    const pa = String(a ?? '').split('.').map(n => parseInt(n, 10) || 0);
    const pb = String(b ?? '').split('.').map(n => parseInt(n, 10) || 0);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
      const x = pa[i] || 0, y = pb[i] || 0;
      if (x !== y) return x > y;
    }
    return false;
  }

  /* ------------------------------------------------------------------ *
   * State
   * ------------------------------------------------------------------ */
  const THEMES = ['default', 'xanhThan', 'pastel', 'best'];
  const THEME_LABEL = { default: 'Mặc định (tím)', xanhThan: 'Xanh than', pastel: 'Pastel nhẹ', best: 'BEST Express' };
  function loadTheme() { try { const t = localStorage.getItem('tbc_theme'); return THEMES.includes(t) ? t : 'default'; } catch { return 'default'; } }
  function saveTheme(t) { try { localStorage.setItem('tbc_theme', t); } catch { /* bỏ qua */ } }
  const KPI_DEFAULT = { main: 75, tt: 70 };
  function loadKpi() {
    try {
      const o = JSON.parse(localStorage.getItem('tbc_kpi') || 'null');
      if (o && Number.isFinite(+o.main) && Number.isFinite(+o.tt)) return { main: +o.main, tt: +o.tt };
    } catch { /* bỏ qua */ }
    return Object.assign({}, KPI_DEFAULT);
  }
  function saveKpi(k) { try { localStorage.setItem('tbc_kpi', JSON.stringify(k)); } catch { /* bỏ qua */ } }

  const state = {
    office: null,            // { name, id, code }
    busy: false, err: '', note: '',
    calc: null, at: 0,
    discBusy: false, discList: null, discErr: '',
    bl5codes: [],
    tiktok: 'TTVN',
    statusMap: {},
    open: false,
    theme: 'default', kpi: Object.assign({}, KPI_DEFAULT),
    showTheme: false, showSettings: false,
    // Xem ngày cũ: viewDate null = đang xem hôm nay (live); khác null = đang xem ngày đó (Kanban)
    viewDate: null, showDatePick: false, oldBusy: false, oldErr: '', oldNote: '', oldCalc: null, oldAt: 0,
    newVersion: '', updateDismissed: false,
  };

  /* ------------------------------------------------------------------ *
   * Hàm thuần dùng lại từ tool chính (800best-report.user.js)
   * ------------------------------------------------------------------ */
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function parseNum(v) {
    const s = String(v ?? '').trim();
    if (!s) return NaN;
    const m = s.match(/-?\d[\d.,]*/);
    if (!m) return NaN;
    let t = m[0];
    if (/^-?\d{1,3}([.,]\d{3})+$/.test(t)) t = t.replace(/[.,]/g, '');
    else t = t.replace(',', '.');
    return parseFloat(t);
  }
  const fmt = n => (typeof n === 'number' && Number.isFinite(n)) ? n.toLocaleString('vi-VN') : String(n ?? '');
  const round1 = v => Math.round(v * 10) / 10;
  const ratio = (a, b) => (b > 0 && Number.isFinite(a)) ? a / b * 100 : null;
  const pad = n => String(n).padStart(2, '0');
  const plain = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/\s+/g, ' ').trim();
  const plainOf = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/\s+/g, ' ').trim().toUpperCase();
  const normName = s => String(s ?? '').replace(/\(.*?\)/g, ' ').replace(/[\d.]+/g, ' ')
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
  function shipKeys(s) {
    const raw = String(s ?? '').trim();
    if (!raw) return { code: '', name: '' };
    const m = raw.match(/\d{4,}(?:\.\d+)?/);
    return { code: m ? m[0] : '', name: normName(raw) };
  }
  const DATE_RE = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})|(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/;
  function parseDate(s) {
    const m = String(s ?? '').match(DATE_RE);
    if (!m) return null;
    return m[1] ? { d: +m[1], m: +m[2], y: +m[3] } : { d: +m[6], m: +m[5], y: +m[4] };
  }
  const isToday = (s, now) => { const p = parseDate(s); return !!p && p.d === now.getDate() && p.m === now.getMonth() + 1 && p.y === now.getFullYear(); };
  const dayNum = p => Math.round(Date.UTC(p.y, p.m - 1, p.d) / 86400000);
  function daysBefore(s, now) {
    const p = parseDate(s);
    return p ? dayNum({ y: now.getFullYear(), m: now.getMonth() + 1, d: now.getDate() }) - dayNum(p) : null;
  }
  const colBy = (headers, re) => headers.findIndex(h => re.test(h));
  const parseList = text => String(text || '').split(/[,;\s]+/).map(x => x.trim().toUpperCase()).filter(Boolean);
  const isTikTok = (wb, prefixes) => { const w = String(wb ?? '').trim().toUpperCase(); return !!w && prefixes.some(p => w.startsWith(p)); };
  function minOfDay(s) { const m = String(s ?? '').match(/(?:^|\s|T)(\d{1,2}):(\d{2})/); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : null; }
  function defaultCat(status) {
    const t = String(status ?? '').normalize('NFC').toLowerCase();
    if (/đã ký nhận|giao hàng thành công/.test(t)) return 'ok';
    if (/nhập kho|chưa quét|đã đến|chờ phát|chưa phát/.test(t)) return 'unscanned';
    return 'other';
  }
  function classify(status, map) {
    const s = String(status ?? '').trim();
    const v = map && map[s];
    return (v === 'ok' || v === 'unscanned' || v === 'other') ? v : defaultCat(s);
  }
  const dateKey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const keyToUTC = k => { const [y, m, d] = String(k).split('-').map(Number); return Date.UTC(y, m - 1, d); };
  const utcToKey = t => { const d = new Date(t); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };
  const addDays = (k, n) => utcToKey(keyToUTC(k) + n * 86400000);
  const fmtDM = k => `${k.slice(8, 10)}/${k.slice(5, 7)}`;
  const fmtDMY = k => `${fmtDM(k)}/${k.slice(0, 4)}`;
  function rateClass(v, good) {
    if (typeof v !== 'number' || !Number.isFinite(v)) return '';
    if (good === 'low') return v <= 10 ? 'r-ok' : v <= 20 ? 'r-mid' : 'r-bad';
    return v >= 75 ? 'r-ok' : v >= 60 ? 'r-mid' : 'r-bad';
  }
  const asArray = obj => {
    if (Array.isArray(obj)) return obj;
    if (!obj || typeof obj !== 'object') return [];
    for (const k of ['list', 'records', 'rows', 'items', 'resultList', 'dataList', 'siteList', 'children', 'nodes']) if (Array.isArray(obj[k])) return obj[k];
    return obj.data && typeof obj.data === 'object' ? asArray(obj.data) : [];
  };
  function isAuthError(e) {
    return /HTTP\s*(401|403)\b|đăng nhập|dang nhap|login|log in|logged|unauthor|expired|hết hạn|het han|session|phiên|\btoken\b|Unexpected token|not valid JSON/i.test(String((e && e.message) ?? e ?? ''));
  }

  // ---- Quy tắc tính ship đi làm (giống tool chính) ----
  const DM_RULES_DEFAULT = {
    on: true,
    depts: ['Nhân viên giao nhận - Chính thức', 'Nhân viên giao nhận - Chính thức (CARGO)', 'Nhân viên giao nhận - Thuê ngoài dài hạn', 'Nhân viên giao nhận - Thuê ngoài ngắn hạn'],
    cutoff: '14:00', minOk: 5
  };
  const normDept = s => plain(String(s ?? '').replace(/（/g, '(').replace(/）/g, ')')).replace(/\s*-\s*/g, ' - ').replace(/\s*\(\s*/g, ' (').replace(/\s*\)/g, ')').replace(/\s+/g, ' ').trim();
  const hmToMin = t => { const m = /^(\d{1,2}):(\d{2})$/.exec(String(t ?? '').trim()); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : NaN; };
  function dmRulesOf(r) {
    const o = Object.assign({}, DM_RULES_DEFAULT, r && typeof r === 'object' ? r : {});
    o.depts = Array.isArray(o.depts) ? o.depts.map(x => String(x).trim()).filter(Boolean) : DM_RULES_DEFAULT.depts.slice();
    o.on = o.on !== false; o.minOk = Number.isFinite(+o.minOk) ? +o.minOk : 5;
    if (!Number.isFinite(hmToMin(o.cutoff))) o.cutoff = '14:00';
    return o;
  }
  const SNAP_SLOTS = ['23:30'];
  const slotMin = sl => { const m = /^(\d{2}):(\d{2})$/.exec(String(sl)); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : NaN; };
  const dmFinal = now => now.getHours() * 60 + now.getMinutes() >= slotMin(SNAP_SLOTS[SNAP_SLOTS.length - 1]);

  // Shipper đi làm hôm nay: dòng có ngày hôm nay ở cột ngày (cột thứ 4); để trắng = không đi làm
  function workersOf(emp, now) {
    if (!emp || !emp.rows.length) return null;
    const H = emp.headers.map(h => String(h ?? '').normalize('NFC')), R = emp.rows;
    const find = l => { const i = H.findIndex(h => h === l); return i >= 0 ? i : H.findIndex(h => h.includes(l)); };
    const needCol = find('Đơn cần phát'), doneCol = find('Đã quét phát');
    let nameCol = find('Nhân viên giao nhận'); if (nameCol < 0) nameCol = 0;
    const dateCnt = H.map((_, c) => R.filter(r => parseDate(r[c])).length);
    let workCol = dateCnt[3] > 0 ? 3 : dateCnt.findIndex(n => n > 0);
    let list, note = '';
    if (workCol >= 0) {
      list = R.filter(r => isToday(r[workCol], now));
      if (!list.length) {
        list = R.filter(r => String(r[workCol] ?? '').trim() !== '');
        note = `Không dòng nào có ngày hôm nay ở cột "${H[workCol]}"; tạm tính các dòng có ngày.`;
      }
    } else {
      workCol = Math.min(3, H.length - 1);
      list = R.filter(r => String(r[workCol] ?? '').trim() !== '');
      note = 'Không thấy cột ngày; tạm tính các dòng có dữ liệu ở cột thứ 4.';
    }
    return { list, nameCol, needCol, doneCol, workCol, deptCol: find('Chức vụ'), workHeader: H[workCol] || `Cột ${workCol + 1}`, note };
  }

  // ---- Dữ liệu API -> bảng giống bảng trên trang ----
  const STATUS_TXT = { 2: 'Đã nhập kho nhưng chưa quét phát', 3: 'Đang phát hàng', 4: 'Giao hàng thành công', 5: 'Giao hàng thất bại', 22: 'Chờ duyệt', 23: 'Chưa in bill hoàn', 36: 'Từ chối hoàn' };
  const vnParts = ms => { const d = new Date(ms + 7 * 3600000); return { d: pad(d.getUTCDate()), m: pad(d.getUTCMonth() + 1), y: d.getUTCFullYear(), h: pad(d.getUTCHours()), n: pad(d.getUTCMinutes()), s: pad(d.getUTCSeconds()) }; };
  const msDMY = ms => { if (!Number.isFinite(+ms) || +ms < 1e11) return ''; const t = vnParts(+ms); return `${t.d}/${t.m}/${t.y}`; };
  const msDMYHMS = ms => { if (!Number.isFinite(+ms) || +ms < 1e11) return ''; const t = vnParts(+ms); return `${t.d}/${t.m}/${t.y} ${t.h}:${t.n}:${t.s}`; };
  const SIGN_KEYS = ['signTime', 'signingTime', 'signDate', 'signedTime', 'signOffTime', 'finishTime', 'finishDeliveryTime', 'deliveryFinishTime', 'deliveredTime', 'deliverTime', 'deliveryTime', 'latestScanTime', 'lastScanTime', 'scanTime'];
  const msOf = v => { const n = (typeof v === 'string' && /^\d{10,13}$/.test(v.trim())) ? Number(v) : v; return (typeof n === 'number' && Number.isFinite(n) && n >= 1e11 && n < 4e12) ? n : null; };
  function signMsOf(it) { for (const k of SIGN_KEYS) { const m = msOf(it && it[k]); if (m !== null) return m; } return null; }
  function apiToTables(items, empList) {
    const seen = new Set(), rows = [], unknown = new Set(), codeById = new Map();
    let okN = 0, signN = 0, dupN = 0, emptyN = 0;
    for (const it of (items || [])) {
      if (it && it.dispDelivererId && it.delivererCode && !codeById.has(it.dispDelivererId)) codeById.set(it.dispDelivererId, it.delivererCode);
      const wb = String(it && it.billCode || '').trim().toUpperCase();
      if (!wb) emptyN++; else if (seen.has(wb)) dupN++; else seen.add(wb);
      const code = String(it.parcelStatus ?? '');
      let st = STATUS_TXT[code];
      if (!st) { st = `Trạng thái ${code}`; unknown.add(code); }
      const signMs = code === '4' ? signMsOf(it) : null;
      if (code === '4') { okN++; if (signMs !== null) signN++; }
      rows.push([wb, msDMY(it.predictTimeByDeparture), st, String(it.latestScanMan || '').trim(), msDMY(it.forwardAgingEndDate), signMs !== null ? msDMYHMS(signMs) : '']);
    }
    const orders = { headers: ['Mã vận đơn', 'Ngày cần phát hàng', 'Trạng thái đơn hàng', 'Người quét mới nhất', 'Thời hạn giao hàng nội bộ', 'Giờ giao thành công'], rows, wbCol: 0 };
    const eRows = (empList || []).map(e => {
      const c = codeById.get(e.dispDelivererId);
      return [`${String(e.delivererName || '').trim()}${c ? `(${c})` : ''}`, e.shouldDispCount ?? 0, e.todayDispCount ?? 0, msDMYHMS(e.firstDispScanTime), e.signCount ?? 0, String(e.departmentName ?? '').trim()];
    });
    const emp = { headers: ['Nhân viên giao nhận', 'Đơn cần phát', 'Đã quét phát', 'Quét phát đầu tiên', 'Đã ký nhận', 'Chức vụ'], rows: eRows };
    return { orders, emp, unknown: [...unknown], okN, signN, dupN, emptyN };
  }

  /* ------------------------------------------------------------------ *
   * Tính dashboard (giống computeDashboard của tool chính, có thêm phần mở rộng
   * cho "Trưởng Bưu Cục": Đơn mới / POD2400 / Đơn đang giao theo từng shipper)
   * ------------------------------------------------------------------ */
  function computeDashboard({ orders, emp, tiktok, statusMap, now, rules }) {
    const H = orders.headers.map(h => String(h ?? '').normalize('NFC')), wb = orders.wbCol, rows = orders.rows;
    const stCol = colBy(H, /trạng thái/i);
    const scCol = colBy(H, /người quét|nhân viên|shipper|người phát|giao nhận/i);
    const dCol = colBy(H, /ngày cần phát/i);
    const oCol = colBy(H, /(thời\s*(hạn|gian)|hạn).*nội\s*bộ|odr/i);
    const tCol = colBy(H, /giờ giao thành công/i);
    const bl = { total: 0, wait: 0, fail: 0, unscanned: 0, d5: 0 };
    const bl5 = { total: 0, wait: 0, fail: 0, unscanned: 0, codes: [] };
    let odr = 0;
    const mkAgg = () => ({ total: 0, ok: 0, unscanned: 0, newTotal: 0, newOk: 0, newUns: 0, hoan: 0, orphan: 0, orphanNew: 0 });
    const all = mkAgg(), tt = mkAgg(), sme = mkAgg();
    const byCode = new Map(), byName = new Map(), stats = [];

    const getStat = cell => {
      const k = shipKeys(cell);
      if (!k.code && !k.name) return null;
      let o = k.code ? byCode.get(k.code) : byName.get(k.name);
      if (!o) { o = { label: String(cell).trim(), total: 0, ok: 0, ttTotal: 0, ttOk: 0, newTotal: 0, newOk: 0, transit: 0 }; stats.push(o); }
      if (k.code) byCode.set(k.code, o);
      if (k.name && !byName.has(k.name)) byName.set(k.name, o);
      return o;
    };

    for (const r of rows) {
      const isTT = isTikTok(r[wb], tiktok);
      const cat = stCol >= 0 ? classify(r[stCol], statusMap) : 'other';
      const dateStr = dCol >= 0 ? r[dCol] : '';
      const isNew = dCol >= 0 && isToday(dateStr, now);
      const back = dCol >= 0 ? daysBefore(dateStr, now) : null;
      const isHoan = stCol >= 0 && /hoàn/i.test(String(r[stCol] ?? ''));
      for (const a of (isTT ? [all, tt] : [all, sme])) {
        a.total++;
        if (cat === 'ok') a.ok++; else if (cat === 'unscanned') a.unscanned++;
        if (isNew) { a.newTotal++; if (cat === 'ok') a.newOk++; else if (cat === 'unscanned') a.newUns++; }
        if (isHoan) a.hoan++;
      }
      // Backlog 3 ngày (giữ để đối chiếu) + Backlog 5 ngày (mới, có danh sách mã đơn để tải)
      if (back !== null && back >= 3 && cat !== 'ok') {
        bl.total++; if (back >= 5) bl.d5++;
        const st = String(stCol >= 0 ? r[stCol] : '').normalize('NFC').toLowerCase();
        if (/chờ duyệt/.test(st)) bl.wait++; else if (/thất bại|không thành công/.test(st)) bl.fail++; else if (cat === 'unscanned') bl.unscanned++;
      }
      if (back !== null && back >= 5 && cat !== 'ok') {
        bl5.total++;
        const st = String(stCol >= 0 ? r[stCol] : '').normalize('NFC').toLowerCase();
        if (/chờ duyệt/.test(st)) bl5.wait++; else if (/thất bại|không thành công/.test(st)) bl5.fail++; else if (cat === 'unscanned') bl5.unscanned++;
        if (r[wb]) bl5.codes.push(String(r[wb]));
      }
      if (isTT && oCol >= 0 && isToday(r[oCol], now)) odr++;
      // Đơn đã quét giao được tính cho người quét mới nhất
      let attributed = false;
      if (cat !== 'unscanned' && scCol >= 0) {
        const o = getStat(r[scCol]);
        if (o) {
          attributed = true;
          o.total++;
          if (cat === 'ok') o.ok++;
          if (isTT) { o.ttTotal++; if (cat === 'ok') o.ttOk++; }
          if (isNew) { o.newTotal++; if (cat === 'ok') o.newOk++; }
          if (String(stCol >= 0 ? r[stCol] : '').trim() === 'Đang phát hàng') o.transit++;
        }
      }
      // "Mồ côi": đơn ĐÃ quét (cat !== 'unscanned') nhưng không gắn được vào shipper nào — ô
      // "người quét" trống/không đọc được tên, hoặc cột đó không tồn tại trong bảng đơn. Trước
      // đây số này biến mất hoàn toàn khỏi bảng (không nằm trong "(Chưa quét giao)" vì không
      // phải chưa quét, cũng không nằm trong o.total nào) khiến TỔNG bảng luôn có thể thấp hơn
      // thẻ "Tổng đơn" một chút mà không rõ lý do — giờ gộp riêng thành 1 dòng để không mất đơn.
      if (cat !== 'unscanned' && !attributed) {
        for (const a of (isTT ? [all, tt] : [all, sme])) { a.orphan++; if (isNew) a.orphanNew++; }
      }
    }

    const W = workersOf(emp, now);
    const ftIdx = W ? emp.headers.findIndex(h => h.includes('Quét phát đầu tiên')) : -1;
    let nWork = W ? W.list.length : stats.filter(o => o.total > 0).length;
    const DM = dmRulesOf(rules);
    const used = new Set();
    let shippers;
    if (W) {
      // Khử trùng: danh sách "đi làm hôm nay" (emp sheet) đôi khi liệt kê CÙNG 1 tài khoản
      // nhiều hơn 1 dòng (nhiều ca, dữ liệu trùng…) — nếu map thẳng từng dòng thành 1 shipper
      // thì đơn của người đó bị cộng 2 lần vào TỔNG của bảng (TỔNG > thẻ "Tổng đơn"). Chỉ giữ
      // 1 dòng cho mỗi người (ưu tiên dòng có "Tổng đơn giao" lớn hơn).
      const seenO = new Map(), seenName = new Map();
      shippers = [];
      for (const row of W.list) {
        const cell = row[W.nameCol];
        const k = shipKeys(cell);
        const o = (k.code && byCode.get(k.code)) || (k.name && byName.get(k.name)) || null;
        // CHỈ dùng số tự đếm lại từ danh sách đơn (o.total/o.ok) — cùng nguồn với thẻ "Tổng đơn"
        // phía trên, đảm bảo cộng lại LUÔN khớp đúng bằng toán học, không phụ thuộc dữ liệu thực
        // tế. Không dùng số "Đã quét phát"/"Đã ký nhận" báo sẵn trong bảng nhân sự nữa — số đó
        // đến từ nguồn khác (site tự đếm riêng), không đối chiếu được nên từng gây TỔNG bảng lệch
        // thẻ dù đã hết trùng dòng nhân viên. Không khớp được người nào thì hiện 0 (trung thực,
        // còn hơn hiện số không đối chiếu được).
        const s = {
          name: cell, matched: !!o, scanned: o ? o.total : 0,
          ok: o ? o.ok : 0,
          ttTotal: o ? o.ttTotal : 0, ttOk: o ? o.ttOk : 0,
          newTotal: o ? o.newTotal : 0, newOk: o ? o.newOk : 0, transit: o ? o.transit : 0,
          ft: (ftIdx >= 0 && isToday(row[ftIdx], now)) ? minOfDay(row[ftIdx]) : null, _row: row
        };
        const dupMap = o ? seenO : seenName, dupKey = o || normName(cell);
        const prev = dupKey ? dupMap.get(dupKey) : undefined;
        if (prev) {
          if (s.scanned > prev.scanned) { const i = shippers.indexOf(prev); if (i >= 0) shippers[i] = s; dupMap.set(dupKey, s); }
          continue;
        }
        if (dupKey) dupMap.set(dupKey, s);
        shippers.push(s);
      }
      shippers.forEach(s => { const k = shipKeys(s.name); const o = (k.code && byCode.get(k.code)) || (k.name && byName.get(k.name)) || null; if (o) used.add(o); });
      if (DM.on) {
        const hasDept = W.deptCol >= 0 && W.list.some(r => String(r[W.deptCol] ?? '').trim() !== '');
        const allow = new Set(DM.depts.map(normDept)), cut = hmToMin(DM.cutoff), fin = dmFinal(now);
        shippers = shippers.filter(s => {
          if (hasDept && allow.size && !allow.has(normDept(s._row[W.deptCol]))) return false;
          if (Number.isFinite(cut) && Number.isFinite(s.ft) && s.ft > cut) return false;
          if (fin && DM.minOk > 0 && !((Number.isFinite(s.ok) ? s.ok : 0) >= DM.minOk)) return false;
          return true;
        });
      }
      shippers.forEach(s => { delete s._row; });
      // Các tài khoản có quét đơn nhưng KHÔNG khớp được tên với ai trong danh sách
      // "đi làm hôm nay" (emp sheet) — trước đây bị âm thầm loại bỏ hoàn toàn khỏi bảng,
      // khiến tổng "Đơn mới"/"Tổng đơn giao" của bảng thấp hơn hẳn thẻ phía trên. Gộp thêm
      // vào bảng (đánh dấu matched:false) để không mất đơn.
      const extra = stats.filter(o => !used.has(o) && o.total > 0).map(o => ({
        name: o.label, matched: false, scanned: o.total, ok: o.ok, ttTotal: o.ttTotal, ttOk: o.ttOk,
        newTotal: o.newTotal, newOk: o.newOk, transit: o.transit
      }));
      if (extra.length) shippers = shippers.concat(extra);
      nWork = shippers.length;
    } else {
      shippers = stats.filter(o => o.total > 0).map(o => ({
        name: o.label, matched: true, scanned: o.total, ok: o.ok, ttTotal: o.ttTotal, ttOk: o.ttOk,
        newTotal: o.newTotal, newOk: o.newOk, transit: o.transit
      }));
    }
    // Đơn CHƯA ai quét giao thì không có tên shipper để gán — gộp thành 1 dòng riêng
    // "(Chưa quét giao)" trong bảng, để TỔNG của bảng khớp thẳng với thẻ "Tổng đơn" ở trên
    // mà không cần cộng thêm số ở thẻ "Chưa quét giao".
    if (all.unscanned > 0) {
      shippers.push({
        name: '(Chưa quét giao)', matched: false, scanned: all.unscanned, ok: 0,
        ttTotal: tt.unscanned, ttOk: 0, newTotal: all.newUns, newOk: 0, transit: 0
      });
    }
    // Đơn ĐÃ quét nhưng ô "người quét" trống/không đọc được tên (hoặc bảng đơn không có cột đó)
    // — gộp riêng dòng này để TỔNG bảng = thẻ "Tổng đơn" LUÔN đúng, không phụ thuộc dữ liệu.
    if (all.orphan > 0) {
      shippers.push({
        name: '(Không khớp người quét)', matched: false, scanned: all.orphan, ok: 0,
        ttTotal: tt.orphan, ttOk: 0, newTotal: all.orphanNew, newOk: 0, transit: 0
      });
    }
    shippers.forEach(s => {
      s.rateTrip = ratio(s.ok, s.scanned);
      s.rateTT = ratio(s.ttOk, s.ttTotal);
      s.pod = ratio(s.newOk, s.newTotal);
    });

    const sum = k => shippers.reduce((a, s) => a + (Number.isFinite(s[k]) ? s[k] : 0), 0);
    const totals = { name: 'TỔNG', scanned: sum('scanned'), ok: sum('ok'), ttTotal: sum('ttTotal'), ttOk: sum('ttOk'), newTotal: sum('newTotal'), newOk: sum('newOk'), transit: sum('transit') };
    totals.rateTrip = ratio(totals.ok, totals.scanned);
    totals.rateTT = ratio(totals.ttOk, totals.ttTotal);
    totals.pod = ratio(totals.newOk, totals.newTotal);

    const card = a => ({
      pod: ratio(a.newOk, a.newTotal), newOk: a.newOk, newTotal: a.newTotal,
      total: a.total, unscanned: a.unscanned, scanned: a.total - a.unscanned, ok: a.ok,
      scanRate: ratio(a.total - a.unscanned, a.total),
      gtcTrip: ratio(a.ok, a.total - a.unscanned),
      gtcTotal: ratio(a.ok, a.total),
      hoan: a.hoan, hoanPct: ratio(a.hoan, a.total)
    });
    return {
      cards: { all: card(all), tt: card(tt), sme: card(sme) },
      extra: { backlog: dCol >= 0 ? bl : null, backlog5: dCol >= 0 ? bl5 : null, odrTT: oCol >= 0 ? odr : null },
      hasTT: tt.total > 0, nWork, W, shippers, totals
    };
  }

  // Số đơn mới cần giao thành công thêm để đạt KPI% POD2400 (0 nếu đã đạt hoặc không tính được)
  function needForKpi(newOk, newTotal, kpiPct) {
    if (!Number.isFinite(newTotal) || newTotal <= 0 || !Number.isFinite(kpiPct)) return null;
    return Math.max(0, Math.ceil(kpiPct / 100 * newTotal - (Number.isFinite(newOk) ? newOk : 0) - 1e-9));
  }
  // Thẻ chỉ số: [nhãn, giá trị, kiểu, màu, chi tiết?, cờ?]
  function cardList(c, kpi) {
    const K = Object.assign({}, KPI_DEFAULT, kpi || {});
    const a = c.cards.all, t = c.cards.tt, s = c.cards.sme, e = c.extra || {}, bl5 = e.backlog5;
    const needMain = needForKpi(a.newOk, a.newTotal, K.main), needTT = needForKpi(t.newOk, t.newTotal, K.tt);
    return {
      main: [
        ['Tổng đơn', a.total, 'int', 'blue'], ['Tỷ lệ quét giao', a.scanRate, 'pct', 'green'],
        ['POD2400', a.pod, 'pct', 'amber', [['Giao thành công', a.newOk], ['Tổng đơn mới hôm nay', a.newTotal]]],
        ['Chưa quét giao', a.unscanned, 'int', 'purple', [['SME', s.unscanned], ['TikTok', t.unscanned]]],
        [`Cần đạt KPI ${fmt(K.main)}%`, needMain, 'int', 'teal', [['Đã giao', a.newOk], ['Tổng đơn mới', a.newTotal]]],
        ['GTC Total', a.gtcTotal, 'pct', 'orange'],
        ['Backlog 5 ngày', bl5 ? bl5.total : null, 'int', 'red',
          bl5 ? [['Chờ duyệt', bl5.wait], ['Giao thất bại', bl5.fail], ['Nhập kho chưa quét phát', bl5.unscanned]] : null, 'bl5dl'],
        ['Tỉ lệ hoàn', a.hoanPct, 'pct', 'slate', [['Số đơn hoàn', a.hoan]]]
      ],
      tts: [
        ['Tổng đơn TikTok', t.total, 'int', 'pink'], ['Tỷ lệ quét TikTok', t.scanRate, 'pct', 'green'],
        ['TikTok POD2400', t.pod, 'pct', 'amber', [['Giao thành công', t.newOk], ['Tổng đơn mới hôm nay', t.newTotal]]],
        ['TikTok chưa quét giao', t.unscanned, 'int', 'purple'],
        [`TikTok cần đạt KPI ${fmt(K.tt)}%`, needTT, 'int', 'teal', [['Đã giao', t.newOk], ['Tổng đơn mới', t.newTotal]]],
        ['GTC Total TikTok', t.gtcTotal, 'pct', 'orange'],
        ['TikTok đến hạn ODR', e.odrTT, 'int', 'red']
      ]
    };
  }

  function sortRows(rows, key, dir) {
    if (!key) return rows;
    const isN = v => typeof v === 'number' && Number.isFinite(v);
    return [...rows].sort((a, b) => {
      const x = a[key], y = b[key];
      if (isN(x) && isN(y)) return (x - y) * dir;
      if (isN(x)) return -1;
      if (isN(y)) return 1;
      return String(x ?? '').localeCompare(String(y ?? ''), 'vi') * dir;
    });
  }

  /* ------------------------------------------------------------------ *
   * Mạng: bắt header xác thực của trang, gọi API bằng cookie của trang
   * ------------------------------------------------------------------ */
  const apiHdr = (window.__tbcHdr = window.__tbcHdr || {});
  function ensureHdrHook() {
    try {
      const P = XMLHttpRequest.prototype;
      if (P.setRequestHeader && P.setRequestHeader.__tbc) return;
      const cur = P.setRequestHeader;
      const ours = function (k, v) {
        try { if (/^(token|lang-type|langType|TIME-ZONE|no-offset-timeZone|data-masking-N|Accept)$/i.test(k)) apiHdr[k] = v; } catch { /* bỏ qua */ }
        return cur.apply(this, arguments);
      };
      ours.__tbc = true;
      P.setRequestHeader = ours;
    } catch { /* bỏ qua */ }
  }
  const API_DET = '/v9-order/amTodayDispBoard/getTodayDetail', API_EMP = '/v9-order/amTodayDispBoard/getTodayEmployeeSum', API_SITE = '/sea-man/node/pageQuery';
  const API_PU_DET = '/v9-order/dc/kanban/am/get-srec-detail';
  // ---- API Kanban: dùng để xem lại NGÀY CŨ (API đơn hôm nay ở trên chỉ có đúng ngày hôm nay, không tra được ngày khác) ----
  const API_KB_DET = '/v9-order/dc/kanban/am/vn/site-kpi-kanban/get-sdisp-detail', API_KB_SUM = '/v9-order/dc/kanban/am/vn/site-kpi-kanban/get-sdisp';
  const KB_DAYS_BACK = 9; // giữ 9 ngày gần nhất để chọn xem, giống tool chính
  async function apiPost(path, body) {
    const hdr = Object.assign({ Accept: 'application/json, text/plain, */*', 'lang-type': 'vi-VN', langType: 'vi-VN', 'TIME-ZONE': 'GMT+7', 'data-masking-N': 'mark', 'no-offset-timeZone': 'true', token: '', 'X-Requested-With': 'XMLHttpRequest' }, apiHdr);
    for (const k of Object.keys(hdr)) if (/^lang-?type$/i.test(k)) delete hdr[k];
    hdr['lang-type'] = 'vi-VN'; hdr.langType = 'vi-VN';
    for (const k of Object.keys(hdr)) if (typeof hdr[k] === 'boolean') hdr[k] = String(hdr[k]);
    hdr['Content-Type'] = 'application/json;charset=UTF-8';
    const r = await fetch(location.origin + path, { method: 'POST', credentials: 'include', headers: hdr, body: JSON.stringify(body) });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const j = JSON.parse(await r.text());
    if (!j || j.success === false || (j.success === undefined && j.code !== undefined && Number(j.code) !== 200)) throw new Error((j && (j.errMsg || j.message || j.msg)) || 'API báo thất bại');
    return j.data;
  }

  /* ---- Tự dò bưu cục đang quản lý ---- */
  const DISC_SKIP_TYPE = ['DC', 'KA'];
  const discName = x => { const v = x && x.nameLocale; const t = typeof v === 'string' ? v : (v && typeof v === 'object' ? (Object.values(v).find(q => typeof q === 'string') || '') : (x && x.name) || ''); return String(t || '').trim().replace(/\s+/g, ' '); };
  function discPick(list) {
    const cand = [], seen = new Set(), seenName = new Set();
    for (const x of Array.isArray(list) ? list : []) {
      if (!x || typeof x !== 'object') continue;
      const name = discName(x), id = x.id, code = x.code;
      if (!name || !id || code === undefined || code === null || code === '') continue;
      const key = String(id); if (seen.has(key)) continue; seen.add(key);
      if (x.status !== 'ENABLE') continue;
      if (seenName.has(name)) continue; seenName.add(name);
      if (DISC_SKIP_TYPE.includes(String(x.type || '').toUpperCase()) || String(x.logisticsType || '').toUpperCase() === 'CARGO') continue;
      cand.push({ name, id, code: String(code) });
    }
    return cand;
  }
  async function discList() {
    const all = []; let total = Infinity;
    for (let pg = 1; pg <= 20 && all.length < total; pg++) {
      const d = await apiPost(API_SITE, { pageNumber: pg, pageSize: 200, fuzzyNameOrCode: '', limitSiteByAuth: true, statusList: ['ENABLE', 'SHUTDOWN'] });
      const arr = asArray(d);
      const t = Number(d && d.total); total = Number.isFinite(t) && t >= 0 ? t : arr.length;
      if (!arr.length) break;
      all.push(...arr);
    }
    return all;
  }
  async function discProbe(c) {
    const d = await apiPost(API_DET, { isContainsSecSite: true, siteId: c.id, siteCode: c.code, shouldDispType: '1', searchType: '1', pageSize: 20, currentPage: 1 });
    return Math.max(asArray(d).length, Number(d && d.total) || 0);
  }
  // Trả về danh sách bưu cục có đơn hôm nay (= đang được tài khoản này quản lý)
  async function discoverMyOffices() {
    const list = await discList();
    const cand = discPick(list);
    const ok = []; let p = 0;
    const worker = async () => { while (p < cand.length) { const c = cand[p++]; try { const n = await discProbe(c); if (n > 0) ok.push(c); } catch { /* bỏ qua */ } } };
    await Promise.all([worker(), worker(), worker()]);
    ok.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
    return ok;
  }

  function loadSavedOffice() {
    try { const o = JSON.parse(localStorage.getItem('tbc_office') || 'null'); if (o && o.id && o.code && o.name) return o; } catch { /* bỏ qua */ }
    return null;
  }
  function saveOffice(o) { state.office = o; try { localStorage.setItem('tbc_office', JSON.stringify(o)); } catch { /* bỏ qua */ } }

  /* ---- Lấy toàn bộ đơn + nhân viên của 1 bưu cục hôm nay ---- */
  async function fetchOfficeToday(site) {
    const base = { isContainsSecSite: true, siteId: site.id, siteCode: site.code, shouldDispType: '1' };
    const det = Object.assign({}, base, { searchType: '1' });
    const empP = apiPost(API_EMP, base);
    empP.catch(() => { /* lỗi sẽ báo khi đợi bên dưới */ });
    const sizes = [1000, 500, 200, 100, 50];
    let first = null, lastErr = null;
    for (const ps of sizes) {
      try { first = await apiPost(API_DET, Object.assign({}, det, { pageSize: ps, currentPage: 1 })); break; }
      catch (e) { lastErr = e; }
    }
    if (!first) throw lastErr || new Error('API đơn không trả dữ liệu');
    const empData = await empP;
    let items = ((first && first.list) || []).slice();
    const total = Number(first && first.total) || 0, per = items.length;
    if (per && items.length < total) {
      const pages = Math.min(400, Math.ceil(total / per));
      const nums = []; for (let p = 2; p <= pages; p++) nums.push(p);
      for (let i = 0; i < nums.length; i += 6) {
        const got = await Promise.all(nums.slice(i, i + 6).map(p => apiPost(API_DET, Object.assign({}, det, { pageSize: per, currentPage: p })).then(d => (d && d.list) || [])));
        got.forEach(g => items.push(...g));
      }
      const sg = new Set(), uniq = [];
      for (const it of items) { let k; try { k = JSON.stringify(it); } catch { k = null; } if (k !== null) { if (sg.has(k)) continue; sg.add(k); } uniq.push(it); }
      items = uniq;
    }
    const t = apiToTables(items, asArray(empData));
    return t;
  }

  /* ---- Lấy toàn bộ đơn lấy hàng (pickup) của 1 bưu cục hôm nay ---- */
  async function fetchPickupToday(code, date) {
    const base = { siteCode: code, isSelfData: 0, isCargo: 0, isExpress: 0, nationCode: 'VN', statisticsDate: `${date} 00:00:00`, dataKey: 'srecAhCnt' };
    const seen = new Map(); let size = 1000, total = null, got = 0;
    for (let page = 1; page <= 60; page++) {
      const d = await apiPost(API_PU_DET, Object.assign({}, base, { page, size }));
      const list = d && Array.isArray(d.list) ? d.list : [];
      if (total === null) total = Number(d && d.total) || 0;
      for (const r of list) { got++; seen.set(r && r.billCode ? String(r.billCode) : `#${got}`, r); }
      if (!list.length || got >= total) break;
    }
    return [...seen.values()];
  }
  // Gộp theo shipper (smallManName): tổng đơn lấy / đã lấy thành công (có inNetworkTime) / đang lấy (isHistory=0, chưa có kết quả)
  function pickupByShipper(rows) {
    const m = new Map();
    for (const r of rows || []) {
      if (!r || typeof r !== 'object') continue;
      const nm = String(r.smallManName ?? '').trim(); if (!nm) continue;
      const key = normName(nm);
      let o = m.get(key); if (!o) m.set(key, o = { total: 0, ok: 0, pending: 0 });
      o.total++;
      if (r.inNetworkTime) o.ok++;
      else if (Number(r.isHistory) === 0) o.pending++;
    }
    return m;
  }

  /* ------------------------------------------------------------------ *
   * Xem NGÀY CŨ: nguồn Kanban (get-sdisp-detail), mỗi đơn 1 dòng nhưng ít trường hơn
   * API "hôm nay" của bảng điều phối, nên bảng ngày cũ chỉ có: Tổng đơn / Đơn mới /
   * POD2400 / Tỉ lệ giao TC (không có Chưa quét giao / Đang giao / Tỉ lệ hoàn / Backlog
   * vì Kanban không trả trạng thái chi tiết hay bộ phận). Lấy hàng (pickup) thì vẫn
   * tra được API cũ (fetchPickupToday ở trên nhận date tuỳ ý).
   * ------------------------------------------------------------------ */
  async function fetchKbDay(code, date) {
    const base = { siteCode: code, isSelfData: 0, isCargo: 0, isExpress: 0, nationCode: 'VN', statisticsDate: `${date} 00:00:00` };
    const items = []; let size = 200, total = null;
    for (let page = 1; page <= 80; page++) {
      let d;
      try { d = await apiPost(API_KB_DET, Object.assign({}, base, { dataKey: 'allCnt', page, size })); }
      catch (e) { if (page === 1 && size > 50) { size = 50; page--; continue; } throw e; }
      const list = d && Array.isArray(d.list) ? d.list : [];
      if (total === null) total = Number(d && d.total) || 0;
      items.push(...list);
      if (!list.length || items.length >= total) break;
    }
    let sumTotal = null;
    try {
      const s = await apiPost(API_KB_SUM, Object.assign({}, base, { timeFrom: `${date}T00:00:00.000+07:00` }));
      sumTotal = s && Number.isFinite(Number(s.topicIndicatorValue)) ? Number(s.topicIndicatorValue) : null;
    } catch { /* không đối chiếu được thì thôi, vẫn dùng số đã đếm */ }
    return { items, total: Number.isFinite(total) ? total : items.length, sumTotal };
  }
  // Tính dashboard rút gọn cho 1 ngày cũ từ danh sách đơn Kanban
  function computeOldDay(items, tiktok) {
    const mkAgg = () => ({ total: 0, newTotal: 0, newOk: 0, ok: 0 });
    const all = mkAgg(), tt = mkAgg();
    const byName = new Map(), stats = [];
    // Đơn chưa có tên shipper (chưa gán người quét/lấy) vẫn phải tính vào bảng, gộp chung
    // vào 1 dòng "(Chưa gán shipper)" — trước đây bị bỏ qua hoàn toàn (getStat trả null khi
    // tên rỗng) nên TỔNG của bảng thấp hơn thẻ chỉ số phía trên dù số liệu thẻ vẫn đúng.
    const getStat = nm => {
      const raw = String(nm ?? '').trim();
      const key = raw ? (normName(raw) || '__unassigned__') : '__unassigned__';
      let o = byName.get(key);
      if (!o) { o = { label: raw || '(Chưa gán shipper)', total: 0, newTotal: 0, newOk: 0, ok: 0, ttTotal: 0, ttOk: 0 }; byName.set(key, o); stats.push(o); }
      return o;
    };
    for (const x of (items || [])) {
      if (!x || typeof x !== 'object') continue;
      const isTT = isTikTok(x.billCode, tiktok);
      const isNew = Number(x.isTodaySdisp) === 1;
      const signed = !!x.isSign;
      const okNew = Number(x.is2400PodSign) === 1;
      for (const a of (isTT ? [all, tt] : [all])) {
        a.total++; if (signed) a.ok++;
        if (isNew) { a.newTotal++; if (okNew) a.newOk++; }
      }
      const o = getStat(x.smallManName);
      if (o) {
        o.total++; if (signed) o.ok++;
        if (isTT) { o.ttTotal++; if (signed) o.ttOk++; }
        if (isNew) { o.newTotal++; if (okNew) o.newOk++; }
      }
    }
    const shippers = stats.filter(o => o.total > 0).map(o => Object.assign({}, o, {
      scanned: o.total, rateTrip: ratio(o.ok, o.total), rateTT: ratio(o.ttOk, o.ttTotal), pod: ratio(o.newOk, o.newTotal)
    }));
    const sum = k => shippers.reduce((a, s) => a + (Number.isFinite(s[k]) ? s[k] : 0), 0);
    const totals = { name: 'TỔNG', scanned: sum('scanned'), ok: sum('ok'), ttTotal: sum('ttTotal'), ttOk: sum('ttOk'), newTotal: sum('newTotal'), newOk: sum('newOk') };
    totals.rateTrip = ratio(totals.ok, totals.scanned); totals.rateTT = ratio(totals.ttOk, totals.ttTotal); totals.pod = ratio(totals.newOk, totals.newTotal);
    const card = a => ({ total: a.total, ok: a.ok, newTotal: a.newTotal, newOk: a.newOk, pod: ratio(a.newOk, a.newTotal), gtcTotal: ratio(a.ok, a.total) });
    return { cards: { all: card(all), tt: card(tt) }, hasTT: tt.total > 0, shippers, totals };
  }
  async function loadOldDay(date) {
    if (state.oldBusy || !state.office) return;
    state.oldBusy = true; state.oldErr = ''; state.viewDate = date; render();
    try {
      const kb = await fetchKbDay(state.office.code, date);
      if (!kb.items.length) throw new Error('Không có dữ liệu đơn cho ngày này (có thể quá cũ hoặc Kanban chưa chốt số)');
      if (Number.isFinite(kb.sumTotal) && kb.sumTotal !== kb.items.length) {
        state.oldNote = `Lưu ý: đếm được ${fmt(kb.items.length)} đơn nhưng Kanban báo tổng ${fmt(kb.sumTotal)} đơn — số liệu có thể lệch nhẹ.`;
      } else state.oldNote = '';
      const tk = parseList(state.tiktok);
      const calc = computeOldDay(kb.items, tk);
      let pkRows = [];
      try { pkRows = await fetchPickupToday(state.office.code, date); } catch { /* bỏ qua, cột lấy hàng để trống */ }
      const pkMap = pickupByShipper(pkRows);
      for (const s of calc.shippers) {
        const pk = pkMap.get(normName(s.label));
        s.pkTotal = pk ? pk.total : 0; s.pkOk = pk ? pk.ok : 0; s.pkPending = pk ? pk.pending : 0; s.pkRate = pk ? ratio(pk.ok, pk.total) : null;
      }
      const ptt = { total: 0, ok: 0, pending: 0 };
      for (const o of pkMap.values()) { ptt.total += o.total; ptt.ok += o.ok; ptt.pending += o.pending; }
      calc.totals.pkTotal = ptt.total; calc.totals.pkOk = ptt.ok; calc.totals.pkPending = ptt.pending; calc.totals.pkRate = ratio(ptt.ok, ptt.total);
      state.oldCalc = calc; state.oldAt = Date.now();
    } catch (e) {
      state.oldErr = `Không tải được số liệu ngày ${fmtDM(date)} (${e && e.message || e}).`;
    } finally { state.oldBusy = false; render(); }
  }

  /* ------------------------------------------------------------------ *
   * Tải số liệu
   * ------------------------------------------------------------------ */
  async function loadData() {
    if (state.busy || !state.office) return;
    state.busy = true; state.err = ''; render();
    const now = new Date();
    try {
      const t = await fetchOfficeToday(state.office);
      // rules:{on:false} — bảng này cần liệt kê TẤT CẢ tài khoản đã quét hàng trong ngày,
      // không áp quy tắc lọc "ship đi làm" (bộ phận / giờ quét đầu / tối thiểu đơn thành công)
      // của tool chính (ANM), vì quy tắc đó sau 23:30 sẽ ẩn hẳn những ship có <5 đơn thành công
      // khỏi bảng dù họ vẫn có đơn mới/đã quét — gây lệch tổng so với thẻ chỉ số đầu trang.
      const calc = computeDashboard({ orders: t.orders, emp: t.emp, tiktok: parseList(state.tiktok), statusMap: state.statusMap, now, rules: { on: false } });
      let pkRows = [];
      try { pkRows = await fetchPickupToday(state.office.code, dateKey(now)); }
      catch (e) { state.note = `Không tải được số liệu lấy hàng (${e && e.message || e}); bảng shipper tạm để trống nhóm "Chỉ số lấy hàng".`; }
      const pkMap = pickupByShipper(pkRows);
      for (const s of calc.shippers) {
        const pk = pkMap.get(normName(s.name));
        s.pkTotal = pk ? pk.total : 0; s.pkOk = pk ? pk.ok : 0; s.pkPending = pk ? pk.pending : 0;
        s.pkRate = pk ? ratio(pk.ok, pk.total) : null;
      }
      const tt = { total: 0, ok: 0, pending: 0 };
      for (const o of pkMap.values()) { tt.total += o.total; tt.ok += o.ok; tt.pending += o.pending; }
      calc.totals.pkTotal = tt.total; calc.totals.pkOk = tt.ok; calc.totals.pkPending = tt.pending; calc.totals.pkRate = ratio(tt.ok, tt.total);
      state.calc = calc;
      state.bl5codes = (calc.extra.backlog5 && calc.extra.backlog5.codes) || [];
      state.at = Date.now();
    } catch (e) {
      state.err = `Không tải được số liệu (${e && e.message || e}).`;
    } finally { state.busy = false; render(); }
  }

  async function initOffice() {
    const saved = loadSavedOffice();
    if (saved) { state.office = saved; await loadData(); return; }
    state.discBusy = true; state.discErr = ''; render();
    try {
      const found = await discoverMyOffices();
      if (!found.length) { state.discErr = 'Không tìm thấy bưu cục nào bạn quản lý có đơn hôm nay. Bấm "Thử lại" hoặc kiểm tra tài khoản đăng nhập.'; }
      else if (found.length === 1) { saveOffice(found[0]); await loadData(); }
      else { state.discList = found; }
    } catch (e) {
      state.discErr = `Không dò được bưu cục (${e && e.message || e}). Bấm "Thử lại".`;
    } finally { state.discBusy = false; render(); }
  }

  /* ------------------------------------------------------------------ *
   * Giao diện
   * ------------------------------------------------------------------ */
  const CARD_ABBR = { blue: 'c-blue', green: 'c-green', amber: 'c-amber', purple: 'c-purple', teal: 'c-teal', orange: 'c-orange', red: 'c-red', pink: 'c-pink', slate: 'c-slate' };
  function fmtCardVal(v, kind) {
    if (v === null || v === undefined || !Number.isFinite(v)) return '—';
    if (kind === 'pct') return round1(v).toLocaleString('vi-VN') + '%';
    return Math.round(v).toLocaleString('vi-VN');
  }
  function cardHtml([label, val, kind, color, sub, flag]) {
    const dl = flag === 'bl5dl' ? `<button class="tbc-dl" data-a="dlBl5" title="Tải danh sách mã đơn Backlog 5 ngày">↓</button>` : '';
    return `<div class="tbc-card" style="position:relative"><span>${esc(label)}</span><b class="${CARD_ABBR[color] || ''}">${fmtCardVal(val, kind)}</b>${sub ? `<small>${sub.map(([k, v2]) => `${esc(k)} <i>${fmt(v2)}</i>`).join(' · ')}</small>` : ''}${dl}</div>`;
  }
  function cardsRowHtml(title, barColor, list) {
    return `<div class="tbc-sec" style="--c:${barColor}"><span class="bar"></span>${esc(title)}</div><div class="tbc-cards" style="grid-template-columns:repeat(${list.length},1fr)">${list.map(cardHtml).join('')}</div>`;
  }
  function tableHtml(calc) {
    const hasTT = calc.hasTT;
    const rows = sortRows(calc.shippers, 'ok', -1);
    const row = (s, isTotal) => `<tr${isTotal ? ' class="total"' : ''}>
      <td class="l">${esc(s.name)}</td>
      <td>${fmt(s.scanned)}</td>
      <td>${fmt(s.newTotal)}</td>
      <td class="${rateClass(s.pod)}">${fmtCardVal(s.pod, 'pct')}</td>
      <td>${fmt(s.transit)}</td>
      <td class="${rateClass(s.rateTrip)}">${fmtCardVal(s.rateTrip, 'pct')}</td>
      <td>${fmt(s.pkTotal)}</td>
      <td class="${rateClass(s.pkRate)}">${fmtCardVal(s.pkRate, 'pct')}</td>
      <td>${fmt(s.pkPending)}</td>
      ${hasTT ? `<td>${fmt(s.ttTotal)}</td><td>${fmt(s.ttOk)}</td><td class="${rateClass(s.rateTT)}">${fmtCardVal(s.rateTT, 'pct')}</td>` : ''}
    </tr>`;
    return `<div class="tbc-tw"><table>
      <thead>
        <tr>
          <th class="l" rowspan="2">Tên NV</th>
          <th colspan="5" class="g-gh">Chỉ số giao hàng</th>
          <th colspan="3" class="g-pk">Chỉ số lấy hàng</th>
          ${hasTT ? '<th colspan="3" class="g-tt">Chỉ số TikTok Shop</th>' : ''}
        </tr>
        <tr>
          <th>Tổng đơn giao</th><th>Đơn mới</th><th>POD2400</th><th>Đang giao</th><th>Tỉ lệ giao TC</th>
          <th>Tổng đơn lấy</th><th>Tỉ lệ lấy TC</th><th>Đang lấy</th>
          ${hasTT ? '<th>Tổng đơn TikTok</th><th>GTC TikTok</th><th>%GTC TikTok</th>' : ''}
        </tr>
      </thead>
      <tbody>${rows.map(s => row(s, false)).join('')}${row(calc.totals, true)}</tbody>
    </table></div>
    <div class="tbc-note">Đơn mới = đơn nhận trong ngày hôm nay mà shipper đã quét giao · Đang giao/Đang lấy = chưa có kết quả cuối tính đến lúc xem · "(Chưa quét giao)" = đơn chưa ai quét · "(Không khớp người quét)" = đơn đã quét nhưng không đọc được tên người quét. TỔNG = đúng bằng thẻ "Tổng đơn" phía trên (mọi đơn đều được gộp vào đúng 1 dòng).${(() => {
      const gap = (calc.cards.all.total || 0) - (calc.totals.scanned || 0);
      return gap !== 0 ? ` · Lệch ${fmt(Math.abs(gap))} đơn so với thẻ phía trên — báo lại giúp mình, đây là trường hợp chưa lường tới.` : '';
    })()}</div>`;
  }
  function dashboardHtml() {
    const calc = state.calc;
    if (!calc) return '';
    const c = cardList(calc, state.kpi);
    return cardsRowHtml('🏷️ Giao thường', 'var(--barGT)', c.main) + cardsRowHtml('🎵 TikTok Shop', '#db2777', c.tts) + tableHtml(calc);
  }
  // ---- Ngày cũ: thẻ/bảng rút gọn (Kanban không trả đủ trường như API hôm nay) ----
  function oldCardList(c, kpi) {
    const K = Object.assign({}, KPI_DEFAULT, kpi || {});
    const a = c.cards.all, t = c.cards.tt;
    const needMain = needForKpi(a.newOk, a.newTotal, K.main), needTT = needForKpi(t.newOk, t.newTotal, K.tt);
    return {
      main: [
        ['Tổng đơn', a.total, 'int', 'blue'],
        ['POD2400', a.pod, 'pct', 'amber', [['Giao thành công', a.newOk], ['Tổng đơn mới', a.newTotal]]],
        [`Cần đạt KPI ${fmt(K.main)}%`, needMain, 'int', 'teal', [['Đã giao', a.newOk], ['Tổng đơn mới', a.newTotal]]],
        ['GTC Total', a.gtcTotal, 'pct', 'orange']
      ],
      tts: c.hasTT ? [
        ['Tổng đơn TikTok', t.total, 'int', 'pink'],
        ['TikTok POD2400', t.pod, 'pct', 'amber', [['Giao thành công', t.newOk], ['Tổng đơn mới', t.newTotal]]],
        [`TikTok cần đạt KPI ${fmt(K.tt)}%`, needTT, 'int', 'teal', [['Đã giao', t.newOk], ['Tổng đơn mới', t.newTotal]]],
        ['GTC Total TikTok', t.gtcTotal, 'pct', 'orange']
      ] : []
    };
  }
  function oldTableHtml(calc) {
    const hasTT = calc.hasTT;
    const rows = sortRows(calc.shippers, 'ok', -1);
    const row = (s, isTotal) => `<tr${isTotal ? ' class="total"' : ''}>
      <td class="l">${esc(s.label ?? s.name)}</td>
      <td>${fmt(s.scanned)}</td>
      <td>${fmt(s.newTotal)}</td>
      <td class="${rateClass(s.pod)}">${fmtCardVal(s.pod, 'pct')}</td>
      <td class="${rateClass(s.rateTrip)}">${fmtCardVal(s.rateTrip, 'pct')}</td>
      <td>${fmt(s.pkTotal)}</td>
      <td class="${rateClass(s.pkRate)}">${fmtCardVal(s.pkRate, 'pct')}</td>
      <td>${fmt(s.pkPending)}</td>
      ${hasTT ? `<td>${fmt(s.ttTotal)}</td><td>${fmt(s.ttOk)}</td><td class="${rateClass(s.rateTT)}">${fmtCardVal(s.rateTT, 'pct')}</td>` : ''}
    </tr>`;
    return `<div class="tbc-tw"><table>
      <thead>
        <tr>
          <th class="l" rowspan="2">Tên NV</th>
          <th colspan="4" class="g-gh">Chỉ số giao hàng</th>
          <th colspan="3" class="g-pk">Chỉ số lấy hàng</th>
          ${hasTT ? '<th colspan="3" class="g-tt">Chỉ số TikTok Shop</th>' : ''}
        </tr>
        <tr>
          <th>Tổng đơn giao</th><th>Đơn mới</th><th>POD2400</th><th>Tỉ lệ giao TC</th>
          <th>Tổng đơn lấy</th><th>Tỉ lệ lấy TC</th><th>Đang lấy</th>
          ${hasTT ? '<th>Tổng đơn TikTok</th><th>GTC TikTok</th><th>%GTC TikTok</th>' : ''}
        </tr>
      </thead>
      <tbody>${rows.map(s => row(s, false)).join('')}${row(calc.totals, true)}</tbody>
    </table></div>
    <div class="tbc-note">Xem lại từ Kanban nên không có: Chưa quét giao / Đang giao / Tỉ lệ hoàn / Backlog (Kanban không trả trạng thái chi tiết hay bộ phận cho ngày đã qua).${(() => {
      const gap = (calc.cards.all.total || 0) - (calc.totals.scanned || 0);
      return gap > 0 ? ` · Lệch ${fmt(gap)} đơn so với thẻ phía trên, đã kiểm tra lại nguyên nhân khi gặp.` : '';
    })()}${state.oldNote ? ' · ' + esc(state.oldNote) : ''}</div>`;
  }
  function oldDashboardHtml() {
    const calc = state.oldCalc;
    if (!calc) return '';
    const c = oldCardList(calc, state.kpi);
    const banner = `<div class="tbc-sec" style="--c:#64748b;margin-top:0"><span class="bar"></span>📅 Đang xem ngày ${esc(fmtDMY(state.viewDate))} (dữ liệu Kanban) — <span data-a="today" style="cursor:pointer;text-decoration:underline">quay về hôm nay</span></div>`;
    return banner + cardsRowHtml('🏷️ Giao thường', 'var(--barGT)', c.main) + (c.tts.length ? cardsRowHtml('🎵 TikTok Shop', '#db2777', c.tts) : '') + oldTableHtml(calc);
  }
  function pickerHtml() {
    return `<div class="tbc-pick"><h3>Bưu cục bạn đang quản lý</h3>
      <p class="muted">Tài khoản này quản lý nhiều hơn 1 bưu cục. Chọn 1 bưu cục để theo dõi (có thể đổi lại bằng cách xóa dữ liệu trình duyệt rồi mở lại):</p>
      ${state.discList.map((o, i) => `<div class="tbc-pickrow" data-a="pick" data-i="${i}">${esc(o.name)}</div>`).join('')}
    </div>`;
  }
  function bodyHtml() {
    if (state.discBusy) return `<div class="tbc-empty">Đang dò bưu cục bạn quản lý…</div>`;
    if (state.discErr) return `<div class="tbc-empty">${esc(state.discErr)}<div style="margin-top:10px"><button class="tbc-btn" data-a="retryDisc">Thử lại</button></div></div>`;
    if (state.discList) return pickerHtml();
    if (state.viewDate) {
      if (state.oldBusy && !state.oldCalc) return `<div class="tbc-empty">Đang tải số liệu ngày ${esc(fmtDMY(state.viewDate))}…</div>`;
      if (state.oldErr) return `<div class="tbc-empty">${esc(state.oldErr)}<div style="margin-top:10px"><button class="tbc-btn" data-a="refresh">Thử lại</button> <button class="tbc-btn" data-a="today">Về hôm nay</button></div></div>`;
      return oldDashboardHtml();
    }
    if (state.busy && !state.calc) return `<div class="tbc-empty">Đang tải số liệu…</div>`;
    if (state.err) return `<div class="tbc-empty">${esc(state.err)}<div style="margin-top:10px"><button class="tbc-btn" data-a="refresh">Thử lại</button></div></div>`;
    return dashboardHtml();
  }
  function themePopHtml() {
    if (!state.showTheme) return '';
    return `<div class="tbc-pop">${THEMES.map(t => `<div class="tbc-swrow ${state.theme === t ? 'on' : ''}" data-a="themePick" data-t="${t}"><span class="sw" data-theme-sw="${t}"></span>${esc(THEME_LABEL[t])}</div>`).join('')}</div>`;
  }
  function settingsPopHtml() {
    if (!state.showSettings) return '';
    return `<div class="tbc-pop wide">
      <label>KPI POD2400 Giao thường (%) <input type="number" id="tbc-kpi-main" min="0" max="100" step="1" value="${esc(state.kpi.main)}"></label>
      <label>KPI POD2400 TikTok (%) <input type="number" id="tbc-kpi-tt" min="0" max="100" step="1" value="${esc(state.kpi.tt)}"></label>
      <button class="tbc-btn solid" data-a="kpiSave" style="margin-top:6px">Lưu</button>
    </div>`;
  }
  function datePickHtml() {
    if (!state.showDatePick) return '';
    const today = dateKey(new Date());
    const opts = [`<div class="tbc-swrow ${!state.viewDate ? 'on' : ''}" data-a="today">📍 Hôm nay (trực tiếp)</div>`];
    for (let i = 1; i <= KB_DAYS_BACK; i++) {
      const d = addDays(today, -i);
      opts.push(`<div class="tbc-swrow ${state.viewDate === d ? 'on' : ''}" data-a="daySel" data-date="${d}">${esc(fmtDMY(d))}</div>`);
    }
    return `<div class="tbc-pop">${opts.join('')}</div>`;
  }
  function headerHtml() {
    const now = new Date();
    const off = state.office ? state.office.name : '';
    const upd = state.viewDate
      ? (state.oldAt ? `Số liệu ngày ${fmtDMY(state.viewDate)} · tải lúc ${new Date(state.oldAt).toLocaleTimeString('vi-VN')}` : `Đang xem ngày ${fmtDMY(state.viewDate)}`)
      : (state.at ? `Cập nhật lúc ${new Date(state.at).toLocaleTimeString('vi-VN')} ngày ${now.toLocaleDateString('vi-VN')}` : '');
    return `<div class="tbc-top">
      <div class="tbc-off"><span class="pin">📍</span><div><div class="nm">${esc(off || 'Trưởng Bưu Cục')}</div><div class="sub">${esc(upd)}</div></div></div>
      <div class="tbc-acts">
        <button class="tbc-btn solid" data-a="refresh" ${state.busy || state.oldBusy ? 'disabled' : ''}>${(state.busy || state.oldBusy) ? '⟳ Đang tải…' : (state.viewDate ? '⟳ Tải lại ngày này' : '⟳ Tải số liệu mới nhất')}</button>
        <span style="position:relative"><button class="tbc-btn" data-a="datepick" title="Xem ngày khác">📅 ${state.viewDate ? esc(fmtDM(state.viewDate)) : 'Ngày khác'}</button>${datePickHtml()}</span>
        <button class="tbc-btn" data-a="xlsx">↓ Xuất Excel</button>
        <span style="position:relative"><button class="tbc-btn" data-a="theme" title="Đổi màu nền">🎨</button>${themePopHtml()}</span>
        <span style="position:relative"><button class="tbc-btn" data-a="settings" title="Cài đặt KPI">⚙</button>${settingsPopHtml()}</span>
        <button class="tbc-btn" data-a="close">Đóng ×</button>
      </div>
    </div>`;
  }
  function updateBannerHtml() {
    if (!state.newVersion || state.updateDismissed) return '';
    return `<div class="tbc-updbar">
      <span>🆕 Đã có bản mới <b>v${esc(state.newVersion)}</b> (đang dùng v${esc(SCRIPT_VERSION)}) — bấm Cập nhật để tự cài bản mới.</span>
      <span class="tbc-updacts">
        <button class="tbc-btn solid" data-a="doUpdate">⬆ Cập nhật</button>
        <button class="tbc-btn" data-a="dismissUpdate">Để sau</button>
      </span>
    </div>`;
  }
  function appHtml() {
    return `${updateBannerHtml()}${headerHtml()}<div class="tbc-body">${bodyHtml()}</div>${WM_HTML}`;
  }

  function $(s, root = document) { return root.querySelector(s); }
  function render() {
    const app = $('#tbc-app'); if (!app) return;
    app.style.display = state.open ? 'flex' : 'none';
    app.setAttribute('data-theme', state.theme);
    updateLaunchBtn();
    if (!state.open) return;
    app.innerHTML = appHtml();
  }

  function downloadText(name, text) {
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const u = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(u), 4000);
  }
  function exportXlsx() {
    const isOld = !!state.viewDate, calc = isOld ? state.oldCalc : state.calc;
    if (!calc || typeof XLSX === 'undefined') return;
    const hasTT = calc.hasTT;
    const head = ['Tên NV', 'Tổng đơn giao', 'Đơn mới', 'POD2400(%)'].concat(isOld ? [] : ['Đang giao'], ['Tỉ lệ giao TC(%)', 'Tổng đơn lấy', 'Tỉ lệ lấy TC(%)', 'Đang lấy']).concat(hasTT ? ['Tổng đơn TikTok', 'GTC TikTok', '%GTC TikTok'] : []);
    const rows = sortRows(calc.shippers, 'ok', -1).map(s => [s.name ?? s.label, s.scanned, s.newTotal, round1(s.pod ?? NaN)].concat(isOld ? [] : [s.transit], [round1(s.rateTrip ?? NaN), s.pkTotal, round1(s.pkRate ?? NaN), s.pkPending]).concat(hasTT ? [s.ttTotal, s.ttOk, round1(s.rateTT ?? NaN)] : []));
    const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shipper');
    XLSX.writeFile(wb, `TruongBuuCuc_${isOld ? state.viewDate : dateKey(new Date())}.xlsx`);
  }

  function onClick(e) {
    const el = e.target.closest('[data-a]');
    if (!el) {
      // Bấm ra ngoài popover thì tự đóng (bấm vào trong popover, ví dụ ô nhập KPI, thì không đóng)
      if ((state.showTheme || state.showSettings || state.showDatePick) && !e.target.closest('.tbc-pop')) { state.showTheme = false; state.showSettings = false; state.showDatePick = false; render(); }
      return;
    }
    const a = el.dataset.a;
    if (a === 'open') { state.open = true; render(); if (!state.office && !state.discBusy && !state.discList) initOffice(); }
    else if (a === 'close') { state.open = false; render(); }
    else if (a === 'refresh') { if (state.viewDate) loadOldDay(state.viewDate); else loadData(); }
    else if (a === 'retryDisc') { state.discErr = ''; initOffice(); }
    else if (a === 'pick') { const i = Number(el.dataset.i); const o = state.discList[i]; state.discList = null; saveOffice(o); loadData(); }
    else if (a === 'dlBl5') { downloadText(`Backlog5ngay_${dateKey(new Date())}.txt`, (state.bl5codes || []).join('\n') || 'Không có đơn backlog 5 ngày.'); }
    else if (a === 'xlsx') { exportXlsx(); }
    else if (a === 'datepick') { state.showDatePick = !state.showDatePick; state.showTheme = false; state.showSettings = false; render(); }
    else if (a === 'daySel') { state.showDatePick = false; loadOldDay(el.dataset.date); }
    else if (a === 'today') { state.showDatePick = false; state.viewDate = null; state.oldErr = ''; render(); }
    else if (a === 'theme') { state.showTheme = !state.showTheme; state.showSettings = false; render(); }
    else if (a === 'settings') { state.showSettings = !state.showSettings; state.showTheme = false; render(); }
    else if (a === 'themePick') { state.theme = el.dataset.t; saveTheme(state.theme); state.showTheme = false; render(); }
    else if (a === 'doUpdate') { window.open(RAW_URL, '_blank'); }
    else if (a === 'dismissUpdate') { state.updateDismissed = true; render(); }
    else if (a === 'kpiSave') {
      const m = $('#tbc-kpi-main'), tt = $('#tbc-kpi-tt');
      const km = m ? parseNum(m.value) : NaN, ktt = tt ? parseNum(tt.value) : NaN;
      state.kpi = { main: Number.isFinite(km) ? Math.max(0, Math.min(100, km)) : state.kpi.main, tt: Number.isFinite(ktt) ? Math.max(0, Math.min(100, ktt)) : state.kpi.tt };
      saveKpi(state.kpi); state.showSettings = false; render();
    }
  }

  /* ------------------------------------------------------------------ *
   * Gắn vào trang
   * ------------------------------------------------------------------ */
  const CSS = `
#tbc-launch{position:fixed;left:50%;transform:translateX(-50%);bottom:10px;z-index:999998;background:#4c3a8f;color:#fff;border:0;border-radius:10px;padding:10px 15px;font-weight:700;cursor:pointer;box-shadow:0 4px 18px #0005;font:17px/1 "Segoe UI",Roboto,Arial,sans-serif}
#tbc-app{position:fixed;inset:2vh 2vw;z-index:1000000;display:none;flex-direction:column;border-radius:16px;overflow:hidden;box-shadow:0 20px 80px #0007;font:15px/1.45 "Segoe UI",Roboto,Arial,sans-serif;color:#1f2937;
  --hd1:#6c4fc2;--hd2:#8a5cdb;--pageBg:#f3f4f8;--barGT:#6d4fc4;--gShip:#1e3a8a;--gRate:#4c3a8f;--gGh:#3b82f6;--gPk:#0d9488;--gTt:#f59e0b;--primaryTxt:#6c4fc2}
#tbc-app[data-theme="xanhThan"]{--hd1:#1e293b;--hd2:#334155;--pageBg:#eef2f6;--barGT:#334155;--gShip:#0f172a;--gRate:#334155;--gGh:#1d4ed8;--gPk:#0e7490;--gTt:#b45309;--primaryTxt:#334155}
#tbc-app[data-theme="pastel"]{--hd1:#8b7cd8;--hd2:#a996e8;--pageBg:#f6f4fc;--barGT:#8b7cd8;--gShip:#3730a3;--gRate:#7c6fd1;--gGh:#2563eb;--gPk:#0d9488;--gTt:#f59e0b;--primaryTxt:#6d5bd0}
#tbc-app[data-theme="best"]{--hd1:#c81e1e;--hd2:#e2231a;--pageBg:#fff6f6;--barGT:#e2231a;--gShip:#7f1d1d;--gRate:#991b1b;--gGh:#e2231a;--gPk:#b91c1c;--gTt:#111827;--primaryTxt:#c81e1e}
#tbc-app{background:var(--pageBg)}
#tbc-app *{box-sizing:border-box}
#tbc-app .tbc-top{background:linear-gradient(90deg,var(--hd1),var(--hd2));padding:10px 18px;display:flex;align-items:center;gap:14px;color:#fff;flex-wrap:wrap}
#tbc-app .tbc-off{display:flex;align-items:center;gap:8px}
#tbc-app .tbc-off .nm{font-weight:800;font-size:20px}
#tbc-app .tbc-off .sub{font-size:13px;opacity:.85}
#tbc-app .tbc-acts{display:flex;gap:8px;margin-left:auto;flex-wrap:wrap}
#tbc-app .tbc-btn{border:0;border-radius:8px;padding:9px 14px;font-weight:700;cursor:pointer;background:#ffffff33;color:#fff;font-size:15px}
#tbc-app .tbc-btn.solid{background:#fff;color:var(--primaryTxt)}
#tbc-app .tbc-btn:disabled{opacity:.7;cursor:default}
#tbc-app .tbc-body{flex:1;min-height:0;overflow:hidden;padding:12px 16px;display:flex;flex-direction:column}
#tbc-app .tbc-sec{flex:none;display:flex;align-items:center;gap:8px;font-weight:700;color:#374151;margin:10px 0 6px;font-size:15px}
#tbc-app .tbc-sec .bar{width:4px;height:16px;background:var(--c,var(--barGT));border-radius:2px;display:inline-block}
#tbc-app .tbc-cards{flex:none;display:grid;grid-template-columns:repeat(7,1fr);gap:8px}
#tbc-app .tbc-card{background:#fff;border-radius:10px;padding:9px 11px;box-shadow:0 1px 5px #0001;min-width:0}
#tbc-app .tbc-card span{display:block;font-size:13px;color:#6b7280;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#tbc-app .tbc-card b{display:block;font-size:25px;margin-top:2px;line-height:1.15}
#tbc-app .tbc-card small{display:block;font-size:12px;color:#6b7280;margin-top:2px}
#tbc-app .tbc-card small i{font-style:normal;font-weight:700;color:#374151}
#tbc-app .c-blue{color:#2563eb}#tbc-app .c-green{color:#16a34a}#tbc-app .c-amber{color:#d97706}#tbc-app .c-purple{color:#7c3aed}#tbc-app .c-teal{color:#0d9488}#tbc-app .c-orange{color:#ea580c}#tbc-app .c-red{color:#dc2626}#tbc-app .c-pink{color:#db2777}#tbc-app .c-slate{color:#475569}
#tbc-app .tbc-dl{position:absolute;top:6px;right:6px;border:0;background:#fee2e2;color:#dc2626;border-radius:7px;width:22px;height:22px;font-size:13px;cursor:pointer}
#tbc-app .tbc-tw{flex:1;min-height:0;background:#fff;border-radius:10px;overflow:auto;margin-top:10px;box-shadow:0 1px 5px #0001}
#tbc-app .tbc-tw table{width:100%;border-collapse:collapse;font-size:14px}
#tbc-app .tbc-tw th,#tbc-app .tbc-tw td{padding:7px 10px;text-align:center;border-bottom:1px solid #e5e7eb;white-space:nowrap}
#tbc-app .tbc-tw .l{text-align:left}
#tbc-app .tbc-tw thead th{position:sticky;top:0;color:#fff;font-weight:600;background:var(--gShip);z-index:1}
#tbc-app .tbc-tw thead tr:nth-child(2) th{top:34px}
#tbc-app .tbc-tw .g-gh{background:var(--gGh)}#tbc-app .tbc-tw .g-pk{background:var(--gPk)}#tbc-app .tbc-tw .g-tt{background:var(--gTt)}
#tbc-app .tbc-tw tbody tr:nth-child(even){background:#f9fafb}
#tbc-app .tbc-tw tr.total td{font-weight:700;background:#eef2f7;position:sticky;bottom:0}
#tbc-app .r-ok{color:#16a34a;font-weight:700}#tbc-app .r-mid{color:#d97706;font-weight:700}#tbc-app .r-bad{color:#dc2626;font-weight:700}
#tbc-app .tbc-note{flex:none;margin-top:6px;font-size:13px;color:#6b7280}
#tbc-app .tbc-empty{background:#fff;border-radius:12px;padding:28px;text-align:center;color:#4b5563}
#tbc-app .tbc-pick{background:#fff;border-radius:12px;padding:22px;max-width:560px;margin:20px auto}
#tbc-app .tbc-pick h3{margin:0 0 6px}
#tbc-app .tbc-pickrow{padding:11px 14px;border:1px solid #e5e7eb;border-radius:8px;margin-top:8px;cursor:pointer;font-weight:600}
#tbc-app .tbc-pickrow:hover{background:#ede9fe;border-color:#8a5cdb}
#tbc-app .muted{color:#6b7280}
#tbc-app .tbc-wm{position:absolute;right:14px;bottom:10px;display:flex;align-items:center;gap:6px;background:#ffffffd9;border:1px solid #0000001f;border-radius:999px;padding:6px 14px 6px 11px;box-shadow:0 2px 8px #0002;pointer-events:none;user-select:none;z-index:999999}
#tbc-app .tbc-wm .wm-ic{font-size:14px;line-height:1}
#tbc-app .tbc-wm .wm-txt{display:flex;flex-direction:column;line-height:1.2}
#tbc-app .tbc-wm .wm-name{font-size:14px;font-weight:800;color:var(--primaryTxt);letter-spacing:.2px}
#tbc-app .tbc-wm .wm-mail{font-size:10px;font-weight:600;color:var(--primaryTxt);opacity:.75}
#tbc-app .tbc-pop{position:absolute;right:0;top:36px;z-index:50;background:#fff;color:#1f2937;border-radius:10px;box-shadow:0 8px 24px #0004;padding:9px;min-width:230px}
#tbc-app .tbc-pop.wide{min-width:280px;display:flex;flex-direction:column;gap:8px}
#tbc-app .tbc-pop label{font-size:14px;color:#4b5563;display:flex;flex-direction:column;gap:3px}
#tbc-app .tbc-pop input{padding:6px 8px;border:1px solid #cbd5e1;border-radius:6px;font-size:14px}
#tbc-app .tbc-swrow{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:7px;cursor:pointer;font-size:14px;white-space:nowrap}
#tbc-app .tbc-swrow:hover{background:#f3f4f6}
#tbc-app .tbc-swrow.on{background:#ede9fe;font-weight:700}
#tbc-app .sw{width:14px;height:14px;border-radius:50%;display:inline-block;flex:none;box-shadow:inset 0 0 0 1px #0002}
#tbc-app .sw[data-theme-sw="default"]{background:linear-gradient(90deg,#6c4fc2,#8a5cdb)}
#tbc-app .sw[data-theme-sw="xanhThan"]{background:linear-gradient(90deg,#1e293b,#334155)}
#tbc-app .sw[data-theme-sw="pastel"]{background:linear-gradient(90deg,#8b7cd8,#a996e8)}
#tbc-app .sw[data-theme-sw="best"]{background:linear-gradient(90deg,#c81e1e,#e2231a)}
@media(max-width:1100px){#tbc-app .tbc-cards{grid-template-columns:repeat(4,1fr)}}
#tbc-app .tbc-updbar{display:flex;align-items:center;justify-content:space-between;gap:12px;background:linear-gradient(90deg,#0f766e,#14b8a6);color:#fff;padding:9px 16px;font-size:13px;font-weight:600;flex:0 0 auto}
#tbc-app .tbc-updbar b{font-weight:800}
#tbc-app .tbc-updacts{display:flex;gap:8px;flex:0 0 auto}
#tbc-app .tbc-updacts .tbc-btn{background:#ffffffe6;color:#0f766e;border:none}
#tbc-app .tbc-updacts .tbc-btn.solid{background:#fff;font-weight:800}
#tbc-launch.tbc-hasupd{box-shadow:0 0 0 3px #14b8a6a0,0 4px 14px #0004}
`;

  // Watermark không thể xóa, dạng huy hiệu (PA B đã chọn): icon + tên + email, nổi trên nền trắng mờ
  const WM_HTML = `<div class="tbc-wm"><span class="wm-ic">✍️</span><span class="wm-txt"><span class="wm-name">Design By Erik Nguyễn</span><span class="wm-mail">dungvietam@gmail.com</span></span></div>`;
  function ensureWatermark() {
    const app = $('#tbc-app'); if (!app) return;
    let wm = app.querySelector('.tbc-wm');
    if (!wm) { app.insertAdjacentHTML('beforeend', WM_HTML); }
  }

  /* ------------------------------------------------------------------ *
   * Kiểm tra bản mới trên GitHub, không cần mở Tampermonkey dashboard.
   * Tìm dòng "// @version   X.Y.Z" trong file raw, so với SCRIPT_VERSION.
   * Nếu mới hơn: hiện banner trong app + chấm sáng trên nút mở app.
   * ------------------------------------------------------------------ */
  async function checkForUpdate() {
    try {
      const res = await fetch(RAW_URL + '?t=' + Date.now(), { cache: 'no-store' });
      if (!res.ok) return;
      const text = await res.text();
      const m = text.match(/@version\s+([\d.]+)/);
      if (!m) return;
      const latest = m[1];
      if (verNewer(latest, SCRIPT_VERSION)) {
        state.newVersion = latest;
        render(); updateLaunchBtn();
      }
    } catch (e) { /* mạng lỗi hoặc repo chưa có file — bỏ qua, thử lại lần sau */ }
  }
  function updateLaunchBtn() {
    const btn = $('#tbc-launch'); if (!btn) return;
    const has = !!(state.newVersion && !state.updateDismissed);
    btn.classList.toggle('tbc-hasupd', has);
    btn.textContent = has ? `📊 Trưởng Bưu Cục 🆕` : '📊 Trưởng Bưu Cục';
  }

  function mount() {
    if ($('#tbc-launch')) return;
    state.theme = loadTheme(); state.kpi = loadKpi();
    const style = document.createElement('style'); style.textContent = CSS; document.head.appendChild(style);
    const btn = document.createElement('button'); btn.id = 'tbc-launch'; btn.setAttribute('data-a', 'open'); btn.textContent = '📊 Trưởng Bưu Cục';
    const app = document.createElement('div'); app.id = 'tbc-app';
    document.body.appendChild(btn); document.body.appendChild(app);
    document.body.addEventListener('click', onClick);
    ensureHdrHook(); setInterval(ensureHdrHook, 1000);
    // Watermark không thể xóa: canh lại ngay nếu bị gỡ khỏi DOM, và kiểm tra định kỳ
    const mo = new MutationObserver(() => { if (state.open) ensureWatermark(); });
    mo.observe(app, { childList: true, subtree: false });
    setInterval(() => { if (state.open) ensureWatermark(); }, 1500);
    // Kiểm tra bản mới: ngay khi mở trang, rồi lặp lại theo chu kỳ UPDATE_CHECK_MS
    checkForUpdate();
    setInterval(checkForUpdate, UPDATE_CHECK_MS);
  }
  function safeMount() { try { mount(); } catch (e) { console.error('[Trưởng Bưu Cục] lỗi khởi động:', e); } }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', safeMount, { once: true });
  else safeMount();
})();
