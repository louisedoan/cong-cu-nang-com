/* Giao diện công cụ Quản lý Kênh Viettel. Mọi phép tính tiền do máy chủ (app.py) làm; giao diện chỉ hiển thị. */
'use strict';
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const ST_CLS = { da_len_don: 'info', giao_mot_phan: 'warn', da_giao_du: 'ok', huy: 'mute' };

let S = null, M = {};
let draft = null;
let lastQuote = null;
const filters = {};

// ---------- tiện ích ----------
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = n => Math.round(Number(n) || 0).toLocaleString('vi-VN');
const fd = d => d ? String(d).slice(0, 10).split('-').reverse().join('/') : '';
const pct = x => x == null || x === '' ? '' : `${+(x * 100).toFixed(2)}%`;
const monthLabel = m => m ? `${m.slice(5, 7)}/${m.slice(0, 4)}` : '';
const sumBy = (arr, k) => arr.reduce((a, x) => a + (Number(x[k]) || 0), 0);

async function api(path, body) {
  const res = await fetch(path, { method: body === undefined ? 'GET' : 'POST', headers: { 'Content-Type': 'application/json' }, body: body === undefined ? undefined : JSON.stringify(body) });
  const data = await res.json();
  if (res.status === 403 && data.license) showLicenseScreen(data.license);  // license hết hạn giữa chừng
  if (!res.ok) throw new Error(data.error || 'Có lỗi xảy ra');
  return data;
}
function toast(msg, err = false, html = false) {
  const t = document.createElement('div');
  t.className = 'toast' + (err ? ' err' : '');
  t[html ? 'innerHTML' : 'textContent'] = msg;
  $('#toast-root').append(t);
  setTimeout(() => t.remove(), err ? 7000 : 9000);
  return t;
}
async function run(fn, okMsg) {
  try { const r = await fn(); if (okMsg) toast(okMsg); return r; }
  catch (e) { toast(e.message, true); throw e; }
}
function openFileLink(path, label) { return `<a data-open="${esc(path)}">${esc(label)}</a>`; }
document.addEventListener('click', e => {
  const a = e.target.closest('[data-open]');
  if (a) { e.preventDefault(); api('/api/open', { path: a.dataset.open }).catch(err => toast(err.message, true)); }
});

// ---------- hộp thoại ----------
function modal(title, body, foot = '', wide = false) {
  const root = $('#modal-root');
  root.innerHTML = `<div class="modal-bg"><div class="modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
    <div class="modal-h"><h3>${esc(title)}</h3><button class="x" data-close aria-label="Đóng">×</button></div>
    <div class="modal-b">${body}</div>${foot ? `<div class="modal-f">${foot}</div>` : ''}</div></div>`;
  root.querySelector('[data-close]').onclick = closeModal;
  root.querySelector('.modal-bg').onclick = e => { if (e.target.classList.contains('modal-bg')) closeModal(); };
  $$('[data-close2]', root).forEach(b => b.onclick = closeModal);
  setTimeout(() => root.querySelector('input, select, textarea')?.focus(), 30);
  return root;
}
function closeModal() { $('#modal-root').innerHTML = ''; }
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
function confirmDialog(title, text, onOk) {
  modal(title, `<p>${esc(text)}</p>`, `<button class="btn ghost" data-close2>Không</button><button class="btn danger" id="c-yes">Đồng ý</button>`);
  $('#c-yes').onclick = async () => { closeModal(); await onOk(); };
}

// ---------- dữ liệu ----------
async function load() {
  if (!(await checkLicense())) return false;
  S = await api('/api/state');
  M.prod = Object.fromEntries(S.products.map(p => [p.id, p]));
  M.combos = S.products.filter(p => p.la_combo && p.dang_ban);
  M.sellable = S.products.filter(p => !p.la_combo && p.dang_ban);
  $('#today').textContent = 'Hôm nay ' + fd(S.today);
  $('#side-foot').innerHTML = `File dữ liệu: ${openFileLink(S.data_file, 'mở bằng Excel')}<br>Hoa hồng: ${esc(S.settings.ten_hh_cp)} ${pct(S.settings.cp_pct)} · Tỉnh ${pct(S.settings.tinh_pct)}`;
  return true;
}

// ---------- license (offline 48 giờ) ----------
let LIC = null;
async function checkLicense() {
  LIC = await (await fetch('/api/license', { cache: 'no-store' })).json();
  if (!LIC.valid) { showLicenseScreen(LIC); return false; }
  $('#license-root').innerHTML = '';
  const low = LIC.remaining_seconds < 6 * 3600;
  $('#side-lic').innerHTML = `<button class="lic-badge ${low ? 'low' : ''}" id="lic-open" title="Xem thông tin license">
    <span>License còn <b>${esc(LIC.remaining_text)}</b></span><span>hết hạn ${esc(LIC.expires_text)}</span></button>`;
  $('#lic-open').onclick = licenseInfo;
  return true;
}
function licenseForm(lic) {
  return `<div class="field"><label>Mã thiết bị của máy này — gửi cho quản trị viên để xin mã</label>
      <div class="devid"><code>${esc(lic.device_id || '—')}</code>${lic.device_id ? '<button class="btn ghost sm" id="lic-copy-dev">Sao chép</button>' : ''}</div></div>
    <div class="field"><label for="lic-code">Mã kích hoạt (dán nguyên mã, bắt đầu bằng NC)</label>
      <textarea id="lic-code" class="lic-code" rows="4" spellcheck="false" placeholder="NC…-XXXX-XXXX-…" ${lic.can_activate ? '' : 'disabled'}></textarea></div>
    <div id="lic-err"></div>`;
}
function bindLicenseForm(onOk) {
  $('#lic-copy-dev')?.addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(LIC.device_id); toast('Đã sao chép Mã thiết bị.'); } catch { toast('Không sao chép được — hãy bôi đen rồi Ctrl+C.', true); }
  });
  $('#lic-activate').onclick = async () => {
    const btn = $('#lic-activate'); btn.disabled = true;
    try {
      const res = await fetch('/api/license/activate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ code: $('#lic-code').value }) });
      const data = await res.json();
      if (!res.ok) { $('#lic-err').innerHTML = `<div class="alert bad">${esc(data.error || 'Không kích hoạt được.')}</div>`; return; }
      toast(data.license.message);
      await onOk();
    } catch (e) { $('#lic-err').innerHTML = `<div class="alert bad">Không kết nối được công cụ: ${esc(e.message)}</div>`; }
    finally { btn.disabled = false; }
  };
}
const policyBlock = lic => `<details class="lic-policy"><summary>Chính sách license offline</summary><pre>${esc(lic.policy)}</pre></details>`;
function showLicenseScreen(lic) {
  LIC = lic;
  closeModal();
  $('#side-lic').innerHTML = '';
  if ($('#license-root').dataset.state === lic.state && $('#lic-code')) return;  // đang hiện, giữ mã đang gõ
  $('#license-root').dataset.state = lic.state;
  $('#license-root').innerHTML = `<div class="lic-bg"><div class="lic-card" role="dialog" aria-modal="true" aria-labelledby="lic-title">
    <div class="brand-mark">NC</div>
    <h2 id="lic-title">${lic.state === 'expired' ? 'License đã hết hạn' : 'Kích hoạt license'}</h2>
    <div class="alert ${lic.state === 'expired' ? 'warn' : 'bad'}">${esc(lic.message)}</div>
    ${licenseForm(lic)}
    <div class="actions"><button class="btn" id="lic-activate" ${lic.can_activate ? '' : 'disabled'}>Kích hoạt</button></div>
    <div class="small muted">Dữ liệu của bạn được giữ nguyên, không bị xóa khi license hết hạn: ${esc(lic.data_file)}</div>
    ${policyBlock(lic)}</div></div>`;
  bindLicenseForm(async () => { $('#license-root').innerHTML = ''; delete $('#license-root').dataset.state; await refresh(); });
  setTimeout(() => $('#lic-code')?.focus(), 30);
}
function licenseInfo() {
  const l = LIC;
  modal('Thông tin license', `
    <div class="fields small"><div><span class="muted">Trạng thái</span><br><b style="color:var(--ok)">Đang hoạt động</b></div>
      <div><span class="muted">Còn lại</span><br><b>${esc(l.remaining_text)}</b></div>
      <div><span class="muted">Phát hành</span><br><b>${esc(l.issued_text)}</b></div><div><span class="muted">Hết hạn</span><br><b>${esc(l.expires_text)}</b></div>
      <div><span class="muted">License ID</span><br><code>${esc(l.license_id)}</code></div></div>
    <div class="alert warn small">48 giờ tính từ lúc quản trị viên phát hành mã. Khi hết hạn, công cụ khóa các chức năng cho tới khi nhập mã mới; dữ liệu giữ nguyên.</div>
    <b>Nhập mã mới</b> <span class="small muted">(khi quản trị viên đã cấp mã cho chu kỳ tiếp theo)</span>
    ${licenseForm(l)}${policyBlock(l)}`,
    `<button class="btn ghost" data-close2>Đóng</button><button class="btn" id="lic-activate">Kích hoạt mã mới</button>`);
  bindLicenseForm(async () => { closeModal(); await refresh(); });
}
setInterval(() => checkLicense().catch(() => {}), 60_000);  // cập nhật thời gian còn lại, khóa ngay khi hết hạn
function ordersOf(kenh) { return S.orders.filter(o => o.kenh === kenh); }
function newDraft(kenh) { return { kenh, id: null, khach: '', tinh: '', ngay_dat: S.today, ck_pct: 0, cong_no: 0, phi_vc: 0, chi_phi_khac: 0, pay_status: 'chua_tt', ghi_chu: '', lines: {}, combos: [], httm: [] }; }

// ---------- điều hướng ----------
const SUB = {
  'tong-quan': ['Tổng quan', 'Doanh thu & lợi nhuận theo tháng, theo tỉnh', renderDashboard],
  'len-don': ['Lên đơn hàng', 'Nhập đơn, tự tính tiền; combo chọn 1 trong 3; lưu & xuất đơn + phiếu xuất kho', renderOrderForm],
  'don-hang': ['Đơn hàng', 'Lịch sử đơn, lọc theo tỉnh/tháng, chia đợt giao, xuất file', renderOrders],
};
function route() {
  const hash = (location.hash || '#buu_cuc/tong-quan').slice(1);
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.page === hash));
  if (hash === 'danh-muc') {
    $('#page-title').textContent = 'Danh mục & thiết lập';
    $('#page-sub').textContent = 'Sản phẩm, giá, combo, tải trọng, tỉnh, hoa hồng — sửa trực tiếp';
    return renderCatalog($('#page'));
  }
  const [kenh, sub] = hash.split('/');
  const def = SUB[sub] || SUB['tong-quan'];
  $('#page-title').textContent = `${S.channels[kenh] || ''} · ${def[0]}`;
  $('#page-sub').textContent = def[1];
  def[2]($('#page'), kenh);
}
window.addEventListener('hashchange', route);
async function refresh() { if (await load()) route(); }

// =====================================================================
// TỔNG QUAN
// =====================================================================
function renderDashboard(el, kenh) {
  const orders = ordersOf(kenh).filter(o => o.trang_thai !== 'huy');
  const month = S.today.slice(0, 7);
  const mset = orders.filter(o => (o.ngay_dat || '').slice(0, 7) === month);
  const sum = (arr, k) => arr.reduce((a, o) => a + (o.profit[k] || 0), 0);
  const rev = sum(mset, 'doanh_thu'), prof = sum(mset, 'loi_nhuan'), gv = sum(mset, 'gia_von');
  const sl = sum(mset, 'san_luong');
  const hhCp = sum(mset, 'tien_cp'), hhTinh = sum(mset, 'tien_tinh');
  const owing = orders.filter(o => o.pay_status !== 'da_tt').reduce((a, o) => a + (o.cong_no_tien || 0), 0);
  el.innerHTML = `
  <div class="kpis">
    <div class="kpi main"><div class="l">Doanh thu ${monthLabel(month)}</div><div class="v">${fmt(rev)}</div><div class="s">${mset.length} đơn</div></div>
    <div class="kpi ${prof < 0 ? 'bad' : 'ok'}"><div class="l">Lợi nhuận ${monthLabel(month)}</div><div class="v">${fmt(prof)}</div><div class="s">biên ${rev ? (prof / rev * 100).toFixed(1) : 0}%</div></div>
    <div class="kpi"><div class="l">Giá vốn ${monthLabel(month)}</div><div class="v">${fmt(gv)}</div></div>
    <div class="kpi"><div class="l">Sản lượng ${monthLabel(month)}</div><div class="v">${fmt(sl)}</div><div class="s">chai/hũ/phần</div></div>
    <div class="kpi"><div class="l">HH ${esc(S.settings.ten_hh_cp)} ${monthLabel(month)}</div><div class="v">${fmt(hhCp)}</div></div>
    <div class="kpi"><div class="l">HH Bưu cục ${monthLabel(month)}</div><div class="v">${fmt(hhTinh)}</div></div>
    <div class="kpi warn"><div class="l">Chưa thu (còn công nợ)</div><div class="v">${fmt(owing)}</div></div>
  </div>
  <div class="card"><div class="card-h"><h2>Doanh thu & lợi nhuận 6 tháng gần nhất</h2></div><div class="card-b">${revenueChart(orders)}</div></div>
  <div class="grid g2">
    <div class="card"><div class="card-h"><h2>Theo tỉnh (${monthLabel(month)})</h2></div><div class="tbl-wrap">${provinceTable(kenh, month)}</div></div>
    <div class="card"><div class="card-h"><h2>Theo tháng</h2></div><div class="tbl-wrap">${monthTable(orders)}</div></div>
  </div>
  <div class="card"><div class="card-b"><button class="btn ghost" id="btn-report">Xuất báo cáo Excel (${S.channels[kenh]})</button></div></div>`;
  $('#btn-report').onclick = async () => {
    const t = toast('Đang dựng báo cáo Excel…');
    try { const r = await api('/api/report/export', { kenh }); t.remove(); toast(`Đã xuất báo cáo. ${openFileLink(r.file, 'Mở file')}`, false, true); }
    catch (e) { t.remove(); toast(e.message, true); }
  };
  loadReport(kenh, month);
}
function revenueChart(orders) {
  const months = [];
  const d = new Date(S.today + 'T00:00:00');
  for (let i = 5; i >= 0; i--) { const x = new Date(d.getFullYear(), d.getMonth() - i, 1); months.push(`${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}`); }
  const data = months.map(m => {
    const set = orders.filter(o => (o.ngay_dat || '').slice(0, 7) === m);
    return { m, rev: set.reduce((a, o) => a + o.profit.doanh_thu, 0), prof: set.reduce((a, o) => a + o.profit.loi_nhuan, 0) };
  });
  const max = Math.max(1, ...data.map(x => x.rev));
  return `<div class="bars">${data.map(x => `<div class="bar" title="${monthLabel(x.m)}: DT ${fmt(x.rev)} · LN ${fmt(x.prof)}">
    <div class="val">${x.rev ? (x.rev / 1e6).toFixed(0) + ' tr' : '–'}</div>
    <div class="stack" style="height:${Math.max(2, x.rev / max * 150)}px"><span style="height:100%;background:#2F5597"></span></div>
    <div class="lab">${monthLabel(x.m)}</div></div>`).join('')}</div>
  <div class="legend"><span><i style="background:#2F5597"></i>Doanh thu (cột). Di chuột xem lợi nhuận.</span></div>`;
}
function provinceTable(kenh, month) {
  const rows = (M.report && M.report.kenh === kenh && M.report.month === month) ? M.report.province : null;
  return `<table id="prov-tbl"><thead><tr><th>Tỉnh</th><th class="n">Số đơn</th><th class="n">Doanh thu</th><th class="n">Giá vốn</th><th class="n">HH ${esc(S.settings.ten_hh_cp)}</th><th class="n">HH Bưu cục</th><th class="n">Lợi nhuận</th><th class="n">Biên</th></tr></thead>
    <tbody data-prov="${kenh}|${month}">${rows ? provRows(rows) : '<tr><td colspan="8" class="empty">Đang tải…</td></tr>'}</tbody></table>`;
}
function provRows(rows) {
  if (!rows.length) return '<tr><td colspan="8" class="empty">Chưa có đơn.</td></tr>';
  return rows.map(r => `<tr><td>${esc(r.tinh)}</td><td class="n">${r.so_don}</td><td class="n">${fmt(r.doanh_thu)}</td><td class="n">${fmt(r.gia_von)}</td><td class="n">${fmt(r.tien_cp)}</td><td class="n">${fmt(r.tien_tinh)}</td>
    <td class="n" style="color:${r.loi_nhuan < 0 ? 'var(--bad)' : 'inherit'}">${fmt(r.loi_nhuan)}</td><td class="n">${(r.bien_ln * 100).toFixed(1)}%</td></tr>`).join('');
}
function monthTable(orders) {
  const by = {};
  for (const o of orders) { const m = (o.ngay_dat || '').slice(0, 7); if (!m) continue; (by[m] ||= { rev: 0, prof: 0, n: 0 }); by[m].rev += o.profit.doanh_thu; by[m].prof += o.profit.loi_nhuan; by[m].n++; }
  const months = Object.keys(by).sort().reverse();
  if (!months.length) return '<div class="empty">Chưa có đơn.</div>';
  return `<table><thead><tr><th>Tháng</th><th class="n">Số đơn</th><th class="n">Doanh thu</th><th class="n">Lợi nhuận</th><th class="n">Biên</th></tr></thead><tbody>
    ${months.map(m => `<tr><td>${monthLabel(m)}</td><td class="n">${by[m].n}</td><td class="n">${fmt(by[m].rev)}</td>
      <td class="n" style="color:${by[m].prof < 0 ? 'var(--bad)' : 'inherit'}">${fmt(by[m].prof)}</td><td class="n">${by[m].rev ? (by[m].prof / by[m].rev * 100).toFixed(1) : 0}%</td></tr>`).join('')}</tbody></table>`;
}

// =====================================================================
// LÊN ĐƠN
// =====================================================================
function renderOrderForm(el, kenh) {
  if (!draft || draft.kenh !== kenh) draft = newDraft(kenh);
  const editing = draft.id ? ordersOf(kenh).find(o => o.id === draft.id) : null;
  el.innerHTML = `
  <div class="grid">
    <div class="card"><div class="card-h"><h2>${editing ? 'Sửa đơn ' + esc(editing.ma) : 'Thông tin đơn'}</h2>${editing ? '<button class="btn ghost sm" id="cancel-edit">Lên đơn mới</button>' : ''}</div>
      <div class="card-b"><div class="fields">
        <div class="field" style="grid-column:span 2"><label>Khách hàng / Kho</label><input id="f-khach" list="khach-list" value="${esc(draft.khach)}" placeholder="VD: Chi nhánh Viettel Post Cà Mau"><datalist id="khach-list">${[...new Set(ordersOf(kenh).map(o => o.khach).filter(Boolean))].map(k => `<option value="${esc(k)}">`).join('')}</datalist></div>
        <div class="field"><label>Tỉnh/Thành</label><select id="f-tinh"><option value="">— Chọn tỉnh —</option>${S.provinces.map(t => `<option ${t === draft.tinh ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select></div>
        <div class="field"><label>Ngày đặt</label><input type="date" id="f-date" value="${draft.ngay_dat}"></div>
        <div class="field"><label>Chiết khấu (%)</label><input type="number" step="0.5" id="f-ck" value="${+(draft.ck_pct * 100).toFixed(2)}"></div>
        <div class="field"><label>Công nợ (tiền chưa thu)</label><input type="number" id="f-cn" value="${draft.cong_no || ''}"></div>
        <div class="field"><label>Phí vận chuyển</label><input type="number" id="f-vc" value="${draft.phi_vc || ''}"></div>
        <div class="field"><label>Chi phí khác</label><input type="number" id="f-cpk" value="${draft.chi_phi_khac || ''}"></div>
        ${kenh === 'tm' ? `<div class="field"><label>Trạng thái thanh toán</label><select id="f-pay">${Object.entries(S.pay_status).map(([k, v]) => `<option value="${k}" ${k === draft.pay_status ? 'selected' : ''}>${v}</option>`).join('')}</select></div>` : ''}
        <div class="field" style="grid-column:span 2"><label>Ghi chú</label><input id="f-note" value="${esc(draft.ghi_chu)}"></div>
      </div></div></div>
    <div class="card"><div class="card-h"><div><h2>Combo & Hộp quà</h2><div class="sub">Combo "chọn 1 trong 3" phải chọn món trước khi nhập số phần.</div></div></div>
      <div class="tbl-wrap"><table><thead><tr><th>Combo</th><th>Chọn món (nếu có)</th><th class="n">Giá</th><th class="n">Số phần/hộp</th></tr></thead><tbody id="combo-body"></tbody></table></div></div>
    <div class="card"><div class="card-h"><div><h2>Mặt hàng lẻ</h2></div><input id="f-search" placeholder="Tìm mặt hàng…" style="width:220px"></div>
      <div class="tbl-wrap"><table id="prod-table"><thead><tr><th>Tên hàng</th><th>ĐVT</th><th class="n">QC</th><th class="n">Giá chai</th><th class="n">SL chai</th><th class="n">Thành tiền</th></tr></thead><tbody id="prod-body"></tbody></table></div></div>
    <div class="card"><div class="card-h"><div><h2>Hỗ trợ thị trường</h2><div class="sub">Hàng tặng NPP – không tính tiền.</div></div><button class="btn ghost sm" id="httm-add">+ Thêm dòng</button></div>
      <div class="card-b" id="httm-box"></div></div>
    <div class="card sumbar" id="sum-box"></div>
  </div>`;
  $('#f-khach').oninput = e => draft.khach = e.target.value;
  $('#f-tinh').onchange = e => draft.tinh = e.target.value;
  $('#f-date').onchange = e => draft.ngay_dat = e.target.value;
  $('#f-ck').oninput = e => { draft.ck_pct = (Number(e.target.value) || 0) / 100; requote(); };
  $('#f-cn').oninput = e => { draft.cong_no = Number(e.target.value) || 0; requote(); };
  $('#f-vc').oninput = e => draft.phi_vc = Number(e.target.value) || 0;
  $('#f-cpk').oninput = e => draft.chi_phi_khac = Number(e.target.value) || 0;
  $('#f-pay') && ($('#f-pay').onchange = e => draft.pay_status = e.target.value);
  $('#f-note').oninput = e => draft.ghi_chu = e.target.value;
  $('#cancel-edit')?.addEventListener('click', () => { draft = newDraft(kenh); renderOrderForm(el, kenh); });
  $('#f-search').oninput = e => { const q = e.target.value.toLowerCase(); $$('#prod-body tr').forEach(tr => tr.hidden = q && !tr.dataset.name.includes(q)); };
  $('#httm-add').onclick = () => { draft.httm.push({ sp: '', thung: 0 }); renderHttm(); };
  renderCombos();
  renderProds();
  renderHttm();
  requote();
}
function renderHttm() {
  const box = $('#httm-box'); if (!box) return;
  const opts = M.sellable;
  box.innerHTML = draft.httm.length ? `<table><thead><tr><th>Mặt hàng</th><th class="n">SL chai</th><th></th></tr></thead><tbody>
    ${draft.httm.map((h, i) => `<tr><td><select data-hs="${i}" style="width:100%"><option value="">— Chọn hàng —</option>${opts.map(p => `<option value="${p.id}" ${p.id === h.sp ? 'selected' : ''}>${esc(p.ten)}</option>`).join('')}</select></td>
      <td class="n"><input type="number" min="0" step="1" data-hq="${i}" value="${h.chai || ''}" style="width:84px"></td>
      <td><button class="btn danger sm" data-hd="${i}">Xóa</button></td></tr>`).join('')}</tbody></table>`
    : '<div class="muted small">Chưa có hàng HTTM. Bấm "+ Thêm dòng" nếu muốn tặng.</div>';
  $$('[data-hs]', box).forEach(s => s.onchange = () => { draft.httm[s.dataset.hs].sp = s.value; });
  $$('[data-hq]', box).forEach(i => i.oninput = () => { draft.httm[i.dataset.hq].chai = Number(i.value) || 0; });
  $$('[data-hd]', box).forEach(b => b.onclick = () => { draft.httm.splice(b.dataset.hd, 1); renderHttm(); });
}
let comboRows = [];
function renderCombos() {
  const body = $('#combo-body');
  comboRows = [];
  for (const c of M.combos) {
    const bom = S.combos[c.id] || {};
    if (bom.choice && bom.choice.length) {
      // 1 dòng cho MỖI lựa chọn -> cho phép nhập nhiều món cùng combo (VD Muối Ruốc + Mắm Nêm)
      for (const o of bom.choice) comboRows.push({ c, choice_sp: o.sp || null, choiceTen: o.ten, hasChoice: true });
    } else {
      comboRows.push({ c, choice_sp: null, choiceTen: '', hasChoice: false });
    }
  }
  body.innerHTML = comboRows.map((r, i) => {
    const line = draft.combos.find(x => x.combo_id === r.c.id && (x.choice_sp || null) === (r.choice_sp || null)) || {};
    return `<tr><td class="combo-name">${esc(r.c.ten)}</td>
      <td>${r.hasChoice ? `<span class="pill info">${esc(r.choiceTen)}</span>` : '<span class="muted small">—</span>'}</td>
      <td class="n">${fmt(r.c.gia_ban)}</td>
      <td class="n"><input type="number" min="0" step="1" class="qty ${line.phan ? 'has' : ''}" data-ci="${i}" value="${line.phan || ''}" style="width:90px"></td></tr>`;
  }).join('') || '<tr><td colspan="4" class="empty">Chưa có combo nào.</td></tr>';
  $$('[data-ci]', body).forEach(inp => inp.oninput = () => { const r = comboRows[inp.dataset.ci]; setCombo(r.c.id, r.choice_sp, Number(inp.value) || 0); inp.classList.toggle('has', !!inp.value); });
}
function setCombo(id, choiceSp, phan) {
  draft.combos = draft.combos.filter(x => !(x.combo_id === id && (x.choice_sp || null) === (choiceSp || null)));
  if (phan > 0) draft.combos.push({ combo_id: id, choice_sp: choiceSp || null, phan });
  requote();
}
function renderProds() {
  const body = $('#prod-body');
  body.innerHTML = M.sellable.map(p => {
    const ln = draft.lines[p.id] || {};
    return `<tr data-name="${esc(p.ten.toLowerCase())}"><td>${esc(p.ten)}</td><td class="small">${esc(p.dvt)}</td><td class="n">${p.quy_cach}</td>
      <td class="n"><input type="number" min="0" class="price" data-price="${p.id}" value="${ln.gia_chai ?? ''}" placeholder="${fmt(p.gia_ban)}" style="width:96px"></td>
      <td class="n"><input type="number" min="0" step="1" class="qty ${ln.chai ? 'has' : ''}" data-sp="${p.id}" value="${ln.chai || ''}" style="width:84px"></td>
      <td class="n" data-tt="${p.id}"></td></tr>`;
  }).join('');
  const lineOf = sp => (draft.lines[sp] ||= {});
  const tidy = sp => { const l = draft.lines[sp]; if (l && !l.chai && l.gia_chai == null) delete draft.lines[sp]; };
  $$('.qty[data-sp]', body).forEach(inp => inp.oninput = () => { lineOf(inp.dataset.sp).chai = Number(inp.value) || 0; inp.classList.toggle('has', !!inp.value); tidy(inp.dataset.sp); requote(); });
  $$('.price', body).forEach(inp => inp.oninput = () => { const v = inp.value === '' ? null : Number(inp.value) || 0; lineOf(inp.dataset.price).gia_chai = v; tidy(inp.dataset.price); requote(); });
}
// Giao diện nhập theo CHAI; máy chủ dùng đơn vị thùng -> quy đổi thùng = chai / quy cách (có thể lẻ).
function toThung(sp, chai) { const qc = M.prod[sp]?.quy_cach || 1; return (Number(chai) || 0) / qc; }
function draftBody() {
  return {
    kenh: draft.kenh, id: draft.id, khach: draft.khach, tinh: draft.tinh, ngay_dat: draft.ngay_dat,
    ck_pct: draft.ck_pct, cong_no: draft.cong_no, phi_vc: draft.phi_vc, chi_phi_khac: draft.chi_phi_khac, pay_status: draft.pay_status, ghi_chu: draft.ghi_chu,
    lines: Object.entries(draft.lines).filter(([, l]) => l.chai > 0).map(([sp, l]) => ({ sp, thung: toThung(sp, l.chai), gia_chai: l.gia_chai ?? null })),
    combos: draft.combos.filter(c => c.phan > 0).map(c => ({ combo_id: c.combo_id, choice_sp: c.choice_sp || null, phan: c.phan })),
    httm: draft.httm.filter(h => h.sp && h.chai).map(h => ({ sp: h.sp, thung: toThung(h.sp, h.chai) })),
  };
}
let qTimer = null;
function requote() {
  clearTimeout(qTimer);
  qTimer = setTimeout(async () => {
    try { lastQuote = await api('/api/quote', draftBody()); } catch (e) { toast(e.message, true); return; }
    for (const l of lastQuote.lines) { const td = $(`[data-tt="${l.sp}"]`); if (td) td.textContent = fmt(l.thanh_tien); }
    $$('#prod-body [data-tt]').forEach(td => { if (!lastQuote.lines.find(l => l.sp === td.dataset.tt)) td.textContent = ''; });
    renderSummary();
  }, 120);
}
function renderSummary() {
  const box = $('#sum-box'); if (!box) return;
  const q = lastQuote;
  if (!q || (!q.lines.length && !q.combos.length)) { box.innerHTML = '<div class="muted">Chọn combo hoặc mặt hàng — tổng đơn hiện ở đây.</div>'; return; }
  // kiểm tra combo chọn 1/3 đã chọn chưa
  const missing = q.combos.filter(c => { const bom = S.combos[c.combo_id]; return bom && bom.choice && bom.choice.length && !c.choice_sp; });
  box.innerHTML = `<div class="sb-row">
      <div class="sb-cell"><span>Tổng tiền hàng</span><b>${fmt(q.tong_tien_hang)}</b></div>
      <div class="sb-cell"><span>Chiết khấu ${pct(q.ck_pct)}</span><b>− ${fmt(q.chiet_khau)}</b></div>
      <div class="sb-cell"><span>Sau chiết khấu</span><b>${fmt(q.sau_ck)}</b></div>
      <div class="sb-cell cn"><span>Công nợ</span><b>${fmt(q.cong_no)}</b></div>
      <div class="sb-cell big"><span>Tổng thanh toán</span><b>${fmt(q.tong_thanh_toan)}</b></div>
      <div class="sb-btns"><button class="btn" id="save-order">${draft.id ? 'Lưu thay đổi' : 'Lưu đơn'}</button>
        <button class="btn accent" id="save-export">Lưu &amp; xuất đơn + PXK</button></div></div>
    ${missing.length ? `<div class="alert bad">Còn ${missing.length} combo "chọn 1 trong 3" chưa chọn món — hãy chọn trước khi lưu.</div>` : ''}`;
  $('#save-order').onclick = () => saveOrder(false, missing.length);
  $('#save-export').onclick = () => saveOrder(true, missing.length);
}
async function saveOrder(exportToo, missing) {
  if (missing) { toast('Hãy chọn món cho combo "chọn 1 trong 3" trước.', true); return; }
  const btns = $$('#sum-box .btn'); btns.forEach(b => b.disabled = true);
  try {
    const r = await run(() => api('/api/order/save', draftBody()));
    toast(`Đã lưu đơn ${r.order.ma}.`);
    if (exportToo) await exportOrder(r.order.id);
    draft = newDraft(draft.kenh);
    await refresh();
  } catch (e) { btns.forEach(b => b.disabled = false); }
}
async function exportOrder(id) {
  const t = toast('Đang điền mẫu Excel… (khoảng 10 giây)');
  try {
    const r = await api('/api/order/export', { id });
    t.remove();
    const warn = Object.keys(r.lech || {}).length ? ' <b>(!) Số trên file lệch, kiểm tra lại.</b>' : '';
    toast(`Đã xuất. ${openFileLink(r.file, 'Mở Excel')} · ${openFileLink(r.pdf, 'Mở PDF')}${warn}`, false, true);
  } catch (e) { t.remove(); toast(e.message, true); }
}

// =====================================================================
// ĐƠN HÀNG
// =====================================================================
function renderOrders(el, kenh) {
  const f = (filters[kenh] ||= { tinh: '', month: '', q: '' });
  const all = ordersOf(kenh);
  const months = [...new Set(all.map(o => (o.ngay_dat || '').slice(0, 7)).filter(Boolean))].sort().reverse();
  const rows = all.filter(o => (!f.tinh || o.tinh === f.tinh) && (!f.month || (o.ngay_dat || '').startsWith(f.month))
    && (!f.q || (o.khach || '').toLowerCase().includes(f.q.toLowerCase()) || (o.ma || '').toLowerCase().includes(f.q.toLowerCase())))
    .sort((a, b) => (b.ngay_dat || '').localeCompare(a.ngay_dat || '') || (b.ma || '').localeCompare(a.ma || ''));
  el.innerHTML = `
  <div class="card"><div class="card-b toolbar">
    <select id="o-tinh"><option value="">Mọi tỉnh</option>${S.provinces.map(t => `<option ${f.tinh === t ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>
    <select id="o-month"><option value="">Mọi tháng</option>${months.map(m => `<option value="${m}" ${f.month === m ? 'selected' : ''}>Tháng ${monthLabel(m)}</option>`).join('')}</select>
    <input id="o-q" class="grow" placeholder="Tìm khách hàng hoặc mã đơn…" value="${esc(f.q)}">
    <a class="btn" href="#${kenh}/len-don">+ Lên đơn mới</a></div></div>
  <div class="card"><div class="tbl-wrap"><table><thead><tr><th>Mã đơn</th><th>Ngày</th><th>Khách</th><th>Tỉnh</th><th class="n">Doanh thu</th><th class="n">HH ${esc(S.settings.ten_hh_cp)}</th><th class="n">HH Bưu cục</th><th class="n">Lợi nhuận</th><th>Trạng thái</th><th>Thao tác</th></tr></thead><tbody>
  ${rows.length ? rows.map(o => `<tr><td class="nowrap"><b>${esc(o.ma)}</b></td><td>${fd(o.ngay_dat)}</td><td>${esc(o.khach)}</td><td class="small">${esc(o.tinh)}</td>
    <td class="n">${fmt(o.profit.doanh_thu)}</td><td class="n">${fmt(o.profit.tien_cp)}</td><td class="n">${fmt(o.profit.tien_tinh)}</td><td class="n" style="color:${o.profit.loi_nhuan < 0 ? 'var(--bad)' : 'inherit'}">${fmt(o.profit.loi_nhuan)}</td>
    <td><span class="pill ${ST_CLS[o.trang_thai]}">${S.status[o.trang_thai]}</span><br><span class="pill ${o.pay_status === 'da_tt' ? 'ok' : 'warn'}" style="margin-top:3px">${S.pay_status[o.pay_status] || ''}</span></td>
    <td><div class="actions">
      <button class="btn ghost sm" data-view="${o.id}">Xem</button>
      ${o.trang_thai !== 'huy' ? `<button class="btn ghost sm" data-status="${o.id}">Trạng thái</button><button class="btn sm" data-deliver="${o.id}">Giao hàng</button><button class="btn ghost sm" data-export="${o.id}">Xuất file</button>` : ''}
      ${o.trang_thai === 'da_len_don' ? `<button class="btn ghost sm" data-edit="${o.id}">Sửa</button><button class="btn danger sm" data-cancel="${o.id}">Hủy</button>` : ''}
    </div></td></tr>`).join('') : '<tr><td colspan="10" class="empty">Không có đơn nào.</td></tr>'}
  </tbody></table></div></div>`;
  $('#o-tinh').onchange = e => { f.tinh = e.target.value; renderOrders(el, kenh); };
  $('#o-month').onchange = e => { f.month = e.target.value; renderOrders(el, kenh); };
  $('#o-q').oninput = e => { f.q = e.target.value; clearTimeout(qTimer); qTimer = setTimeout(() => { renderOrders(el, kenh); const i = $('#o-q'); i.focus(); i.setSelectionRange(i.value.length, i.value.length); }, 250); };
  $$('[data-view]', el).forEach(b => b.onclick = () => viewOrder(b.dataset.view));
  $$('[data-export]', el).forEach(b => b.onclick = () => exportOrder(b.dataset.export));
  $$('[data-edit]', el).forEach(b => b.onclick = () => editOrder(b.dataset.edit, kenh));
  $$('[data-cancel]', el).forEach(b => b.onclick = () => confirmDialog('Hủy đơn này?', 'Đơn hủy không tính vào doanh thu/lợi nhuận.', async () => { await run(() => api('/api/order/cancel', { id: b.dataset.cancel }), 'Đã hủy đơn.'); refresh(); }));
  $$('[data-deliver]', el).forEach(b => b.onclick = () => deliverDialog(b.dataset.deliver));
  $$('[data-status]', el).forEach(b => b.onclick = () => statusDialog(b.dataset.status));
}
function statusDialog(id) {
  const o = S.orders.find(x => x.id === id);
  modal(`Cập nhật trạng thái — đơn ${o.ma}`, `
    <div class="fields">
      <div class="field"><label>Trạng thái đơn</label><select id="s-trang">${Object.entries(S.status).map(([k, v]) => `<option value="${k}" ${k === o.trang_thai ? 'selected' : ''}>${v}</option>`).join('')}</select></div>
      <div class="field"><label>Thanh toán</label><select id="s-pay">${Object.entries(S.pay_status).map(([k, v]) => `<option value="${k}" ${k === o.pay_status ? 'selected' : ''}>${v}</option>`).join('')}</select></div>
    </div>
    <div class="small muted">Lưu ý: nếu sau đó anh/chị ghi nhận đợt giao thì "Trạng thái đơn" sẽ tự cập nhật lại theo tiến độ giao.</div>`,
    `<button class="btn ghost" data-close2>Hủy</button><button class="btn" id="s-ok">Lưu</button>`);
  $('#s-ok').onclick = async () => {
    const trang_thai = $('#s-trang').value, pay_status = $('#s-pay').value;
    closeModal();
    await run(() => api('/api/order/set-status', { id, trang_thai, pay_status }), 'Đã cập nhật trạng thái.');
    refresh();
  };
}
function viewOrder(id) {
  const o = S.orders.find(x => x.id === id);
  const dv = o.delivery;
  modal(`Đơn ${o.ma} — ${o.khach}`, `
    <div class="fields small"><div><span class="muted">Ngày đặt</span><br><b>${fd(o.ngay_dat)}</b></div><div><span class="muted">Tỉnh</span><br><b>${esc(o.tinh)}</b></div>
      <div><span class="muted">Trạng thái</span><br><span class="pill ${ST_CLS[o.trang_thai]}">${S.status[o.trang_thai]}</span></div></div>
    ${o.ghi_chu ? `<div class="alert warn">${esc(o.ghi_chu)}</div>` : ''}
    <h4>Hàng đặt</h4>
    <div class="tbl-wrap"><table><thead><tr><th>Mặt hàng</th><th class="n">SL</th><th class="n">Thành tiền</th></tr></thead><tbody>
      ${o.combos.map(c => { const p = M.prod[c.combo_id]; const ch = c.choice_sp ? ` <span class="pill info">${esc(M.prod[c.choice_sp]?.ten.split(' ').slice(0, 3).join(' '))}</span>` : ''; return `<tr><td>${esc(p?.ten || '')}${ch}</td><td class="n">${fmt(c.phan)} phần</td><td class="n">${fmt(c.phan * (p?.gia_ban || 0))}</td></tr>`; }).join('')}
      ${o.lines.map(l => { const p = M.prod[l.sp]; return `<tr><td>${esc(p?.ten)}</td><td class="n">${fmt(l.thung * (p?.quy_cach || 1))} chai</td><td class="n">${fmt(l.thung * (l.gia_chai || p?.gia_ban || 0) * (p?.quy_cach || 1))}</td></tr>`; }).join('')}
    </tbody></table></div>
    <div class="sum-row"><span>Doanh thu (giá niêm yết)</span><b>${fmt(o.profit.doanh_thu)}</b></div>
    <div class="sum-row"><span>Giá vốn</span><span>${fmt(o.profit.gia_von)}</span></div>
    <div class="sum-row"><span>Chiết khấu ${pct(o.ck_pct)}</span><span>− ${fmt(o.profit.chiet_khau)}</span></div>
    <div class="sum-row"><span>Hoa hồng ${esc(S.settings.ten_hh_cp)} + Tỉnh</span><span>− ${fmt(o.profit.tien_cp + o.profit.tien_tinh)}</span></div>
    <div class="sum-row ${o.profit.loi_nhuan < 0 ? 'neg' : ''}"><span>Lợi nhuận</span><b>${fmt(o.profit.loi_nhuan)}</b></div>
    <h4>Tiến độ giao (đã tách combo)</h4>
    <div class="tbl-wrap"><table><thead><tr><th>SP</th><th class="n">Đặt</th><th class="n">Đã giao</th><th class="n">Còn nợ</th></tr></thead><tbody>
      ${Object.entries(dv.per_sp).map(([sp, v]) => `<tr><td>${esc(M.prod[sp]?.ten || sp)}</td><td class="n">${fmt(v.ordered)}</td><td class="n">${fmt(v.delivered)}</td><td class="n">${v.remaining > 0 ? `<b>${fmt(v.remaining)}</b>` : '✓'}</td></tr>`).join('')}
    </tbody></table></div>`,
    `<button class="btn ghost" data-close2>Đóng</button><button class="btn" data-open-export="${o.id}">Xuất file</button>`, true);
  $('[data-open-export]').onclick = () => { closeModal(); exportOrder(o.id); };
}
function editOrder(id, kenh) {
  const o = S.orders.find(x => x.id === id);
  draft = { kenh, id: o.id, khach: o.khach, tinh: o.tinh, ngay_dat: o.ngay_dat, ck_pct: o.ck_pct, cong_no: o.cong_no, phi_vc: o.phi_vc, chi_phi_khac: o.chi_phi_khac, pay_status: o.pay_status, ghi_chu: o.ghi_chu || '',
    lines: Object.fromEntries(o.lines.map(l => [l.sp, { chai: Math.round(l.thung * (M.prod[l.sp]?.quy_cach || 1)), gia_chai: l.gia_chai ?? null }])),
    combos: o.combos.map(c => ({ combo_id: c.combo_id, choice_sp: c.choice_sp, phan: c.phan })),
    httm: (o.httm || []).map(h => ({ sp: h.sp, chai: Math.round(h.thung * (M.prod[h.sp]?.quy_cach || 1)) })) };
  location.hash = `#${kenh}/len-don`;
}
function deliverDialog(id) {
  const o = S.orders.find(x => x.id === id);
  const dv = o.delivery;
  const pending = Object.entries(dv.per_sp).filter(([, v]) => v.remaining > 0);
  if (!pending.length) { toast('Đơn đã giao đủ.'); return; }
  modal(`Giao hàng — đơn ${o.ma} (đợt ${dv.dot.length + 1})`, `
    <p>Nhập số lượng (chai/hũ/phần) giao đợt này cho từng SP. Bỏ trống = không giao SP đó đợt này.</p>
    <div class="field" style="max-width:220px"><label>Ngày giao</label><input type="date" id="d-date" value="${S.today}"></div>
    <div class="tbl-wrap"><table><thead><tr><th>SP</th><th class="n">Còn nợ</th><th class="n">Giao đợt này</th></tr></thead><tbody>
      ${pending.map(([sp, v]) => `<tr><td>${esc(M.prod[sp]?.ten || sp)}</td><td class="n">${fmt(v.remaining)}</td>
        <td class="n"><input type="number" min="0" max="${v.remaining}" data-giao="${sp}" value="${v.remaining}" style="width:100px"></td></tr>`).join('')}
    </tbody></table></div>`,
    `<button class="btn ghost" data-close2>Hủy</button><button class="btn" id="d-ok">Xác nhận giao đợt ${dv.dot.length + 1}</button>`, true);
  $('#d-ok').onclick = async () => {
    const items = $$('[data-giao]').map(i => ({ sp: i.dataset.giao, qty: Number(i.value) || 0 })).filter(x => x.qty > 0);
    if (!items.length) { toast('Chưa nhập số lượng giao.', true); return; }
    closeModal();
    await run(() => api('/api/order/deliver', { id, ngay: $('#d-date')?.value || S.today, items }), 'Đã ghi nhận đợt giao.');
    refresh();
  };
}

// =====================================================================
// DANH MỤC & THIẾT LẬP
// =====================================================================
let catTab = 'sp';
function renderCatalog(el) {
  el.innerHTML = `
  <div class="card"><div class="tabs">
    ${[['sp', 'Sản phẩm & giá'], ['combo', 'Combo (công thức)'], ['tt', 'Tải trọng'], ['tinh', 'Tỉnh'], ['hh', 'Hoa hồng']].map(([k, l]) => `<button data-tab="${k}" class="${catTab === k ? 'active' : ''}">${l}</button>`).join('')}
  </div><div class="card-b" id="cat-body"></div></div>`;
  $$('[data-tab]', el).forEach(b => b.onclick = () => { catTab = b.dataset.tab; renderCatalog(el); });
  const body = $('#cat-body');
  if (catTab === 'sp') catProducts(body);
  else if (catTab === 'combo') catCombos(body);
  else if (catTab === 'tt') catWeights(body);
  else if (catTab === 'tinh') catProvinces(body);
  else catCommission(body);
}
function prodRowHtml(p) {
  p = p || { id: '', ten: '', ma_vach: '', dvt: '', quy_cach: 1, gia_von: 0, gia_ban: 0, dang_ban: true };
  return `<tr data-id="${p.id}"><td><input value="${esc(p.ten)}" data-f="ten" style="min-width:200px"></td>
    <td><input value="${esc(p.ma_vach)}" data-f="ma_vach" style="width:120px"></td>
    <td><input value="${esc(p.dvt)}" data-f="dvt" style="width:70px"></td><td class="n"><input type="number" value="${p.quy_cach}" data-f="quy_cach" style="width:60px"></td>
    <td class="n"><input type="number" value="${p.gia_von}" data-f="gia_von" style="width:100px"></td><td class="n"><input type="number" value="${p.gia_ban}" data-f="gia_ban" style="width:100px"></td>
    <td class="c"><input type="checkbox" data-f="dang_ban" ${p.dang_ban ? 'checked' : ''}></td>
    <td class="c"><button class="btn danger sm sp-del" type="button" title="Xóa">×</button></td></tr>`;
}
function catProducts(body) {
  body.innerHTML = `<div class="sub" style="margin-bottom:8px">Thêm / sửa / xóa sản phẩm (tên, mã vạch, ĐVT, quy cách, giá) rồi bấm Lưu. Không xóa được sản phẩm đang dùng trong đơn hoặc combo.</div>
    <div class="tbl-wrap"><table><thead><tr><th>Tên hàng</th><th>Mã vạch</th><th>ĐVT</th><th class="n">QC</th><th class="n">Giá vốn</th><th class="n">Giá bán</th><th class="c">Bán</th><th class="c"></th></tr></thead><tbody id="sp-tbody">
    ${S.products.filter(p => !p.la_combo).map(prodRowHtml).join('')}
    </tbody></table></div>
    <div style="margin-top:10px"><button class="btn ghost" id="add-sp" type="button">+ Thêm sản phẩm</button>
      <button class="btn" id="save-sp" style="margin-left:8px">Lưu sản phẩm</button></div>`;
  $('#add-sp').onclick = () => $('#sp-tbody').insertAdjacentHTML('beforeend', prodRowHtml(null));
  $('#sp-tbody').onclick = e => { const d = e.target.closest('.sp-del'); if (d) d.closest('tr').remove(); };
  $('#save-sp').onclick = async () => {
    const products = $$('#sp-tbody tr[data-id]').map(tr => { const o = { id: tr.dataset.id }; $$('[data-f]', tr).forEach(i => o[i.dataset.f] = i.type === 'checkbox' ? i.checked : i.value); return o; })
      .filter(o => o.id || (o.ten || '').trim());  // bỏ dòng mới để trống
    await run(() => api('/api/catalog/save', { products }), 'Đã lưu sản phẩm.'); refresh();
  };
}
function comboRowHtml(prodOpts, sp, qty, withQty) {
  const opts = '<option value="">— chọn SP —</option>' + prodOpts.map(o => `<option value="${o.id}" ${o.id === sp ? 'selected' : ''}>${esc(o.ten)}</option>`).join('');
  return `<div class="ce-row" style="display:flex;gap:6px;margin:3px 0;align-items:center"><select class="ce-sp" style="flex:1;min-width:180px">${opts}</select>
    ${withQty ? `<input type="number" class="ce-qty" min="1" value="${qty || 1}" style="width:60px">` : '<input type="hidden" class="ce-qty" value="1">'}
    <button class="btn danger sm ce-del" type="button">×</button></div>`;
}
function catCombos(body) {
  const prodOpts = S.products.filter(p => !p.la_combo).map(p => ({ id: p.id, ten: p.ten }));
  const combos = S.products.filter(p => p.la_combo);
  const comboCard = (c) => { const b = (c && S.combos[c.id]) || {}; const id = c ? c.id : ''; return `<div class="card" data-cid="${id}" data-new="${c ? '' : '1'}" style="margin-bottom:10px"><div class="card-h">
      ${c ? `<b>${esc(c.ten)}</b>` : `<input class="combo-ten" placeholder="Tên combo mới…" style="min-width:320px">`}
      <div style="display:flex;gap:10px;align-items:center">
        <label class="small" style="display:flex;gap:5px;align-items:center">Giá bán <input type="number" class="combo-gb" value="${c ? c.gia_ban : 0}" style="width:100px"></label>
        <label class="small" style="display:flex;gap:5px;align-items:center"><input type="checkbox" class="combo-on" ${!c || c.dang_ban ? 'checked' : ''}>Đang bán</label>
        <button class="btn danger sm combo-del" type="button">Xóa combo</button></div></div><div class="card-b">
      <div class="small muted">Thành phần (× số lượng / 1 combo)</div><div data-sec="components">${(b.components || []).map(x => comboRowHtml(prodOpts, x.sp, x.qty, true)).join('')}</div><button class="btn ghost sm" type="button" data-add="components">+ thành phần</button>
      <div class="small muted" style="margin-top:8px">Chọn 1 trong các món (để trống nếu không có)</div><div data-sec="choice">${(b.choice || []).map(x => comboRowHtml(prodOpts, x.sp, 1, false)).join('')}</div><button class="btn ghost sm" type="button" data-add="choice">+ lựa chọn</button>
      <div class="small muted" style="margin-top:8px">Hàng tặng</div><div data-sec="gift">${(b.gift || []).map(x => comboRowHtml(prodOpts, x.sp, x.qty, true)).join('')}</div><button class="btn ghost sm" type="button" data-add="gift">+ tặng</button>
    </div></div>`; };
  body.innerHTML = `<div class="sub" style="margin-bottom:8px">Thêm / sửa / xóa combo. Tạo combo mới: đặt tên + giá bán + thành phần. Bỏ tích "Đang bán" để ẩn combo không dùng. Không xóa được combo đang dùng trong đơn. Thành phần sẽ tách ra ở Phiếu Xuất Kho.</div>`
    + `<div id="combo-list">` + (combos.map(comboCard).join('') || '') + `</div>`
    + `<div style="margin-top:4px"><button class="btn ghost" id="add-combo" type="button">+ Thêm combo</button>`
    + `<button class="btn" id="save-combo" style="margin-left:8px">Lưu combo</button></div>`;
  $('#add-combo').onclick = () => $('#combo-list').insertAdjacentHTML('beforeend', comboCard(null));
  body.onclick = e => {
    const add = e.target.closest('[data-add]');
    if (add) { add.previousElementSibling.insertAdjacentHTML('beforeend', comboRowHtml(prodOpts, '', 1, add.dataset.add !== 'choice')); return; }
    const del = e.target.closest('.ce-del');
    if (del) { del.closest('.ce-row').remove(); return; }
    const cdel = e.target.closest('.combo-del');
    if (cdel) cdel.closest('[data-cid]').remove();
  };
  $('#save-combo').onclick = async () => {
    const combos = {}, combos_meta = [];
    let ok = true;
    $$('[data-cid]', body).forEach(card => {
      const get = sec => $$(`[data-sec="${sec}"] .ce-row`, card).map(r => ({ sp: $('.ce-sp', r).value || null, qty: Number($('.ce-qty', r).value) || 1 })).filter(x => x.sp);
      const choice = get('choice');
      const bom = { components: get('components'), choice: choice.length ? choice.map(x => ({ sp: x.sp })) : null, gift: get('gift') };
      const isNew = card.dataset.new === '1';
      const ten = isNew ? ($('.combo-ten', card).value || '').trim() : null;
      if (isNew && !ten) { ok = false; return; }  // combo mới chưa đặt tên: bỏ qua
      const meta = { id: card.dataset.cid || '', dang_ban: $('.combo-on', card).checked, gia_ban: Number($('.combo-gb', card).value) || 0 };
      if (isNew) meta.ten = ten;
      combos_meta.push(meta);
    });
    // id của combo mới do máy chủ sinh; gửi BOM theo id hiện có, BOM combo mới sẽ lưu ở lần làm mới sau
    $$('[data-cid]', body).forEach(card => { if (card.dataset.cid) { const get = sec => $$(`[data-sec="${sec}"] .ce-row`, card).map(r => ({ sp: $('.ce-sp', r).value || null, qty: Number($('.ce-qty', r).value) || 1 })).filter(x => x.sp); const choice = get('choice'); combos[card.dataset.cid] = { components: get('components'), choice: choice.length ? choice.map(x => ({ sp: x.sp })) : null, gift: get('gift') }; } });
    if (!ok) { toast('Combo mới cần có tên.', true); return; }
    await run(() => api('/api/catalog/save', { combos, combos_meta, combos_full: true }), 'Đã lưu combo.'); refresh();
  };
}
function catWeights(body) {
  body.innerHTML = `<div class="sub" style="margin-bottom:8px">Quy cách & KG/thùng dùng để tính <b>tải trọng</b> (khác quy cách đơn hàng). SP thiếu KG sẽ tính tải trọng = 0.</div>
    <div class="tbl-wrap"><table><thead><tr><th>Tên hàng</th><th class="n">Quy cách (tải trọng)</th><th class="n">KG/thùng</th></tr></thead><tbody>
    ${S.products.filter(p => !p.la_combo).map(p => { const w = S.weights[p.id] || {}; return `<tr data-id="${p.id}"><td>${esc(p.ten)}</td>
      <td class="n"><input type="number" value="${w.quy_cach_tt ?? ''}" data-f="quy_cach_tt" style="width:90px"></td>
      <td class="n"><input type="number" step="0.1" value="${w.kg_thung ?? ''}" data-f="kg_thung" style="width:90px"></td></tr>`; }).join('')}
    </tbody></table></div><div style="margin-top:12px"><button class="btn" id="save-tt">Lưu tải trọng</button></div>`;
  $('#save-tt').onclick = async () => {
    const weights = {};
    $$('#cat-body tr[data-id]').forEach(tr => { const qc = $('[data-f=quy_cach_tt]', tr).value, kg = $('[data-f=kg_thung]', tr).value; if (qc || kg) weights[tr.dataset.id] = { quy_cach_tt: Number(qc) || 0, kg_thung: Number(kg) || 0 }; });
    await run(() => api('/api/catalog/save', { weights }), 'Đã lưu tải trọng.'); refresh();
  };
}
function catProvinces(body) {
  body.innerHTML = `<div class="field"><label>Danh sách tỉnh (mỗi dòng 1 tỉnh)</label><textarea class="msg" id="prov-ta" style="min-height:300px">${esc(S.provinces.join('\n'))}</textarea></div>
    <div style="margin-top:12px"><button class="btn" id="save-prov">Lưu tỉnh</button></div>`;
  $('#save-prov').onclick = async () => {
    const provinces = $('#prov-ta').value.split('\n').map(s => s.trim()).filter(Boolean);
    await run(() => api('/api/catalog/save', { provinces }), 'Đã lưu tỉnh.'); refresh();
  };
}
function catCommission(body) {
  body.innerHTML = `<div class="fields" style="max-width:420px">
    <div class="field"><label>Người nhận hoa hồng giới thiệu</label><input id="hh-ten" value="${esc(S.settings.ten_hh_cp)}"></div>
    <div class="field"><label>Hoa hồng ${esc(S.settings.ten_hh_cp)} (%)</label><input type="number" step="0.5" id="hh-cp" value="${+(S.settings.cp_pct * 100).toFixed(2)}"></div>
    <div class="field"><label>Hoa hồng Bưu cục/Tỉnh (%)</label><input type="number" step="0.5" id="hh-tinh" value="${+(S.settings.tinh_pct * 100).toFixed(2)}"></div>
    </div><div class="small muted" style="margin:8px 0">Luôn trừ cả 2 mức khi tính lợi nhuận mỗi đơn.</div>
    <button class="btn" id="save-hh">Lưu hoa hồng</button>`;
  $('#save-hh').onclick = async () => {
    await run(() => api('/api/catalog/save', { settings: { ten_hh_cp: $('#hh-ten').value, cp_pct: (Number($('#hh-cp').value) || 0) / 100, tinh_pct: (Number($('#hh-tinh').value) || 0) / 100 } }), 'Đã lưu hoa hồng.'); refresh();
  };
}

// ---------- báo cáo theo tỉnh (tải khi mở dashboard) ----------
async function loadReport(kenh, month) {
  try {
    const rep = await api('/api/report', { kenh, month });
    M.report = { kenh, month, province: rep.province };
    const tb = $(`[data-prov="${kenh}|${month}"]`);
    if (tb) tb.innerHTML = provRows(rep.province);
  } catch (e) { /* bỏ qua */ }
}
// ---------- khởi động ----------
(async function boot() {
  try { if (await load()) route(); }
  catch (e) { $('#page').innerHTML = `<div class="alert bad">Không kết nối được máy chủ: ${esc(e.message)}</div>`; }
})();
