/* Giao diện công cụ Doanh thu – Đơn hàng – Công nợ. Mọi phép tính tiền đều do máy chủ (app.py) làm, giao diện chỉ hiển thị. */
'use strict';
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const REG_COLORS = { 'Miền Bắc': '#7C3AED', 'Miền Trung': '#2F5597', 'Đông Nam Bộ': '#C2620A', 'Tây Nam Bộ': '#157F3D' };
const STATUS = { cho_giao: ['Chờ giao', 'warn'], da_giao: ['Đã giao', 'ok'], huy: ['Đã hủy', 'mute'], lich_su: ['Lịch sử (file cũ)', 'info'] };

let S = null;                 // toàn bộ dữ liệu từ máy chủ
let M = {};                   // các bảng tra cứu dựng từ S
let draft = newDraft();       // đơn đang lên
let lastQuote = null;
const filters = { don: { status: 'all', month: '', region: '', q: '' }, cn: { region: '', gsm: '', owing: true }, so: { month: '', region: '' }, dm: 'npp', gia: 'Miền Trung' };

// ---------- tiện ích ----------
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = n => Math.round(Number(n) || 0).toLocaleString('vi-VN');
const fq = n => (Math.round((Number(n) || 0) * 100) / 100).toLocaleString('vi-VN');
const fd = d => d ? d.slice(0, 10).split('-').reverse().join('/') : '';
const pct = x => x == null || x === '' ? '' : `${+(x * 100).toFixed(2)}%`;
const monthLabel = m => m ? `${m.slice(5, 7)}/${m.slice(0, 4)}` : '';
const parseMoney = s => Number(String(s ?? '').replace(/[^\d-]/g, '')) || 0;
const pill = ([t, c]) => `<span class="pill ${c}">${esc(t)}</span>`;
const sumBy = (arr, k) => arr.reduce((a, x) => a + (x[k] || 0), 0);
function newDraft() { return { id: null, npp_id: '', lines: {}, httm: [], ngay_dat: new Date().toISOString().slice(0, 10), ghi_chu: '', ck: null, ck_km: null, cn_pct: null, region: '' }; }
function orderMonth(o) { return (o.trang_thai === 'lich_su' ? o.thang : o.ngay_giao || '').slice(0, 7); }
function isRevenue(o) { return o.trang_thai === 'da_giao' || o.trang_thai === 'lich_su'; }
function nppRegion(n) { return n && S.regions.includes(n.khu_vuc) ? n.khu_vuc : S.regions[0]; }

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

async function load() {
  if (!(await checkLicense())) return false;
  S = await api('/api/state');
  M.npp = Object.fromEntries(S.npps.map(n => [n.id, n]));
  M.sp = Object.fromEntries(S.products.map(p => [p.id, p]));
  M.paid = {};
  for (const it of S.debt_items) M.paid[it.order_id] = it.da_tra;
  $('#today').textContent = 'Hôm nay ' + fd(S.today);
  $('#side-foot').innerHTML = `File dữ liệu: ${openFileLink(S.data_file, 'mở bằng Excel')}<br>Công nợ chốt lấy ngày ${fd(S.settings.ngay_cn_chot)}<br>Nhắc nợ: từ ngày ${S.settings.nhac_tu_ngay} tới hết mỗi tháng`;
  return true;
}

// ---------- license (offline 20 ngày) ----------
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
    <div class="fields small"><div><span class="muted">Trạng thái</span><br>${pill(['Đang hoạt động', 'ok'])}</div>
      <div><span class="muted">Còn lại</span><br><b>${esc(l.remaining_text)}</b></div>
      <div><span class="muted">Phát hành</span><br><b>${esc(l.issued_text)}</b></div><div><span class="muted">Hết hạn</span><br><b>${esc(l.expires_text)}</b></div>
      <div><span class="muted">License ID</span><br><code>${esc(l.license_id)}</code></div></div>
    <div class="alert warn small">20 ngày tính từ lúc quản trị viên phát hành mã. Khi hết hạn, công cụ khóa các chức năng cho tới khi nhập mã mới; dữ liệu giữ nguyên.</div>
    <b>Nhập mã mới</b> <span class="small muted">(khi quản trị viên đã cấp mã cho chu kỳ tiếp theo)</span>
    ${licenseForm(l)}${policyBlock(l)}`,
    `<button class="btn ghost" data-close2>Đóng</button><button class="btn" id="lic-activate">Kích hoạt mã mới</button>`);
  bindLicenseForm(async () => { closeModal(); await refresh(); });
}
setInterval(() => checkLicense().catch(() => {}), 60_000);  // cập nhật thời gian còn lại, khóa ngay khi hết hạn

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

// ---------- điều hướng ----------
const PAGES = {
  'tong-quan': ['Tổng quan', 'Doanh thu tháng, nhắc nợ cuối tháng và đơn chờ giao', renderDashboard],
  'len-don': ['Lên đơn hàng', 'Giá & khuyến mãi theo vùng của NPP, chiết khấu nhập tay, hàng hỗ trợ thị trường', renderOrderForm],
  'don-hang': ['Đơn hàng', 'Theo dõi đơn: chờ giao → đã giao (vào sổ doanh thu)', renderOrders],
  'cong-no': ['Công nợ', 'NPP đã trả chưa, còn nợ bao nhiêu – nhắc mỗi cuối tháng tới khi hết nợ', renderDebts],
  'so-doanh-thu': ['Sổ doanh thu', 'Sổ theo dõi doanh thu theo tháng (đơn đã giao)', renderLedger],
  'danh-muc': ['Danh mục & cài đặt', 'NPP, bảng giá – khuyến mãi theo vùng, sản phẩm, cài đặt nhắc nợ', renderCatalog],
};
function route() {
  const page = (location.hash || '#tong-quan').slice(1);
  const [title, sub, fn] = PAGES[page] || PAGES['tong-quan'];
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.page === page));
  $('#page-title').textContent = title;
  $('#page-sub').textContent = sub;
  fn($('#page'));
}
window.addEventListener('hashchange', route);
$('#btn-ledger').onclick = async () => {
  const r = await run(() => api('/api/export/ledger', {}));
  toast(`Đã xuất sổ Excel. ${openFileLink(r.file, 'Mở file')}`, false, true);
};
async function refresh() { if (await load()) route(); }

// =====================================================================
// TỔNG QUAN
// =====================================================================
function debtTotals() {
  const t = { chot: 0, don: 0, truoc: 0, trong: 0, n: 0 };
  for (const st of Object.values(S.debt_npp)) {
    t.chot += st.cn_chot_con; t.don += st.con_no_don; t.truoc += st.no_thang_truoc; t.trong += st.no_trong_thang;
    if (st.tong_no > 0.5) t.n++;
  }
  return t;
}
function owingByGsm() {
  const groups = {};
  for (const n of S.npps) {
    const st = S.debt_npp[n.id];
    if (st && st.tong_no > 0.5) (groups[n.gsm || 'Chưa có quản lý'] ||= []).push({ n, st });
  }
  for (const list of Object.values(groups)) list.sort((a, b) => b.st.tong_no - a.st.tong_no);
  return Object.entries(groups).sort((a, b) => b[1].reduce((x, r) => x + r.st.tong_no, 0) - a[1].reduce((x, r) => x + r.st.tong_no, 0));
}
function renderDashboard(el) {
  const month = S.today.slice(0, 7);
  const rev = S.orders.filter(o => isRevenue(o) && orderMonth(o) === month);
  const paysMonth = S.payments.filter(p => p.ngay.startsWith(month));
  const pend = S.orders.filter(o => o.trang_thai === 'cho_giao');
  const t = debtTotals();
  const nh = S.nhac;
  el.innerHTML = `
  <div class="kpis">
    <div class="kpi main"><div class="l">Doanh thu tháng ${monthLabel(month)}</div><div class="v">${fmt(sumBy(rev, 'sau_ck'))}</div><div class="s">${rev.length} đơn đã giao · sau chiết khấu</div></div>
    <div class="kpi ok"><div class="l">Đã thu trong tháng</div><div class="v">${fmt(sumBy(paysMonth, 'so_tien'))}</div><div class="s">${paysMonth.length} lần NPP thanh toán</div></div>
    <div class="kpi"><div class="l">Tổng công nợ</div><div class="v">${fmt(t.chot + t.don)}</div><div class="s">${t.n} NPP còn nợ · chốt ${fmt(t.chot)}</div></div>
    <div class="kpi bad"><div class="l">Nợ đơn tồn từ tháng trước</div><div class="v">${fmt(t.truoc)}</div><div class="s">đã qua ít nhất 1 kỳ nhắc</div></div>
    <div class="kpi warn"><div class="l">Nợ đơn phát sinh trong tháng</div><div class="v">${fmt(t.trong)}</div></div>
    <div class="kpi"><div class="l">Đơn chờ giao</div><div class="v">${pend.length}</div><div class="s">${fmt(sumBy(pend, 'sau_ck'))} đ</div></div>
  </div>
  <div class="card"><div class="card-h"><div><h2>Nhắc nợ cuối tháng ${esc(nh.thang)}</h2>
      <div class="sub">Kỳ nhắc: ${fd(nh.tu)} → ${fd(nh.den)}. Nhắc mọi NPP còn nợ, tháng nào cũng nhắc cho tới khi công nợ = 0.</div></div>
      ${nh.dang_nhac ? pill(['Đang trong kỳ nhắc nợ', 'bad']) : pill([`Chưa tới kỳ – bắt đầu ${fd(nh.tu)}`, 'info'])}</div>
    <div class="card-b grid">${nh.dang_nhac ? '' : '<div class="alert warn small">Chưa tới kỳ nhắc nợ cuối tháng. Danh sách dưới đây để xem trước; tới ngày ' + fd(nh.tu) + ' công cụ sẽ báo đỏ.</div>'}${remindBlocks()}</div></div>
  <div class="grid g2">
    <div class="card"><div class="card-h"><div><h2>Doanh thu 6 tháng gần nhất</h2><div class="sub">Sau chiết khấu, theo vùng</div></div><div class="legend">${Object.entries(REG_COLORS).map(([r, c]) => `<span><i style="background:${c}"></i>${r}</span>`).join('')}<span><i style="background:#98A2B3"></i>Khách khác (không phải NPP)</span></div></div>
      <div class="card-b">${revenueChart()}</div></div>
    <div class="card"><div class="card-h"><h2>Đơn chờ giao</h2><a href="#don-hang" class="small">Xem tất cả</a></div>
      <div class="tbl-wrap">${pend.length ? `<table><thead><tr><th>Mã</th><th>NPP</th><th>Ngày đặt</th><th class="n">Sau CK</th><th></th></tr></thead><tbody>
      ${pend.slice(0, 8).map(o => `<tr><td class="nowrap">${esc(o.ma)}</td><td>${esc(M.npp[o.npp_id]?.ten)}</td><td>${fd(o.ngay_dat)}</td><td class="n">${fmt(o.sau_ck)}</td>
        <td><button class="btn sm" data-deliver="${o.id}">Đã giao</button></td></tr>`).join('')}</tbody></table>` : '<div class="empty">Không có đơn nào đang chờ giao.</div>'}</div></div>
  </div>`;
  bindCommon(el);
}
function remindBlocks() {
  const groups = owingByGsm();
  if (!groups.length) return '<div class="alert ok">Không NPP nào còn nợ. 👍</div>';
  return groups.map(([gsm, list]) => {
    const total = list.reduce((a, r) => a + r.st.tong_no, 0), old = list.reduce((a, r) => a + r.st.no_thang_truoc, 0);
    return `<details class="card remind" ${S.nhac.dang_nhac ? 'open' : ''}><summary class="card-h"><div><h2>${esc(gsm)}</h2><div class="sub">${list.length} NPP còn nợ · tổng ${fmt(total)} đ${old > 0.5 ? ` · <b style="color:var(--bad)">nợ đơn tồn tháng trước ${fmt(old)} đ</b>` : ''}</div></div>
      <button class="btn accent sm" data-zalo="${esc(gsm)}">Soạn tin Zalo</button></summary>
      <div class="tbl-wrap"><table><thead><tr><th>NPP</th><th class="n">Công nợ chốt còn</th><th class="n">Nợ đơn tháng trước</th><th class="n">Nợ đơn trong tháng</th><th class="n">Tổng còn nợ</th><th>Nợ đơn từ</th><th></th></tr></thead><tbody>
      ${list.map(({ n, st }) => `<tr><td><b>${esc(n.ten)}</b><div class="small muted">${esc(n.tinh)} · ${esc(n.sdt)}</div></td>
        <td class="n">${fmt(st.cn_chot_con)}</td><td class="n">${st.no_thang_truoc > 0.5 ? `<b style="color:var(--bad)">${fmt(st.no_thang_truoc)}</b>` : '–'}</td>
        <td class="n">${st.no_trong_thang > 0.5 ? fmt(st.no_trong_thang) : '–'}</td><td class="n"><b>${fmt(st.tong_no)}</b></td>
        <td class="nowrap">${st.no_tu_thang ? 'tháng ' + monthLabel(st.no_tu_thang) : '<span class="muted small">chỉ công nợ chốt</span>'}</td>
        <td><div class="actions"><button class="btn ghost sm" data-pay="${n.id}">Ghi nhận thu</button><button class="btn ghost sm" data-npp-detail="${n.id}">Chi tiết</button></div></td></tr>`).join('')}
      </tbody></table></div></details>`;
  }).join('');
}
function zaloMessage(gsm) {
  const list = (owingByGsm().find(([g]) => g === gsm) || [, []])[1];
  const lines = list.map(({ n, st }, i) => {
    const parts = [];
    if (st.cn_chot_con > 0.5) parts.push(`công nợ chốt ${fmt(st.cn_chot_con)}đ`);
    if (st.no_thang_truoc > 0.5) parts.push(`nợ đơn từ tháng ${monthLabel(st.no_tu_thang)} chưa trả ${fmt(st.no_thang_truoc)}đ`);
    if (st.no_trong_thang > 0.5) parts.push(`nợ đơn trong tháng ${fmt(st.no_trong_thang)}đ`);
    return `${i + 1}. ${n.ten} (${n.tinh}) – SĐT ${n.sdt}\n   Tổng còn nợ: ${fmt(st.tong_no)}đ (${parts.join('; ')})`;
  });
  return `Chào anh ${gsm},\nEm gửi danh sách công nợ NPP cần nhắc cuối tháng ${S.nhac.thang} (cập nhật ${fd(S.today)}):\n\n${lines.join('\n\n')}\n\nTổng cần thu: ${fmt(list.reduce((a, r) => a + r.st.tong_no, 0))}đ.\n${S.settings.chuyen_khoan ? `Thông tin chuyển khoản: ${S.settings.chuyen_khoan}.\n` : ''}Anh nhắc NPP thanh toán giúp em nhé. Cảm ơn anh!`;
}
function showZalo(gsm) {
  modal(`Tin nhắn Zalo gửi anh ${gsm}`, `<textarea class="msg" id="zalo-msg">${esc(zaloMessage(gsm))}</textarea><div class="small muted">Sửa lại nội dung nếu cần, rồi bấm "Copy" và dán vào Zalo. Công cụ không tự gửi Zalo.</div>${S.settings.chuyen_khoan ? '' : '<div class="alert warn small">Chưa có thông tin chuyển khoản trong tin nhắn — nhập ở Danh mục &amp; cài đặt › Cài đặt.</div>'}`,
    `<button class="btn ghost" data-close2>Đóng</button><button class="btn" id="copy-zalo">Copy nội dung</button>`);
  $('#copy-zalo').onclick = async () => {
    const ta = $('#zalo-msg');
    try { await navigator.clipboard.writeText(ta.value); } catch { ta.select(); document.execCommand('copy'); }
    toast('Đã copy. Mở Zalo và dán (Ctrl+V).');
  };
}
function revenueChart() {
  const months = [];
  const d = new Date(S.today + 'T00:00:00');
  for (let i = 5; i >= 0; i--) { const x = new Date(d.getFullYear(), d.getMonth() - i, 1); months.push(`${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}`); }
  const data = months.map(m => {
    const by = {};
    for (const o of S.orders) if (isRevenue(o) && orderMonth(o) === m) { const r = M.npp[o.npp_id]?.khu_vuc || 'Khác'; by[r] = (by[r] || 0) + o.sau_ck; }
    return { m, by, total: Object.values(by).reduce((a, b) => a + b, 0) };
  });
  const max = Math.max(1, ...data.map(x => x.total));
  return `<div class="bars">${data.map(x => `<div class="bar" title="${monthLabel(x.m)}: ${fmt(x.total)} đ">
    <div class="val">${x.total ? (x.total / 1e6).toLocaleString('vi-VN', { maximumFractionDigits: 0 }) + ' tr' : '–'}</div>
    <div class="stack" style="height:${Math.max(2, x.total / max * 160)}px">${Object.entries(x.by).map(([r, v]) => `<span style="height:${v / x.total * 100}%;background:${REG_COLORS[r] || '#98A2B3'}"></span>`).join('')}</div>
    <div class="lab">${monthLabel(x.m)}</div></div>`).join('')}</div>`;
}

// =====================================================================
// LÊN ĐƠN
// =====================================================================
function activeNpp(n) { return !/thanh lí|nghỉ/i.test(n.tinh_trang || ''); }
function kmRule(p, rule) {
  if (!rule) return '';
  if (rule.km_loai === 'cung_loai') return `Mỗi ${fq(rule.km_n)} thùng tặng ${fq(rule.km_m)} ${esc((p.dvt || '').toLowerCase())}`;
  if (rule.km_loai === 'khac') return `Mỗi thùng tặng ${fq(rule.km_moi_thung_tang)} chai ${esc(M.sp[p.km_sp]?.ten.split(' - ')[0] || '')}`;
  if (rule.km_loai === 'hang_tang') return 'Tự tính theo SP mua';
  return 'Không';
}
function renderOrderForm(el) {
  const npps = S.npps.filter(n => activeNpp(n) || n.id === draft.npp_id).filter(n => !draft.region || n.khu_vuc === draft.region);
  const groups = S.regions.map(r => [r, npps.filter(n => n.khu_vuc === r)]).filter(([, l]) => l.length);
  const npp = M.npp[draft.npp_id];
  const region = nppRegion(npp);
  const price = S.prices[region];
  const prods = S.products.filter(p => p.dang_ban && p.loai !== 'hang_tang');
  const byLabel = {};
  for (const p of prods) (byLabel[p.nhan || 'Khác'] ||= []).push(p);
  const editing = draft.id ? S.orders.find(o => o.id === draft.id) : null;
  el.innerHTML = `
  <div class="grid">
    <div class="card"><div class="card-h"><h2>${editing ? 'Sửa đơn ' + esc(editing.ma) : 'Thông tin đơn hàng'}</h2>${editing ? '<button class="btn ghost sm" id="cancel-edit">Bỏ sửa, lên đơn mới</button>' : ''}</div>
      <div class="card-b grid">
        <div class="fields">
          <div class="field"><label for="f-region">Lọc theo vùng</label><select id="f-region"><option value="">Tất cả vùng</option>${S.regions.map(r => `<option ${r === draft.region ? 'selected' : ''}>${r}</option>`).join('')}</select></div>
          <div class="field" style="grid-column:span 2"><label for="f-npp">Nhà phân phối</label><select id="f-npp"><option value="">— Chọn NPP —</option>
            ${groups.map(([r, l]) => `<optgroup label="${r}">${l.map(n => `<option value="${n.id}" ${n.id === draft.npp_id ? 'selected' : ''}>${esc(n.ten)} – ${esc(n.tinh)}</option>`).join('')}</optgroup>`).join('')}</select></div>
          <div class="field"><label for="f-date">Ngày đặt hàng</label><input type="date" id="f-date" value="${draft.ngay_dat}"></div>
        </div>
        <div id="npp-info"></div>
        <div class="field"><label for="f-note">Ghi chú</label><input id="f-note" value="${esc(draft.ghi_chu)}" placeholder="VD: giao chành, giao 2 đợt…"></div>
      </div></div>
    <div class="card"><div class="card-h"><div><h2>Hàng đặt ${npp ? `<span class="pill info">Giá & KM ${esc(region)}</span>` : ''}</h2><div class="sub">Nhập số thùng. Giá và số tặng lấy theo bảng giá vùng; sửa tay được từng dòng (ô vàng = đã sửa tay).</div></div>
      <input id="f-search" placeholder="Tìm mặt hàng…" style="width:220px"></div>
      <div class="tbl-wrap"><table id="prod-table"><thead><tr><th>Tên hàng</th><th>ĐVT</th><th class="n">QC</th><th class="n">Giá chai</th><th class="n">SL thùng</th><th class="n">Thành tiền</th><th>Khuyến mãi</th><th class="n">Tặng</th><th class="n">Tổng giao</th></tr></thead><tbody>
      ${Object.entries(byLabel).map(([lab, list]) => `<tr class="group"><td colspan="9">${esc(lab)}</td></tr>` + list.map(p => {
        const ln = draft.lines[p.id] || {}, rule = price[p.id] || {};
        return `<tr data-row="${p.id}" data-name="${esc(p.ten.toLowerCase())}"><td>${esc(p.ten)}${p.nhom_ck === 'km' ? ' <span class="pill warn">CK KM khác</span>' : ''}</td><td class="small">${esc(p.dvt)}</td><td class="n">${p.quy_cach}</td>
        <td class="n"><input type="number" min="0" class="price ${ln.gia_chai != null ? 'has' : ''}" data-price="${p.id}" value="${ln.gia_chai ?? ''}" placeholder="${fmt(rule.gia_chai)}" aria-label="Giá chai ${esc(p.ten)}"></td>
        <td class="n"><input type="number" min="0" step="1" class="qty ${ln.thung ? 'has' : ''}" data-sp="${p.id}" value="${ln.thung || ''}" aria-label="Số thùng ${esc(p.ten)}"></td>
        <td class="n" data-tt></td><td class="small km">${esc(rule.km_text || '')}</td>
        <td class="n"><input type="number" min="0" class="gift ${ln.tang != null ? 'has' : ''}" data-gift="${p.id}" value="${ln.tang ?? ''}" aria-label="Số tặng ${esc(p.ten)}"></td>
        <td class="n" data-giao></td></tr>`; }).join('')).join('')}
      </tbody></table></div></div>
    <div class="card"><div class="card-h"><div><h2>Hỗ trợ thị trường</h2><div class="sub">Hàng tặng riêng cho NPP – không tính tiền, in ở phần "Hỗ Trợ Thương Mại" cuối đơn. Số lượng tùy chọn.</div></div>
      <button class="btn ghost sm" id="httm-add">+ Thêm dòng</button></div>
      <div class="card-b" id="httm-box"></div></div>
    <div class="card sumbar" id="sum-box"><div class="muted">Chọn NPP và nhập số thùng.</div></div>
  </div>`;
  $('#f-region').onchange = e => { draft.region = e.target.value; renderOrderForm(el); };
  $('#f-npp').onchange = e => { draft.npp_id = e.target.value; draft.ck = draft.ck_km = draft.cn_pct = null; renderOrderForm(el); };
  $('#f-date').onchange = e => { draft.ngay_dat = e.target.value; };
  $('#f-note').oninput = e => { draft.ghi_chu = e.target.value; };
  $('#f-search').oninput = e => { const q = e.target.value.toLowerCase(); $$('#prod-table tr[data-row]').forEach(tr => { tr.hidden = q && !tr.dataset.name.includes(q); }); };
  const lineOf = sp => (draft.lines[sp] ||= {});
  const tidy = sp => { const l = draft.lines[sp]; if (l && !l.thung && l.gia_chai == null && l.tang == null) delete draft.lines[sp]; };
  $$('.qty', el).forEach(inp => inp.oninput = () => { const v = Math.max(0, Number(inp.value) || 0); lineOf(inp.dataset.sp).thung = v || undefined; inp.classList.toggle('has', !!v); tidy(inp.dataset.sp); requote(); });
  $$('.price', el).forEach(inp => inp.oninput = () => { const v = inp.value === '' ? null : Math.max(0, Number(inp.value) || 0); lineOf(inp.dataset.price).gia_chai = v; inp.classList.toggle('has', v != null); tidy(inp.dataset.price); requote(); });
  $$('.gift', el).forEach(inp => inp.oninput = () => { const v = inp.value === '' ? null : Math.max(0, Number(inp.value) || 0); lineOf(inp.dataset.gift).tang = v; inp.classList.toggle('has', v != null); tidy(inp.dataset.gift); requote(); });
  $('#cancel-edit')?.addEventListener('click', () => { draft = newDraft(); renderOrderForm(el); });
  $('#httm-add').onclick = () => { draft.httm.push({ sp: '', thung: '' }); renderHttm(); };
  renderNppInfo();
  renderHttm();
  requote();
}
function renderHttm() {
  const box = $('#httm-box');
  if (!box) return;
  const opts = S.products.filter(p => p.loai !== 'hang_tang');
  const giao = h => h.sp && h.thung ? fmt(h.thung * (M.sp[h.sp]?.quy_cach || 0)) + ' ' + esc((M.sp[h.sp]?.dvt || '').toLowerCase()) : '';
  box.innerHTML = draft.httm.length ? `<table><thead><tr><th>Mặt hàng</th><th class="n">SL thùng</th><th class="n">Tổng giao</th><th></th></tr></thead><tbody>
    ${draft.httm.map((h, i) => `<tr><td><select data-httm-sp="${i}" style="width:100%"><option value="">— Chọn hàng —</option>${opts.map(p => `<option value="${p.id}" ${p.id === h.sp ? 'selected' : ''}>${esc(p.ten)}</option>`).join('')}</select></td>
      <td class="n"><input type="number" min="0" step="0.5" class="qty ${h.thung ? 'has' : ''}" data-httm-q="${i}" value="${h.thung}"></td>
      <td class="n" data-httm-giao="${i}">${giao(h)}</td>
      <td><button class="btn danger sm" data-httm-del="${i}">Xóa</button></td></tr>`).join('')}</tbody></table>`
    : '<div class="muted small">Chưa có hàng hỗ trợ thị trường. Bấm "+ Thêm dòng" nếu muốn tặng thêm cho NPP.</div>';
  $$('[data-httm-sp]', box).forEach(s => s.onchange = () => { draft.httm[s.dataset.httmSp].sp = s.value; renderHttm(); requote(); });
  $$('[data-httm-q]', box).forEach(inp => inp.oninput = () => {
    const i = inp.dataset.httmQ, h = draft.httm[i];
    h.thung = Math.max(0, Number(inp.value) || 0) || '';
    inp.classList.toggle('has', !!h.thung);
    $(`[data-httm-giao="${i}"]`).innerHTML = giao(h);
    requote();
  });
  $$('[data-httm-del]', box).forEach(b => b.onclick = () => { draft.httm.splice(b.dataset.httmDel, 1); renderHttm(); requote(); });
}
function renderNppInfo() {
  const box = $('#npp-info');
  const n = M.npp[draft.npp_id];
  if (!box) return;
  if (!n) { box.innerHTML = ''; return; }
  const st = S.debt_npp[n.id] || {};
  const ck = draft.ck ?? n.ck, ckkm = draft.ck_km ?? S.settings.ck_km, cn = draft.cn_pct ?? (n.cn_pct ?? S.settings.cn_pct);
  box.innerHTML = `<div class="npp-card">
    <div class="row"><b>${esc(n.ten)}</b>${pill([n.khu_vuc || 'Chưa có vùng', 'info'])}</div>
    <div class="small muted">${esc(n.dia_chi)} · ${esc(n.sdt)} · Quản lý: <b>${esc(n.gsm || '—')}</b></div>
    <div class="small">Đang nợ <b>${fmt(st.tong_no)}</b> (công nợ chốt ${fmt(st.cn_chot_con)} · nợ đơn ${fmt(st.con_no_don)})</div>
    <div class="fields" style="margin-top:6px">
      <div class="field"><label for="o-ck">Chiết khấu thường (%) – nhập</label><input type="number" step="0.5" id="o-ck" value="${+(ck * 100).toFixed(2)}"></div>
      <div class="field"><label for="o-ckkm">CK KM khác – Sọc Đỏ/Sọc Phấn 520 (%)</label><input type="number" step="0.5" id="o-ckkm" value="${+(ckkm * 100).toFixed(2)}"></div>
      <div class="field"><label for="o-cn">Công nợ (%) in trên đơn</label><input type="number" step="5" id="o-cn" value="${+(cn * 100).toFixed(2)}"></div>
    </div>
    ${st.no_thang_truoc > 0.5 ? `<div class="alert bad">NPP còn <b>${fmt(st.no_thang_truoc)} đ</b> nợ đơn từ tháng ${monthLabel(st.no_tu_thang)} chưa trả. Cân nhắc đòi nợ trước khi giao đơn mới.</div>` : ''}</div>`;
  $('#o-ck').oninput = e => { draft.ck = (Number(e.target.value) || 0) / 100; requote(); };
  $('#o-ckkm').oninput = e => { draft.ck_km = (Number(e.target.value) || 0) / 100; requote(); };
  $('#o-cn').oninput = e => { draft.cn_pct = (Number(e.target.value) || 0) / 100; requote(); };
}
function draftBody() {
  return {
    id: draft.id, npp_id: draft.npp_id, ngay_dat: draft.ngay_dat, ghi_chu: draft.ghi_chu, ck: draft.ck, ck_km: draft.ck_km, cn_pct: draft.cn_pct,
    lines: Object.entries(draft.lines).filter(([, l]) => l.thung).map(([sp, l]) => ({ sp, thung: l.thung, gia_chai: l.gia_chai ?? null, tang: l.tang ?? null })),
    httm: draft.httm.filter(h => h.sp && h.thung).map(h => ({ sp: h.sp, thung: h.thung })),
  };
}
let qTimer = null;
function requote() {
  clearTimeout(qTimer);
  qTimer = setTimeout(async () => {
    $$('#prod-table [data-tt], #prod-table [data-giao]').forEach(td => td.textContent = '');
    $$('#prod-table [data-gift]').forEach(inp => { inp.placeholder = ''; });
    if (!draft.npp_id) { lastQuote = null; renderSummary(); return; }
    try { lastQuote = await api('/api/quote', draftBody()); } catch (e) { toast(e.message, true); return; }
    const gifts = Object.fromEntries(lastQuote.gifts.map(g => [g.cho, g]));
    for (const l of lastQuote.lines) {
      const tr = $(`#prod-table tr[data-row="${l.sp}"]`);
      if (!tr) continue;
      tr.querySelector('[data-tt]').textContent = fmt(l.thanh_tien);
      const g = gifts[l.sp];
      tr.querySelector('[data-gift]').placeholder = fq(l.tang_tu_dong) || '0';
      tr.querySelector('[data-giao]').innerHTML = fmt(l.tong_giao) + (g ? `<div class="small" style="color:var(--ok)">+${fmt(g.tang)} ${esc(g.ten.split(' - ')[0].replace('Nước chấm Nàng Cơm ', ''))}</div>` : '');
    }
    renderSummary();
  }, 120);
}
function renderSummary() {
  const box = $('#sum-box');
  if (!box) return;
  const q = lastQuote;
  if (!q || (!q.lines.length && !q.httm.length)) { box.innerHTML = `<div class="muted">${draft.npp_id ? 'Nhập số thùng cho các mặt hàng — tổng đơn hiện ở đây.' : 'Chọn NPP và nhập số thùng — tổng đơn hiện ở đây.'}</div>`; return; }
  const t = q.terms;
  const cell = (l, v, cls = '') => `<div class="sb-cell ${cls}"><span>${l}</span><b>${v}</b></div>`;
  box.innerHTML = `<div class="sb-row">
      ${cell('Tổng tiền hàng', fmt(q.tong_tien_hang))}
      ${cell(`CK thường ${pct(t.ck)}`, '− ' + fmt(q.ck_thuong))}
      ${q.ck_km ? cell(`CK KM khác ${pct(t.ck_km)}`, '− ' + fmt(q.ck_km)) : ''}
      ${cell('Sau chiết khấu', fmt(q.sau_ck))}
      ${cell(`Công nợ ${pct(t.cn_pct)}`, fmt(q.cong_no), 'cn')}
      ${cell('Tổng thanh toán', fmt(q.thanh_toan), 'big')}
      <div class="sb-btns"><button class="btn" id="save-order">${draft.id ? 'Lưu thay đổi' : 'Lưu đơn (chờ giao)'}</button>
        <button class="btn accent" id="save-export">Lưu &amp; xuất Đơn đặt hàng</button></div></div>
    <div class="small muted">${q.lines.length} mặt hàng · tổng giao ${fmt(q.tong_chai_giao)} chai/hũ${q.chai_httm ? ` (gồm ${fmt(q.chai_httm)} hỗ trợ thị trường)` : ''}${q.gifts.length ? ' · ' + q.gifts.map(g => `<b style="color:var(--ok)">tặng ${fmt(g.tang)} ${esc(g.dvt.toLowerCase())} ${esc(g.ten.split(' - ')[0])}</b>`).join(' · ') : ''} · giá & KM vùng ${esc(q.region)}</div>`;
  $('#save-order').onclick = () => saveOrder(false);
  $('#save-export').onclick = () => saveOrder(true);
}
async function saveOrder(exportToo) {
  const btns = $$('#sum-box .btn'); btns.forEach(b => b.disabled = true);
  try {
    const r = await run(() => api('/api/order/save', draftBody()));
    toast(`Đã lưu đơn ${r.order.ma} – chờ giao.`);
    if (exportToo) await exportOrder(r.order.id);
    draft = newDraft();
    await refresh();
  } finally { btns.forEach(b => b.disabled = false); }
}
async function exportOrder(id) {
  const t = toast('Đang điền mẫu Đơn đặt hàng bằng Excel… (khoảng 10 giây)');
  try {
    const r = await api('/api/order/export', { id });
    t.remove();
    const warn = Object.keys(r.lech || {}).length ? ' <b>(!) Số trên file lệch với công cụ, kiểm tra lại.</b>' : '';
    toast(`Đã xuất đơn. ${openFileLink(r.file, 'Mở Excel')} · ${openFileLink(r.pdf, 'Mở PDF')}${warn}`, false, true);
  } catch (e) { t.remove(); toast(e.message, true); }
}

// =====================================================================
// ĐƠN HÀNG
// =====================================================================
function renderOrders(el) {
  const f = filters.don;
  const months = [...new Set(S.orders.map(o => (o.trang_thai === 'lich_su' ? o.thang : o.ngay_dat || '').slice(0, 7)).filter(Boolean))].sort().reverse();
  const rows = S.orders.filter(o => (f.status === 'all' || o.trang_thai === f.status)
    && (!f.month || (o.trang_thai === 'lich_su' ? o.thang : o.ngay_dat).startsWith(f.month))
    && (!f.region || M.npp[o.npp_id]?.khu_vuc === f.region)
    && (!f.q || (M.npp[o.npp_id]?.ten || o.npp_ten || '').toLowerCase().includes(f.q.toLowerCase()) || (o.ma || '').toLowerCase().includes(f.q.toLowerCase())))
    .sort((a, b) => (b.ngay_dat || '').localeCompare(a.ngay_dat || '') || (b.ma || '').localeCompare(a.ma || ''));
  const counts = s => S.orders.filter(o => o.trang_thai === s).length;
  el.innerHTML = `
  <div class="card"><div class="card-b toolbar">
    <select id="o-status"><option value="all">Tất cả trạng thái</option>${['cho_giao', 'da_giao', 'huy', 'lich_su'].map(s => `<option value="${s}" ${f.status === s ? 'selected' : ''}>${STATUS[s][0]} (${counts(s)})</option>`).join('')}</select>
    <select id="o-month"><option value="">Mọi tháng</option>${months.map(m => `<option value="${m}" ${f.month === m ? 'selected' : ''}>Tháng ${monthLabel(m)}</option>`).join('')}</select>
    <select id="o-region"><option value="">Mọi vùng</option>${S.regions.map(r => `<option ${f.region === r ? 'selected' : ''}>${r}</option>`).join('')}</select>
    <input id="o-q" class="grow" placeholder="Tìm theo tên NPP hoặc mã đơn…" value="${esc(f.q)}">
    <a class="btn" href="#len-don">+ Lên đơn mới</a></div></div>
  <div class="card"><div class="tbl-wrap"><table><thead><tr><th>Mã đơn</th><th>Ngày đặt</th><th>Ngày giao</th><th>NPP</th><th>Vùng</th><th class="n">Sau CK</th><th class="n">Đã thu</th><th class="n">Còn thu</th><th class="n">HTTM</th><th>Trạng thái</th><th>Thao tác</th></tr></thead><tbody>
  ${rows.length ? rows.map(o => {
    const n = M.npp[o.npp_id];
    const paid = o.trang_thai === 'lich_su' ? o.da_tt : (M.paid[o.id] || 0);
    const left = o.trang_thai === 'lich_su' ? o.chua_tt : o.trang_thai === 'da_giao' ? o.sau_ck - paid : 0;
    const httm = sumBy(o.httm || [], 'tong_giao');
    const acts = o.trang_thai === 'lich_su' ? '' : [
      `<button class="btn ghost sm" data-view="${o.id}">Xem</button>`,
      o.trang_thai === 'cho_giao' ? `<button class="btn sm" data-deliver="${o.id}">Đã giao</button><button class="btn ghost sm" data-edit="${o.id}">Sửa</button>` : '',
      o.trang_thai !== 'huy' ? `<button class="btn ghost sm" data-export="${o.id}">Xuất file</button>` : '',
      o.trang_thai === 'cho_giao' ? `<button class="btn danger sm" data-cancel="${o.id}">Hủy</button>` : '',
      o.trang_thai === 'da_giao' ? `<button class="btn ghost sm" data-undeliver="${o.id}">Hoàn tác giao</button>` : ''].join('');
    return `<tr><td class="nowrap"><b>${esc(o.ma)}</b></td><td>${fd(o.ngay_dat)}</td><td>${fd(o.ngay_giao)}</td><td>${esc(n?.ten || o.npp_ten)}<div class="small muted">${esc(n?.gsm || '')}</div></td><td class="small">${esc(n?.khu_vuc || '')}</td>
      <td class="n">${fmt(o.sau_ck)}</td><td class="n">${fmt(paid)}</td><td class="n">${left > 0.5 ? `<b>${fmt(left)}</b>` : '–'}</td><td class="n small">${httm ? fmt(httm) + ' chai' : ''}</td><td>${pill(STATUS[o.trang_thai])}</td><td><div class="actions">${acts}</div></td></tr>`;
  }).join('') : '<tr><td colspan="11" class="empty">Không có đơn nào khớp bộ lọc.</td></tr>'}
  </tbody></table></div></div>`;
  $('#o-status').onchange = e => { f.status = e.target.value; renderOrders(el); };
  $('#o-month').onchange = e => { f.month = e.target.value; renderOrders(el); };
  $('#o-region').onchange = e => { f.region = e.target.value; renderOrders(el); };
  $('#o-q').oninput = e => { f.q = e.target.value; clearTimeout(qTimer); qTimer = setTimeout(() => { renderOrders(el); const i = $('#o-q'); i.focus(); i.setSelectionRange(i.value.length, i.value.length); }, 250); };
  bindCommon(el);
}
function bindCommon(el) {
  $$('[data-deliver]', el).forEach(b => b.onclick = () => deliverDialog(b.dataset.deliver));
  $$('[data-undeliver]', el).forEach(b => b.onclick = async () => { await run(() => api('/api/order/undeliver', { id: b.dataset.undeliver }), 'Đã đưa đơn về trạng thái chờ giao.'); refresh(); });
  $$('[data-cancel]', el).forEach(b => b.onclick = () => confirmDialog('Hủy đơn này?', 'Đơn bị hủy sẽ không vào sổ doanh thu và công nợ.', async () => { await run(() => api('/api/order/cancel', { id: b.dataset.cancel }), 'Đã hủy đơn.'); refresh(); }));
  $$('[data-export]', el).forEach(b => b.onclick = () => exportOrder(b.dataset.export));
  $$('[data-edit]', el).forEach(b => b.onclick = () => editOrder(b.dataset.edit));
  $$('[data-view]', el).forEach(b => b.onclick = () => viewOrder(b.dataset.view));
  $$('[data-pay]', el).forEach(b => b.onclick = e => { e.preventDefault(); paymentDialog(b.dataset.pay); });
  $$('[data-zalo]', el).forEach(b => b.onclick = e => { e.preventDefault(); showZalo(b.dataset.zalo); });
  $$('[data-npp-detail]', el).forEach(b => b.onclick = () => nppDebtDetail(b.dataset.nppDetail));
}
function deliverDialog(id) {
  const o = S.orders.find(x => x.id === id), n = M.npp[o.npp_id];
  modal(`Xác nhận đã giao đơn ${o.ma}`, `
    <p><b>${esc(n.ten)}</b> · sau chiết khấu <b>${fmt(o.sau_ck)} đ</b>. Đơn sẽ vào <b>sổ doanh thu</b> và được tính vào công nợ của NPP cho tới khi NPP trả đủ (nhắc mỗi cuối tháng).</p>
    <div class="field" style="max-width:220px"><label for="d-date">Ngày giao hàng</label><input type="date" id="d-date" value="${S.today}"></div>`,
    `<button class="btn ghost" data-close2>Để sau</button><button class="btn" id="d-ok">Xác nhận đã giao</button>`);
  $('#d-ok').onclick = async () => { const ngay = $('#d-date').value; closeModal(); await run(() => api('/api/order/deliver', { id, ngay_giao: ngay }), `Đơn ${o.ma} đã vào sổ doanh thu.`); refresh(); };
}
function editOrder(id) {
  const o = S.orders.find(x => x.id === id);
  draft = { id: o.id, npp_id: o.npp_id, ngay_dat: o.ngay_dat, ghi_chu: o.ghi_chu || '', ck: o.terms.ck, ck_km: o.terms.ck_km, cn_pct: o.terms.cn_pct, region: '',
    lines: Object.fromEntries(o.lines.map(l => {
      const g = o.gifts.find(x => x.cho === l.sp);
      return [l.sp, { thung: l.thung, gia_chai: l.gia_tay ? l.gia_chai : null, tang: g ? (g.tang_tay ? g.tang : null) : (l.tang_tay ? l.tang : null) }];
    })),
    httm: (o.httm || []).map(h => ({ sp: h.sp, thung: h.thung })) };
  if (location.hash === '#len-don') route(); else location.hash = '#len-don';
}
function viewOrder(id) {
  const o = S.orders.find(x => x.id === id), n = M.npp[o.npp_id], t = o.terms;
  const paid = M.paid[o.id] || 0;
  modal(`Đơn ${o.ma} – ${n.ten}`, `
    <div class="fields small"><div><span class="muted">Ngày đặt</span><br><b>${fd(o.ngay_dat)}</b></div><div><span class="muted">Ngày giao</span><br><b>${fd(o.ngay_giao) || '—'}</b></div>
      <div><span class="muted">Trạng thái</span><br>${pill(STATUS[o.trang_thai])}</div><div><span class="muted">Quản lý</span><br><b>${esc(n.gsm)}</b></div></div>
    ${o.ghi_chu ? `<div class="alert warn">${esc(o.ghi_chu)}</div>` : ''}
    <div class="tbl-wrap"><table><thead><tr><th>Mặt hàng</th><th class="n">Thùng</th><th class="n">Giá chai</th><th class="n">Thành tiền</th><th class="n">Tặng</th><th class="n">Tổng giao</th></tr></thead><tbody>
      ${o.lines.map(l => `<tr><td>${esc(l.ten)}</td><td class="n">${fq(l.thung)}</td><td class="n">${fmt(l.gia_chai)}${l.gia_tay ? ' <span class="pill warn">sửa tay</span>' : ''}</td><td class="n">${fmt(l.thanh_tien)}</td><td class="n">${l.tang ? fmt(l.tang) : ''}</td><td class="n">${fmt(l.tong_giao)}</td></tr>`).join('')}
      ${o.gifts.map(g => `<tr><td>${esc(g.ten)} <span class="pill ok">Hàng tặng KM</span></td><td class="n">${fq(g.thung)}</td><td></td><td></td><td class="n">${fmt(g.tang)}</td><td class="n">${fmt(g.tong_giao)}</td></tr>`).join('')}
      ${(o.httm || []).map(h => `<tr><td>${esc(h.ten)} <span class="pill info">Hỗ trợ thị trường</span></td><td class="n">${fq(h.thung)}</td><td></td><td class="n">–</td><td></td><td class="n">${fmt(h.tong_giao)}</td></tr>`).join('')}
    </tbody><tfoot><tr><td colspan="3">Tổng tiền hàng</td><td class="n">${fmt(o.tong_tien_hang)}</td><td></td><td class="n">${fmt(o.tong_chai_giao)}</td></tr></tfoot></table></div>
    <div class="grid g2"><div>
      <div class="sum-row"><span>Chiết khấu thường (${pct(t.ck)})</span><span>− ${fmt(o.ck_thuong)}</span></div>
      <div class="sum-row"><span>CK KM khác (${pct(t.ck_km)})</span><span>− ${fmt(o.ck_km)}</span></div>
      <div class="sum-row"><span>Sau chiết khấu</span><b>${fmt(o.sau_ck)}</b></div>
      <div class="sum-row cn"><span>Công nợ (${pct(t.cn_pct)})</span><span>${fmt(o.cong_no)}</span></div>
      <div class="sum-row big"><span>Tổng thanh toán</span><span>${fmt(o.thanh_toan)}</span></div></div>
    <div>${o.trang_thai === 'da_giao' ? `<div class="sum-row"><span>NPP đã trả cho đơn này</span><b>${fmt(paid)}</b></div><div class="sum-row"><span>Còn phải thu</span><b style="color:${o.sau_ck - paid > 0.5 ? 'var(--bad)' : 'var(--ok)'}">${fmt(o.sau_ck - paid)}</b></div>`
      : '<div class="muted small">Đơn tính vào công nợ khi xác nhận đã giao.</div>'}</div></div>`,
    `<button class="btn ghost" data-close2>Đóng</button>${o.trang_thai !== 'huy' ? `<button class="btn accent" id="v-export">Xuất Đơn đặt hàng</button>` : ''}`, true);
  $('#v-export')?.addEventListener('click', () => exportOrder(o.id));
}

// =====================================================================
// CÔNG NỢ
// =====================================================================
function renderDebts(el) {
  const f = filters.cn;
  const gsms = [...new Set(S.npps.map(n => n.gsm).filter(Boolean))].sort();
  const t = debtTotals();
  const rows = S.npps.map(n => ({ n, st: S.debt_npp[n.id] })).filter(({ n, st }) => (!f.region || n.khu_vuc === f.region) && (!f.gsm || n.gsm === f.gsm) && (!f.owing || st.tong_no > 0.5))
    .sort((a, b) => b.st.no_thang_truoc - a.st.no_thang_truoc || b.st.tong_no - a.st.tong_no);
  const tot = k => rows.reduce((a, r) => a + r.st[k], 0);
  el.innerHTML = `
  <div class="kpis">
    <div class="kpi main"><div class="l">Tổng còn nợ</div><div class="v">${fmt(t.chot + t.don)}</div><div class="s">${t.n} NPP còn nợ</div></div>
    <div class="kpi"><div class="l">Công nợ chốt còn</div><div class="v">${fmt(t.chot)}</div><div class="s">số dư cũ từ file "Công nợ chốt"</div></div>
    <div class="kpi bad"><div class="l">Nợ đơn tồn từ tháng trước</div><div class="v">${fmt(t.truoc)}</div></div>
    <div class="kpi warn"><div class="l">Nợ đơn trong tháng</div><div class="v">${fmt(t.trong)}</div></div>
  </div>
  <div class="card"><div class="card-b toolbar">
    <select id="c-region"><option value="">Mọi vùng</option>${S.regions.map(r => `<option ${f.region === r ? 'selected' : ''}>${r}</option>`).join('')}</select>
    <select id="c-gsm"><option value="">Mọi quản lý (GSM)</option>${gsms.map(g => `<option ${f.gsm === g ? 'selected' : ''}>${esc(g)}</option>`).join('')}</select>
    <label class="small"><input type="checkbox" id="c-owing" ${f.owing ? 'checked' : ''}> Chỉ NPP còn nợ</label>
    <span class="grow"></span><button class="btn" id="c-pay">+ Ghi nhận NPP thanh toán</button></div></div>
  <div class="card"><div class="tbl-wrap"><table><thead><tr><th>NPP</th><th>Vùng</th><th>Quản lý</th><th class="n">Công nợ chốt còn</th><th class="n">Nợ đơn tháng trước</th><th class="n">Nợ đơn trong tháng</th><th class="n">Tổng còn nợ</th><th>Nợ đơn từ</th><th></th></tr></thead><tbody>
  ${rows.length ? rows.map(({ n, st }) => `<tr><td><b>${esc(n.ten)}</b><div class="small muted">${esc(n.tinh)} · ${esc(n.sdt)}</div></td><td class="small">${esc(n.khu_vuc)}</td><td class="small">${esc(n.gsm)}</td>
    <td class="n">${fmt(st.cn_chot_con)}</td><td class="n">${st.no_thang_truoc > 0.5 ? `<b style="color:var(--bad)">${fmt(st.no_thang_truoc)}</b>` : '–'}</td>
    <td class="n">${st.no_trong_thang > 0.5 ? fmt(st.no_trong_thang) : '–'}</td>
    <td class="n"><b>${fmt(st.tong_no)}</b>${st.du > 0.5 ? `<div class="small" style="color:var(--ok)">trả dư ${fmt(st.du)}</div>` : ''}</td>
    <td class="nowrap small">${st.no_tu_thang ? 'tháng ' + monthLabel(st.no_tu_thang) : ''}</td>
    <td><div class="actions"><button class="btn sm" data-pay="${n.id}">Thu tiền</button><button class="btn ghost sm" data-npp-detail="${n.id}">Chi tiết</button></div></td></tr>`).join('')
    : '<tr><td colspan="9" class="empty">Không có NPP nào khớp bộ lọc.</td></tr>'}
  </tbody><tfoot><tr><td colspan="3">Tổng (${rows.length} NPP)</td><td class="n">${fmt(tot('cn_chot_con'))}</td><td class="n">${fmt(tot('no_thang_truoc'))}</td><td class="n">${fmt(tot('no_trong_thang'))}</td><td class="n">${fmt(tot('tong_no'))}</td><td colspan="2"></td></tr></tfoot></table></div></div>`;
  $('#c-region').onchange = e => { f.region = e.target.value; renderDebts(el); };
  $('#c-gsm').onchange = e => { f.gsm = e.target.value; renderDebts(el); };
  $('#c-owing').onchange = e => { f.owing = e.target.checked; renderDebts(el); };
  $('#c-pay').onclick = () => paymentDialog('');
  bindCommon(el);
}
function paymentDialog(nid) {
  const opts = S.regions.map(r => `<optgroup label="${r}">${S.npps.filter(n => n.khu_vuc === r).map(n => `<option value="${n.id}" ${n.id === nid ? 'selected' : ''}>${esc(n.ten)}</option>`).join('')}</optgroup>`).join('');
  modal('Ghi nhận NPP thanh toán', `
    <div class="fields">
      <div class="field" style="grid-column:span 2"><label for="p-npp">Nhà phân phối</label><select id="p-npp"><option value="">— Chọn NPP —</option>${opts}</select></div>
      <div class="field"><label for="p-date">Ngày nhận tiền</label><input type="date" id="p-date" value="${S.today}"></div>
      <div class="field"><label for="p-amt">Số tiền (đ)</label><input id="p-amt" inputmode="numeric" placeholder="VD: 15.000.000"></div>
    </div>
    <div id="p-info"></div>
    <div class="field"><label>Trừ vào</label>
      <label class="small"><input type="radio" name="p-for" value="don" checked> Nợ theo đơn (trừ đơn giao sớm nhất trước; còn dư thì trừ tiếp công nợ chốt)</label>
      <label class="small"><input type="radio" name="p-for" value="cn_chot"> Công nợ chốt (số dư cũ)</label></div>
    <div class="field"><label for="p-note">Ghi chú</label><input id="p-note" placeholder="VD: chuyển khoản ngân hàng"></div>`,
    `<button class="btn ghost" data-close2>Hủy</button><button class="btn" id="p-ok">Lưu thanh toán</button>`);
  const info = () => { const st = S.debt_npp[$('#p-npp').value]; $('#p-info').innerHTML = st ? `<div class="alert ${st.no_thang_truoc > 0.5 ? 'bad' : 'ok'} small">Đang nợ <b>${fmt(st.tong_no)}</b> · công nợ chốt ${fmt(st.cn_chot_con)} · nợ đơn ${fmt(st.con_no_don)}</div>` : ''; };
  $('#p-npp').onchange = info; info();
  $('#p-amt').oninput = e => { const v = parseMoney(e.target.value); e.target.value = v ? fmt(v) : ''; };
  $('#p-ok').onclick = async () => {
    const body = { npp_id: $('#p-npp').value, ngay: $('#p-date').value, so_tien: parseMoney($('#p-amt').value), tra_cho: $('input[name=p-for]:checked').value, ghi_chu: $('#p-note').value };
    if (!body.npp_id) return toast('Chọn NPP.', true);
    await run(() => api('/api/payment/add', body), `Đã ghi nhận ${fmt(body.so_tien)} đ.`);
    closeModal(); refresh();
  };
}
function nppDebtDetail(nid) {
  const n = M.npp[nid], st = S.debt_npp[nid];
  const items = S.debt_items.filter(i => i.npp_id === nid).sort((a, b) => b.ngay_giao.localeCompare(a.ngay_giao));
  const pays = S.payments.filter(p => p.npp_id === nid).sort((a, b) => b.ngay.localeCompare(a.ngay));
  modal(`Công nợ – ${n.ten}`, `
    <div class="kpis"><div class="kpi main"><div class="l">Tổng còn nợ</div><div class="v">${fmt(st.tong_no)}</div></div>
      <div class="kpi"><div class="l">Công nợ chốt còn</div><div class="v">${fmt(st.cn_chot_con)}</div><div class="s">đầu kỳ ${fmt(st.cn_chot_dau)}${st.cn_chot_tu_du > 0.5 ? ` · đã trừ ${fmt(st.cn_chot_tu_du)} từ tiền trả dư` : ''}</div></div>
      <div class="kpi"><div class="l">Nợ theo đơn</div><div class="v">${fmt(st.con_no_don)}</div></div></div>
    <b>Đơn đã giao</b>
    <div class="tbl-wrap">${items.length ? `<table><thead><tr><th>Đơn</th><th>Ngày giao</th><th class="n">Sau CK</th><th class="n">Đã trả</th><th class="n">Còn lại</th><th></th></tr></thead><tbody>
      ${items.map(i => `<tr><td>${esc(i.ma)}</td><td>${fd(i.ngay_giao)}</td><td class="n">${fmt(i.so_tien)}</td><td class="n">${fmt(i.da_tra)}</td><td class="n"><b>${fmt(i.con_lai)}</b></td>
        <td>${i.con_lai <= 0.5 ? pill(['Đã trả đủ', 'ok']) : i.tu_thang_truoc ? pill(['Tồn từ tháng trước', 'bad']) : pill(['Trong tháng', 'warn'])}</td></tr>`).join('')}</tbody></table>` : '<div class="empty">Chưa có đơn nào đã giao trong công cụ.</div>'}</div>
    <b>Lịch sử thanh toán</b>
    <div class="tbl-wrap">${pays.length ? `<table><thead><tr><th>Ngày</th><th>Trừ vào</th><th class="n">Số tiền</th><th>Ghi chú</th><th></th></tr></thead><tbody>
      ${pays.map(p => `<tr><td>${fd(p.ngay)}</td><td>${p.tra_cho === 'cn_chot' ? 'Công nợ chốt' : 'Nợ theo đơn'}</td><td class="n">${fmt(p.so_tien)}</td><td>${esc(p.ghi_chu)}</td><td><button class="btn danger sm" data-delpay="${p.id}">Xóa</button></td></tr>`).join('')}</tbody></table>` : '<div class="empty">Chưa ghi nhận thanh toán nào.</div>'}</div>`,
    `<button class="btn ghost" data-close2>Đóng</button><button class="btn" id="dd-pay">Thu tiền</button>`, true);
  $('#dd-pay').onclick = () => paymentDialog(nid);
  $$('[data-delpay]').forEach(b => b.onclick = () => confirmDialog('Xóa lần thanh toán này?', 'Số tiền sẽ được tính lại vào công nợ.', async () => { await run(() => api('/api/payment/delete', { id: b.dataset.delpay }), 'Đã xóa.'); refresh(); }));
}

// =====================================================================
// SỔ DOANH THU
// =====================================================================
function renderLedger(el) {
  const f = filters.so;
  const months = [...new Set(S.orders.filter(isRevenue).map(orderMonth).filter(Boolean))].sort().reverse();
  if (!f.month || !months.includes(f.month)) f.month = months.includes(S.today.slice(0, 7)) ? S.today.slice(0, 7) : months[0] || '';
  const rows = S.orders.filter(o => isRevenue(o) && orderMonth(o) === f.month && (!f.region || M.npp[o.npp_id]?.khu_vuc === f.region))
    .sort((a, b) => (S.regions.indexOf(M.npp[a.npp_id]?.khu_vuc) - S.regions.indexOf(M.npp[b.npp_id]?.khu_vuc)) || (a.ngay_giao || '').localeCompare(b.ngay_giao || ''));
  const val = o => { const paid = o.trang_thai === 'lich_su' ? o.da_tt : (M.paid[o.id] || 0); return { paid, left: o.trang_thai === 'lich_su' ? o.chua_tt : o.sau_ck - paid }; };
  const byReg = {};
  for (const o of rows) { const r = M.npp[o.npp_id]?.khu_vuc || 'Khác'; byReg[r] = (byReg[r] || 0) + o.sau_ck; }
  const sum = k => rows.reduce((a, o) => a + (o[k] || 0), 0);
  el.innerHTML = `
  <div class="card"><div class="card-b toolbar">
    <select id="s-month">${months.map(m => `<option value="${m}" ${m === f.month ? 'selected' : ''}>Tháng ${monthLabel(m)}</option>`).join('')}</select>
    <select id="s-region"><option value="">Mọi vùng</option>${S.regions.map(r => `<option ${f.region === r ? 'selected' : ''}>${r}</option>`).join('')}</select>
    <span class="grow"></span><button class="btn" id="s-export">Xuất sổ Excel (đúng mẫu theo dõi)</button></div></div>
  <div class="kpis">${Object.entries(byReg).map(([r, v]) => `<div class="kpi"><div class="l"><i style="display:inline-block;width:9px;height:9px;border-radius:2px;background:${REG_COLORS[r] || '#98A2B3'}"></i> ${esc(r)}</div><div class="v">${fmt(v)}</div><div class="s">${rows.filter(o => (M.npp[o.npp_id]?.khu_vuc || 'Khác') === r).length} đơn</div></div>`).join('') || '<div class="muted">Tháng này chưa có đơn đã giao.</div>'}</div>
  <div class="card"><div class="tbl-wrap"><table><thead><tr><th>Vùng</th><th>GSM</th><th>Tỉnh</th><th>Tên NPP</th><th class="n">CK</th><th>Ngày đặt</th><th>Ngày giao</th><th class="n">Trước CK</th><th class="n">Chiết khấu</th><th class="n">Sau CK</th><th class="n">Đã TT</th><th class="n">Chưa TT</th><th class="n">HTTM</th><th>Ghi chú</th></tr></thead><tbody>
  ${rows.length ? rows.map(o => { const n = M.npp[o.npp_id] || {}; const v = val(o); const h = sumBy(o.httm || [], 'tong_giao'); return `<tr><td class="small">${esc(n.khu_vuc || '')}</td><td class="small">${esc(n.gsm || '')}</td><td class="small">${esc(n.tinh || '')}</td><td><b>${esc(n.ten || o.npp_ten)}</b></td>
    <td class="n">${pct(o.terms?.ck)}</td><td>${fd(o.ngay_dat)}</td><td>${fd(o.ngay_giao)}</td><td class="n">${fmt(o.truoc_ck)}</td><td class="n">${fmt(o.chiet_khau)}</td><td class="n"><b>${fmt(o.sau_ck)}</b></td>
    <td class="n">${fmt(v.paid)}</td><td class="n">${v.left > 0.5 ? `<b style="color:var(--bad)">${fmt(v.left)}</b>` : '–'}</td><td class="n small">${h ? fmt(h) + ' chai' : ''}</td>
    <td class="small">${o.trang_thai === 'lich_su' ? '<span class="pill info">file cũ</span> ' : esc(o.ma) + ' '}${esc(o.ghi_chu || '')}</td></tr>`; }).join('')
    : '<tr><td colspan="14" class="empty">Không có dữ liệu.</td></tr>'}
  </tbody><tfoot><tr><td colspan="7">Tổng tháng ${monthLabel(f.month)} (${rows.length} đơn)</td><td class="n">${fmt(sum('truoc_ck'))}</td><td class="n">${fmt(sum('chiet_khau'))}</td><td class="n">${fmt(sum('sau_ck'))}</td>
    <td class="n">${fmt(rows.reduce((a, o) => a + val(o).paid, 0))}</td><td class="n">${fmt(rows.reduce((a, o) => a + Math.max(0, val(o).left), 0))}</td><td></td><td></td></tr></tfoot></table></div></div>
  <div class="small muted">Đơn "file cũ" lấy từ sheet tháng trong file "Doanh thu, công nợ NPP 2026" – chỉ để xem doanh thu, không tính vào công nợ (công nợ cũ đã nằm trong "Công nợ chốt").</div>`;
  $('#s-month').onchange = e => { f.month = e.target.value; renderLedger(el); };
  $('#s-region').onchange = e => { f.region = e.target.value; renderLedger(el); };
  $('#s-export').onclick = () => $('#btn-ledger').click();
}

// =====================================================================
// DANH MỤC
// =====================================================================
function renderCatalog(el) {
  const tab = filters.dm;
  el.innerHTML = `<div class="card"><div class="tabs" role="tablist">
      <button data-tab="npp" class="${tab === 'npp' ? 'active' : ''}">Nhà phân phối (${S.npps.length})</button>
      <button data-tab="gia" class="${tab === 'gia' ? 'active' : ''}">Bảng giá &amp; KM theo vùng</button>
      <button data-tab="sp" class="${tab === 'sp' ? 'active' : ''}">Sản phẩm</button>
      <button data-tab="cd" class="${tab === 'cd' ? 'active' : ''}">Cài đặt</button></div><div id="tab-body"></div></div>`;
  $$('[data-tab]', el).forEach(b => b.onclick = () => { filters.dm = b.dataset.tab; renderCatalog(el); });
  const body = $('#tab-body');
  if (tab === 'npp') {
    body.innerHTML = `<div class="card-b toolbar"><span class="muted small">CK = chiết khấu thường gợi ý sẵn khi lên đơn (nhập lại được trên từng đơn). Giá & KM lấy theo vùng của NPP.</span><span class="grow"></span><button class="btn" id="npp-add">+ Thêm NPP</button></div>
    <div class="tbl-wrap"><table><thead><tr><th>Vùng</th><th>GSM</th><th>Tỉnh</th><th>NPP</th><th>SĐT</th><th class="n">CK</th><th class="n">Công nợ %</th><th class="n">CN chốt đầu kỳ</th><th>Tình trạng</th><th></th></tr></thead><tbody>
    ${[...S.npps].sort((a, b) => S.regions.indexOf(a.khu_vuc) - S.regions.indexOf(b.khu_vuc) || a.ten.localeCompare(b.ten)).map(n => `<tr><td class="small">${esc(n.khu_vuc)}</td><td class="small">${esc(n.gsm)}</td><td class="small">${esc(n.tinh)}</td><td><b>${esc(n.ten)}</b></td><td class="small">${esc(n.sdt)}</td>
      <td class="n">${pct(n.ck)}</td><td class="n">${n.cn_pct == null ? '<span class="muted">mặc định</span>' : pct(n.cn_pct)}</td>
      <td class="n">${fmt(n.cn_chot)}</td><td>${pill([n.tinh_trang, activeNpp(n) ? 'ok' : 'mute'])}</td><td><button class="btn ghost sm" data-npp="${n.id}">Sửa</button></td></tr>`).join('')}</tbody></table></div>`;
    $('#npp-add').onclick = () => nppDialog(null);
    $$('[data-npp]', body).forEach(b => b.onclick = () => nppDialog(b.dataset.npp));
  } else if (tab === 'gia') {
    renderPriceTab(body);
  } else if (tab === 'sp') {
    body.innerHTML = `<div class="card-b small muted">Danh mục sản phẩm (lấy từ file mẫu đơn đặt hàng). Giá và khuyến mãi sửa ở tab "Bảng giá & KM theo vùng".</div>
    <div class="tbl-wrap"><table><thead><tr><th>Nhãn</th><th>Tên hàng</th><th>ĐVT</th><th class="n">Quy cách</th><th>Nhóm CK</th><th>Đang bán</th><th></th></tr></thead><tbody>
    ${S.products.map(p => `<tr><td class="small">${esc(p.nhan)}</td><td>${esc(p.ten)}</td><td class="small">${esc(p.dvt)}</td><td class="n">${p.quy_cach}</td><td class="small">${p.nhom_ck === 'km' ? 'KM khác' : 'Thường'}</td>
      <td>${p.loai === 'hang_tang' ? '<span class="muted small">dòng hàng tặng</span>' : p.dang_ban ? '✓' : '–'}</td><td>${p.loai === 'hang_tang' ? '' : `<button class="btn ghost sm" data-sp="${p.id}">Sửa</button>`}</td></tr>`).join('')}</tbody></table></div>`;
    $$('[data-sp]', body).forEach(b => b.onclick = () => productDialog(b.dataset.sp));
  } else {
    const s = S.settings;
    body.innerHTML = `<div class="card-b grid" style="max-width:720px">
      <div class="fields">
        <div class="field"><label for="s-nhac">Nhắc nợ cuối tháng bắt đầu từ ngày</label><input type="number" min="1" max="31" id="s-nhac" value="${s.nhac_tu_ngay}"></div>
        <div class="field"><label for="s-ckkm">CK KM khác gợi ý (%)</label><input type="number" id="s-ckkm" value="${+(s.ck_km * 100).toFixed(2)}"></div>
        <div class="field"><label for="s-cn">Công nợ % gợi ý trên đơn</label><input type="number" id="s-cn" value="${+(s.cn_pct * 100).toFixed(2)}"></div>
      </div>
      <div class="field"><label for="s-ck">Thông tin chuyển khoản (in trong tin nhắn Zalo nhắc nợ)</label><input id="s-ck" value="${esc(s.chuyen_khoan || '')}" placeholder="VD: STK … – Ngân hàng … – Tên đơn vị nhận tiền">
        <span class="small muted">Chỉ lưu trong file dữ liệu trên máy này, không nằm trong mã nguồn.</span></div>
      <div class="small muted">Từ ngày nhắc tới hết tháng, trang Tổng quan báo đỏ và liệt kê mọi NPP còn nợ (công nợ chốt + nợ theo đơn), chia theo GSM. Tháng nào cũng nhắc cho tới khi NPP trả hết.</div>
      <div><button class="btn" id="s-save">Lưu cài đặt</button></div></div>`;
    $('#s-save').onclick = async () => {
      await run(() => api('/api/settings/save', { nhac_tu_ngay: $('#s-nhac').value, cn_pct: (Number($('#s-cn').value) || 0) / 100, ck_km: (Number($('#s-ckkm').value) || 0) / 100, chuyen_khoan: $('#s-ck').value }), 'Đã lưu cài đặt.');
      refresh();
    };
  }
}
function renderPriceTab(body) {
  const reg = filters.gia;
  const price = S.prices[reg];
  body.innerHTML = `<div class="card-b toolbar">
      <label class="small" for="g-reg">Vùng</label><select id="g-reg">${S.regions.map(r => `<option ${r === reg ? 'selected' : ''}>${r}</option>`).join('')}</select>
      <span class="muted small">KM "mỗi N thùng tặng M": chỉ tặng khi đủ trọn N thùng (sửa được số tặng trên từng đơn).</span><span class="grow"></span>
      <select id="g-from">${S.regions.filter(r => r !== reg).map(r => `<option>${r}</option>`).join('')}</select><button class="btn ghost sm" id="g-copy">Chép bảng giá từ vùng này</button>
      <button class="btn" id="g-save">Lưu bảng giá ${esc(reg)}</button></div>
    <div class="tbl-wrap"><table id="price-table"><thead><tr><th>Tên hàng</th><th class="n">QC</th><th class="n">Giá chai</th><th class="n">Giá thùng</th><th>Chương trình KM (in trên đơn)</th><th class="n">Mỗi N thùng</th><th class="n">Tặng M</th><th>Quy tắc</th></tr></thead><tbody>
    ${S.products.filter(p => p.loai !== 'hang_tang').map(p => { const r = price[p.id] || {}; const other = r.km_loai === 'khac'; return `<tr data-p="${p.id}"><td>${esc(p.ten)}</td><td class="n">${p.quy_cach}</td>
      <td class="n"><input type="number" class="price" data-k="gia_chai" value="${r.gia_chai ?? ''}"></td><td class="n" data-gt>${fmt((r.gia_chai || 0) * p.quy_cach)}</td>
      <td><input data-k="km_text" value="${esc(r.km_text || '')}" style="width:100%"></td>
      ${other ? `<td class="n" colspan="2"><input type="number" class="qty" data-k="km_moi_thung_tang" value="${r.km_moi_thung_tang ?? ''}" title="Mỗi thùng tặng bao nhiêu chai ${esc(M.sp[p.km_sp]?.ten || '')}"> <span class="small muted">chai/thùng</span></td>`
        : `<td class="n"><input type="number" class="qty" data-k="km_n" value="${r.km_n ?? ''}" placeholder="–"></td><td class="n"><input type="number" class="qty" data-k="km_m" value="${r.km_m ?? ''}" placeholder="–"></td>`}
      <td class="small">${kmRule(p, r)}</td></tr>`; }).join('')}</tbody></table></div>`;
  $('#g-reg').onchange = e => { filters.gia = e.target.value; renderPriceTab(body); };
  $$('#price-table [data-k="gia_chai"]').forEach(inp => inp.oninput = () => { const tr = inp.closest('tr'); tr.querySelector('[data-gt]').textContent = fmt((Number(inp.value) || 0) * M.sp[tr.dataset.p].quy_cach); });
  $('#g-save').onclick = async () => {
    const rows = $$('#price-table tr[data-p]').map(tr => Object.fromEntries([['sp', tr.dataset.p], ...$$('[data-k]', tr).map(i => [i.dataset.k, i.value])]));
    await run(() => api('/api/prices/save', { region: reg, rows }), `Đã lưu bảng giá ${reg}.`);
    refresh();
  };
  $('#g-copy').onclick = () => { const from = $('#g-from').value; confirmDialog(`Chép bảng giá ${from} sang ${reg}?`, `Toàn bộ giá & khuyến mãi của ${reg} sẽ được thay bằng của ${from}.`, async () => { await run(() => api('/api/prices/copy', { from, to: reg }), 'Đã chép bảng giá.'); refresh(); }); };
}
function nppDialog(id) {
  const n = id ? M.npp[id] : { ten: '', khu_vuc: S.regions[0], gsm: '', asm: '', tinh: '', dia_chi: '', sdt: '', tinh_trang: 'Đang hoạt động', ck: 0, cn_pct: null, cn_chot: 0, ghi_chu: '' };
  const fld = (k, label, type = 'text', v = n[k]) => `<div class="field"><label for="n-${k}">${label}</label><input id="n-${k}" type="${type}" value="${esc(v ?? '')}"></div>`;
  modal(id ? `Sửa NPP – ${n.ten}` : 'Thêm NPP', `<div class="fields">
    ${fld('ten', 'Tên NPP')}<div class="field"><label for="n-khu_vuc">Vùng miền (quyết định bảng giá)</label><select id="n-khu_vuc">${S.regions.map(r => `<option ${r === n.khu_vuc ? 'selected' : ''}>${r}</option>`).join('')}</select></div>
    ${fld('gsm', 'GSM (quản lý)')}${fld('asm', 'ASM')}${fld('tinh', 'Tỉnh')}${fld('sdt', 'SĐT')}
    <div class="field" style="grid-column:1/-1"><label for="n-dia_chi">Địa chỉ</label><input id="n-dia_chi" value="${esc(n.dia_chi)}"></div>
    ${fld('ck', 'Chiết khấu thường gợi ý (%)', 'number', +(n.ck * 100).toFixed(2))}
    ${fld('cn_pct', 'Công nợ % in trên đơn – trống = mặc định', 'number', n.cn_pct == null ? '' : +(n.cn_pct * 100).toFixed(2))}
    ${fld('cn_chot', 'Công nợ chốt đầu kỳ (đ)', 'number')}
    ${fld('tinh_trang', 'Tình trạng')}
    <div class="field" style="grid-column:1/-1"><label for="n-ghi_chu">Ghi chú</label><input id="n-ghi_chu" value="${esc(n.ghi_chu)}"></div></div>`,
    `<button class="btn ghost" data-close2>Hủy</button><button class="btn" id="n-ok">Lưu</button>`);
  $('#n-ok').onclick = async () => {
    const v = k => $('#n-' + k).value;
    const data = { ...(id ? n : {}), id, ten: v('ten').trim(), khu_vuc: v('khu_vuc'), gsm: v('gsm'), asm: v('asm'), tinh: v('tinh'), sdt: v('sdt'), dia_chi: v('dia_chi'),
      ck: (Number(v('ck')) || 0) / 100, cn_pct: v('cn_pct') === '' ? '' : Number(v('cn_pct')) / 100, cn_chot: v('cn_chot'), tinh_trang: v('tinh_trang'), ghi_chu: v('ghi_chu') };
    if (!data.ten) return toast('Nhập tên NPP.', true);
    await run(() => api('/api/npp/save', { npp: data }), 'Đã lưu NPP.');
    closeModal(); refresh();
  };
}
function productDialog(id) {
  const p = M.sp[id];
  modal(`Sửa sản phẩm – ${p.ten}`, `<div class="fields">
    <div class="field" style="grid-column:1/-1"><label for="p-ten">Tên hàng</label><input id="p-ten" value="${esc(p.ten)}"></div>
    <div class="field"><label for="p-dvt">ĐVT</label><input id="p-dvt" value="${esc(p.dvt)}"></div>
    <div class="field"><label for="p-qc">Quy cách (chai/thùng)</label><input type="number" id="p-qc" value="${p.quy_cach}"></div>
    <div class="field"><label><input type="checkbox" id="p-ban" ${p.dang_ban ? 'checked' : ''}> Đang bán (hiện khi lên đơn)</label></div></div>`,
    `<button class="btn ghost" data-close2>Hủy</button><button class="btn" id="p-ok2">Lưu</button>`);
  $('#p-ok2').onclick = async () => {
    await run(() => api('/api/product/save', { product: { id, ten: $('#p-ten').value, dvt: $('#p-dvt').value, quy_cach: $('#p-qc').value, dang_ban: $('#p-ban').checked } }), 'Đã lưu sản phẩm.');
    closeModal(); refresh();
  };
}

// ---------- khởi động ----------
(async () => {
  try { if (await load()) route(); }
  catch (e) { $('#page').innerHTML = `<div class="alert bad">Không kết nối được công cụ: ${esc(e.message)}. Hãy chạy lại "Mo-cong-cu.bat".</div>`; }
})();
