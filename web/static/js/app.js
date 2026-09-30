const TITLES = {
  new: ["신규 상품", "어제(직전 날짜 CSV)에는 없던 SKU입니다."],
  poship: ["발주/쉽먼트", "공급사 허브에서 발주확정 후 쉽먼트 엑셀을 올립니다."],
  order: ["발주서 만들기", "시트 로켓발주로 INBOX부터 발주서 채우기까지 실행합니다."],
  logs: ["실행 로그", "지금 돌아가는 작업의 출력입니다."],
  settings: ["설정", "내 발주 시트와 드라이브 폴더 주소를 저장합니다."],
};

const state = { tab: "new", es: null };

const $ = (id) => document.getElementById(id);
const fmt = (n) => (n == null ? "-" : Number(n).toLocaleString("ko-KR"));

function esc(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
function badge(status) {
  const s = status || "미기재";
  return `<span class="badge ${esc(s)}">${esc(s)}</span>`;
}
function imgCell(url) {
  if (!url) return `<div class="im"></div>`;
  return `<img class="im" src="${esc(url)}" alt="" loading="lazy" onerror="this.style.opacity=.2">`;
}

async function api(path, opt) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...opt,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok && data.ok === false) throw new Error(data.message || data.error || "요청 실패");
  return data;
}

function switchTab(tab) {
  state.tab = tab;
  document.querySelectorAll("nav button[data-tab]").forEach((b) => {
    b.classList.toggle("active", b.dataset.tab === tab);
  });
  document.querySelectorAll(".panel").forEach((p) => p.classList.add("hidden"));
  $(`tab-${tab}`).classList.remove("hidden");
  const [t, s] = TITLES[tab];
  $("pageTitle").textContent = t;
  $("pageSub").textContent = s;
  if (tab === "new") loadNew();
  if (tab === "poship") loadPoShip();
  if (tab === "settings") loadSettings();
  if (tab === "logs") loadLogs();
}

function applyJobUI(st) {
  const running = !!st.running;
  $("jobBanner").classList.toggle("hidden", !running);
  $("jobBannerText").textContent = running ? `${st.label || "작업"} 실행 중` : "";
  $("loginBanner").classList.toggle("hidden", !st.need_login);
  $("deniedBanner").classList.toggle("hidden", !st.denied);
  document.querySelectorAll(".run").forEach((b) => (b.disabled = running));
  $("jobLogMeta").textContent = running
    ? `${st.label} 실행 중`
    : st.ended_at
      ? `종료 ${st.ended_at} / 코드 ${st.returncode}`
      : "대기";
}

async function loadStatus() {
  const d = await api("/api/status");
  const pill = $("chromePill");
  pill.textContent = d.chrome ? "허브 Chrome 연결됨" : "허브 Chrome 꺼짐";
  pill.className = "chrome-pill " + (d.chrome ? "on" : "off");
  const s = d.settings || {};
  $("setupBanner").classList.toggle("hidden", !!(s.order_sheet_url && s.inbox_folder));
  applyJobUI(d.job || {});
}

// ── 신규 상품 ──
async function loadNew() {
  const d = await api("/api/new");
  $("csvMeta").textContent = d.csv_mtime
    ? `최신 CSV ${d.csv_mtime} · 전체 ${fmt(d.all_total)}개`
    : "아직 수집한 CSV가 없습니다.";
  $("newMeta").textContent = d.previous
    ? `${d.previous} 대비 ${fmt(d.total)}개`
    : "비교할 직전 파일이 없습니다.";
  $("newBody").innerHTML = d.rows.length
    ? d.rows
        .map(
          (r) => `<tr>
        <td>${imgCell(r.image)}</td>
        <td class="sku-cell" title="클릭하면 복사">${esc(r.sku)}</td>
        <td class="nm-cell" title="${esc(r.name)}">${esc(r.name)}</td>
        <td>${esc(r.barcode)}</td>
        <td>${badge(r.status)}</td>
        <td>${esc(r.bm)}</td>
        <td>${esc(r.scm)}</td>
        <td>${esc(r.moq)}</td>
      </tr>`
        )
        .join("")
    : `<tr><td colspan="8" class="empty">신규 상품이 없습니다.</td></tr>`;
}

// ── 발주/쉽먼트 ──
function renderShipFiles(files) {
  const kind = files.kind === "drive" ? "구글 드라이브" : `PC 폴더 ${files.dir || ""}`;
  const extra = files.fetched_at ? ` · ${files.fetched_at}` : "";
  $("shipMeta").textContent = `${kind} · ${fmt((files.files || []).length)}개${extra}`;
  $("shipBody").innerHTML = (files.files || []).length
    ? files.files
        .map(
          (f) => `<tr>
            <td>${esc(f.name)}</td>
            <td>${esc(f.mtime || "")}</td>
            <td>${f.size_kb != null && f.size_kb !== "" ? esc(f.size_kb) + " KB" : ""}</td>
          </tr>`
        )
        .join("")
    : `<tr><td colspan="3" class="empty">${files.kind === "drive" ? "목록 불러오기를 누르세요." : "이 폴더에 쉽먼트 xlsx가 없습니다."}</td></tr>`;
}

async function loadPoShip() {
  const [po, files, settings] = await Promise.all([
    api("/api/po"),
    api("/api/shipment/files"),
    api("/api/settings"),
  ]);
  if (settings.inbound_date) $("inboundDate").value = settings.inbound_date;
  const rows = po.pending_rows && po.pending_rows.length ? po.pending_rows : po.rows || [];
  $("poMeta").textContent = po.fetched_at
    ? `${po.fetched_at} · 전체 ${fmt(po.total)} / 미확정 ${fmt(po.pending)}`
    : "서플라이어 허브 로그인 후 1번을 누르세요.";
  $("poBody").innerHTML = rows.length
    ? rows
        .map(
          (r) => `<tr>
            <td class="sku-cell">${esc(r.orderId)}</td>
            <td class="sku-cell">${esc(r.sku)}</td>
            <td class="nm-cell" title="${esc(r.name)}">${esc(r.name)}</td>
            <td>${esc(r.center)}</td>
            <td>${esc(r.qty)}</td>
            <td>${esc(r.edd)}</td>
            <td>${esc(r.status)}</td>
          </tr>`
        )
        .join("")
    : `<tr><td colspan="7" class="empty">조회된 발주가 없습니다.</td></tr>`;
  renderShipFiles(files);
}

async function saveInboundDate() {
  const day = $("inboundDate").value.trim();
  await api("/api/settings", { method: "POST", body: JSON.stringify({ inbound_date: day }) });
}

// ── 설정 ──
async function loadSettings() {
  const s = await api("/api/settings");
  $("setOrderSheet").value = s.order_sheet_url || "";
  $("setShipFolder").value = s.ship_folder || "";
  $("setInboxFolder").value = s.inbox_folder || "";
  $("setBarcodeUrl").value = s.barcode_app_url || "";
}

async function saveSettings() {
  const d = await api("/api/settings", {
    method: "POST",
    body: JSON.stringify({
      order_sheet_url: $("setOrderSheet").value.trim(),
      ship_folder: $("setShipFolder").value.trim(),
      inbox_folder: $("setInboxFolder").value.trim(),
      barcode_app_url: $("setBarcodeUrl").value.trim(),
    }),
  });
  $("setOrderSheet").value = d.order_sheet_url || "";
  $("setShipFolder").value = d.ship_folder || "";
  $("setInboxFolder").value = d.inbox_folder || "";
  $("setBarcodeUrl").value = d.barcode_app_url || "";
  await loadStatus();
  alert("설정을 저장했습니다.");
}

// ── 작업 실행 / 로그 ──
async function loadLogs() {
  const d = await api("/api/job/log");
  $("jobLog").textContent = (d.lines || []).join("\n") || "아직 이번 세션에서 실행한 작업이 없습니다.";
  applyJobUI(d.status || {});
  $("jobLog").scrollTop = $("jobLog").scrollHeight;
}

function appendLive(line) {
  const box = $("liveLog");
  const jobLog = $("jobLog");
  if (box.textContent.includes("실행하면 진행 로그")) box.textContent = "";
  const parts = ((box.textContent ? box.textContent + "\n" : "") + line).split("\n");
  box.textContent = parts.slice(-80).join("\n");
  box.scrollTop = box.scrollHeight;
  jobLog.textContent = jobLog.textContent.startsWith("아직") || !jobLog.textContent
    ? line
    : jobLog.textContent + "\n" + line;
  jobLog.scrollTop = jobLog.scrollHeight;
}

function watchJob() {
  if (state.es) state.es.close();
  const es = new EventSource("/api/job/stream");
  state.es = es;
  es.addEventListener("log", (e) => {
    const d = JSON.parse(e.data);
    if (d.line) appendLive(d.line);
  });
  es.addEventListener("status", (e) => applyJobUI(JSON.parse(e.data)));
  es.addEventListener("done", async () => {
    es.close();
    state.es = null;
    await loadStatus();
    if (state.tab === "poship") loadPoShip();
    if (state.tab === "new") loadNew();
  });
  es.onerror = () => {};
}

const NEEDS_DATE = ["po_list", "po_confirm", "po_ship"];

async function startJob(mode) {
  try {
    if (NEEDS_DATE.includes(mode)) {
      if (!$("inboundDate").value.trim()) {
        alert("입고예정일을 선택하세요.");
        switchTab("poship");
        $("inboundDate").focus();
        return;
      }
      await saveInboundDate();
    }
    const d = await api("/api/job", { method: "POST", body: JSON.stringify({ mode }) });
    if (!d.ok) {
      alert(d.message || "이미 실행 중입니다.");
      return;
    }
    $("liveLog").textContent = "";
    $("jobLog").textContent = "";
    applyJobUI(d);
    if (mode !== "collect") switchTab("logs");
    watchJob();
  } catch (e) {
    alert(e.message || "실행에 실패했습니다.");
  }
}

async function openTarget(target) {
  try {
    await api("/api/open", { method: "POST", body: JSON.stringify({ target }) });
    if (target === "hub_chrome") setTimeout(loadStatus, 1500);
  } catch (e) {
    alert(e.message);
  }
}

function copySku(e) {
  const cell = e.target.closest(".sku-cell");
  if (!cell) return;
  navigator.clipboard.writeText(cell.textContent.trim()).catch(() => {});
  cell.title = "복사됨";
  setTimeout(() => (cell.title = "클릭하면 복사"), 800);
}

document.querySelectorAll("nav button[data-tab]").forEach((b) => {
  b.addEventListener("click", () => switchTab(b.dataset.tab));
});
document.querySelectorAll("[data-goto]").forEach((b) => {
  b.addEventListener("click", () => switchTab(b.dataset.goto));
});
document.querySelectorAll(".run").forEach((b) => {
  b.addEventListener("click", () => startJob(b.dataset.mode));
});
document.querySelectorAll("[data-open]").forEach((b) => {
  b.addEventListener("click", () => openTarget(b.dataset.open));
});
$("btnStop").addEventListener("click", () => api("/api/job/stop", { method: "POST" }));
$("btnFolder").addEventListener("click", () => openTarget("folder"));
$("btnSheet").addEventListener("click", () => openTarget("sheet"));
$("btnOrderSheet").addEventListener("click", () => openTarget("sheet"));
$("btnInboxFolder").addEventListener("click", () => openTarget("inbox_folder"));
$("btnPoPage").addEventListener("click", () => openTarget("po"));
$("btnShipFolder").addEventListener("click", () => openTarget("ship_folder"));
$("btnDriveList").addEventListener("click", () => startJob("drive_list"));
$("btnHubChrome").addEventListener("click", () => openTarget("hub_chrome"));
async function openBarcode() {
  try {
    await api("/api/barcode/open", { method: "POST", body: "{}" });
  } catch (e) {
    alert(e.message);
    switchTab("settings");
  }
}
$("btnBarcode").addEventListener("click", openBarcode);
$("btnBarcodeTop").addEventListener("click", openBarcode);
$("btnSaveSettings").addEventListener("click", () => saveSettings().catch((e) => alert(e.message)));
$("inboundDate").addEventListener("change", () => saveInboundDate().catch(() => {}));
document.body.addEventListener("click", copySku);

switchTab("new");
loadStatus().then(() => api("/api/job")).then((st) => {
  if (st && st.running) watchJob();
});
setInterval(loadStatus, 15000);
