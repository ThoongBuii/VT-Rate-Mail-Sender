let state = null;
let pollTimer = null;
let currentView = "compose";
let previewIndex = 0;
/** HTML template gốc (từ Outlook sync) — không qua TinyMCE để giữ chữ ký. */
let templateHtml = "";

async function api(path, options = {}) {
  const res = await fetch(path, options);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || res.statusText);
  return data;
}

function setOutlookStatus(text) {
  document.getElementById("outlookStatus").textContent = text;
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function frameDoc(html) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      html,body{margin:0;padding:12px;background:#fff;color:#222}
      body{font-family:Calibri,"Segoe UI",Arial,sans-serif;font-size:14px;line-height:1.4}
      table{border-collapse:collapse}
      img{max-width:100%;height:auto}
    </style></head><body>${html || ""}</body></html>`;
}

function setComposeHtml(html) {
  templateHtml = html || "";
  document.getElementById("composeFrame").srcdoc = frameDoc(templateHtml);
}

function getComposeHtml() {
  return templateHtml || "";
}

function setView(view) {
  currentView = view;
  document.getElementById("viewCompose").classList.toggle("hidden", view !== "compose");
  document.getElementById("viewPreview").classList.toggle("hidden", view !== "preview");
  document.getElementById("tabCompose").classList.toggle("active", view === "compose");
  document.getElementById("tabPreview").classList.toggle("active", view === "preview");
  if (view === "preview") {
    loadPreview(previewIndex);
  }
}

function renderList() {
  const box = document.getElementById("mailList");
  box.innerHTML = "";
  if (!state || !state.mails.length) {
    box.innerHTML = `<div class="mail-item"><div class="meta"><div class="company">Chưa có dữ liệu</div></div></div>`;
    return;
  }
  state.mails.forEach((m, i) => {
    const active =
      currentView === "preview" ? i === previewIndex : i === state.selected_index;
    const el = document.createElement("div");
    el.className = "mail-item" + (active ? " active" : "");
    el.innerHTML = `
      <div class="idx">${i + 1}</div>
      <div class="meta">
        <div class="company">${escapeHtml(m.agency_company || "(no company)")}</div>
        <div class="email">${escapeHtml(m.account_mail || "")}</div>
      </div>
      <div class="st st-${m.status}">${m.status}</div>`;
    el.onclick = () => selectMail(i);
    box.appendChild(el);
  });
}

function applyState(s, forceHtml = false) {
  state = s;
  document.getElementById("subject").value = s.subject || "";
  document.getElementById("attachmentName").value = s.attachment
    ? s.attachment.split(/[/\\]/).pop()
    : "";
  document.getElementById("delayMin").value = s.delay_min;
  document.getElementById("delayMax").value = s.delay_max;
  document.getElementById("stats").textContent =
    `${s.stats.total} agency · ready ${s.stats.ready} · sent ${s.stats.sent} · fail ${s.stats.failed}`;
  if (s.outlook_ready) {
    setOutlookStatus(`Outlook sẵn sàng · ${s.outlook_account || ""}`);
  } else {
    setOutlookStatus("Outlook: chưa kết nối");
  }
  if (forceHtml || !templateHtml) {
    setComposeHtml(s.template_html || "");
  }
  previewIndex = Math.min(previewIndex, Math.max(0, (s.mails?.length || 1) - 1));
  renderList();
  applyProgress(s.progress);
  if (currentView === "preview") loadPreview(previewIndex);
}

function applyProgress(p) {
  if (!p) return;
  document.getElementById("progressFill").style.width = `${p.percent || 0}%`;
  document.getElementById("progressText").textContent = p.message || "";
  const running = !!p.is_running;
  document.getElementById("btnStart").disabled = running;
  document.getElementById("btnPause").disabled = !running;
  document.getElementById("btnStop").disabled = !running;
  document.getElementById("btnPause").textContent = p.is_paused ? "▶ Resume" : "⏸ Pause";
  const log = document.getElementById("logBox");
  log.textContent = (p.logs || [])
    .map(
      (e) =>
        `[${e.time}] ${String(e.status).toUpperCase().padEnd(7)} | ${e.agency} | ${e.mail} | ${e.message}`
    )
    .join("\n");
  if (log.textContent) log.scrollTop = log.scrollHeight;
}

async function refreshState(forceHtml = false) {
  const s = await api("/api/state");
  applyState(s, forceHtml);
}

async function saveCompose(extra = {}) {
  const payload = {
    subject: document.getElementById("subject").value,
    template_html: getComposeHtml(),
    delay_min: Number(document.getElementById("delayMin").value || 10),
    delay_max: Number(document.getElementById("delayMax").value || 20),
    selected_index: state ? state.selected_index : 0,
    ...extra,
  };
  await api("/api/compose", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

async function selectMail(index) {
  if (!state) return;
  state.selected_index = index;
  previewIndex = index;
  await saveCompose({ selected_index: index });
  renderList();
  if (currentView === "preview") await loadPreview(index);
}

async function loadPreview(index) {
  if (!state?.mails?.length) {
    document.getElementById("previewMeta").innerHTML = "<p>Chưa có danh sách agency.</p>";
    document.getElementById("previewFrame").srcdoc = "";
    document.getElementById("previewCounter").textContent = "0 / 0";
    return;
  }
  index = Math.max(0, Math.min(index, state.mails.length - 1));
  previewIndex = index;
  await saveCompose({ selected_index: index });
  const p = await api(`/api/preview?index=${index}`);
  document.getElementById("previewCounter").textContent = `${p.index + 1} / ${p.total}`;
  const att = p.attachment_ok
    ? escapeHtml(p.attachment)
    : `<span class="err">${escapeHtml(p.attachment_error || p.attachment)}</span>`;
  document.getElementById("previewMeta").innerHTML = `
    <div><b>From</b><span>${escapeHtml(p.from || "—")}</span></div>
    <div><b>To</b><span>${escapeHtml(p.to || "—")}</span></div>
    <div><b>CC</b><span>${escapeHtml(p.cc || "(không)")}</span></div>
    <div><b>Subject</b><span>${escapeHtml(p.subject || "—")}</span></div>
    <div><b>File</b><span>${att}</span></div>
    <div><b>Agency</b><span>${escapeHtml(p.agency_company || "")} · ${escapeHtml(p.account_name || "")}</span></div>
    <div><b>Status</b><span class="st st-${p.status}">${escapeHtml(p.status || "")}</span></div>
  `;
  document.getElementById("previewFrame").srcdoc = frameDoc(p.body_html || "");
  renderList();
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(async () => {
    try {
      const p = await api("/api/progress");
      applyProgress(p);
      if (p && p.is_running) {
        const s = await api("/api/state");
        state.mails = s.mails;
        state.stats = s.stats;
        document.getElementById("stats").textContent =
          `${s.stats.total} agency · ready ${s.stats.ready} · sent ${s.stats.sent} · fail ${s.stats.failed}`;
        renderList();
      }
    } catch (_) {}
  }, 1000);
}

document.getElementById("tabCompose").onclick = () => setView("compose");
document.getElementById("btnToPreview").onclick = async () => {
  await saveCompose();
  setView("preview");
};
document.getElementById("tabPreview").onclick = async () => {
  await saveCompose();
  setView("preview");
};
document.getElementById("btnBackCompose").onclick = () => setView("compose");

document.getElementById("btnPrevMail").onclick = () => {
  if (!state?.mails?.length) return;
  loadPreview(Math.max(0, previewIndex - 1));
};
document.getElementById("btnNextMail").onclick = () => {
  if (!state?.mails?.length) return;
  loadPreview(Math.min(state.mails.length - 1, previewIndex + 1));
};

document.getElementById("btnOutlook").onclick = async () => {
  try {
    const res = await api("/api/outlook/open", { method: "POST" });
    setOutlookStatus(`Outlook sẵn sàng · ${res.account || ""}`);
    alert(res.message || "Outlook sẵn sàng");
  } catch (e) {
    alert(e.message);
  }
};

document.getElementById("btnComposeOutlook").onclick = async () => {
  try {
    await saveCompose();
    const res = await api("/api/outlook/compose-template", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        template_html: getComposeHtml(),
        subject: document.getElementById("subject").value,
      }),
    });
    setOutlookStatus(`Outlook sẵn sàng · ${res.account || ""}`);
    alert(res.message || "Đã mở Outlook");
  } catch (e) {
    alert(e.message);
  }
};

document.getElementById("btnSyncOutlook").onclick = async () => {
  try {
    const res = await api("/api/outlook/sync-template", { method: "POST" });
    setComposeHtml(res.template_html || "");
    if (res.state) applyState(res.state, true);
    await saveCompose();
    alert(res.message || "Đã đồng bộ từ Outlook");
  } catch (e) {
    alert(e.message);
  }
};

document.getElementById("fileImport").onchange = async (ev) => {
  const file = ev.target.files?.[0];
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  try {
    await saveCompose();
    const res = await api("/api/import", { method: "POST", body: fd });
    previewIndex = 0;
    applyState(res.state, true);
    alert(`Đã nạp ${res.count} agency`);
  } catch (e) {
    alert(e.message);
  } finally {
    ev.target.value = "";
  }
};

document.getElementById("fileAttach").onchange = async (ev) => {
  const file = ev.target.files?.[0];
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  try {
    const res = await api("/api/attachment", { method: "POST", body: fd });
    document.getElementById("attachmentName").value = res.name;
  } catch (e) {
    alert(e.message);
  } finally {
    ev.target.value = "";
  }
};

document.getElementById("btnValidate").onclick = async () => {
  try {
    await saveCompose();
    const res = await api("/api/validate", { method: "POST" });
    applyState(res.state);
    alert(`Sẵn sàng: ${res.ready}\nCần sửa: ${res.bad}`);
  } catch (e) {
    alert(e.message);
  }
};

document.getElementById("btnSuggest").onclick = async () => {
  if (!state) return;
  document.getElementById("subject").value = state.suggested_subject;
  setComposeHtml(state.suggested_html || "");
  await saveCompose();
  await refreshState(true);
};

document.getElementById("btnStart").onclick = async () => {
  try {
    await saveCompose();
    const payload = {
      subject: document.getElementById("subject").value,
      template_html: getComposeHtml(),
      delay_min: Number(document.getElementById("delayMin").value || 10),
      delay_max: Number(document.getElementById("delayMax").value || 20),
    };
    if (!getComposeHtml().trim()) {
      alert("Chưa có nội dung mail. Hãy Soạn trong Outlook → Đồng bộ từ Outlook.");
      return;
    }
    if (!confirm(`Đã kiểm tra Preview?\nGửi Semi-Auto · Delay ${payload.delay_min}–${payload.delay_max}s`))
      return;
    const res = await api("/api/send/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    alert(`Đang gửi ${res.count} mail…`);
  } catch (e) {
    alert(e.message);
  }
};

document.getElementById("btnPause").onclick = async () => {
  try {
    if (state?.progress?.is_paused) await api("/api/send/resume", { method: "POST" });
    else await api("/api/send/pause", { method: "POST" });
  } catch (e) {
    alert(e.message);
  }
};

document.getElementById("btnStop").onclick = async () => {
  try {
    await api("/api/send/stop", { method: "POST" });
  } catch (e) {
    alert(e.message);
  }
};

(async function boot() {
  const s = await api("/api/state");
  applyState(s, true);
  startPolling();
})();
